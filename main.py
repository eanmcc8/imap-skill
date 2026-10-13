"""imap-skill — Universal IMAP/SMTP client, CLI and provider schemas.

Debloated: direct IMAP connection per client (no pool/keepalive thread),
fixed STARTTLS + SMTP defaults + HTML escaping, deduplicated reply/forward/draft paths.
Public API unchanged: ImapClient, dispatch_tool, canonical_tools_for_provider,
compile_tools, CLI tools, DEFAULT_MAX_BODY_CHARS.
"""

import getpass
import html
import imaplib as _imaplib_mod  # noqa: F401  (kept for Time2Internaldate check path)
import json
import logging
import mimetypes
import os
import re
import ssl
import sys
import time
import uuid
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, getaddresses
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import imaplib
    import email
    HAVE_IMAP = True
except Exception:  # pragma: no cover - interpreter without imaplib/email
    HAVE_IMAP = False
    imaplib = None  # type: ignore
    email = None  # type: ignore


# ---------- bootstrap: .env + logging ----------

_STRIP_COMMENT_KEYS = {
    "IMAP_HOST", "IMAP_PORT", "IMAP_USE_SSL", "IMAP_TIMEOUT",
    "IMAP_REQUIRE_ID", "IMAP_BACKUP_DIR",
    "SMTP_HOST", "SMTP_PORT", "SMTP_USE_SSL",
    "MAX_BODY_CHARS", "LOG_LEVEL",
}


def _env_token(name: str, default: str = "") -> str:
    """First whitespace-separated token of an env var, ignoring '#…' comments."""
    raw = os.getenv(name)
    if raw is None:
        return default
    text = str(raw).split("#", 1)[0].strip()
    if not text:
        return default
    return text.split()[0]


def _load_dotenv() -> None:
    """Minimal .env loader (no dependency): KEY=VALUE lines, no override of real env."""
    try:
        base = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base, ".env")
        if not os.path.isfile(path):
            return
        with open(path, "r", encoding="utf-8") as fh:
            for raw in fh:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip()
                if key in _STRIP_COMMENT_KEYS:
                    val = val.split("#", 1)[0].strip().strip('"').strip("'")
                elif key == "IMAP_USER":
                    val = re.split(r"\s+#", val, maxsplit=1)[0].strip().strip('"').strip("'")
                elif key in ("IMAP_PASSWORD", "SMTP_PASSWORD", "SMTP_USER"):
                    # Secrets may legitimately contain '#', so never split on it.
                    # But support a quoted secret with trailing comment:
                    #   IMAP_PASSWORD="s#cret"  # my note  ->  s#cret
                    if val[:1] in ("'", '"'):
                        q = val[0]
                        end = val.find(q, 1)
                        if end > 0:
                            val = val[1:end]
                        else:
                            val = val.strip('"').strip("'")
                else:
                    val = val.strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception:
        pass


_load_dotenv()


def _env_int(name: str, default: int) -> int:
    token = _env_token(name, "")
    if not token:
        return default
    try:
        return int(token)
    except ValueError:
        logging.getLogger("email_skill").warning(
            "Invalid %s=%r; using default %d", name, os.getenv(name), default)
        return default


LOG_LEVEL = _env_token("LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
LOG = logging.getLogger("email_skill")

DEFAULT_MAX_BODY_CHARS = _env_int("MAX_BODY_CHARS", 8000)
DEFAULT_TIMEOUT = _env_int("IMAP_TIMEOUT", 30)
MAX_FETCH_LIMIT = 100
TRUNCATION_SUFFIX = "… (truncated)"

IMAP_REQUIRE_ID = _env_token("IMAP_REQUIRE_ID", "false").lower() in ("1", "true", "yes", "y")

SMTP_HOST = _env_token("SMTP_HOST", "")
SMTP_PORT = _env_int("SMTP_PORT", 465)
SMTP_USE_SSL = _env_token("SMTP_USE_SSL", "true").lower() in ("1", "true", "yes", "y")
SMTP_USER = _env_token("SMTP_USER", "")
SMTP_PASSWORD = _env_token("SMTP_PASSWORD", "")

BACKUP_DIR = _env_token("IMAP_BACKUP_DIR", str(Path.home() / "imap_backup"))


# ---------- helpers ----------

def _uid_to_str(uid: Any) -> str:
    if isinstance(uid, bytes):
        return uid.decode("ascii", "replace")
    return str(uid)


def _decode_mime(s: Any) -> str:
    """Safe decoding of MIME encoded headers."""
    if not s:
        return ""
    try:
        parts = decode_header(str(s))
    except Exception:
        return str(s)
    out: List[str] = []
    for text, enc in parts:
        try:
            if isinstance(text, str):
                out.append(text.strip())
            else:
                out.append((text or b"").decode(enc or "utf-8", "replace").strip())
        except Exception:
            try:
                out.append((text or b"").decode("utf-8", "replace").strip())  # type: ignore
            except Exception:
                continue
    return " ".join(p for p in out if p)


def _extract_body_part(msg: "email.message.Message", content_type: str) -> str:
    """First non-attachment part of content_type; '' when none."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.is_multipart():
                    continue
                if part.get_content_type() != content_type:
                    continue
                if part.get_content_disposition() == "attachment":
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                try:
                    return payload.decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    ).strip()
                except Exception:
                    continue
            return ""
        if content_type != "text/plain" and msg.get_content_type() != content_type:
            return ""
        payload = msg.get_payload(decode=True)
        if payload is None:
            return str(msg.get_payload() or "").strip()
        try:
            return payload.decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            ).strip()
        except Exception:
            return ""
    except Exception:
        return ""


def _extract_plain_text(msg: "email.message.Message") -> str:
    return _extract_body_part(msg, "text/plain")


def _extract_html_body(msg: "email.message.Message") -> str:
    return _extract_body_part(msg, "text/html")


def _parse_mailbox_line(line: bytes) -> Optional[str]:
    """Robustly parse one LIST response line into a folder name."""
    try:
        text = line.decode("utf-8", "replace") if isinstance(line, bytes) else str(line)
    except Exception:
        return None
    text = text.strip()
    if not text:
        return None
    quoted = re.findall(r'"([^"]*)"', text)
    if quoted:
        return quoted[-1]
    tokens = text.split()
    if not tokens:
        return None
    return tokens[-1].strip('"').strip()


def _parse_fetch_uid(meta: bytes) -> Optional[str]:
    try:
        m = re.search(rb"UID\s+(\d+)", meta or b"")
        if m:
            return m.group(1).decode("ascii")
    except Exception:
        pass
    return None


# ---------- IMAP ID (RFC 2971, for 126/163) ----------

def _send_imap_id(conn, require: bool = False) -> None:
    if not (require or IMAP_REQUIRE_ID):
        return
    try:
        cap = getattr(conn, "capability", None)
        if cap is not None:
            cap()
    except Exception:
        pass
    try:
        sender = getattr(conn, "_simple_command", None)
        if sender is None:
            return
        typ, resp = sender(b"ID", b'("name" "imap-skill" "version" "2.0")')
        if typ != "OK":
            LOG.warning("IMAP ID command failed: %r", resp)
    except Exception as e:
        LOG.warning("IMAP ID command error: %s", e)


# ---------- SMTP ----------

def _get_smtp_config() -> Dict[str, Any]:
    return {
        "host": SMTP_HOST or "smtp.gmail.com",
        "port": SMTP_PORT,
        "use_ssl": SMTP_USE_SSL,
        "user": SMTP_USER or os.getenv("IMAP_USER", ""),
        "password": SMTP_PASSWORD or os.getenv("IMAP_PASSWORD", ""),
    }


def _create_smtp_client(config: Dict[str, Any]):
    import smtplib

    host, port = config["host"], config["port"]
    if config["use_ssl"]:
        client = smtplib.SMTP_SSL(host, port, timeout=DEFAULT_TIMEOUT)
        client.ehlo()
    else:
        client = smtplib.SMTP(host, port, timeout=DEFAULT_TIMEOUT)
        client.ehlo()
        client.starttls(context=ssl.create_default_context())
        client.ehlo()
    client.login(config["user"], config["password"])
    return client


def _send_smtp_message(msg: MIMEMultipart, config: Dict[str, Any]) -> None:
    client = _create_smtp_client(config)
    try:
        client.send_message(msg)
    finally:
        try:
            client.quit()
        except Exception:
            pass


def _build_message(
    from_addr: str,
    to_addrs: List[str],
    subject: str,
    body: str,
    cc_addrs: Optional[List[str]] = None,
    in_reply_to: Optional[str] = None,
    references: Optional[str] = None,
) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = f"<{uuid.uuid4()}@imap-skill>"
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = f"{references or ''} {in_reply_to}".strip()
    msg.attach(MIMEText(body, "plain", "utf-8"))
    return msg


def _extract_email_address(addr_str: str) -> str:
    if not addr_str:
        return ""
    m = re.search(r"<([^>]+)>", addr_str)
    if m:
        return m.group(1).strip()
    return addr_str.strip() if "@" in addr_str else addr_str


def _extract_email_addresses(header_value: str) -> List[str]:
    if not header_value:
        return []
    return [addr for _, addr in getaddresses([header_value]) if addr]


def _quote_original(body: str, from_label: str) -> str:
    quoted = "\n".join(f"> {line}" for line in body.split("\n"))
    return f"\n---- {from_label} wrote ----\n{quoted}\n--------------------------------"


# ---------- Presets ----------

PRESETS = {
    "gmail": {"name": "Gmail", "imap_host": "imap.gmail.com", "imap_port": 993,
              "imap_use_ssl": True, "imap_require_id": False,
              "smtp_host": "smtp.gmail.com", "smtp_port": 587, "smtp_use_ssl": False},
    "outlook": {"name": "Outlook / Hotmail / Office 365", "imap_host": "outlook.office365.com",
                "imap_port": 993, "imap_use_ssl": True, "imap_require_id": False,
                "smtp_host": "smtp.office365.com", "smtp_port": 587, "smtp_use_ssl": False},
    "qq": {"name": "QQ Mail", "imap_host": "imap.qq.com", "imap_port": 993,
           "imap_use_ssl": True, "imap_require_id": False,
           "smtp_host": "smtp.qq.com", "smtp_port": 465, "smtp_use_ssl": True},
    "163": {"name": "163 NetEase", "imap_host": "imap.163.com", "imap_port": 993,
            "imap_use_ssl": True, "imap_require_id": True,
            "smtp_host": "smtp.163.com", "smtp_port": 465, "smtp_use_ssl": True},
    "126": {"name": "126 NetEase", "imap_host": "imap.126.com", "imap_port": 993,
            "imap_use_ssl": True, "imap_require_id": True,
            "smtp_host": "smtp.126.com", "smtp_port": 465, "smtp_use_ssl": True},
    "tencent": {"name": "Tencent Enterprise Mail", "imap_host": "imap.exmail.qq.com",
                "imap_port": 993, "imap_use_ssl": True, "imap_require_id": False,
                "smtp_host": "smtp.exmail.qq.com", "smtp_port": 465, "smtp_use_ssl": True},
    "yahoo": {"name": "Yahoo Mail", "imap_host": "imap.mail.yahoo.com", "imap_port": 993,
              "imap_use_ssl": True, "imap_require_id": False,
              "smtp_host": "smtp.mail.yahoo.com", "smtp_port": 465, "smtp_use_ssl": True},
    "icloud": {"name": "iCloud Mail", "imap_host": "imap.mail.me.com", "imap_port": 993,
               "imap_use_ssl": True, "imap_require_id": False,
               "smtp_host": "smtp.mail.me.com", "smtp_port": 587, "smtp_use_ssl": False},
}


def _run_init_wizard() -> int:
    """Interactive configuration wizard."""
    print("=" * 60)
    print("  imap-skill Configuration Wizard")
    print("=" * 60 + "\n")

    base = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(base, ".env")
    if os.path.exists(env_path):
        if input(f".env already exists at {env_path}. Overwrite? [y/N]: ").strip().lower() not in ("y", "yes"):
            print("Cancelled.")
            return 1

    keys = list(PRESETS.keys())
    print("Select email provider preset:")
    for i, key in enumerate(keys):
        print(f"  {i + 1}. {PRESETS[key]['name']} ({key})")
    print(f"  {len(keys) + 1}. Custom / Other\n")
    preset = None
    try:
        idx = int(input(f"Enter choice [1-{len(keys) + 1}]: ").strip()) - 1
        if 0 <= idx < len(keys):
            preset = PRESETS[keys[idx]]
            print(f"\nSelected: {preset['name']}")
        else:
            print("\nCustom configuration selected.")
    except (ValueError, IndexError):
        print("\nInvalid choice, using custom configuration.")

    def ask(prompt: str, default: str = "") -> str:
        suffix = f" [{default}]" if default else ""
        val = input(f"{prompt}{suffix}: ").strip()
        return val or default

    print("\n--- Credentials ---")
    email_addr = ask("Email address")
    if not email_addr:
        print("Error: Email address is required.")
        return 1
    print("\nApp password / authorization code (not your login password):")
    print("  - Gmail: Google Account > Security > App Passwords")
    print("  - Outlook: Account > Security > App passwords")
    print("  - QQ/163/126: Settings > POP3/SMTP/IMAP > Generate auth code")
    password = getpass.getpass("Password: ").strip()
    if not password:
        print("Error: Password is required.")
        return 1

    if preset:
        imap_host, imap_port = preset["imap_host"], str(preset["imap_port"])
        imap_ssl, imap_id = str(preset["imap_use_ssl"]).lower(), str(preset["imap_require_id"]).lower()
        smtp_host, smtp_port = preset["smtp_host"], str(preset["smtp_port"])
        smtp_ssl = str(preset["smtp_use_ssl"]).lower()
    else:
        imap_host = ask("IMAP Host")
        imap_port = ask("IMAP Port", "993")
        imap_ssl, imap_id = "true", "false"
        smtp_host = ask("SMTP Host")
        smtp_port = ask("SMTP Port", "465")
        smtp_ssl = "true"

    lines = [
        "# imap-skill configuration",
        f"# Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "# IMAP Settings",
        f"IMAP_HOST={imap_host}",
        f"IMAP_PORT={imap_port}",
        f"IMAP_USE_SSL={imap_ssl}",
        f"IMAP_REQUIRE_ID={imap_id}",
        f"IMAP_USER={email_addr}",
        f"IMAP_PASSWORD={password}",
        "",
        "# SMTP Settings",
        f"SMTP_HOST={smtp_host}",
        f"SMTP_PORT={smtp_port}",
        f"SMTP_USE_SSL={smtp_ssl}",
        f"SMTP_USER={email_addr}",
        f"SMTP_PASSWORD={password}",
        "",
        "# Optional Settings",
        "MAX_BODY_CHARS=8000",
        "LOG_LEVEL=INFO",
        "IMAP_BACKUP_DIR=~/imap_backup",
    ]
    print("\n--- Preview ---")
    print("\n".join(lines if "PASSWORD" not in l else l.split("=")[0] + "=***" for l in lines))
    print("---")
    if input("\nSave to .env? [Y/n]: ").strip().lower() not in ("", "y", "yes"):
        print("\nCancelled.")
        return 1
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    try:
        os.chmod(env_path, 0o600)
    except Exception:
        pass
    print(f"\nConfiguration saved to {env_path}")
    return 0


# ---------- IMAP open ----------

def _open_imap(host: str, port: int, use_ssl: bool, timeout: int):
    """Create an IMAP connection (SSL or STARTTLS). Fixes missing _upgrade_to_tls."""
    if not HAVE_IMAP or imaplib is None:
        raise RuntimeError("imaplib/email unavailable in this interpreter")
    if use_ssl:
        return imaplib.IMAP4_SSL(host, port, timeout=timeout)
    conn = imaplib.IMAP4(host, port, timeout=timeout)
    try:
        conn.starttls(ssl.create_default_context())
    except Exception as e:
        LOG.warning("STARTTLS failed on %s:%s: %s", host, port, e)
    return conn


# ---------- Offline backup ----------

def _extract_attachments(msg: "email.message.Message", backup_dir: Path, uid: str) -> List[Dict[str, Any]]:
    attachments: List[Dict[str, Any]] = []
    backup_dir.mkdir(parents=True, exist_ok=True)
    for i, part in enumerate(msg.walk()):
        if part.is_multipart() or part.get_content_disposition() != "attachment":
            continue
        fname = _decode_mime(part.get_filename() or f"unnamed-attachment-{i}")
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        safe = re.sub(r'[<>:"/\\|?*]', "_", fname)
        ext = Path(safe).suffix or mimetypes.guess_extension(part.get_content_type()) or ""
        save_name = f"{uid}_{i}{ext}"
        save_path = backup_dir / save_name
        try:
            save_path.write_bytes(payload)
            attachments.append({"original_name": fname, "saved_as": save_name,
                                "size": len(payload), "content_type": part.get_content_type(),
                                "path": str(save_path)})
        except Exception as e:
            LOG.warning("Failed to save attachment %s: %s", fname, e)
    return attachments


def _save_email_as_html(msg, mailbox: str, uid: str, backup_dir: Path,
                        include_attachments: bool = True) -> Dict[str, Any]:
    backup_dir.mkdir(parents=True, exist_ok=True)
    subject = _decode_mime(msg.get("Subject", ""))
    from_addr = _decode_mime(msg.get("From", ""))
    to_addr = _decode_mime(msg.get("To", ""))
    date = msg.get("Date", "")
    message_id = msg.get("Message-ID", "")

    html_body = _extract_html_body(msg)
    plain_body = _extract_plain_text(msg)
    if not html_body and plain_body:
        html_body = f"<pre>{html.escape(plain_body)}</pre>"
    elif not html_body:
        html_body = "<p><em>No body content</em></p>"

    attachments = _extract_attachments(msg, backup_dir, uid) if include_attachments else []
    att_html = ""
    if attachments:
        links = "".join(
            f'<div class="attachment"><a href="{html.escape(a["saved_as"])}">'
            f'{html.escape(a["original_name"])}</a> ({a["size"]} bytes)</div>'
            for a in attachments)
        att_html = f'<div class="attachments"><h2>Attachments ({len(attachments)})</h2>{links}</div>'

    page = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>{html.escape(subject)}</title>
<style>body{{font-family:sans-serif;max-width:800px;margin:2rem auto;padding:0 1rem}}
.header{{border-bottom:1px solid #ddd;padding-bottom:1rem;margin-bottom:1rem}}
.meta{{color:#666;font-size:.9rem}}.body{{line-height:1.6}}
.attachments{{margin-top:2rem;padding-top:1rem;border-top:1px solid #ddd}}
.attachment{{display:inline-block;margin:.5rem;padding:.5rem;background:#f5f5f5;border-radius:4px}}</style>
</head>
<body><div class="header"><h1>{html.escape(subject)}</h1>
<div class="meta"><div><strong>From:</strong> {html.escape(from_addr)}</div>
<div><strong>To:</strong> {html.escape(to_addr)}</div>
<div><strong>Date:</strong> {html.escape(date)}</div>
<div><strong>Message-ID:</strong> {html.escape(message_id)}</div>
<div><strong>Mailbox:</strong> {html.escape(mailbox)}</div>
<div><strong>UID:</strong> {html.escape(uid)}</div></div></div>
<div class="body">{html_body}</div>{att_html}</body></html>"""

    safe_subject = re.sub(r'[<>:"/\\|?*]', "_", subject)[:100]
    html_path = backup_dir / f"{uid}_{safe_subject}.html"
    html_path.write_text(page, encoding="utf-8")
    return {"html_file": str(html_path), "subject": subject, "from": from_addr,
            "to": to_addr, "date": date, "message_id": message_id,
            "mailbox": mailbox, "uid": uid, "attachments": attachments,
            "has_html_body": bool(_extract_html_body(msg)), "has_plain_body": bool(plain_body)}


class ImapClient:
    """Direct IMAP transport (one connection per client, no pool)."""

    def __init__(self):
        self.host = _env_token("IMAP_HOST", "imap.gmail.com")
        self.port = _env_int("IMAP_PORT", 993)
        mode = _env_token("IMAP_USE_SSL", "true").lower()
        if mode == "auto":
            self.use_ssl = (self.port == 993)
        elif mode in ("1", "true", "yes", "y", "on"):
            self.use_ssl = True
        elif mode in ("0", "false", "no", "n", "off"):
            self.use_ssl = False
        else:
            raise RuntimeError("IMAP_USE_SSL must be 'true', 'false' or 'auto'")
        self.user = os.getenv("IMAP_USER")
        self.password = os.getenv("IMAP_PASSWORD")
        self.timeout = DEFAULT_TIMEOUT
        self._conn = None

    # -- lifecycle --
    def connect(self) -> None:
        if self._conn:
            return
        if not HAVE_IMAP:
            raise RuntimeError("imaplib/email unavailable in this interpreter")
        if not self.user or not self.password:
            raise RuntimeError("IMAP_USER and IMAP_PASSWORD must be set (env or .env)")
        if any(c.isspace() for c in self.user):
            raise RuntimeError("IMAP_USER must be the bare email address with no spaces")
        conn = _open_imap(self.host, self.port, self.use_ssl, self.timeout)
        try:
            conn.login(self.user, self.password)
        except Exception:
            try:
                conn.logout()
            except Exception:
                pass
            raise
        _send_imap_id(conn)
        try:
            if getattr(conn, "capability", None):
                conn.capability()
        except Exception:
            pass
        self._conn = conn

    def close(self) -> None:
        if self._conn is not None:
            try:
                try:
                    self._conn.close()
                except Exception:
                    pass
                self._conn.logout()
            except Exception:
                pass
            self._conn = None

    def __enter__(self) -> "ImapClient":
        self.connect()
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    @classmethod
    def connect_once(cls) -> "ImapClient":
        c = cls()
        c.connect()
        return c

    @classmethod
    def close_pool(cls) -> None:
        """No-op kept for backward compatibility (pool removed)."""

    # -- internals --
    def _require_conn(self) -> None:
        if not self._conn:
            raise RuntimeError("Not connected to IMAP (call connect() first)")

    def _caps(self) -> set:
        caps = set()
        for c in getattr(self._conn, "capabilities", None) or ():
            caps.add(c.decode("ascii", "ignore").upper() if isinstance(c, bytes) else str(c).upper())
        return caps

    def _has_cap(self, name: str) -> bool:
        try:
            if not self._caps():
                try:
                    self._conn.capability()
                except Exception:
                    pass
            return name.upper() in self._caps()
        except Exception:
            return False

    def _select(self, mailbox: str, readonly: bool) -> None:
        self._require_conn()
        typ, _ = self._conn.select(mailbox, readonly=readonly)
        if typ != "OK":
            raise RuntimeError(f"Cannot select mailbox {mailbox!r}")

    def _fetch_original(self, mailbox: str, uid: str):
        """Fetch full message by UID (BODY.PEEK: no Seen side effect)."""
        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw = first[1] if isinstance(first, tuple) else first
        if not raw:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw, str):
            raw = raw.encode("utf-8", "replace")
        return email.message_from_bytes(raw)

    def _reply_addrs(self, orig_msg, reply_all: bool):
        config = _get_smtp_config()
        from_addr = config["user"]
        orig_from = _decode_mime(orig_msg.get("From", ""))
        to_addrs = [_extract_email_address(orig_from)]
        cc_addrs: List[str] = []
        if reply_all:
            wanted = _extract_email_addresses(_decode_mime(orig_msg.get("To", ""))) \
                + _extract_email_addresses(_decode_mime(orig_msg.get("Cc", "")))
            me = (from_addr or "").lower()
            sender = _extract_email_address(orig_from).lower()
            seen = set()
            for addr in wanted:
                low = addr.lower()
                if low in seen or low == me or low == sender:
                    continue
                seen.add(low)
                cc_addrs.append(addr)
        return config, from_addr, to_addrs, cc_addrs

    def _append_mime(self, mailbox: str, msg: MIMEMultipart, flags: str = "(\\Draft)") -> None:
        if HAVE_IMAP and imaplib is not None:
            internal_date = imaplib.Time2Internaldate(time.time())
        else:
            internal_date = time.strftime("%d-%b-%Y %H:%M:%S +0000", time.gmtime())
        typ, _ = self._conn.append(mailbox, flags, internal_date, msg.as_bytes())
        if typ != "OK":
            raise RuntimeError("Draft append failed")

    @staticmethod
    def _with_prefix(subject: str, prefix: str) -> str:
        return subject if subject.upper().startswith(prefix.upper()) else f"{prefix} {subject}"

    # ---------- tools ----------
    def list_mailboxes(self) -> List[str]:
        self._require_conn()
        typ, data = self._conn.list()
        if typ != "OK":
            raise RuntimeError("Unable to list mailboxes")
        out = []
        for line in data or []:
            if not line:
                continue
            name = _parse_mailbox_line(line)
            if name:
                out.append(name)
        return out

    def search_messages(self, mailbox: str = "INBOX", criteria: str = "ALL",
                        limit: int = 20) -> List[Dict[str, Any]]:
        """UID SEARCH, most recent `limit`; single batched header FETCH."""
        self._require_conn()
        try:
            limit = int(limit)
        except (TypeError, ValueError):
            raise RuntimeError("limit must be an integer")
        if limit <= 0:
            return []
        limit = min(limit, MAX_FETCH_LIMIT)
        criteria = criteria or "ALL"

        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("SEARCH", None, criteria)
        if typ != "OK":
            raise RuntimeError(f"Search {criteria!r} failed.")
        raw_ids = (data[0] if data else None) or b""
        if isinstance(raw_ids, str):
            raw_ids = raw_ids.encode("ascii", "ignore")
        uid_list = raw_ids.split()
        if not uid_list:
            return []
        wanted = uid_list[-limit:]
        typ, raw = self._conn.uid(
            "FETCH", b",".join(wanted),
            "(UID BODY.PEEK[HEADER.FIELDS (SUBJECT FROM TO DATE)])",
        )
        if typ != "OK":
            raise RuntimeError("Header fetch failed.")
        headers_by_uid: Dict[str, bytes] = {}
        order: List[str] = []
        for item in raw or []:
            if not isinstance(item, tuple) or len(item) < 2:
                continue
            meta, header_bytes = item[0], item[1]
            uid = _parse_fetch_uid(meta if isinstance(meta, bytes) else b"")
            if not uid:
                continue
            if isinstance(header_bytes, str):
                header_bytes = header_bytes.encode("utf-8", "replace")
            headers_by_uid[uid] = header_bytes or b""
            order.append(uid)

        ordered = [_uid_to_str(u) for u in wanted if _uid_to_str(u) in headers_by_uid] or order
        results = []
        for uid in ordered:
            try:
                msg = email.message_from_bytes(headers_by_uid.get(uid, b""))
            except Exception:
                continue
            results.append({"uid": uid, "subject": _decode_mime(msg.get("Subject", "")),
                            "from": _decode_mime(msg.get("From", "")),
                            "to": _decode_mime(msg.get("To", "")),
                            "date": msg.get("Date", "") or ""})
        return results

    def fetch_message(self, mailbox: str, uid: str, include_body: bool = True,
                      max_body_chars: Optional[int] = None) -> Dict[str, Any]:
        self._require_conn()
        if max_body_chars is None:
            max_body_chars = DEFAULT_MAX_BODY_CHARS
        try:
            max_body_chars = int(max_body_chars)
        except (TypeError, ValueError):
            raise RuntimeError("max_body_chars must be an integer")
        if isinstance(include_body, str):
            include_body = include_body.strip().lower() in ("1", "true", "yes", "y")

        msg = self._fetch_original(mailbox, uid)
        body = ""
        if include_body:
            full = _extract_plain_text(msg)
            body = full[:max_body_chars] + TRUNCATION_SUFFIX if len(full) > max_body_chars else full
        return {"uid": _uid_to_str(uid), "subject": _decode_mime(msg.get("Subject", "")),
                "from": _decode_mime(msg.get("From", "")), "to": _decode_mime(msg.get("To", "")),
                "date": msg.get("Date", "") or "", "body": body}

    def mark_message(self, mailbox: str, uid: str, flag: str = "\\Seen",
                     add: bool = True) -> Dict[str, Any]:
        self._require_conn()
        if isinstance(add, str):
            add = add.strip().lower() in ("1", "true", "yes", "y")
        self._select(mailbox, readonly=False)
        typ, _ = self._conn.uid("STORE", str(uid), "+FLAGS" if add else "-FLAGS", flag)
        if typ != "OK":
            raise RuntimeError(f"Failed to {'add' if add else 'remove'} flag {flag}")
        return {"status": "ok", "uid": _uid_to_str(uid), "flag": flag, "added": bool(add)}

    def move_message(self, mailbox: str = "", uid: str = "", dest: str = "",
                     **aliases: Any) -> Dict[str, Any]:
        src = mailbox or aliases.get("mailbox_src") or aliases.get("source") or ""
        dst = dest or aliases.get("mailbox_dest") or aliases.get("destination") or ""
        if not src or not uid or not dst:
            raise RuntimeError("move_message requires mailbox, uid and dest")
        self._require_conn()
        self._select(src, readonly=False)
        if self._has_cap("MOVE"):
            typ, _ = self._conn.uid("MOVE", str(uid), dst)
            if typ != "OK":
                raise RuntimeError(f"MOVE to {dst!r} failed")
        else:
            typ, _ = self._conn.uid("COPY", str(uid), dst)
            if typ != "OK":
                raise RuntimeError("COPY failed; message not moved")
            typ, _ = self._conn.uid("STORE", str(uid), "+FLAGS", "\\Deleted")
            if typ != "OK":
                raise RuntimeError("COPY succeeded but flagging original \\Deleted failed")
            if self._has_cap("UIDPLUS"):
                typ, _ = self._conn.uid("EXPUNGE", str(uid))
            else:
                typ, _ = self._conn.expunge()
            if typ != "OK":
                raise RuntimeError("Copy ok, but EXPUNGE failed")
        return {"status": "ok", "uid": _uid_to_str(uid), "source": src, "destination": dst}

    def delete_message(self, mailbox: str, uid: str, expunge: bool = False) -> Dict[str, Any]:
        self._require_conn()
        if isinstance(expunge, str):
            expunge = expunge.strip().lower() in ("1", "true", "yes", "y")
        self._select(mailbox, readonly=False)
        typ, _ = self._conn.uid("STORE", str(uid), "+FLAGS", "\\Deleted")
        if typ != "OK":
            raise RuntimeError("Marking as deleted failed")
        expunged = False
        if expunge:
            if self._has_cap("UIDPLUS"):
                typ, _ = self._conn.uid("EXPUNGE", str(uid))
            else:
                typ, _ = self._conn.expunge()
            if typ != "OK":
                raise RuntimeError("Expunge failed")
            expunged = True
        return {"status": "ok", "uid": _uid_to_str(uid), "expunged": expunged}

    def append_draft(self, mailbox: str, raw_message: str) -> Dict[str, Any]:
        self._require_conn()
        if not raw_message:
            raise RuntimeError("raw_message must not be empty")
        if HAVE_IMAP and imaplib is not None:
            internal_date = imaplib.Time2Internaldate(time.time())
        else:
            internal_date = time.strftime("%d-%b-%Y %H:%M:%S +0000", time.gmtime())
        typ, _ = self._conn.append(mailbox, "(\\Draft)", internal_date, raw_message.encode("utf-8"))
        if typ != "OK":
            raise RuntimeError("Draft append failed")
        return {"status": "ok", "mailbox": mailbox}

    def list_attachments(self, mailbox: str, uid: str, max_links: int = 3) -> Dict[str, Any]:
        self._require_conn()
        try:
            max_links = int(max_links)
        except (TypeError, ValueError):
            raise RuntimeError("max_links must be an integer")
        max_links = max(0, min(max_links, 50))
        msg = self._fetch_original(mailbox, uid)
        names = [_decode_mime(p.get_filename() or f"unnamed-attachment-{i}")
                 for i, p in enumerate(msg.walk())
                 if not p.is_multipart() and p.get_content_disposition() == "attachment"]
        return {"uid": _uid_to_str(uid), "mailbox": mailbox,
                "attachment_count": len(names), "attachments": names[:max_links] if max_links else []}

    def save_email_as_html(self, mailbox: str, uid: str, include_attachments: bool = True,
                           backup_dir: Optional[str] = None) -> Dict[str, Any]:
        self._require_conn()
        if isinstance(include_attachments, str):
            include_attachments = include_attachments.strip().lower() in ("1", "true", "yes", "y")
        path = Path(backup_dir).expanduser() if backup_dir else Path(BACKUP_DIR).expanduser()
        return _save_email_as_html(self._fetch_original(mailbox, uid), mailbox,
                                   _uid_to_str(uid), path, include_attachments)

    # ---------- SMTP-based tools ----------
    def send_message(self, to: str, subject: str, body: str,
                     cc: Optional[str] = None) -> Dict[str, Any]:
        config = _get_smtp_config()
        to_addrs = [a.strip() for a in to.split(",") if a.strip()]
        cc_addrs = [a.strip() for a in cc.split(",") if a.strip()] if cc else []
        if not to_addrs:
            raise RuntimeError("At least one recipient (to) is required")
        if not subject:
            raise RuntimeError("Subject is required")
        if not body:
            raise RuntimeError("Body is required")
        _send_smtp_message(_build_message(config["user"], to_addrs, subject, body, cc_addrs), config)
        return {"status": "sent", "from": config["user"], "to": to_addrs,
                "cc": cc_addrs, "subject": subject}

    def reply_message(self, mailbox: str, uid: str, body: str,
                      reply_all: bool = False) -> Dict[str, Any]:
        self._require_conn()
        if not body:
            raise RuntimeError("Reply body is required")
        if isinstance(reply_all, str):
            reply_all = reply_all.strip().lower() in ("1", "true", "yes", "y")
        orig = self._fetch_original(mailbox, uid)
        orig_from = _decode_mime(orig.get("From", ""))
        config, from_addr, to_addrs, cc_addrs = self._reply_addrs(orig, reply_all)
        subject = self._with_prefix(_decode_mime(orig.get("Subject", "")), "Re:")
        full_body = body + _quote_original(_extract_plain_text(orig), orig_from)
        msg_id = orig.get("Message-ID", "")
        _send_smtp_message(_build_message(from_addr, to_addrs, subject, full_body, cc_addrs,
                                          in_reply_to=msg_id,
                                          references=orig.get("References", "")), config)
        return {"status": "sent", "from": from_addr, "to": to_addrs,
                "cc": cc_addrs, "subject": subject, "in_reply_to": msg_id}

    def forward_message(self, mailbox: str, uid: str, to: str, body: str = "") -> Dict[str, Any]:
        self._require_conn()
        to_addrs = [a.strip() for a in to.split(",") if a.strip()]
        if not to_addrs:
            raise RuntimeError("At least one recipient (to) is required")
        orig = self._fetch_original(mailbox, uid)
        orig_from = _decode_mime(orig.get("From", ""))
        orig_subject = _decode_mime(orig.get("Subject", ""))
        config = _get_smtp_config()
        subject = self._with_prefix(orig_subject, "Fwd:")
        header = (f"---------- Forwarded message ----------\nFrom: {orig_from}\n"
                  f"Date: {orig.get('Date', '')}\nSubject: {orig_subject}\n\n")
        orig_body = _extract_plain_text(orig)
        full_body = f"{body}\n{header}{orig_body}" if body else header + orig_body
        _send_smtp_message(_build_message(config["user"], to_addrs, subject, full_body), config)
        return {"status": "sent", "from": config["user"], "to": to_addrs, "subject": subject}

    def draft_message(self, to: str, subject: str, body: str,
                      cc: Optional[str] = None, mailbox: str = "Drafts") -> Dict[str, Any]:
        self._require_conn()
        to_addrs = [a.strip() for a in to.split(",") if a.strip()]
        cc_addrs = [a.strip() for a in cc.split(",") if a.strip()] if cc else []
        if not to_addrs:
            raise RuntimeError("At least one recipient (to) is required")
        if not subject or not body:
            raise RuntimeError("Subject and body are required")
        config = _get_smtp_config()
        self._append_mime(mailbox, _build_message(config["user"], to_addrs, subject, body, cc_addrs))
        return {"status": "ok", "mailbox": mailbox, "from": config["user"],
                "to": to_addrs, "cc": cc_addrs, "subject": subject}

    def draft_reply(self, mailbox: str, uid: str, body: str, reply_all: bool = False,
                    draft_mailbox: str = "Drafts") -> Dict[str, Any]:
        self._require_conn()
        if not body:
            raise RuntimeError("Reply body is required")
        if isinstance(reply_all, str):
            reply_all = reply_all.strip().lower() in ("1", "true", "yes", "y")
        orig = self._fetch_original(mailbox, uid)
        orig_from = _decode_mime(orig.get("From", ""))
        config, from_addr, to_addrs, cc_addrs = self._reply_addrs(orig, reply_all)
        subject = self._with_prefix(_decode_mime(orig.get("Subject", "")), "Re:")
        full_body = body + _quote_original(_extract_plain_text(orig), orig_from)
        msg_id = orig.get("Message-ID", "")
        self._append_mime(draft_mailbox, _build_message(
            from_addr, to_addrs, subject, full_body, cc_addrs,
            in_reply_to=msg_id, references=orig.get("References", "")))
        return {"status": "ok", "mailbox": draft_mailbox, "from": from_addr,
                "to": to_addrs, "cc": cc_addrs, "subject": subject, "in_reply_to": msg_id}

    # ---------- uniform dispatch ----------
    def dispatch_tool(self, name: str, args: Optional[Dict[str, Any]] = None) -> Any:
        args = dict(args or {})
        table = {
            "list_mailboxes": lambda: self.list_mailboxes(),
            "search_messages": lambda: self.search_messages(
                mailbox=args.get("mailbox", "INBOX"), criteria=args.get("criteria", "ALL"),
                limit=args.get("limit", 20)),
            "fetch_message": lambda: self.fetch_message(
                mailbox=args["mailbox"], uid=args["uid"],
                include_body=args.get("include_body", True),
                max_body_chars=args.get("max_body_chars", DEFAULT_MAX_BODY_CHARS)),
            "mark_message": lambda: self.mark_message(
                mailbox=args["mailbox"], uid=args["uid"],
                flag=args.get("flag", "\\Seen"), add=args.get("add", True)),
            "move_message": lambda: self.move_message(
                mailbox=args.get("mailbox", args.get("mailbox_src", "")),
                uid=args["uid"], dest=args.get("dest", args.get("mailbox_dest", "")),
                **{k: v for k, v in args.items()
                   if k in ("mailbox_src", "mailbox_dest", "source", "destination")}),
            "delete_message": lambda: self.delete_message(
                mailbox=args["mailbox"], uid=args["uid"], expunge=args.get("expunge", False)),
            "append_draft": lambda: self.append_draft(
                mailbox=args.get("mailbox", "Drafts"), raw_message=args["raw_message"]),
            "list_attachments": lambda: self.list_attachments(
                mailbox=args["mailbox"], uid=args["uid"], max_links=args.get("max_links", 3)),
            "save_email_as_html": lambda: self.save_email_as_html(
                mailbox=args["mailbox"], uid=args["uid"],
                include_attachments=args.get("include_attachments", True),
                backup_dir=args.get("backup_dir")),
            "send_message": lambda: self.send_message(
                to=args["to"], subject=args["subject"], body=args["body"], cc=args.get("cc")),
            "reply_message": lambda: self.reply_message(
                mailbox=args["mailbox"], uid=args["uid"], body=args["body"],
                reply_all=args.get("reply_all", False)),
            "forward_message": lambda: self.forward_message(
                mailbox=args["mailbox"], uid=args["uid"], to=args["to"], body=args.get("body", "")),
            "draft_message": lambda: self.draft_message(
                to=args["to"], subject=args["subject"], body=args["body"], cc=args.get("cc"),
                mailbox=args.get("mailbox", "Drafts")),
            "draft_reply": lambda: self.draft_reply(
                mailbox=args["mailbox"], uid=args["uid"], body=args["body"],
                reply_all=args.get("reply_all", False),
                draft_mailbox=args.get("draft_mailbox", "Drafts")),
        }
        try:
            return table[name]()
        except KeyError:
            raise RuntimeError(f"Unknown tool: {name}")


def canonical_tools_for_provider(provider: str) -> List[Dict[str, Any]]:
    """Single canonical JSON schema shared by every provider."""
    return [
        {"name": "list_mailboxes", "description": "List all available mailboxes.",
         "parameters": {"type": "object", "properties": {}, "required": []}},
        {"name": "search_messages", "description": "Search messages with IMAP SEARCH syntax.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string", "default": "INBOX"},
                                       "criteria": {"type": "string", "default": "ALL"},
                                       "limit": {"type": "integer", "default": 20}},
                        "required": []}},
        {"name": "fetch_message", "description": "Retrieve full details (headers + plain body) by UID.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "include_body": {"type": "boolean", "default": True},
                                       "max_body_chars": {"type": "integer", "default": DEFAULT_MAX_BODY_CHARS}},
                        "required": ["mailbox", "uid"]}},
        {"name": "mark_message", "description": "Add or remove IMAP flag on a message.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "flag": {"type": "string", "default": "\\Seen"},
                                       "add": {"type": "boolean", "default": True}},
                        "required": ["mailbox", "uid"]}},
        {"name": "move_message", "description": "Move a message from one mailbox to another.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "dest": {"type": "string"}},
                        "required": ["mailbox", "uid", "dest"]}},
        {"name": "delete_message", "description": "Flag message as deleted and optionally expunge.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "expunge": {"type": "boolean", "default": False}},
                        "required": ["mailbox", "uid"]}},
        {"name": "append_draft", "description": "Create a new draft with raw RFC822 content.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string", "default": "Drafts"},
                                       "raw_message": {"type": "string"}},
                        "required": ["raw_message"]}},
        {"name": "list_attachments",
         "description": "List attachment names for a message (count + up to max_links names).",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "max_links": {"type": "integer", "default": 3}},
                        "required": ["mailbox", "uid"]}},
        {"name": "save_email_as_html",
         "description": "Save an email as HTML file with attachments for offline backup.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "include_attachments": {"type": "boolean", "default": True},
                                       "backup_dir": {"type": "string", "description": "Custom backup directory (optional)"}},
                        "required": ["mailbox", "uid"]}},
        {"name": "send_message", "description": "Send a new email via SMTP.",
         "parameters": {"type": "object",
                        "properties": {"to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                                       "subject": {"type": "string"},
                                       "body": {"type": "string", "description": "Plain text body"},
                                       "cc": {"type": "string", "description": "CC recipients, comma-separated (optional)"}},
                        "required": ["to", "subject", "body"]}},
        {"name": "reply_message", "description": "Reply to an email message.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "body": {"type": "string", "description": "Reply body text"},
                                       "reply_all": {"type": "boolean", "default": False}},
                        "required": ["mailbox", "uid", "body"]}},
        {"name": "forward_message", "description": "Forward an email message.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                                       "body": {"type": "string", "description": "Additional text to prepend (optional)", "default": ""}},
                        "required": ["mailbox", "uid", "to"]}},
        {"name": "draft_message", "description": "Create a draft email in IMAP Drafts folder.",
         "parameters": {"type": "object",
                        "properties": {"to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                                       "subject": {"type": "string"}, "body": {"type": "string", "description": "Plain text body"},
                                       "cc": {"type": "string", "description": "CC recipients, comma-separated (optional)"},
                                       "mailbox": {"type": "string", "default": "Drafts"}},
                        "required": ["to", "subject", "body"]}},
        {"name": "draft_reply", "description": "Create a reply draft in IMAP Drafts folder.",
         "parameters": {"type": "object",
                        "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"},
                                       "body": {"type": "string", "description": "Reply body text"},
                                       "reply_all": {"type": "boolean", "default": False},
                                       "draft_mailbox": {"type": "string", "default": "Drafts"}},
                        "required": ["mailbox", "uid", "body"]}},
    ]


def compile_tools(provider: str) -> List[Dict[str, Any]]:
    """Adapt the canonical schema to each provider's native tool format."""
    canonical = canonical_tools_for_provider(provider)
    p = (provider or "").strip().lower()
    if p in ("openai", "deepseek"):
        return [{"type": "function",
                 "function": {"name": t["name"], "description": t["description"],
                              "parameters": t["parameters"]}} for t in canonical]
    if p == "anthropic":
        return [{"name": t["name"], "description": t["description"],
                 "input_schema": t["parameters"]} for t in canonical]
    if p == "gemini":
        return [{"name": t["name"], "description": t["description"],
                 "parameters": t["parameters"]} for t in canonical]
    if p == "mcp":
        return [{"name": t["name"], "description": t["description"],
                 "inputSchema": t["parameters"]} for t in canonical]
    return canonical


# ---------- CLI ----------

_INT_ARGS = {"limit", "max_body_chars", "max_links"}
_BOOL_ARGS = {"add", "expunge", "include_body", "include_attachments", "reply_all"}
_IMAP_TOOLS = {"list_mailboxes", "search_messages", "fetch_message", "mark_message",
               "move_message", "delete_message", "append_draft", "list_attachments",
               "save_email_as_html", "reply_message", "forward_message",
               "draft_message", "draft_reply"}
_SMTP_ONLY_TOOLS = {"send_message"}

HELP_TEXT = """Universal IMAP Email Skill — CLI

Usage:
  python main.py --help | tools | <tool> [key=value ...] | compile_tools provider=<name>
  python main.py init                           # Interactive configuration wizard

Tools (IMAP - require connection):
  list_mailboxes
  search_messages mailbox=INBOX criteria="UNSEEN" limit=5
  fetch_message mailbox=INBOX uid=123 [include_body=true] [max_body_chars=1500]
  mark_message mailbox=INBOX uid=123 [flag="\\\\Seen"] [add=true]
  move_message mailbox=INBOX uid=123 dest=Archive
  delete_message mailbox=INBOX uid=123 [expunge=false]
  append_draft [mailbox=Drafts] raw_message="Subject: x\\n\\nbody"
  list_attachments mailbox=INBOX uid=123 [max_links=3]
  save_email_as_html mailbox=INBOX uid=123 [include_attachments=true] [backup_dir=path]
  reply_message mailbox=INBOX uid=123 body="reply text" [reply_all=false]
  forward_message mailbox=INBOX uid=123 to="addr@example.com" [body="extra text"]
  draft_message to="addr@example.com" subject="Hi" body="text" [cc="..."] [mailbox=Drafts]
  draft_reply mailbox=INBOX uid=123 body="reply text" [reply_all=false] [draft_mailbox=Drafts]

Tools (SMTP - no IMAP connection needed):
  send_message to="addr@example.com" subject="Hi" body="text" [cc="..."]

Config:
  Providers for compile_tools: openai, anthropic, gemini, deepseek, mcp, generic.
  Presets for init: gmail, outlook, qq, 163, 126, tencent, yahoo, icloud
Output is JSON on stdout; errors are {"error": "<reason>"} on stdout with exit code 1.
"""


def _coerce_arg(key: str, value: str) -> Any:
    if key in _INT_ARGS:
        try:
            return int(value)
        except ValueError:
            raise RuntimeError(f"Argument {key} must be an integer, got {value!r}")
    if key in _BOOL_ARGS:
        v = value.strip().lower()
        if v in ("1", "true", "yes", "y", "on"):
            return True
        if v in ("0", "false", "no", "n", "off"):
            return False
        raise RuntimeError(f"Argument {key} must be a boolean, got {value!r}")
    v = value.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in ("'", '"'):
        return v[1:-1]
    return v


def _parse_kv_args(items: List[str]) -> Dict[str, Any]:
    args: Dict[str, Any] = {}
    for item in items:
        if "=" not in item:
            raise RuntimeError(f"Bad argument {item!r}: expected key=value. See --help.")
        key, _, value = item.partition("=")
        key = key.strip()
        if not key:
            raise RuntimeError(f"Bad argument {item!r}: empty key.")
        args[key] = _coerce_arg(key, value)
    return args


def main(argv: Optional[List[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("--help", "-h", "help"):
        print(HELP_TEXT)
        return 0
    if argv[0] == "tools":
        print(json.dumps(canonical_tools_for_provider("generic"), indent=2))
        return 0
    if argv[0] == "compile_tools":
        try:
            args = _parse_kv_args(argv[1:])
            tools = compile_tools(str(args.get("provider", "generic")))
        except Exception as e:
            print(json.dumps({"error": str(e)}))
            return 1
        print(json.dumps(tools, indent=2))
        return 0
    if argv[0] == "init":
        return _run_init_wizard()

    tool, rest = argv[0], argv[1:]
    if tool not in _IMAP_TOOLS and tool not in _SMTP_ONLY_TOOLS:
        print(json.dumps({"error": f"Unknown tool: {tool}. See --help."}))
        return 1
    try:
        args = _parse_kv_args(rest)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    client = ImapClient()
    needs_imap = tool not in _SMTP_ONLY_TOOLS
    if needs_imap:
        try:
            client.connect()
        except Exception as e:
            print(json.dumps({"error": str(e)}))
            return 1
    try:
        try:
            result = client.dispatch_tool(tool, args)
        except TypeError as e:
            print(json.dumps({"error": f"Bad arguments for {tool}: {e}"}))
            return 1
        except KeyError as e:
            print(json.dumps({"error": f"Missing required argument {e} for {tool}"}))
            return 1
        print(json.dumps(result, indent=2, default=str))
        return 0
    except Exception as e:
        LOG.debug("tool %s failed", tool, exc_info=True)
        print(json.dumps({"error": str(e)}))
        return 1
    finally:
        client.close()


if __name__ == "__main__":
    raise SystemExit(main())
