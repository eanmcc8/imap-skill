import os
import re
import json
import ssl
import sys
import time
import logging
import uuid
import threading
import mimetypes
import getpass
import socket
from typing import Any, Dict, List, Optional
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import parsedate_to_datetime, getaddresses, formatdate
from pathlib import Path

try:
    import imaplib, email
    HAVE_IMAP = True
except Exception:  # pragma: no cover - interpreter without imaplib/email
    HAVE_IMAP = False
    imaplib = None  # type: ignore
    email = None  # type: ignore


# ---------- bootstrap: .env + logging ----------

# Keys whose values are safe to strip a trailing "# comment" from.
# Secrets (IMAP_USER/IMAP_PASSWORD/…) are deliberately excluded: a password
# may legitimately contain '#'.
_STRIP_COMMENT_KEYS = {
    "IMAP_HOST", "IMAP_PORT", "IMAP_USE_SSL", "IMAP_TIMEOUT",
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
                val = val.strip().strip('"').strip("'")
                if key in _STRIP_COMMENT_KEYS:
                    # env.example lines carry trailing "# comment" text; a copy/paste
                    # into .env must not poison numeric/boolean settings.
                    # (Never applied to secrets like IMAP_PASSWORD, which may
                    # legitimately contain '#'.)
                    val = val.split("#", 1)[0].strip().strip('"').strip("'")
                elif key == "IMAP_USER":
                    # Email logins never contain whitespace: tolerate a pasted
                    # trailing comment ("user@gmail.com  # my account").
                    val = re.split(r"\s+#", val, maxsplit=1)[0].strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception:
        pass


_load_dotenv()


def _env_int(name: str, default: int) -> int:
    """Parse an int env var tolerantly ('8000  # comment' → 8000)."""
    from logging import getLogger as _getLogger
    token = _env_token(name, "")
    if not token:
        return default
    try:
        return int(token)
    except ValueError:
        _getLogger("email_skill").warning(
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
MAX_FETCH_LIMIT = 100  # clamp for search_messages limit
TRUNCATION_SUFFIX = "… (truncated)"

# Connection pooling settings
CONNECTION_POOL_SIZE = _env_int("IMAP_POOL_SIZE", 2)
CONNECTION_IDLE_TIMEOUT = _env_int("IMAP_IDLE_TIMEOUT", 300)  # 5 minutes

# IMAP ID support (for 126/163 NetEase)
IMAP_REQUIRE_ID = _env_token("IMAP_REQUIRE_ID", "false").lower() in ("1", "true", "yes", "y")

# SMTP settings
SMTP_HOST = _env_token("SMTP_HOST", "")
SMTP_PORT = _env_int("SMTP_PORT", 465)
SMTP_USE_SSL = _env_token("SMTP_USE_SSL", "true").lower() in ("1", "true", "yes", "y")
SMTP_USER = _env_token("SMTP_USER", "")
SMTP_PASSWORD = _env_token("SMTP_PASSWORD", "")

# Offline backup settings
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


def _extract_plain_text(msg: "email.message.Message") -> str:
    """Extract first non-attachment text/plain body; '' when none."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.is_multipart():
                    continue  # container, not content
                if part.get_content_type() != "text/plain":
                    continue
                if part.get_content_disposition() == "attachment":
                    continue  # real attachment, not body
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
        payload = msg.get_payload(decode=True)
        if payload is None:
            # not encoded — may already be str
            raw = msg.get_payload()
            return str(raw or "").strip()
        try:
            return payload.decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            ).strip()
        except Exception:
            return ""
    except Exception:
        return ""


def _parse_mailbox_line(line: bytes) -> Optional[str]:
    """Robustly parse one LIST response line into a folder name."""
    try:
        text = line.decode("utf-8", "replace") if isinstance(line, bytes) else str(line)
    except Exception:
        return None
    text = text.strip()
    if not text:
        return None
    # Prefer last quoted string: (\HasNoChildren) "/" "INBOX.Sent"
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


# ---------- Structured logging with correlation IDs ----------

class CorrelationLogger:
    """Logger with automatic correlation ID injection."""
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.correlation_id: Optional[str] = None

    def set_correlation_id(self, corr_id: Optional[str] = None) -> str:
        self.correlation_id = corr_id or str(uuid.uuid4())[:8]
        return self.correlation_id

    def _log(self, level: int, msg: str, *args, **kwargs):
        extra = kwargs.pop("extra", {})
        if self.correlation_id:
            extra["correlation_id"] = self.correlation_id
        self.logger.log(level, msg, *args, extra=extra, **kwargs)

    def debug(self, msg: str, *args, **kwargs): self._log(logging.DEBUG, msg, *args, **kwargs)
    def info(self, msg: str, *args, **kwargs): self._log(logging.INFO, msg, *args, **kwargs)
    def warning(self, msg: str, *args, **kwargs): self._log(logging.WARNING, msg, *args, **kwargs)
    def error(self, msg: str, *args, **kwargs): self._log(logging.ERROR, msg, *args, **kwargs)
    def exception(self, msg: str, *args, **kwargs): self._log(logging.ERROR, msg, *args, exc_info=True, **kwargs)


LOG = CorrelationLogger(logging.getLogger("email_skill"))


# ---------- IMAP ID Support (RFC 2971 for 126/163) ----------

def _send_imap_id(conn, host: str, port: int) -> None:
    """Send IMAP ID command (RFC 2971) for providers like 126/163 that require it."""
    if not IMAP_REQUIRE_ID:
        return
    
    # Send CAPABILITY first to check support
    try:
        typ, _ = conn.capability()
    except Exception:
        pass
    
    # Send ID command with client identification
    id_params = '("name" "imap-skill") ("version" "1.0") ("vendor" "kilo-code")'
    try:
        typ, resp = conn._simple_command(b"ID", id_params.encode())
        if typ != "OK":
            LOG.warning(f"IMAP ID command failed: {resp}")
    except Exception as e:
        LOG.warning(f"IMAP ID command error: {e}")


# ---------- SMTP Client ----------

def _get_smtp_config() -> Dict[str, Any]:
    """Get SMTP configuration from environment."""
    return {
        "host": SMTP_HOST or _env_token("IMAP_HOST", "smtp.gmail.com"),
        "port": SMTP_PORT,
        "use_ssl": SMTP_USE_SSL,
        "user": SMTP_USER or os.getenv("IMAP_USER", ""),
        "password": SMTP_PASSWORD or os.getenv("IMAP_PASSWORD", ""),
    }


def _create_smtp_client(config: Dict[str, Any]):
    """Create and authenticate SMTP client."""
    import smtplib
    
    host = config["host"]
    port = config["port"]
    user = config["user"]
    password = config["password"]
    use_ssl = config["use_ssl"]
    
    if use_ssl:
        client = smtplib.SMTP_SSL(host, port, timeout=DEFAULT_TIMEOUT)
    else:
        client = smtplib.SMTP(host, port, timeout=DEFAULT_TIMEOUT)
        client.starttls(context=ssl.create_default_context())
    
    client.login(user, password)
    return client


def _send_smtp_message(msg: MIMEMultipart, config: Dict[str, Any]) -> None:
    """Send message via SMTP."""
    client = _create_smtp_client(config)
    try:
        client.send_message(msg)
    finally:
        client.quit()


def _build_message(
    from_addr: str,
    to_addrs: List[str],
    subject: str,
    body: str,
    cc_addrs: Optional[List[str]] = None,
    in_reply_to: Optional[str] = None,
    references: Optional[str] = None,
) -> MIMEMultipart:
    """Build MIME message with proper headers."""
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
        refs = references or ""
        msg["References"] = f"{refs} {in_reply_to}".strip()
    
    msg.attach(MIMEText(body, "plain", "utf-8"))
    return msg


def _extract_email_address(addr_str: str) -> str:
    """Extract email address from 'Name <addr>' format."""
    if not addr_str:
        return ""
    # Try to extract from <addr> format
    import re
    m = re.search(r"<([^>]+)>", addr_str)
    if m:
        return m.group(1)
    # If no brackets, check if it looks like an email
    if "@" in addr_str:
        return addr_str.strip()
    return addr_str


def _extract_email_addresses(header_value: str) -> List[str]:
    """Extract all email addresses from a header value."""
    if not header_value:
        return []
    addrs = getaddresses([header_value])
    return [addr for _, addr in addrs if addr]


def _quote_original(body: str, from_label: str) -> str:
    """Format original message body as inline quote."""
    lines = body.split("\n")
    quoted = "\n".join(f"> {line}" for line in lines)
    return f"\n---- {from_label} wrote ----\n{quoted}\n--------------------------------"


# ---------- Preset Configurations ----------

PRESETS = {
    "gmail": {
        "name": "Gmail",
        "imap_host": "imap.gmail.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": False,
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
        "smtp_use_ssl": False,  # STARTTLS
    },
    "outlook": {
        "name": "Outlook / Hotmail / Office 365",
        "imap_host": "outlook.office365.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": False,
        "smtp_host": "smtp.office365.com",
        "smtp_port": 587,
        "smtp_use_ssl": False,  # STARTTLS
    },
    "qq": {
        "name": "QQ Mail",
        "imap_host": "imap.qq.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": False,
        "smtp_host": "smtp.qq.com",
        "smtp_port": 465,
        "smtp_use_ssl": True,
    },
    "163": {
        "name": "163 NetEase",
        "imap_host": "imap.163.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": True,
        "smtp_host": "smtp.163.com",
        "smtp_port": 465,
        "smtp_use_ssl": True,
    },
    "126": {
        "name": "126 NetEase",
        "imap_host": "imap.126.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": True,
        "smtp_host": "smtp.126.com",
        "smtp_port": 465,
        "smtp_use_ssl": True,
    },
    "tencent": {
        "name": "Tencent Enterprise Mail",
        "imap_host": "imap.exmail.qq.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": False,
        "smtp_host": "smtp.exmail.qq.com",
        "smtp_port": 465,
        "smtp_use_ssl": True,
    },
    "yahoo": {
        "name": "Yahoo Mail",
        "imap_host": "imap.mail.yahoo.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": False,
        "smtp_host": "smtp.mail.yahoo.com",
        "smtp_port": 465,
        "smtp_use_ssl": True,
    },
    "icloud": {
        "name": "iCloud Mail",
        "imap_host": "imap.mail.me.com",
        "imap_port": 993,
        "imap_use_ssl": True,
        "imap_require_id": False,
        "smtp_host": "smtp.mail.me.com",
        "smtp_port": 587,
        "smtp_use_ssl": False,  # STARTTLS
    },
}


def _apply_preset(preset_key: str) -> Dict[str, str]:
    """Apply preset configuration, return env var mappings."""
    preset = PRESETS.get(preset_key.lower())
    if not preset:
        return {}
    return {
        "IMAP_HOST": preset["imap_host"],
        "IMAP_PORT": str(preset["imap_port"]),
        "IMAP_USE_SSL": str(preset["imap_use_ssl"]).lower(),
        "IMAP_REQUIRE_ID": str(preset["imap_require_id"]).lower(),
        "SMTP_HOST": preset["smtp_host"],
        "SMTP_PORT": str(preset["smtp_port"]),
        "SMTP_USE_SSL": str(preset["smtp_use_ssl"]).lower(),
    }


# ---------- Interactive Init Wizard ----------

def _run_init_wizard() -> int:
    """Interactive configuration wizard."""
    print("=" * 60)
    print("  imap-skill Configuration Wizard")
    print("=" * 60)
    print()
    
    # Check if .env exists
    base = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(base, ".env")
    
    if os.path.exists(env_path):
        overwrite = input(f".env already exists at {env_path}. Overwrite? [y/N]: ").strip().lower()
        if overwrite not in ("y", "yes"):
            print("Cancelled.")
            return 1
    
    print("\nSelect email provider preset:")
    preset_keys = list(PRESETS.keys())
    for i, key in enumerate(preset_keys):
        print(f"  {i + 1}. {PRESETS[key]['name']} ({key})")
    print(f"  {len(preset_keys) + 1}. Custom / Other")
    print()
    
    choice = input(f"Enter choice [1-{len(preset_keys) + 1}]: ").strip()
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(preset_keys):
            preset = PRESETS[preset_keys[idx]]
            print(f"\nSelected: {preset['name']}")
        else:
            preset = None
            print("\nCustom configuration selected.")
    except (ValueError, IndexError):
        preset = None
        print("\nInvalid choice, using custom configuration.")
    
    # Collect credentials
    print("\n--- Credentials ---")
    email_addr = input("Email address: ").strip()
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
    
    display_name = input("Display name (optional): ").strip()
    
    # Build .env content
    lines = [
        "# imap-skill configuration",
        f"# Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "# IMAP Settings",
        f"IMAP_HOST={preset['imap_host'] if preset else input('IMAP Host: ').strip()}",
        f"IMAP_PORT={preset['imap_port'] if preset else input('IMAP Port [993]: ').strip() or '993'}",
        f"IMAP_USE_SSL={str(preset['imap_use_ssl']).lower() if preset else 'true'}",
        f"IMAP_REQUIRE_ID={str(preset['imap_require_id']).lower() if preset else 'false'}",
        f"IMAP_USER={email_addr}",
        f"IMAP_PASSWORD={password}",
        "",
        "# SMTP Settings",
        f"SMTP_HOST={preset['smtp_host'] if preset else input('SMTP Host: ').strip()}",
        f"SMTP_PORT={preset['smtp_port'] if preset else input('SMTP Port [465]: ').strip() or '465'}",
        f"SMTP_USE_SSL={str(preset['smtp_use_ssl']).lower() if preset else 'true'}",
        f"SMTP_USER={email_addr}",
        f"SMTP_PASSWORD={password}",
        "",
        "# Optional Settings",
        f"MAX_BODY_CHARS=8000",
        f"LOG_LEVEL=INFO",
        f"IMAP_POOL_SIZE=2",
        f"IMAP_IDLE_TIMEOUT=300",
        f"IMAP_BACKUP_DIR=~/imap_backup",
    ]
    
    if display_name:
        lines.insert(3, f"# Display Name: {display_name}")
    
    print("\n--- Preview ---")
    print("\n".join(lines))
    print("---")
    
    confirm = input("\nSave to .env? [Y/n]: ").strip().lower()
    if confirm in ("", "y", "yes"):
        with open(env_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        # Secure permissions
        try:
            os.chmod(env_path, 0o600)
        except Exception:
            pass
        print(f"\n✓ Configuration saved to {env_path}")
        print("  You can now use the CLI or start the MCP server.")
        return 0
    else:
        print("\nCancelled.")
        return 1


# ---------- Connection Pool ----------

class ImapConnectionPool:
    """Thread-safe IMAP connection pool with keepalive."""

    def __init__(self, host: str, port: int, use_ssl: bool, user: str, password: str, timeout: int, require_id: bool = False):
        self.host = host
        self.port = port
        self.use_ssl = use_ssl
        self.user = user
        self.password = password
        self.timeout = timeout
        self.require_id = require_id
        self._pool: List[imaplib.IMAP4] = []
        self._lock = threading.Lock()
        self._in_use: set = set()
        self._created_at: Dict[imaplib.IMAP4, float] = {}
        self._last_used: Dict[imaplib.IMAP4, float] = {}
        self._keepalive_thread: Optional[threading.Thread] = None
        self._stop_keepalive = threading.Event()

    def _create_connection(self) -> imaplib.IMAP4:
        """Create a new IMAP connection."""
        if self.use_ssl:
            conn = imaplib.IMAP4_SSL(self.host, self.port, timeout=self.timeout)
        else:
            conn = imaplib.IMAP4(self.host, self.port, timeout=self.timeout)
            conn = _upgrade_to_tls(conn, self.host, self.port)
        conn.login(self.user, self.password)
        # Send IMAP ID if required (for 126/163 NetEase)
        _send_imap_id(conn, self.host, self.port)
        return conn

    @staticmethod
    def _read_caps(conn) -> set:
        return _read_caps(conn)

    def _is_connection_valid(self, conn: imaplib.IMAP4) -> bool:
        """Check if connection is still alive with NOOP."""
        try:
            typ, _ = conn.noop()
            return typ == "OK"
        except Exception:
            return False

    def get_connection(self) -> imaplib.IMAP4:
        """Get a connection from the pool or create a new one."""
        with self._lock:
            # Try to reuse an existing connection
            while self._pool:
                conn = self._pool.pop()
                if self._is_connection_valid(conn):
                    self._in_use.add(conn)
                    self._last_used[conn] = time.time()
                    LOG.debug("Reused pooled connection", extra={"pool_size": len(self._pool)})
                    return conn
                else:
                    LOG.debug("Discarded stale connection")
                    try:
                        conn.logout()
                    except Exception:
                        pass

            # Create new connection if pool exhausted
            if len(self._in_use) < CONNECTION_POOL_SIZE:
                conn = self._create_connection()
                self._in_use.add(conn)
                self._created_at[conn] = time.time()
                self._last_used[conn] = time.time()
                LOG.debug("Created new connection", extra={"active": len(self._in_use)})
                return conn

            # Pool exhausted - wait for a connection to be returned
            raise RuntimeError("Connection pool exhausted")

    def return_connection(self, conn: imaplib.IMAP4) -> None:
        """Return a connection to the pool."""
        with self._lock:
            if conn in self._in_use:
                self._in_use.remove(conn)
                if self._is_connection_valid(conn):
                    self._pool.append(conn)
                    self._last_used[conn] = time.time()
                    LOG.debug("Returned connection to pool", extra={"pool_size": len(self._pool)})
                else:
                    LOG.debug("Discarded invalid connection on return")
                    try:
                        conn.logout()
                    except Exception:
                        pass
                    self._cleanup_connection(conn)

    def _cleanup_connection(self, conn: imaplib.IMAP4) -> None:
        """Clean up connection tracking."""
        self._created_at.pop(conn, None)
        self._last_used.pop(conn, None)

    def close_all(self) -> None:
        """Close all connections in the pool."""
        with self._lock:
            for conn in self._pool:
                try:
                    conn.logout()
                except Exception:
                    pass
            self._pool.clear()
            for conn in self._in_use:
                try:
                    conn.logout()
                except Exception:
                    pass
            self._in_use.clear()
            self._created_at.clear()
            self._last_used.clear()
            LOG.info("Closed all connections in pool")

    def start_keepalive(self) -> None:
        """Start background keepalive thread."""
        if self._keepalive_thread and self._keepalive_thread.is_alive():
            return
        self._stop_keepalive.clear()
        self._keepalive_thread = threading.Thread(target=self._keepalive_loop, daemon=True)
        self._keepalive_thread.start()
        LOG.debug("Started keepalive thread")

    def stop_keepalive(self) -> None:
        """Stop background keepalive thread."""
        self._stop_keepalive.set()
        if self._keepalive_thread:
            self._keepalive_thread.join(timeout=5)
        LOG.debug("Stopped keepalive thread")

    def _keepalive_loop(self) -> None:
        """Periodically send NOOP to keep connections alive."""
        while not self._stop_keepalive.is_set():
            time.sleep(60)  # Check every minute
            if self._stop_keepalive.is_set():
                break

            with self._lock:
                now = time.time()
                # Check idle connections
                stale_conns = []
                for conn in self._pool[:]:
                    last_used = self._last_used.get(conn, 0)
                    if now - last_used > CONNECTION_IDLE_TIMEOUT:
                        stale_conns.append(conn)

                for conn in stale_conns:
                    LOG.debug("Closing idle connection")
                    self._pool.remove(conn)
                    try:
                        conn.logout()
                    except Exception:
                        pass
                    self._cleanup_connection(conn)

                # Send NOOP to active connections
                for conn in list(self._in_use):
                    try:
                        conn.noop()
                    except Exception:
                        LOG.debug("NOOP failed, connection may be dead")
                        self._in_use.remove(conn)
                        try:
                            conn.logout()
                        except Exception:
                            pass
                        self._cleanup_connection(conn)


# ---------- Offline Backup Helpers ----------

def _extract_html_body(msg: "email.message.Message") -> str:
    """Extract first non-attachment text/html body; '' when none."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.is_multipart():
                    continue
                if part.get_content_type() != "text/html":
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
        if msg.get_content_type() == "text/html":
            payload = msg.get_payload(decode=True)
            if payload is None:
                raw = msg.get_payload()
                return str(raw or "").strip()
            try:
                return payload.decode(
                    msg.get_content_charset() or "utf-8", errors="replace"
                ).strip()
            except Exception:
                return ""
    except Exception:
        return ""
    return ""


def _extract_attachments(msg: "email.message.Message", backup_dir: Path, uid: str) -> List[Dict[str, Any]]:
    """Extract and save attachments, return list of attachment info."""
    attachments = []
    backup_dir.mkdir(parents=True, exist_ok=True)

    for i, part in enumerate(msg.walk()):
        if part.is_multipart():
            continue
        if part.get_content_disposition() != "attachment":
            continue

        fname = part.get_filename() or f"unnamed-attachment-{i}"
        fname = _decode_mime(fname)

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        # Save attachment
        safe_fname = re.sub(r'[<>:"/\\|?*]', '_', fname)
        ext = Path(safe_fname).suffix
        if not ext:
            ext = mimetypes.guess_extension(part.get_content_type()) or ""
        save_name = f"{uid}_{i}{ext}"
        save_path = backup_dir / save_name

        try:
            save_path.write_bytes(payload)
            attachments.append({
                "original_name": fname,
                "saved_as": save_name,
                "size": len(payload),
                "content_type": part.get_content_type(),
                "path": str(save_path)
            })
        except Exception as e:
            LOG.warning(f"Failed to save attachment {fname}: {e}")

    return attachments


def _save_email_as_html(
    msg: "email.message.Message",
    mailbox: str,
    uid: str,
    backup_dir: Path,
    include_attachments: bool = True
) -> Dict[str, Any]:
    """Save email as HTML file with embedded resources."""
    backup_dir.mkdir(parents=True, exist_ok=True)

    # Extract headers
    subject = _decode_mime(msg.get("Subject", ""))
    from_addr = _decode_mime(msg.get("From", ""))
    to_addr = _decode_mime(msg.get("To", ""))
    date = msg.get("Date", "")
    message_id = msg.get("Message-ID", "")

    # Extract HTML body
    html_body = _extract_html_body(msg)
    plain_body = _extract_plain_text(msg)

    # Use HTML if available, otherwise convert plain text
    if not html_body and plain_body:
        html_body = f"<pre>{html_body.replace('&', '&').replace('<', '<').replace('>', '>')}</pre>"
    elif not html_body:
        html_body = "<p><em>No body content</em></p>"

    # Extract attachments
    attachments = []
    if include_attachments:
        attachments = _extract_attachments(msg, backup_dir, uid)

    # Generate HTML file
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{subject}</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
        .header {{ border-bottom: 1px solid #ddd; padding-bottom: 1rem; margin-bottom: 1rem; }}
        .header h1 {{ margin: 0 0 0.5rem 0; font-size: 1.5rem; }}
        .meta {{ color: #666; font-size: 0.9rem; }}
        .meta div {{ margin: 0.25rem 0; }}
        .body {{ line-height: 1.6; }}
        .attachments {{ margin-top: 2rem; padding-top: 1rem; border-top: 1px solid #ddd; }}
        .attachment {{ display: inline-block; margin: 0.5rem; padding: 0.5rem; background: #f5f5f5; border-radius: 4px; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>{subject}</h1>
        <div class="meta">
            <div><strong>From:</strong> {from_addr}</div>
            <div><strong>To:</strong> {to_addr}</div>
            <div><strong>Date:</strong> {date}</div>
            <div><strong>Message-ID:</strong> {message_id}</div>
            <div><strong>Mailbox:</strong> {mailbox}</div>
            <div><strong>UID:</strong> {uid}</div>
        </div>
    </div>
    <div class="body">
        {html_body}
    </div>
    {f'<div class="attachments"><h2>Attachments ({len(attachments)})</h2>' + ''.join(f'<div class="attachment"><a href="{a["saved_as"]}">{a["original_name"]}</a> ({a["size"]} bytes)</div>' for a in attachments) + '</div>' if attachments else ''}
</body>
</html>"""

    # Save HTML file
    safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:100]
    html_filename = f"{uid}_{safe_subject}.html"
    html_path = backup_dir / html_filename
    html_path.write_text(html_content, encoding="utf-8")

    return {
        "html_file": str(html_path),
        "subject": subject,
        "from": from_addr,
        "to": to_addr,
        "date": date,
        "message_id": message_id,
        "mailbox": mailbox,
        "uid": uid,
        "attachments": attachments,
        "has_html_body": bool(_extract_html_body(msg)),
        "has_plain_body": bool(plain_body)
    }


class ImapClient:
    """
    Real IMAP transport backed by environment variables.
    Uses connection pool for efficient connection reuse.
    Throws network/connection exceptions; no simulated fallback intended.
    Usable as a context manager: `with ImapClient.connect_once() as c: ...`
    """
    _pool: Optional["ImapConnectionPool"] = None
    _pool_lock = threading.Lock()

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
        self._pool_conn = None
        self._corr_id = None

    @classmethod
    def _get_pool(cls) -> "ImapConnectionPool":
        """Get or create the global connection pool."""
        with cls._pool_lock:
            if cls._pool is None:
                if not HAVE_IMAP:
                    raise RuntimeError("imaplib/email unavailable in this interpreter")
                if not os.getenv("IMAP_USER") or not os.getenv("IMAP_PASSWORD"):
                    raise RuntimeError("IMAP_USER and IMAP_PASSWORD must be set (env or .env)")
                if any(c.isspace() for c in os.getenv("IMAP_USER", "")):
                    raise RuntimeError("IMAP_USER must be the bare email address with no spaces")
                cls._pool = ImapConnectionPool(
                    host=_env_token("IMAP_HOST", "imap.gmail.com"),
                    port=_env_int("IMAP_PORT", 993),
                    use_ssl=_env_token("IMAP_USE_SSL", "true").lower() in ("1", "true", "yes", "y", "on", "auto"),
                    user=os.getenv("IMAP_USER", ""),
                    password=os.getenv("IMAP_PASSWORD", ""),
                    timeout=DEFAULT_TIMEOUT,
                    require_id=IMAP_REQUIRE_ID,
                )
                cls._pool.start_keepalive()
            return cls._pool

    def _acquire_connection(self) -> None:
        """Acquire a connection from the pool."""
        pool = self._get_pool()
        self._pool_conn = pool.get_connection()
        self._conn = self._pool_conn
        self._corr_id = LOG.set_correlation_id()

    def _release_connection(self) -> None:
        """Release connection back to the pool."""
        if self._pool_conn:
            pool = self._get_pool()
            pool.return_connection(self._pool_conn)
            self._pool_conn = None
            self._conn = None

    # -- lifecycle --
    def connect(self) -> None:
        """Acquire connection from pool; raises RuntimeError on any failure."""
        self._acquire_connection()

    def close(self) -> None:
        self._release_connection()

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
        """Close all connections in the pool (for shutdown)."""
        with cls._pool_lock:
            if cls._pool:
                cls._pool.stop_keepalive()
                cls._pool.close_all()
                cls._pool = None

    # -- internals --
    def _require_conn(self) -> None:
        if not self._conn:
            raise RuntimeError("Not connected to IMAP (call connect() first)")

    def _has_cap(self, name: str) -> bool:
        try:
            caps = getattr(self._conn, "capabilities", None) or ()
            norm = []
            for c in caps:
                if isinstance(c, bytes):
                    norm.append(c.decode("ascii", "ignore").upper())
                else:
                    norm.append(str(c).upper())
            return name.upper() in norm
        except Exception:
            return False

    @staticmethod
    def _read_caps(conn) -> set:
        """Refresh (CAPABILITY command when available) and normalize caps."""
        try:
            cap_cmd = getattr(conn, "capability", None)
            if cap_cmd is not None:
                cap_cmd()
        except Exception:
            pass
        caps = set()
        for c in getattr(conn, "capabilities", None) or ():
            caps.add(c.decode("ascii", "ignore").upper()
                     if isinstance(c, bytes) else str(c).upper())
        return caps

    def _select(self, mailbox: str, readonly: bool) -> None:
        self._require_conn()
        typ, _ = self._conn.select(mailbox, readonly=readonly)
        if typ != "OK":
            raise RuntimeError(f"Cannot select mailbox {mailbox!r}")

    # ---------- Tool implementations ----------
    def list_mailboxes(self) -> List[str]:
        """Return a list of folder names (INBOX, Sent, Drafts, ...)."""
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

    def search_messages(self,
                        mailbox: str = "INBOX",
                        criteria: str = "ALL",
                        limit: int = 20) -> List[Dict[str, Any]]:
        """
        Run UID SEARCH and page results (most recent `limit`).
        Single batched FETCH (1 round-trip) instead of N per-message fetches.
        Returns [{uid, subject, from, to, date}].
        """
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
        uid_set = b",".join(wanted)

        typ, raw = self._conn.uid(
            "FETCH", uid_set,
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

        # Preserve UID order requested (ascending); fall back to server order.
        ordered_uids = [_uid_to_str(u) for u in wanted if _uid_to_str(u) in headers_by_uid]
        if not ordered_uids:
            ordered_uids = order

        results: List[Dict[str, Any]] = []
        for uid in ordered_uids:
            try:
                msg = email.message_from_bytes(headers_by_uid.get(uid, b""))
            except Exception:
                continue
            results.append({
                "uid": uid,
                "subject": _decode_mime(msg.get("Subject", "")),
                "from": _decode_mime(msg.get("From", "")),
                "to": _decode_mime(msg.get("To", "")),
                "date": msg.get("Date", "") or "",
            })
        return results

    def fetch_message(self,
                      mailbox: str,
                      uid: str,
                      include_body: bool = True,
                      max_body_chars: Optional[int] = None) -> Dict[str, Any]:
        """
        Fetch full message metadata + plain-text body.
        BODY.PEEK avoids marking the message Seen as a side effect.
        Body truncated to max_body_chars (env MAX_BODY_CHARS default 8000);
        suffix appended only when actually truncated.
        """
        self._require_conn()
        if max_body_chars is None:
            max_body_chars = DEFAULT_MAX_BODY_CHARS
        try:
            max_body_chars = int(max_body_chars)
        except (TypeError, ValueError):
            raise RuntimeError("max_body_chars must be an integer")
        if isinstance(include_body, str):
            include_body = include_body.strip().lower() in ("1", "true", "yes", "y")

        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw_msg = first[1] if isinstance(first, tuple) else first
        if not raw_msg:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw_msg, str):
            raw_msg = raw_msg.encode("utf-8", "replace")
        msg = email.message_from_bytes(raw_msg)

        body = ""
        if include_body:
            full = _extract_plain_text(msg)  # single pass (was called twice)
            if len(full) > max_body_chars:
                body = full[:max_body_chars] + TRUNCATION_SUFFIX
            else:
                body = full

        return {
            "uid": _uid_to_str(uid),
            "subject": _decode_mime(msg.get("Subject", "")),
            "from": _decode_mime(msg.get("From", "")),
            "to": _decode_mime(msg.get("To", "")),
            "date": msg.get("Date", "") or "",
            "body": body,
        }

    def mark_message(self,
                     mailbox: str,
                     uid: str,
                     flag: str = "\\Seen",
                     add: bool = True) -> Dict[str, Any]:
        """Add or remove IMAP flag via STORE."""
        self._require_conn()
        if isinstance(add, str):
            add = add.strip().lower() in ("1", "true", "yes", "y")
        self._select(mailbox, readonly=False)
        op = "+FLAGS" if add else "-FLAGS"
        typ, _ = self._conn.uid("STORE", str(uid), op, flag)
        if typ != "OK":
            raise RuntimeError(f"Failed to {'add' if add else 'remove'} flag {flag}")
        return {
            "status": "ok",
            "uid": _uid_to_str(uid),
            "flag": flag,
            "added": bool(add),
        }

    def move_message(self,
                     mailbox: str = "",
                     uid: str = "",
                     dest: str = "",
                     **aliases: Any) -> Dict[str, Any]:
        """
        MOVE UID if supported, else COPY + STORE \\Deleted + EXPUNGE.
        Canonical arg names: mailbox, uid, dest.
        Backward-compat aliases: mailbox_src/mailbox_dest/source/destination.
        """
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
                if typ != "OK":
                    raise RuntimeError("Copy ok, but UID EXPUNGE failed")
            else:
                typ, _ = self._conn.expunge()
                if typ != "OK":
                    raise RuntimeError("Copy ok, but EXPUNGE failed")

        return {
            "status": "ok",
            "uid": _uid_to_str(uid),
            "source": src,
            "destination": dst,
        }

    def delete_message(self,
                       mailbox: str,
                       uid: str,
                       expunge: bool = False) -> Dict[str, Any]:
        """STORE +FLAGS \\Deleted; expunge when requested."""
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
                if typ != "OK":
                    raise RuntimeError("Expunge failed")
            else:
                typ, _ = self._conn.expunge()
                if typ != "OK":
                    raise RuntimeError("Expunge failed")
            expunged = True
        return {
            "status": "ok",
            "uid": _uid_to_str(uid),
            "expunged": expunged,
        }

    def append_draft(self,
                     mailbox: str,
                     raw_message: str) -> Dict[str, Any]:
        """Append raw RFC822 message to a mailbox (default Drafts)."""
        self._require_conn()
        if not raw_message:
            raise RuntimeError("raw_message must not be empty")
        if HAVE_IMAP and imaplib is not None:
            internal_date = imaplib.Time2Internaldate(time.time())
        else:
            internal_date = time.strftime("%d-%b-%Y %H:%M:%S +0000", time.gmtime())
        typ, _ = self._conn.append(
            mailbox, "(\\Draft)", internal_date, raw_message.encode("utf-8")
        )
        if typ != "OK":
            raise RuntimeError("Draft append failed")
        return {"status": "ok", "mailbox": mailbox}

    def list_attachments(self,
                         mailbox: str,
                         uid: str,
                         max_links: int = 3) -> Dict[str, Any]:
        """List attachment names (BODY.PEEK: no Seen side effect)."""
        self._require_conn()
        try:
            max_links = int(max_links)
        except (TypeError, ValueError):
            raise RuntimeError("max_links must be an integer")
        max_links = max(0, min(max_links, 50))
        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw_msg = first[1] if isinstance(first, tuple) else first
        if not raw_msg:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw_msg, str):
            raw_msg = raw_msg.encode("utf-8", "replace")
        msg = email.message_from_bytes(raw_msg)
        names: List[str] = []
        for i, part in enumerate(msg.walk()):
            if part.is_multipart():
                continue
            if part.get_content_disposition() != "attachment":
                continue
            fname = part.get_filename() or f"unnamed-attachment-{i}"
            names.append(_decode_mime(fname))
        return {
            "uid": _uid_to_str(uid),
            "mailbox": mailbox,
            "attachment_count": len(names),
            "attachments": names[:max_links] if max_links else [],
        }

    def save_email_as_html(self,
                           mailbox: str,
                           uid: str,
                           include_attachments: bool = True,
                           backup_dir: Optional[str] = None) -> Dict[str, Any]:
        """
        Save an email as an HTML file with embedded resources for offline backup.

        Args:
            mailbox: Source mailbox name
            uid: Message UID
            include_attachments: Whether to extract and save attachments
            backup_dir: Custom backup directory (uses IMAP_BACKUP_DIR env var if not specified)

        Returns:
            Dict with html_file path, metadata, and attachment info
        """
        self._require_conn()
        if isinstance(include_attachments, str):
            include_attachments = include_attachments.strip().lower() in ("1", "true", "yes", "y")

        backup_path = Path(backup_dir) if backup_dir else Path(BACKUP_DIR)

        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw_msg = first[1] if isinstance(first, tuple) else first
        if not raw_msg:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw_msg, str):
            raw_msg = raw_msg.encode("utf-8", "replace")
        msg = email.message_from_bytes(raw_msg)

        return _save_email_as_html(msg, mailbox, uid, backup_path, include_attachments)

    # ---------- SMTP-based Tools ----------
    
    def send_message(self,
                     to: str,
                     subject: str,
                     body: str,
                     cc: Optional[str] = None) -> Dict[str, Any]:
        """
        Send a new email message via SMTP.
        
        Args:
            to: Recipient email address(es), comma-separated
            subject: Email subject
            body: Plain text body
            cc: CC recipients, comma-separated (optional)
            
        Returns:
            Dict with status
        """
        config = _get_smtp_config()
        from_addr = config["user"]
        
        to_addrs = [addr.strip() for addr in to.split(",") if addr.strip()]
        cc_addrs = [addr.strip() for addr in cc.split(",") if addr.strip()] if cc else []
        
        if not to_addrs:
            raise RuntimeError("At least one recipient (to) is required")
        if not subject:
            raise RuntimeError("Subject is required")
        if not body:
            raise RuntimeError("Body is required")
        
        msg = _build_message(from_addr, to_addrs, subject, body, cc_addrs)
        _send_smtp_message(msg, config)
        
        return {
            "status": "sent",
            "from": from_addr,
            "to": to_addrs,
            "cc": cc_addrs,
            "subject": subject,
        }

    def reply_message(self,
                      mailbox: str,
                      uid: str,
                      body: str,
                      reply_all: bool = False) -> Dict[str, Any]:
        """
        Reply to an email message.
        
        Args:
            mailbox: Source mailbox
            uid: Message UID to reply to
            body: Reply body text
            reply_all: Whether to reply to all recipients
            
        Returns:
            Dict with status
        """
        self._require_conn()
        if not body:
            raise RuntimeError("Reply body is required")
        
        # Fetch original message
        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw_msg = first[1] if isinstance(first, tuple) else first
        if not raw_msg:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw_msg, str):
            raw_msg = raw_msg.encode("utf-8", "replace")
        orig_msg = email.message_from_bytes(raw_msg)
        
        # Extract original headers
        orig_from = _decode_mime(orig_msg.get("From", ""))
        orig_to = _decode_mime(orig_msg.get("To", ""))
        orig_cc = _decode_mime(orig_msg.get("Cc", ""))
        orig_subject = _decode_mime(orig_msg.get("Subject", ""))
        orig_message_id = orig_msg.get("Message-ID", "")
        orig_references = orig_msg.get("References", "")
        
        config = _get_smtp_config()
        from_addr = config["user"]
        
        # Build recipient lists
        to_addrs = [_extract_email_address(orig_from)]
        
        cc_addrs = []
        if reply_all:
            # Add original To and Cc (except ourselves and original sender)
            all_recipients = _extract_email_addresses(orig_to) + _extract_email_addresses(orig_cc)
            self_addr = from_addr.lower()
            sender_addr = _extract_email_address(orig_from).lower()
            for addr in all_recipients:
                addr_lower = addr.lower()
                if addr_lower != self_addr and addr_lower != sender_addr:
                    cc_addrs.append(addr)
        
        # Build subject with Re: prefix
        if not orig_subject.upper().startswith("RE:"):
            subject = f"Re: {orig_subject}"
        else:
            subject = orig_subject
        
        # Build body with quoted original
        orig_body = _extract_plain_text(orig_msg)
        full_body = body + _quote_original(orig_body, orig_from)
        
        msg = _build_message(
            from_addr, to_addrs, subject, full_body, cc_addrs,
            in_reply_to=orig_message_id, references=orig_references
        )
        _send_smtp_message(msg, config)
        
        return {
            "status": "sent",
            "from": from_addr,
            "to": to_addrs,
            "cc": cc_addrs,
            "subject": subject,
            "in_reply_to": orig_message_id,
        }

    def forward_message(self,
                        mailbox: str,
                        uid: str,
                        to: str,
                        body: str = "") -> Dict[str, Any]:
        """
        Forward an email message.
        
        Args:
            mailbox: Source mailbox
            uid: Message UID to forward
            to: Recipient email address(es), comma-separated
            body: Additional text to prepend (optional)
            
        Returns:
            Dict with status
        """
        self._require_conn()
        
        to_addrs = [addr.strip() for addr in to.split(",") if addr.strip()]
        if not to_addrs:
            raise RuntimeError("At least one recipient (to) is required")
        
        # Fetch original message
        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw_msg = first[1] if isinstance(first, tuple) else first
        if not raw_msg:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw_msg, str):
            raw_msg = raw_msg.encode("utf-8", "replace")
        orig_msg = email.message_from_bytes(raw_msg)
        
        # Extract original headers
        orig_from = _decode_mime(orig_msg.get("From", ""))
        orig_subject = _decode_mime(orig_msg.get("Subject", ""))
        orig_date = orig_msg.get("Date", "")
        
        config = _get_smtp_config()
        from_addr = config["user"]
        
        # Build subject with Fwd: prefix
        if not orig_subject.upper().startswith("FWD:"):
            subject = f"Fwd: {orig_subject}"
        else:
            subject = orig_subject
        
        # Build body
        orig_body = _extract_plain_text(orig_msg)
        header = f"---------- Forwarded message ----------\nFrom: {orig_from}\nDate: {orig_date}\nSubject: {orig_subject}\n\n"
        full_body = body + "\n" + header + orig_body if body else header + orig_body
        
        msg = _build_message(from_addr, to_addrs, subject, full_body)
        _send_smtp_message(msg, config)
        
        return {
            "status": "sent",
            "from": from_addr,
            "to": to_addrs,
            "subject": subject,
        }

    def draft_message(self,
                      to: str,
                      subject: str,
                      body: str,
                      cc: Optional[str] = None,
                      mailbox: str = "Drafts") -> Dict[str, Any]:
        """
        Create a draft email in IMAP Drafts folder.
        
        Args:
            to: Recipient email address(es), comma-separated
            subject: Email subject
            body: Plain text body
            cc: CC recipients, comma-separated (optional)
            mailbox: Target mailbox (default: Drafts)
            
        Returns:
            Dict with status
        """
        self._require_conn()
        
        to_addrs = [addr.strip() for addr in to.split(",") if addr.strip()]
        cc_addrs = [addr.strip() for addr in cc.split(",") if addr.strip()] if cc else []
        
        if not to_addrs:
            raise RuntimeError("At least one recipient (to) is required")
        if not subject:
            raise RuntimeError("Subject is required")
        if not body:
            raise RuntimeError("Body is required")
        
        config = _get_smtp_config()
        from_addr = config["user"]
        
        msg = _build_message(from_addr, to_addrs, subject, body, cc_addrs)
        
        if HAVE_IMAP and imaplib is not None:
            internal_date = imaplib.Time2Internaldate(time.time())
        else:
            internal_date = time.strftime("%d-%b-%Y %H:%M:%S +0000", time.gmtime())
        
        typ, _ = self._conn.append(
            mailbox, "(\\Draft)", internal_date, msg.as_bytes()
        )
        if typ != "OK":
            raise RuntimeError("Draft append failed")
        
        return {
            "status": "ok",
            "mailbox": mailbox,
            "from": from_addr,
            "to": to_addrs,
            "cc": cc_addrs,
            "subject": subject,
        }

    def draft_reply(self,
                    mailbox: str,
                    uid: str,
                    body: str,
                    reply_all: bool = False,
                    draft_mailbox: str = "Drafts") -> Dict[str, Any]:
        """
        Create a reply draft in IMAP Drafts folder.
        
        Args:
            mailbox: Source mailbox
            uid: Message UID to reply to
            body: Reply body text
            reply_all: Whether to reply to all recipients
            draft_mailbox: Target mailbox for draft (default: Drafts)
            
        Returns:
            Dict with status
        """
        self._require_conn()
        if not body:
            raise RuntimeError("Reply body is required")
        
        # Fetch original message
        self._select(mailbox, readonly=True)
        typ, data = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")
        first = data[0]
        raw_msg = first[1] if isinstance(first, tuple) else first
        if not raw_msg:
            raise RuntimeError(f"UID {uid} not found")
        if isinstance(raw_msg, str):
            raw_msg = raw_msg.encode("utf-8", "replace")
        orig_msg = email.message_from_bytes(raw_msg)
        
        # Extract original headers
        orig_from = _decode_mime(orig_msg.get("From", ""))
        orig_to = _decode_mime(orig_msg.get("To", ""))
        orig_cc = _decode_mime(orig_msg.get("Cc", ""))
        orig_subject = _decode_mime(orig_msg.get("Subject", ""))
        orig_message_id = orig_msg.get("Message-ID", "")
        orig_references = orig_msg.get("References", "")
        
        config = _get_smtp_config()
        from_addr = config["user"]
        
        # Build recipient lists
        to_addrs = [_extract_email_address(orig_from)]
        
        cc_addrs = []
        if reply_all:
            all_recipients = _extract_email_addresses(orig_to) + _extract_email_addresses(orig_cc)
            self_addr = from_addr.lower()
            sender_addr = _extract_email_address(orig_from).lower()
            for addr in all_recipients:
                addr_lower = addr.lower()
                if addr_lower != self_addr and addr_lower != sender_addr:
                    cc_addrs.append(addr)
        
        # Build subject with Re: prefix
        if not orig_subject.upper().startswith("RE:"):
            subject = f"Re: {orig_subject}"
        else:
            subject = orig_subject
        
        # Build body with quoted original
        orig_body = _extract_plain_text(orig_msg)
        full_body = body + _quote_original(orig_body, orig_from)
        
        msg = _build_message(
            from_addr, to_addrs, subject, full_body, cc_addrs,
            in_reply_to=orig_message_id, references=orig_references
        )
        
        if HAVE_IMAP and imaplib is not None:
            internal_date = imaplib.Time2Internaldate(time.time())
        else:
            internal_date = time.strftime("%d-%b-%Y %H:%M:%S +0000", time.gmtime())
        
        typ, _ = self._conn.append(
            draft_mailbox, "(\\Draft)", internal_date, msg.as_bytes()
        )
        if typ != "OK":
            raise RuntimeError("Draft append failed")
        
        return {
            "status": "ok",
            "mailbox": draft_mailbox,
            "from": from_addr,
            "to": to_addrs,
            "cc": cc_addrs,
            "subject": subject,
            "in_reply_to": orig_message_id,
        }

    # ---------- uniform dispatch (LLM / MCP / CLI) ----------
    def dispatch_tool(self, name: str, args: Optional[Dict[str, Any]] = None) -> Any:
        args = dict(args or {})
        if name == "list_mailboxes":
            return self.list_mailboxes()
        if name == "search_messages":
            return self.search_messages(
                mailbox=args.get("mailbox", "INBOX"),
                criteria=args.get("criteria", "ALL"),
                limit=args.get("limit", 20),
            )
        if name == "fetch_message":
            return self.fetch_message(
                mailbox=args["mailbox"], uid=args["uid"],
                include_body=args.get("include_body", True),
                max_body_chars=args.get("max_body_chars", DEFAULT_MAX_BODY_CHARS),
            )
        if name == "mark_message":
            return self.mark_message(
                mailbox=args["mailbox"], uid=args["uid"],
                flag=args.get("flag", "\\Seen"), add=args.get("add", True),
            )
        if name == "move_message":
            return self.move_message(
                mailbox=args.get("mailbox", args.get("mailbox_src", "")),
                uid=args["uid"],
                dest=args.get("dest", args.get("mailbox_dest", "")),
                **{k: v for k, v in args.items()
                   if k in ("mailbox_src", "mailbox_dest", "source", "destination")},
            )
        if name == "delete_message":
            return self.delete_message(
                mailbox=args["mailbox"], uid=args["uid"],
                expunge=args.get("expunge", False),
            )
        if name == "append_draft":
            return self.append_draft(
                mailbox=args.get("mailbox", "Drafts"),
                raw_message=args["raw_message"],
            )
        if name == "list_attachments":
            return self.list_attachments(
                mailbox=args["mailbox"], uid=args["uid"],
                max_links=args.get("max_links", 3),
            )
        if name == "save_email_as_html":
            return self.save_email_as_html(
                mailbox=args["mailbox"], uid=args["uid"],
                include_attachments=args.get("include_attachments", True),
                backup_dir=args.get("backup_dir"),
            )
        if name == "send_message":
            return self.send_message(
                to=args["to"], subject=args["subject"],
                body=args["body"], cc=args.get("cc"),
            )
        if name == "reply_message":
            return self.reply_message(
                mailbox=args["mailbox"], uid=args["uid"],
                body=args["body"], reply_all=args.get("reply_all", False),
            )
        if name == "forward_message":
            return self.forward_message(
                mailbox=args["mailbox"], uid=args["uid"],
                to=args["to"], body=args.get("body", ""),
            )
        if name == "draft_message":
            return self.draft_message(
                to=args["to"], subject=args["subject"],
                body=args["body"], cc=args.get("cc"),
                mailbox=args.get("mailbox", "Drafts"),
            )
        if name == "draft_reply":
            return self.draft_reply(
                mailbox=args["mailbox"], uid=args["uid"],
                body=args["body"], reply_all=args.get("reply_all", False),
                draft_mailbox=args.get("draft_mailbox", "Drafts"),
            )
        raise RuntimeError(f"Unknown tool: {name}")


def canonical_tools_for_provider(provider: str) -> List[Dict[str, Any]]:
    """
    Single canonical JSON schema shared by every provider.
    Optional params carry defaults and are NOT in `required`.
    """
    return [
        {
            "name": "list_mailboxes",
            "description": "List all available mailboxes.",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
        {
            "name": "search_messages",
            "description": "Search messages with IMAP SEARCH syntax.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string", "default": "INBOX"},
                    "criteria": {"type": "string", "default": "ALL"},
                    "limit": {"type": "integer", "default": 20},
                },
                "required": [],
            },
        },
        {
            "name": "fetch_message",
            "description": "Retrieve full details (headers + plain body) by UID.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "include_body": {"type": "boolean", "default": True},
                    "max_body_chars": {"type": "integer", "default": DEFAULT_MAX_BODY_CHARS},
                },
                "required": ["mailbox", "uid"],
            },
        },
        {
            "name": "mark_message",
            "description": "Add or remove IMAP flag on a message.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "flag": {"type": "string", "default": "\\Seen"},
                    "add": {"type": "boolean", "default": True},
                },
                "required": ["mailbox", "uid"],
            },
        },
        {
            "name": "move_message",
            "description": "Move a message from one mailbox to another.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "dest": {"type": "string"},
                },
                "required": ["mailbox", "uid", "dest"],
            },
        },
        {
            "name": "delete_message",
            "description": "Flag message as deleted and optionally expunge.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "expunge": {"type": "boolean", "default": False},
                },
                "required": ["mailbox", "uid"],
            },
        },
        {
            "name": "append_draft",
            "description": "Create a new draft with raw RFC822 content.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string", "default": "Drafts"},
                    "raw_message": {"type": "string"},
                },
                "required": ["raw_message"],
            },
        },
        {
            "name": "list_attachments",
            "description": "List attachment names for a message (count + up to max_links names).",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "max_links": {"type": "integer", "default": 3},
                },
                "required": ["mailbox", "uid"],
            },
        },
        {
            "name": "save_email_as_html",
            "description": "Save an email as HTML file with attachments for offline backup.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "include_attachments": {"type": "boolean", "default": True},
                    "backup_dir": {"type": "string", "description": "Custom backup directory (optional)"},
                },
                "required": ["mailbox", "uid"],
            },
        },
        {
            "name": "send_message",
            "description": "Send a new email via SMTP.",
            "parameters": {
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                    "subject": {"type": "string"},
                    "body": {"type": "string", "description": "Plain text body"},
                    "cc": {"type": "string", "description": "CC recipients, comma-separated (optional)"},
                },
                "required": ["to", "subject", "body"],
            },
        },
        {
            "name": "reply_message",
            "description": "Reply to an email message.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "body": {"type": "string", "description": "Reply body text"},
                    "reply_all": {"type": "boolean", "default": False, "description": "Reply to all recipients"},
                },
                "required": ["mailbox", "uid", "body"],
            },
        },
        {
            "name": "forward_message",
            "description": "Forward an email message.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                    "body": {"type": "string", "description": "Additional text to prepend (optional)", "default": ""},
                },
                "required": ["mailbox", "uid", "to"],
            },
        },
        {
            "name": "draft_message",
            "description": "Create a draft email in IMAP Drafts folder.",
            "parameters": {
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                    "subject": {"type": "string"},
                    "body": {"type": "string", "description": "Plain text body"},
                    "cc": {"type": "string", "description": "CC recipients, comma-separated (optional)"},
                    "mailbox": {"type": "string", "default": "Drafts"},
                },
                "required": ["to", "subject", "body"],
            },
        },
        {
            "name": "draft_reply",
            "description": "Create a reply draft in IMAP Drafts folder.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "uid": {"type": "string"},
                    "body": {"type": "string", "description": "Reply body text"},
                    "reply_all": {"type": "boolean", "default": False},
                    "draft_mailbox": {"type": "string", "default": "Drafts"},
                },
                "required": ["mailbox", "uid", "body"],
            },
        },
    ]


def compile_tools(provider: str) -> List[Dict[str, Any]]:
    """
    Adapt the canonical schema to each provider's native tool format.
    openai/deepseek → {"type":"function","function":{...}}
    anthropic       → {"name","description","input_schema"}
    gemini          → {"name","description","parameters"}
    mcp             → {"name","description","inputSchema"}
    anything else   → canonical [{"name","description","parameters"}]
    """
    canonical = canonical_tools_for_provider(provider)
    p = (provider or "").strip().lower()
    if p in ("openai", "deepseek"):
        return [
            {"type": "function",
             "function": {"name": t["name"], "description": t["description"],
                          "parameters": t["parameters"]}}
            for t in canonical
        ]
    if p == "anthropic":
        return [
            {"name": t["name"], "description": t["description"],
             "input_schema": t["parameters"]}
            for t in canonical
        ]
    if p == "gemini":
        return [
            {"name": t["name"], "description": t["description"],
             "parameters": t["parameters"]}
            for t in canonical
        ]
    if p == "mcp":
        return [
            {"name": t["name"], "description": t["description"],
             "inputSchema": t["parameters"]}
            for t in canonical
        ]
    return canonical


# ---------- CLI ----------

_INT_ARGS = {"limit", "max_body_chars", "max_links"}
_BOOL_ARGS = {"add", "expunge", "include_body", "include_attachments", "reply_all"}
_NEEDS_CONNECTION = {
    "list_mailboxes", "search_messages", "fetch_message", "mark_message",
    "move_message", "delete_message", "append_draft", "list_attachments",
    "save_email_as_html", "send_message", "reply_message", "forward_message",
    "draft_message", "draft_reply", "init",
}

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
            raise RuntimeError(
                f"Bad argument {item!r}: expected key=value. See --help.")
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
    
    # Tools that don't need IMAP connection
    no_imap_tools = {"send_message"}
    
    if tool not in _NEEDS_CONNECTION:
        print(json.dumps({"error": f"Unknown tool: {tool}. See --help."}))
        return 1
    
    try:
        args = _parse_kv_args(rest)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return 1

    # Send message uses SMTP only, no IMAP connection needed
    if tool in no_imap_tools:
        # Create a minimal client just for dispatch
        client = ImapClient()
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

    # All other tools need IMAP connection
    client = ImapClient()
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
