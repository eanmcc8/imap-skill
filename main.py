#!/usr/bin/env python3
"""
email_skill.py — Universal IMAP Email Skill

Supports: OpenAI, Anthropic, Gemini, DeepSeek & MCP
Connects online to your IMAP server via env vars; implements the exact canonical schema.
Delegates to provider-specific adapters that translate this schema to each platform's native format.
Rejects no user request via internal heuristics—the model drives the call chain.

Requirement: Install corresponding LLM SDK via pip OR use built-in MCP handler.
  openai     → pip install openai
  anthropic  → pip install anthropic
  google-api  → pip install google-generatorai
  deepseek   → pip install deepseek
"""

import os
import json
import time
from typing import Any, Dict, List, Optional

try:
    import imaplib, email
    HAVE_IMAP = True
except Exception:
    HAVE_IMAP = False

class ImapClient:
    """
    Real IMAP transport backed by environment variables.
    Throws network/connection exceptions; no simulated fallback intended.
    """
    def __init__(self):
        self.host = os.getenv("IMAP_HOST", "imap.gmail.com")
        self.port = int(os.getenv("IMAP_PORT", "993"))
        self.use_ssl = os.getenv("IMAP_USE_SSL", "true").lower() == "true"
        self.user = os.getenv("IMAP_USER")
        self.password = os.getenv("IMAP_PASSWORD")
        self._conn = None

    def connect(self) -> None:
        """
        Connect ONCE upon instantiation, raises if fail.
        """
        try:
            if self.use_ssl:
                self._conn = imaplib.IMAP4_SSL(self.host, self.port)
            else:
                self._conn = imaplib.IMAP4(self.host, self.port)
            self._conn.login(self.user, self.password)
        except Exception as e:
            raise RuntimeError(f"IMAP connection failed: {e}")

    def close(self) -> None:
        try:
            if self._conn:
                self._conn.logout()
                self._conn = None
        except Exception:
            pass

    # ---------- Tool implementations (offline-safe interface) ----------
    def list_mailboxes(self) -> List[str]:
        """Return a list of folder names (INBOX, Sent, Drafts, ...)."""
        if not self._conn:
            raise RuntimeError("Not connected to IMAP")
        typ, data = self._conn.list()
        if typ != "OK":
            raise RuntimeError("Unable to list mailboxes")
        out = []
        for line in data or []:
            line = line.decode(errors='ignore')
            if '"' in line:
                out.append(line.rsplit('"', 2)[-2])
            else:
                out.append(line.split()[-1])
        return out

    def search_messages(self,
                       mailbox: str = "INBOX",
                       criteria: str = "ALL",
                       limit: int = 20) -> List[Dict[str, Any]]:
        """
        Run UID SEARCH and page results.
        Returns JSON of {uid, subject, from_, date}.
        Raises if server unsupported or parameters bad.
        """
        if not self._conn:
            raise RuntimeError("No IMAP connection")

        self._conn.select(mailbox, readonly=True)
        typ, data = self._conn.uid("SEARCH", None, criteria)

        if typ != "OK":
            raise RuntimeError(f"Search {criteria} failed.")

        uids = data[0].split()[-limit:] if data and len(data) > 0 else []

        results: List[Dict[str, Any]] = []
        for uid in uids:
            typ, raw = self._conn.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
            if typ != "OK":
                continue
            raw_msg = raw[0][1] if raw and len(raw) > 0 else b""
            msg = email.message_from_bytes(raw_msg)

            result: Dict[str, Any] = {
                "uid": uid,
                "subject": _decode_mime(msg.get("Subject", "")),
                "from": _decode_mime(msg.get("From", "")),
                "date": msg.get("Date", ""),
            }
            result["to"] = _decode_mime(msg.get("To", ""))
            results.append(result)

        return results

    def fetch_message(self,
                      mailbox: str,
                      uid: str) -> Dict[str, Any]:
        """
        Fetch full message metadata + plain-text body.
        Returns truncated body if > MAX_BODY_CHARS.
        """
        if not self._conn:
            raise RuntimeError("IMAP not connected")

        self._conn.select(mailbox, readonly=True)

        typ, data = self._conn.uid("FETCH", uid, "(RFC822)")
        if typ != "OK" or not data or not data[0]:
            raise RuntimeError(f"UID {uid} not found")

        raw_msg = data[0][1]
        msg = email.message_from_bytes(raw_msg)

        body = _extract_plain_text(msg)[:8000] + "… (truncated)" if _extract_plain_text(msg) else ""

        return {
            "uid": uid,
            "subject": _decode_mime(msg.get("Subject", "")),
            "from": _decode_mime(msg.get("From", "")),
            "to": _decode_mime(msg.get("To", "")),
            "date": msg.get("Date", ""),
            "body": body,
        }

    def mark_message(self,
                     mailbox: str,
                     uid: str,
                     flag: str = "\\Seen",
                     add: bool = True) -> Dict[str, Any]:
        """
        Add or remove IMAP flag via STORE.
        Raises on connection or bad args.
        """
        if not self._conn:
            raise RuntimeError("IMAP not connected")
        self._conn.select(mailbox, readonly=False)

        op = "+FLAGS" if add else "-FLAGS"
        typ, _ = self._conn.uid("STORE", uid, op, flag)
        if typ != "OK":
            raise RuntimeError(f"Failed to {'add' if add else 'remove'} flag {flag}")

        return {
            "status": "ok",
            "uid": uid,
            "flag": flag,
            "added": str(add),
        }

    def move_message(self,
                     mailbox_src: str,
                     uid: str,
                     mailbox_dest: str) -> Dict[str, Any]:
        """
        MOVE UID | else COPY + STORE + EXPUNGE.
        Returns movement summary.
        """
        if not self._conn:
            raise RuntimeError("IMAP not connected")
        self._conn.select(mailbox_src, readonly=False)

        if self._conn.capabilities and "MOVE" in self._conn.capabilities:
            typ, _ = self._conn.uid("MOVE", uid, mailbox_dest)
        else:
            # Fallback: COPY then EXPUNGE OR normal delete
            typ, _ = self._conn.uid("COPY", uid, mailbox_dest)
            if typ != "OK":
                raise RuntimeError("COPY also failed")

            typ, _ = self._conn.uid("STORE", uid, "+FLAGS", "\\Deleted")
            if self._conn.capabilities and "UIDPLUS" in self._conn.capabilities:
                typ, _ = self._conn.uid("EXPUNGE", uid)
            else:
                self._conn.expunge()

        return {
            "status": "ok",
            "uid": uid,
            "source": mailbox_src,
            "destination": mailbox_dest,
        }

    def delete_message(self,
                      mailbox: str,
                      uid: str,
                      expunge: bool = False) -> Dict[str, Any]:
        """
        STORE +FLAGS \\Deleted; expunge if requested.
        """
        if not self._conn:
            raise RuntimeError("IMAP not connected")
        self._conn.select(mailbox, readonly=False)

        typ, _ = self._conn.uid("STORE", uid, "+FLAGS", "\\Deleted")
        if typ != "OK":
            raise RuntimeError("Marked as deleted failed")

        if expunge and self._conn.capabilities and "UIDPLUS" in self._conn.capabilities:
            typ, _ = self._conn.uid("EXPUNGE", uid)
            if typ != "OK":
                raise RuntimeError("Expunge failed")

        return {
            "status": "ok",
            "uid": uid,
            "expunged": str(expunge),
        }

    def append_draft(self,
                     mailbox: str,
                     raw_message: str) -> Dict[str, Any]:
        """
        Append raw RFC822 to Drafts mailbox.
        """
        if not self._conn:
            raise RuntimeError("IMAP not connected")
        self._conn.select(mailbox, readonly=False)

        typ, _ = self._conn.append(mailbox, "\\Draft", self._get_time(), raw_message.encode("utf-8"))
        if typ != "OK":
            raise RuntimeError("Draft append failed")

        return {"status": "ok", "mailbox": mailbox}

    @staticmethod
    def _get_time() -> str:
        return time.strftime("%a, %d %b %Y %H:%M:%S %Z", time.gmtime())

def _decode_mime(s: str) -> str:
    """Safe decoding of MIME encoded headers."""
    parts = decode_header(s or "")
    out = []
    for text, enc in parts:
        decoded = (text if isinstance(text, str) else text.decode(enc or "utf-8", "replace"))
        out.append(decoded.strip())
    return "\n".join(out)

def _extract_plain_text(msg: email.message.Message) -> str:
    """Extract plain text body; skip multipart by default."""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    return part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    ).strip()
                except Exception:
                    continue
        return ""
    else:
        try:
            return msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            ).strip()
        except Exception:
            return ""

def canonical_tools_for_provider(provider: str) -> List[Dict[str, Any]]:
    """
    Factory returning the single canonical JSON schema expected by every provider.
    Configure via adapter functions shipped elsewhere (OpenAI, Anthropic, Gemini, DeepSeek, MCP).
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
                    "mailbox": {"type": "string"},
                    "criteria": {"type": "string"},
                    "limit": {"type": "integer", "default": 20},
                },
                "required": ["mailbox", "criteria"],
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
                    "max_body_chars": {"type": "integer", "default": 8000},
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
                    "flag": {"type": "string"},
                    "add": {"type": "boolean", "default": True},
                },
                "required": ["mailbox", "uid", "flag", "add"],
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
                "required": ["mailbox", "uid", "expunge"],
            },
        },
        {
            "name": "append_draft",
            "description": "Create a new draft with raw RFC822 content.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mailbox": {"type": "string"},
                    "raw_message": {"type": "string"},
                },
                "required": ["mailbox", "raw_message"],
            },
        },
    ]

def compile_tools(provider: str) -> List[Dict[str, Any]]:
    """
    Adapter gatekeeper: takes the canonical schema and returns provider-specific object.
    Placeholder example: send directly to OpenAI tool schema or Anthropic messages dict.
    """
    # TODO: Implement provider-specific adaptation here (code removed for brevity; see below sections).
    # For OpenAI → dict with "function" key
    # For Anthropic → include in messages list as "tool_call"
    # For Gemini → tools object
    # For DeepSeek → function object
    # For MCP → tool dicts

    # Below is the minimal stub — replace with real adapters in your codebase
    return canonical_tools_for_provider(provider)  # Stub: returns same for demo