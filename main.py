#!/usr/bin/env python3
"""
imap_skill.py — Single-file universal IMAP email management skill.

Works with OpenAI, Anthropic, Gemini, DeepSeek, and MCP through a single
canonical tool schema. Compile the schema to your provider's format using
the matching adapter, then drive the agent loop.

Usage:
    export IMAP_HOST=imap.gmail.com
    export IMAP_USER=you@example.com
    export IMAP_PASSWORD=app-password
    export OPENAI_API_KEY=sk-...          # or ANTHROPIC_API_KEY, etc.

    python imap_skill.py "Show my 5 most recent unread emails"

Requires: pip install openai   (or anthropic / google-generativeai)
"""

from __future__ import annotations

import copy
import email
import imaplib
import json
import os
import sys
from email.header import decode_header
from typing import Any


# ---------------------------------------------------------------------------
# 1. IMAP implementation
# ---------------------------------------------------------------------------

class ImapMail:
    """Thin, UID-based IMAP client. Never exposes raw commands to the model."""

    def __init__(
        self,
        host: str | None = None,
        user: str | None = None,
        password: str | None = None,
        port: int | None = None,
        use_ssl: bool | None = None,
    ) -> None:
        host = host or os.environ["IMAP_HOST"]
        user = user or os.environ["IMAP_USER"]
        password = password or os.environ["IMAP_PASSWORD"]
        port = port or int(os.environ.get("IMAP_PORT", "993"))
        use_ssl = use_ssl if use_ssl is not None else os.environ.get("IMAP_USE_SSL", "true").lower() == "true"

        self.m = imaplib.IMAP4_SSL(host, port) if use_ssl else imaplib.IMAP4(host, port)
        self.m.login(user, password)

    # -- helpers ------------------------------------------------------------

    def _select(self, mailbox: str, readonly: bool = True) -> None:
        typ, data = self.m.select(mailbox, readonly=readonly)
        if typ != "OK":
            raise RuntimeError(f"Mailbox not found: {mailbox}")

    def _has_capability(self, name: str) -> bool:
        caps = getattr(self.m, "capabilities", ())
        return name.upper() in (c.upper() for c in caps)

    # -- tools --------------------------------------------------------------

    def list_mailboxes(self) -> list[str]:
        typ, data = self.m.list()
        if typ != "OK":
            raise RuntimeError("Failed to list mailboxes")
        out = []
        for raw in data or []:
            if raw is None:
                continue
            line = raw.decode(errors="ignore")
            # Extract the quoted folder name at the end of the LIST response.
            if '"' in line:
                out.append(line.rsplit('"', 2)[-2])
            else:
                out.append(line.split()[-1])
        return out

    def search_messages(
        self, mailbox: str = "INBOX", criteria: str = "ALL", limit: int = 20
    ) -> list[dict[str, Any]]:
        self._select(mailbox, readonly=True)
        typ, data = self.m.uid("SEARCH", None, criteria)
        if typ != "OK":
            raise RuntimeError(f"Search failed: {criteria}")
        uids = (data[0] or b"").split()[-limit:]
        results = []
        for uid in uids:
            typ, msg_data = self.m.uid(
                "FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])"
            )
            if typ != "OK" or not msg_data or msg_data[0] is None:
                continue
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            results.append(
                {
                    "uid": uid.decode(),
                    "subject": _decode_mime(msg.get("Subject", "")),
                    "from": _decode_mime(msg.get("From", "")),
                    "date": msg.get("Date", ""),
                }
            )
        return results

    def fetch_message(self, mailbox: str, uid: str) -> dict[str, Any]:
        self._select(mailbox, readonly=True)
        typ, data = self.m.uid("FETCH", uid, "(RFC822)")
        if typ != "OK" or not data or data[0] is None:
            raise RuntimeError(f"UID not found: {uid}")
        msg = email.message_from_bytes(data[0][1])
        body = _extract_plain_text(msg)
        return {
            "uid": uid,
            "subject": _decode_mime(msg.get("Subject", "")),
            "from": _decode_mime(msg.get("From", "")),
            "to": _decode_mime(msg.get("To", "")),
            "date": msg.get("Date", ""),
            "body": body[:8000],
        }

    def mark_message(
        self, mailbox: str, uid: str, flag: str = "\\Seen", add: bool = True
    ) -> dict[str, str]:
        self._select(mailbox, readonly=False)
        op = "+FLAGS" if add else "-FLAGS"
        typ, _ = self.m.uid("STORE", uid, op, flag)
        if typ != "OK":
            raise RuntimeError(f"Failed to {'add' if add else 'remove'} flag {flag}")
        return {"status": "ok", "uid": uid, "flag": flag, "added": str(add)}

    def move_message(self, mailbox: str, uid: str, dest: str) -> dict[str, str]:
        self._select(mailbox, readonly=False)
        if self._has_capability("MOVE"):
            typ, _ = self.m.uid("MOVE", uid, dest)
            if typ != "OK":
                raise RuntimeError(f"Failed to move {uid} to {dest}")
        else:
            typ, _ = self.m.uid("COPY", uid, dest)
            if typ != "OK":
                raise RuntimeError(f"Failed to copy {uid} to {dest}")
            self.m.uid("STORE", uid, "+FLAGS", "\\Deleted")
            if self._has_capability("UIDPLUS"):
                self.m.uid("EXPUNGE", uid)
            else:
                self.m.expunge()
        return {"status": "ok", "uid": uid, "dest": dest}

    def delete_message(
        self, mailbox: str, uid: str, expunge: bool = False
    ) -> dict[str, str]:
        self._select(mailbox, readonly=False)
        self.m.uid("STORE", uid, "+FLAGS", "\\Deleted")
        if expunge:
            if self._has_capability("UIDPLUS"):
                self.m.uid("EXPUNGE", uid)
            else:
                self.m.expunge()
        return {"status": "ok", "uid": uid, "expunged": str(expunge)}

    def append_draft(self, mailbox: str, raw_message: str) -> dict[str, str]:
        typ, _ = self.m.append(
            mailbox,
            "\\Draft",
            imaplib.Time2Internaldate(__import__("time").time()),
            raw_message.encode("utf-8"),
        )
        if typ != "OK":
            raise RuntimeError(f"Failed to append draft to {mailbox}")
        return {"status": "ok", "mailbox": mailbox}

    def close(self) -> None:
        try:
            self.m.logout()
        except Exception:
            pass


def _decode_mime(value: str) -> str:
    parts = decode_header(value or "")
    out = []
    for text, enc in parts:
        if isinstance(text, bytes):
            out.append(text.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def _extract_plain_text(msg: email.message.Message) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    return part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                except Exception:
                    continue
        return ""
    try:
        return msg.get_payload(decode=True).decode(
            msg.get_content_charset() or "utf-8", errors="replace"
        )
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# 2. Canonical tool schema (vendor-agnostic)
# ---------------------------------------------------------------------------

CANONICAL_TOOLS: list[dict[str, Any]] = [
    {
        "name": "list_mailboxes",
        "description": "List all mailbox folders available on the IMAP server.",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "search_messages",
        "description": "Search messages in a mailbox using IMAP criteria. Returns UIDs, subjects, senders, and dates.",
        "parameters": {
            "type": "object",
            "properties": {
                "mailbox": {"type": "string", "description": "Mailbox folder, e.g. INBOX"},
                "criteria": {
                    "type": "string",
                    "description": 'IMAP search criteria, e.g. UNSEEN, FROM "boss@example.com", SINCE 01-Jan-2026',
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of results (default 20)",
                    "nullable": True,
                },
            },
            "required": ["mailbox", "criteria"],
        },
    },
    {
        "name": "fetch_message",
        "description": "Fetch a message by UID, returning headers and plain-text body.",
        "parameters": {
            "type": "object",
            "properties": {"mailbox": {"type": "string"}, "uid": {"type": "string"}},
            "required": ["mailbox", "uid"],
        },
    },
    {
        "name": "mark_message",
        "description": "Add or remove flags on a message (e.g. \\Seen, \\Flagged).",
        "parameters": {
            "type": "object",
            "properties": {
                "mailbox": {"type": "string"},
                "uid": {"type": "string"},
                "flag": {"type": "string", "description": "IMAP flag, e.g. \\Seen or \\Flagged"},
                "add": {"type": "boolean", "description": "true to add, false to remove"},
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
        "description": "Mark a message as deleted and optionally expunge it.",
        "parameters": {
            "type": "object",
            "properties": {
                "mailbox": {"type": "string"},
                "uid": {"type": "string"},
                "expunge": {"type": "boolean", "description": "true to permanently remove"},
            },
            "required": ["mailbox", "uid", "expunge"],
        },
    },
    {
        "name": "append_draft",
        "description": "Append a raw RFC822 message to the Drafts mailbox.",
        "parameters": {
            "type": "object",
            "properties": {
                "mailbox": {"type": "string", "description": "Usually Drafts"},
                "raw_message": {"type": "string", "description": "Full RFC822 message text"},
            },
            "required": ["mailbox", "raw_message"],
        },
    },
]


# ---------------------------------------------------------------------------
# 3. Universal provider adapters
# ---------------------------------------------------------------------------

def _promote_strict(schema: dict[str, Any]) -> None:
    """Promote optional fields to required+nullable for strict-mode providers."""
    if not isinstance(schema, dict) or schema.get("type") != "object":
        return
    props = schema.get("properties", {})
    required = set(schema.get("required", []))
    for name, prop in props.items():
        if name not in required:
            required.add(name)
            t = prop.get("type")
            if isinstance(t, str):
                prop["type"] = [t, "null"]
            elif isinstance(t, list) and "null" not in t:
                prop["type"] = t + ["null"]
            if "enum" in prop and None not in prop["enum"]:
                prop["enum"].append(None)
        if isinstance(prop, dict):
            if prop.get("type") == "object":
                _promote_strict(prop)
            if prop.get("type") == "array" and isinstance(prop.get("items"), dict):
                _promote_strict(prop["items"])
    schema["required"] = sorted(required)
    schema["additionalProperties"] = False


def _to_gemini_schema(schema: Any) -> Any:
    """Convert JSON Schema to Gemini's OpenAPI-compatible dialect."""
    if not isinstance(schema, dict):
        return schema
    out: dict[str, Any] = {}
    t = schema.get("type")
    if t is not None:
        if isinstance(t, list):
            if "null" in t:
                non_null = [x for x in t if x != "null"]
                out["type"] = (non_null[0] if non_null else "string").upper()
                out["nullable"] = True
            else:
                out["type"] = t[0].upper()
        else:
            out["type"] = t.upper()
    if "description" in schema:
        out["description"] = schema["description"]
    if "enum" in schema:
        out["enum"] = [v for v in schema["enum"] if v is not None]
    if "properties" in schema:
        out["properties"] = {k: _to_gemini_schema(v) for k, v in schema["properties"].items()}
    if "required" in schema:
        out["required"] = schema["required"]
    if "items" in schema:
        out["items"] = _to_gemini_schema(schema["items"])
    return out


def to_openai(tool: dict[str, Any]) -> dict[str, Any]:
    params = copy.deepcopy(tool["parameters"])
    _promote_strict(params)
    return {
        "type": "function",
        "function": {
            "name": tool["name"],
            "description": tool["description"],
            "parameters": params,
            "strict": True,
        },
    }


def to_anthropic(tool: dict[str, Any]) -> dict[str, Any]:
    params = copy.deepcopy(tool["parameters"])
    _promote_strict(params)
    return {
        "name": tool["name"],
        "description": tool["description"],
        "input_schema": params,
    }


def to_gemini(tool: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": tool["name"],
        "description": tool["description"],
        "parameters": _to_gemini_schema(copy.deepcopy(tool["parameters"])),
    }


def to_deepseek(tool: dict[str, Any]) -> dict[str, Any]:
    return to_openai(tool)


def to_mcp(tool: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": tool["name"],
        "description": tool["description"],
        "inputSchema": copy.deepcopy(tool["parameters"]),
    }


ADAPTERS = {
    "openai": to_openai,
    "anthropic": to_anthropic,
    "gemini": to_gemini,
    "deepseek": to_deepseek,
    "mcp": to_mcp,
}


def compile_tools(provider: str) -> list[dict[str, Any]]:
    adapter = ADAPTERS.get(provider)
    if adapter is None:
        raise ValueError(f"Unknown provider: {provider}")
    if provider == "anthropic":
        # Anthropic takes tools as a flat list; each entry is already top-level.
        return [adapter(t) for t in CANONICAL_TOOLS]
    if provider == "gemini":
        # Gemini wraps declarations in a single tools object at request time.
        return [adapter(t) for t in CANONICAL_TOOLS]
    return [adapter(t) for t in CANONICAL_TOOLS]


# ---------------------------------------------------------------------------
# 4. Tool dispatcher
# ---------------------------------------------------------------------------

def execute_tool(mail: ImapMail, name: str, args: dict[str, Any]) -> str:
    """Route a model tool call to the ImapMail implementation."""
    try:
        if name == "list_mailboxes":
            result = mail.list_mailboxes()
        elif name == "search_messages":
            result = mail.search_messages(
                mailbox=args.get("mailbox", "INBOX"),
                criteria=args.get("criteria", "ALL"),
                limit=int(args.get("limit") or 20),
            )
        elif name == "fetch_message":
            result = mail.fetch_message(args["mailbox"], args["uid"])
        elif name == "mark_message":
            result = mail.mark_message(
                args["mailbox"], args["uid"], args.get("flag", "\\Seen"), bool(args.get("add", True))
            )
        elif name == "move_message":
            result = mail.move_message(args["mailbox"], args["uid"], args["dest"])
        elif name == "delete_message":
            result = mail.delete_message(
                args["mailbox"], args["uid"], bool(args.get("expunge", False))
            )
        elif name == "append_draft":
            result = mail.append_draft(args["mailbox"], args["raw_message"])
        else:
            return json.dumps({"error": f"Unknown tool: {name}"})
        return json.dumps(result, default=str)
    except Exception as e:  # noqa: BLE001 — surface all errors to the model
        return json.dumps({"error": f"{type(e).__name__}: {e}"})


# ---------------------------------------------------------------------------
# 5. Provider-specific agent loops
# ---------------------------------------------------------------------------

def run_openai_compatible(
    client: Any,
    model: str,
    mail: ImapMail,
    user_message: str,
    max_turns: int = 5,
) -> str:
    """Agent loop for OpenAI and DeepSeek (both use the same SDK shape)."""
    tools = compile_tools("openai")
    messages: list[dict[str, Any]] = [
        {"role": "system", "content": "You manage email over IMAP. Use UIDs. Confirm destructive actions before executing."},
        {"role": "user", "content": user_message},
    ]

    for _ in range(max_turns):
        response = client.chat.completions.create(
            model=model, messages=messages, tools=tools, tool_choice="auto"
        )
        msg = response.choices[0].message
        messages.append(msg)

        if not msg.tool_calls:
            return msg.content or ""

        for tc in msg.tool_calls:
            args = json.loads(tc.function.arguments or "{}")
            result = execute_tool(mail, tc.function.name, args)
            messages.append(
                {"role": "tool", "tool_call_id": tc.id, "content": result}
            )

    return "Reached max tool turns."


def run_anthropic(
    client: Any, model: str, mail: ImapMail, user_message: str, max_turns: int = 5
) -> str:
    tools = compile_tools("anthropic")
    messages: list[dict[str, Any]] = [{"role": "user", "content": user_message}]

    for _ in range(max_turns):
        response = client.messages.create(
            model=model,
            max_tokens=1024,
            system="You manage email over IMAP. Use UIDs. Confirm destructive actions before executing.",
            tools=tools,
            messages=messages,
        )
        messages.append({"role": "assistant", "content": response.content})

        if response.stop_reason != "tool_use":
            text = "".join(b.text for b in response.content if getattr(b, "type", "") == "text")
            return text

        tool_results = []
        for block in response.content:
            if getattr(block, "type", "") == "tool_use":
                result = execute_tool(mail, block.name, dict(block.input or {}))
                tool_results.append(
                    {"type": "tool_result", "tool_use_id": block.id, "content": result}
                )
        messages.append({"role": "user", "content": tool_results})

    return "Reached max tool turns."


def run_gemini(
    client: Any, model: str, mail: ImapMail, user_message: str, max_turns: int = 5
) -> str:
    from google.generativeai import protos  # type: ignore

    declarations = compile_tools("gemini")
    gemini_tools = [protos.Tool(function_declarations=[protos.FunctionDeclaration(**d) for d in declarations])]
    chat = client.GenerativeModel(model_name=model, tools=gemini_tools).start_chat()

    response = chat.send_message(user_message)
    for _ in range(max_turns):
        calls = getattr(response, "function_calls", None) or []
        if not calls:
            return response.text or ""

        parts = []
        for call in calls:
            args = dict(call.args or {})
            result = execute_tool(mail, call.name, args)
            parts.append(
                protos.Part(
                    function_response=protos.FunctionResponse(
                        name=call.name, response={"result": json.loads(result)}
                    )
                )
            )
        response = chat.send_message(parts)

    return "Reached max tool turns."


# ---------------------------------------------------------------------------
# 6. CLI entry point
# ---------------------------------------------------------------------------

def _detect_provider() -> tuple[str, str, Any]:
    """Return (provider, model, client) based on which API keys are set."""
    if os.environ.get("ANTHROPIC_API_KEY"):
        import anthropic  # type: ignore
        return "anthropic", os.environ.get("MODEL", "claude-sonnet-4-5"), anthropic.Anthropic()
    if os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"):
        import google.generativeai as genai  # type: ignore
        genai.configure(api_key=os.environ.get("GEMINI_API_KEY") or os.environ["GOOGLE_API_KEY"])
        return "gemini", os.environ.get("MODEL", "gemini-2.0-flash"), genai
    if os.environ.get("DEEPSEEK_API_KEY"):
        from openai import OpenAI  # type: ignore
        client = OpenAI(
            api_key=os.environ["DEEPSEEK_API_KEY"],
            base_url=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com/beta"),
        )
        return "deepseek", os.environ.get("MODEL", "deepseek-chat"), client
    if os.environ.get("OPENAI_API_KEY"):
        from openai import OpenAI  # type: ignore
        return "openai", os.environ.get("MODEL", "gpt-4o"), OpenAI()
    raise SystemExit("No LLM API key found. Set OPENAI_API_KEY, ANTHROPIC_API_KEY, GEMINI_API_KEY, or DEEPSEEK_API_KEY.")


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python imap_skill.py \"<prompt>\"", file=sys.stderr)
        sys.exit(2)

    user_message = " ".join(sys.argv[1:])
    provider, model, client = _detect_provider()
    mail = ImapMail()

    try:
        if provider in ("openai", "deepseek"):
            answer = run_openai_compatible(client, model, mail, user_message)
        elif provider == "anthropic":
            answer = run_anthropic(client, model, mail, user_message)
        elif provider == "gemini":
            answer = run_gemini(client, model, mail, user_message)
        else:
            raise SystemExit(f"Unsupported provider: {provider}")
        print(answer)
    finally:
        mail.close()


if __name__ == "__main__":
    main()