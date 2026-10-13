"""
imap-skill MCP server (stdio) — wires the IMAP tools into any MCP host
(OpenWork custom/local MCP, Kilo Code, Claude Desktop, …).

Run by the host, NOT by hand:
  python mcp_server.py          # speaks MCP over stdio

Requires (install yourself):  python -m pip install mcp
Credentials: read from the workspace .env by main.py (IMAP_USER, IMAP_PASSWORD…).

Tool calls reuse connections from a shared pool (IMAP_POOL_SIZE, default 2)
for efficient concurrent agent use.
"""

import os
import sys

try:
    from main import DEFAULT_MAX_BODY_CHARS, ImapClient
except ImportError:  # allow launching from a different working directory
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from main import DEFAULT_MAX_BODY_CHARS, ImapClient

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    raise SystemExit(
        "The 'mcp' package is not installed. Run: python -m pip install mcp"
    ) from e

mcp = FastMCP("imap-skill")


@mcp.tool()
def list_mailboxes() -> list:
    """List all available mailboxes (INBOX, Sent, Drafts, …)."""
    with ImapClient.connect_once() as c:
        return c.list_mailboxes()


@mcp.tool()
def search_messages(mailbox: str = "INBOX", criteria: str = "ALL",
                    limit: int = 20) -> list:
    """Search messages with IMAP SEARCH syntax. Returns [{uid, subject, from, to, date}]."""
    with ImapClient.connect_once() as c:
        return c.search_messages(mailbox=mailbox, criteria=criteria, limit=limit)


@mcp.tool()
def fetch_message(mailbox: str, uid: str, include_body: bool = True,
                  max_body_chars: int = DEFAULT_MAX_BODY_CHARS) -> dict:
    """Retrieve full details (headers + plain-text body) for one message by UID."""
    with ImapClient.connect_once() as c:
        return c.fetch_message(mailbox=mailbox, uid=uid,
                               include_body=include_body,
                               max_body_chars=max_body_chars)


@mcp.tool()
def mark_message(mailbox: str, uid: str, flag: str = "\\Seen",
                 add: bool = True) -> dict:
    """Add or remove an IMAP flag (e.g. \\Seen) on a message."""
    with ImapClient.connect_once() as c:
        return c.mark_message(mailbox=mailbox, uid=uid, flag=flag, add=add)


@mcp.tool()
def move_message(mailbox: str, uid: str, dest: str) -> dict:
    """Move a message from one mailbox to another."""
    with ImapClient.connect_once() as c:
        return c.move_message(mailbox=mailbox, uid=uid, dest=dest)


@mcp.tool()
def delete_message(mailbox: str, uid: str, expunge: bool = False) -> dict:
    """Flag a message as deleted, optionally expunging immediately."""
    with ImapClient.connect_once() as c:
        return c.delete_message(mailbox=mailbox, uid=uid, expunge=expunge)


@mcp.tool()
def append_draft(raw_message: str, mailbox: str = "Drafts") -> dict:
    """Create a draft from raw RFC822 content (Subject/To headers + body)."""
    with ImapClient.connect_once() as c:
        return c.append_draft(mailbox=mailbox, raw_message=raw_message)


@mcp.tool()
def list_attachments(mailbox: str, uid: str, max_links: int = 3) -> dict:
    """List attachment names for a message (count + up to max_links names)."""
    with ImapClient.connect_once() as c:
        return c.list_attachments(mailbox=mailbox, uid=uid, max_links=max_links)


@mcp.tool()
def save_email_as_html(mailbox: str, uid: str, include_attachments: bool = True,
                       backup_dir: str = "") -> dict:
    """Save an email as HTML file with attachments for offline backup."""
    with ImapClient.connect_once() as c:
        return c.save_email_as_html(mailbox=mailbox, uid=uid,
                                    include_attachments=include_attachments,
                                    backup_dir=backup_dir if backup_dir else None)


# --- SMTP-based tools (outbound email) ---

@mcp.tool()
def send_message(to: str, subject: str, body: str, cc: str = "") -> dict:
    """Send a new email via SMTP. to/cc: comma-separated addresses."""
    # send_message doesn't need IMAP connection, uses SMTP directly
    client = ImapClient()
    return client.send_message(to=to, subject=subject, body=body, cc=cc if cc else None)


@mcp.tool()
def reply_message(mailbox: str, uid: str, body: str, reply_all: bool = False) -> dict:
    """Reply to an email message. Uses SMTP to send."""
    with ImapClient.connect_once() as c:
        return c.reply_message(mailbox=mailbox, uid=uid, body=body, reply_all=reply_all)


@mcp.tool()
def forward_message(mailbox: str, uid: str, to: str, body: str = "") -> dict:
    """Forward an email message. Uses SMTP to send."""
    with ImapClient.connect_once() as c:
        return c.forward_message(mailbox=mailbox, uid=uid, to=to, body=body)


@mcp.tool()
def draft_message(to: str, subject: str, body: str, cc: str = "", mailbox: str = "Drafts") -> dict:
    """Create a draft email in IMAP Drafts folder."""
    with ImapClient.connect_once() as c:
        return c.draft_message(to=to, subject=subject, body=body, cc=cc if cc else None, mailbox=mailbox)


@mcp.tool()
def draft_reply(mailbox: str, uid: str, body: str, reply_all: bool = False, draft_mailbox: str = "Drafts") -> dict:
    """Create a reply draft in IMAP Drafts folder."""
    with ImapClient.connect_once() as c:
        return c.draft_reply(mailbox=mailbox, uid=uid, body=body, reply_all=reply_all, draft_mailbox=draft_mailbox)


if __name__ == "__main__":
    mcp.run()
