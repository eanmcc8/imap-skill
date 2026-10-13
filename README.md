# imap-skill — Universal IMAP Email Skill

MCP server + CLI for programmatic IMAP/SMTP access. Works with **any MCP host** (Claude Desktop, Kilo Code, OpenWork, etc.) and exposes a **canonical tool schema** for OpenAI, Anthropic, Gemini, DeepSeek.

## Quick Start

```bash
# 1. Clone & install
cd imap-skill
pip install mcp  # for MCP server

# 2. Configure (interactive wizard)
python main.py init

# 3. Test CLI
python main.py list_mailboxes

# 4. Run MCP server (stdio)
python mcp_server.py
```

### Kilo Code Quick Setup

1. Copy `imap-skill` folder to your workspace (or anywhere)
2. Add MCP server config to Kilo Settings → MCP Servers (see below)
3. Set `IMAP_USER` / `IMAP_PASSWORD` in Kilo's environment variables
4. Use `@imap-skill` in chat

## Configuration (`.env`)

| Variable | Default | Description |
|----------|---------|-------------|
| `IMAP_HOST` | `imap.gmail.com` | IMAP server hostname |
| `IMAP_PORT` | `993` | IMAP port (993=SSL, 143=STARTTLS) |
| `IMAP_USE_SSL` | `true` | `true`, `false`, or `auto` |
| `IMAP_REQUIRE_ID` | `false` | Enable IMAP ID (RFC 2971) for 126/163 NetEase |
| `IMAP_USER` | — | Email address |
| `IMAP_PASSWORD` | — | App password (not login password) |
| `IMAP_TIMEOUT` | `30` | Connection timeout (seconds) |
| `MAX_BODY_CHARS` | `8000` | Body truncation limit |
| `LOG_LEVEL` | `INFO` | `DEBUG`, `INFO`, `WARNING`, `ERROR` |
| **Pool / Performance** | | |
| `IMAP_POOL_SIZE` | `2` | Connection pool size |
| `IMAP_IDLE_TIMEOUT` | `300` | Idle connection TTL (seconds) |
| **SMTP (Outbound)** | | |
| `SMTP_HOST` | `IMAP_HOST` | SMTP server hostname |
| `SMTP_PORT` | `465` | SMTP port (465=SSL, 587=STARTTLS) |
| `SMTP_USE_SSL` | `true` | `true` (SSL) or `false` (STARTTLS) |
| `SMTP_USER` | `IMAP_USER` | SMTP login |
| `SMTP_PASSWORD` | `IMAP_PASSWORD` | SMTP password |
| **Offline Backup** | | |
| `IMAP_BACKUP_DIR` | `~/imap_backup` | Backup output directory |

> **Gmail/Outlook**: Use an [App Password](https://support.google.com/accounts/answer/185833), not your login password.
> **126/163**: Set `IMAP_REQUIRE_ID=true` and use authorization code from settings.

## Tools (14 Total)

### IMAP Tools (read/inbox management)
| Tool | Description |
|------|-------------|
| `list_mailboxes` | List all folders (INBOX, Sent, Drafts, …) |
| `search_messages` | IMAP SEARCH with criteria (e.g., `UNSEEN`, `FROM "x"`, `SINCE 2024-01-01`) |
| `fetch_message` | Full message (headers + plain-text body) by UID |
| `mark_message` | Add/remove flags (`\Seen`, `\Flagged`, custom) |
| `move_message` | Move between mailboxes (uses `MOVE` or `COPY`+`STORE`) |
| `delete_message` | Mark `\Deleted`, optionally expunge |
| `append_draft` | Create draft from raw RFC822 |
| `list_attachments` | List attachment names (no download) |
| `save_email_as_html` | Save message as standalone HTML with embedded attachments for offline backup |

### SMTP Tools (outbound email)
| Tool | Description |
|------|-------------|
| `send_message` | Send new email via SMTP (to, subject, body, cc) |
| `reply_message` | Reply to email (fetches original via IMAP, sends via SMTP) |
| `forward_message` | Forward email to new recipients |
| `draft_message` | Create draft in IMAP Drafts folder |
| `draft_reply` | Create reply draft in IMAP Drafts folder |

## MCP Server

```bash
python mcp_server.py
```

Registers 14 tools over stdio. Connection pooling (`IMAP_POOL_SIZE=2`) enables efficient concurrent use by agents.

### Example MCP Config (Claude Desktop)

```json
{
  "mcpServers": {
    "imap-skill": {
      "command": "python",
      "args": ["path/to/imap-skill/mcp_server.py"],
      "env": {
        "IMAP_HOST": "imap.gmail.com",
        "IMAP_PORT": "993",
        "IMAP_USE_SSL": "true",
        "IMAP_USER": "you@gmail.com",
        "IMAP_PASSWORD": "your-app-password",
        "SMTP_HOST": "smtp.gmail.com",
        "SMTP_PORT": "587",
        "SMTP_USE_SSL": "false",
        "SMTP_USER": "you@gmail.com",
        "SMTP_PASSWORD": "your-app-password"
      }
    }
  }
}
```

### Kilo Code / Agent Manager

Add to **Settings → MCP Servers** or `kilo.json`:

```json
{
  "mcpServers": {
    "imap-skill": {
      "command": "python",
      "args": ["${workspaceFolder}/imap-skill/mcp_server.py"],
      "env": {
        "IMAP_HOST": "imap.gmail.com",
        "IMAP_PORT": "993",
        "IMAP_USE_SSL": "true",
        "IMAP_USER": "${env:IMAP_USER}",
        "IMAP_PASSWORD": "${env:IMAP_PASSWORD}",
        "SMTP_HOST": "smtp.gmail.com",
        "SMTP_PORT": "587",
        "SMTP_USE_SSL": "false",
        "SMTP_USER": "${env:IMAP_USER}",
        "SMTP_PASSWORD": "${env:IMAP_PASSWORD}"
      }
    }
  }
}
```

> **Tip**: Use `${env:VAR}` to reference environment variables from your shell or `.env` file. Set `IMAP_USER`/`IMAP_PASSWORD` in your system env or Kilo's env vars UI.

Then in chat:
```
@imap-skill search_messages criteria="UNSEEN" limit=5
@imap-skill send_message to="team@company.com" subject="Update" body="Done"
```

### VSCodium + Kilo Extension

**Config file location** (VSCodium):
- Linux: `~/.config/VSCodium/User/globalStorage/kilo.kilo/mcp.json`
- macOS: `~/Library/Application Support/VSCodium/User/globalStorage/kilo.kilo/mcp.json`
- Windows: `%APPDATA%\VSCodium\User\globalStorage\kilo.kilo\mcp.json`

Or use **Settings → Extensions → Kilo → MCP Servers** (JSON editor).

```json
{
  "mcpServers": {
    "imap-skill": {
      "command": "python",
      "args": ["${workspaceFolder}/imap-skill/mcp_server.py"],
      "env": {
        "IMAP_HOST": "imap.gmail.com",
        "IMAP_PORT": "993",
        "IMAP_USE_SSL": "true",
        "IMAP_USER": "${env:IMAP_USER}",
        "IMAP_PASSWORD": "${env:IMAP_PASSWORD}",
        "SMTP_HOST": "smtp.gmail.com",
        "SMTP_PORT": "587",
        "SMTP_USE_SSL": "false",
        "SMTP_USER": "${env:IMAP_USER}",
        "SMTP_PASSWORD": "${env:IMAP_PASSWORD}"
      }
    }
  }
}
```

**Environment variables** in VSCodium:
- **Option A**: Set in shell before launching: `IMAP_USER=you@gmail.com IMAP_PASSWORD=xxx codium`
- **Option B**: Add to `.env` file in workspace root (Kilo loads it)
- **Option C**: Settings → Extensions → Kilo → Environment Variables

> **Note**: `${workspaceFolder}` resolves to the open folder in VSCodium. If `imap-skill` is outside workspace, use absolute path: `"/home/user/imap-skill/mcp_server.py"` (Linux/macOS) or `"C:\\path\\imap-skill\\mcp_server.py"` (Windows).

## Provider Schemas

```bash
# Canonical (generic JSON)
python main.py compile_tools provider=generic

# OpenAI / DeepSeek function calling
python main.py compile_tools provider=openai

# Anthropic tool use
python main.py compile_tools provider=anthropic

# Gemini function calling
python main.py compile_tools provider=gemini

# MCP (inputSchema)
python main.py compile_tools provider=mcp
```

## CLI Usage

```bash
python main.py --help
python main.py init                          # Interactive configuration wizard
python main.py tools                         # List all tools (canonical)
python main.py list_mailboxes
python main.py search_messages mailbox=INBOX criteria="UNSEEN FROM alice" limit=10
python main.py fetch_message mailbox=INBOX uid=123 include_body=true max_body_chars=2000
python main.py mark_message mailbox=INBOX uid=123 flag="\\Flagged" add=true
python main.py move_message mailbox=INBOX uid=123 dest=Archive
python main.py delete_message mailbox=INBOX uid=123 expunge=true
python main.py append_draft mailbox=Drafts raw_message="Subject: Hi\n\nBody text"
python main.py list_attachments mailbox=INBOX uid=123 max_links=5
python main.py save_email_as_html mailbox=INBOX uid=123 include_attachments=true backup_dir=./backups

# Outbound (SMTP)
python main.py send_message to="a@b.com" subject="Hi" body="Hello" cc="c@d.com"
python main.py reply_message mailbox=INBOX uid=123 body="Thanks!" reply_all=false
python main.py forward_message mailbox=INBOX uid=123 to="a@b.com" body="FYI"
python main.py draft_message to="a@b.com" subject="Hi" body="Draft" mailbox=Drafts
python main.py draft_reply mailbox=INBOX uid=123 body="Reply draft" draft_mailbox=Drafts
```

## Presets (for `init` wizard)

| Preset | IMAP Host | IMAP Port | SMTP Host | SMTP Port | IMAP ID |
|--------|-----------|-----------|-----------|-----------|---------|
| `gmail` | imap.gmail.com | 993 (SSL) | smtp.gmail.com | 587 (STARTTLS) | No |
| `outlook` | outlook.office365.com | 993 (SSL) | smtp.office365.com | 587 (STARTTLS) | No |
| `qq` | imap.qq.com | 993 (SSL) | smtp.qq.com | 465 (SSL) | No |
| `163` | imap.163.com | 993 (SSL) | smtp.163.com | 465 (SSL) | **Yes** |
| `126` | imap.126.com | 993 (SSL) | smtp.126.com | 465 (SSL) | **Yes** |
| `tencent` | imap.exmail.qq.com | 993 (SSL) | smtp.exmail.qq.com | 465 (SSL) | No |
| `yahoo` | imap.mail.yahoo.com | 993 (SSL) | smtp.mail.yahoo.com | 465 (SSL) | No |
| `icloud` | imap.mail.me.com | 993 (SSL) | smtp.mail.me.com | 587 (STARTTLS) | No |

## Architecture

- **`main.py`** — Core IMAP/SMTP client, connection pool, 14 tools, CLI, provider adapters
- **`mcp_server.py`** — Thin MCP stdio wrapper (delegates to `ImapClient`)
- **Connection pooling** — Reuses IMAP connections, keepalive thread with NOOP
- **IMAP ID support** — RFC 2971 for 126/163 NetEase compatibility
- **Structured logging** — Correlation IDs per request
- **Offline backup** — HTML + attachments in `~/imap_backup/UID_subject.html`
- **Stdlib only** — No external dependencies for core (`imaplib`, `email`, `smtplib`)

## Requirements

- Python 3.9+
- `mcp` package (for MCP server only)
- Standard library only for core (`imaplib`, `email`, `smtplib`)

## License

MIT
