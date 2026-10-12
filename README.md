# Universal IMAP Email Management — Deployment Guide

## Mission
Expose ONE consistent tool set so an LLM powered by **OpenAI, Anthropic, Gemini, DeepSeek, or MCP** can perform any email operation without patching code.

Requirements
-------------
- Python 3.9+
- One of:
  - openai (pip install openai)
  - anthropic (pip install anthropic)
  - google-generatorai (pip install google-generatorai)
  - deepseek (pip install deepseek)
- Valid IMAP credentials stored securely via .env or secret management; **NEVER commit raw passwords.**

Directory layout suggestion

project-root/ ├── email_skill.py # code below ├── .env # your credentials └── HOWTO.md # this file


Usage with Real LLM (example OpenAI)
```python
from openai import OpenAI
from email_skill import ImapClient, compile_tools

# instantiate production client (reads .env automatically)
client = ImapClient()

# compile tools per provider; here OpenAI as example
tools = compile_tools(provider="openai")

# OpenAI Chat Completions call snippet
import asyncio, websockets
async def llm_handle(session, user_msg):
    # injected helper - replace with your agent loop
    model = "gpt-4o"
    resp = await session.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": "You manage IMAP emails. Use tools with full arguments. Confirm destructive actions."},
                 {"role": "user", "content": user_msg},
        ],
        tools=tools,
        tool_choice="auto",
        max_tokens=1024,
        temperature=0.1,
    )
    return resp.choices[0].message.content

# In your actual server/client loop...

Use the same pattern for other providers:

Anthropic: wrap in anthropic.AnthropicChatMessages.create(...)
Gemini: feed tools into genai.models.generate_text(...)
MCP: return tool responses from function handlers and let MCP dispatch
Tool Names & Parameters (exact schema)
list_mailboxes               → returns list of folders
search_messages              mailbox=..., criteria=..., limit=...
                                e.g. criteria="FROM \"boss\" SINCE \"01-Jan-2026\""
fetch_message               mailbox=..., uid=...
                                includes headers + plain text body
mark_message                mailbox=..., uid=..., flag=... add=true/false
move_message                mailbox=..., uid=..., dest=
delete_message              mailbox=..., uid=..., expunge=true/false
append_draft                mailbox=..., raw_message=""

No additional fields are introduced; models treat these exactly as described.

Online Behaviour Guarantees
TLS/SSL enforced (IMAP_USE_SSL=true);
Handles real network timeouts/retry;
Returns standardized JSON even on failure ({"error": "<reason>");
Does not rely on local mock data — fetches live mailbox state;
Prevents injection: headers decoded using email.header.decode_header;
Body truncation capped by MAX_BODY_CHARS.
Configuration Tips
OAuth2 tokens preferred: IMAP_TOKEN=Bearer eyJ... hides secrets better.
Rate Limiting: Provider SDK throttling applies; buffer longer responses upstream.
Architecture: Deploy this script behind an API endpoint or MCP server rather than letting models hit raw sockets directly.
HTTPS only: Always set IMAP_USE_SSL=true; avoid STARTTLS unless legacy server forces it.
Region: Some IMAP hosts require region-specific ports/settings — override accordingly.
Sanity Check Script (standalone)
python email_skill.py list_mailboxes | jq .
python email_skill.py search_messages mailbox=INBOX criteria="UNSEEN" limit=5 | jq .