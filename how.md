Optional logging (debug):

LOG_LEVEL=DEBUG

Examples
List all folders

python email_skill.py list_mailboxes

Search unread emails in INBOX

python email_skill.py search_messages mailbox=INBOX criteria="UNSEEN" limit=5

Fetch message by UID with full body

python email_skill.py fetch_message mailbox=INBOX uid=123456 include_body=true max_body_chars=1500

Mark an email as Seen

python email_skill.py mark_message mailbox=INBOX uid=123456 flag=\\Seen add=true

Move email to "Archive"

python email_skill.py move_message mailbox=INBOX uid=123456 dest=Archive

Delete & confirm

python email_skill.py delete_message mailbox=INBOX uid=123456 expunge=true

Create a quick draft in Drafts

python email_skill.py append_draft mailbox=Drafts raw_message="Subject: Meeting recap\n\nDetails below..."

Check attachment list for a message (experimental, shows count only)

python email_skill.py list_attachments mailbox=INBOX uid=123456 max_links=3

