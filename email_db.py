import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
from datetime import datetime
import pymongo

# ========== Configuration ==========
EMAIL = "your_email@gmail.com"
APP_PASSWORD = "app_pass"
IMAP_SERVER = "imap.gmail.com"
MONGO_URI = "mongodb://localhost:27017/"
DB_NAME = "email_logs"
COLLECTION_NAME = "emails"
FOLDERS = ['INBOX']  # You can add others like '[Gmail]/Sent Mail'
MAX_EMAILS_PER_FOLDER = 300

# ========== MongoDB Setup ==========
client = pymongo.MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db[COLLECTION_NAME]

# ========== Connect to IMAP ==========
print("🔌 Connecting to email server...")
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, APP_PASSWORD)

def clean_text(text):
    return ''.join(BeautifulSoup(text, "html.parser").stripped_strings)

def decode_subject(subject):
    if not subject:
        return "No Subject"
    decoded, charset = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(charset or "utf-8", errors="ignore")
    return decoded

def parse_email(msg):
    try:
        subject = decode_subject(msg.get("Subject"))
        date = msg.get("Date")
        from_email = msg.get("From", "")
        message_id = msg.get("Message-ID", "no-id").strip()
        in_reply_to = msg.get("In-Reply-To", None)
        body = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain" and part.get_payload(decode=True):
                    body += part.get_payload(decode=True).decode(errors="ignore")
                elif content_type == "text/html" and not body:
                    html = part.get_payload(decode=True).decode(errors="ignore")
                    body = clean_text(html)
        else:
            body = clean_text(msg.get_payload(decode=True).decode(errors="ignore"))

        return {
            "message_id": message_id,
            "in_reply_to": in_reply_to.strip() if in_reply_to else None,
            "from": from_email,
            "subject": subject,
            "date": date,
            "body": body
        }

    except Exception as e:
        print(f"⚠️ Error parsing email: {e}")
        return None

# ========== Loop through folders ==========
for folder in FOLDERS:
    print(f"\n📂 Processing folder: {folder}")
    try:
        mail.select(folder)
        result, data = mail.search(None, 'ALL')
        if result != 'OK':
            print(f"⚠️ Could not access {folder}")
            continue

        email_ids = data[0].split()
        email_ids = email_ids[-MAX_EMAILS_PER_FOLDER:]  # Limit

        for idx, eid in enumerate(reversed(email_ids), start=1):  # reversed for recent emails first
            try:
                res, msg_data = mail.fetch(eid, '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        email_data = parse_email(msg)

                        if not email_data:
                            continue

                        message_id = email_data["message_id"]

                        if not collection.find_one({"message_id": message_id}):
                            # Insert new email
                            collection.insert_one(email_data)

                        # Check if it's a reply
                        if email_data["in_reply_to"]:
                            parent = collection.find_one({"message_id": email_data["in_reply_to"]})
                            if parent:
                                collection.update_one(
                                    {"message_id": parent["message_id"]},
                                    {
                                        "$inc": {"replied_count": 1},
                                        "$push": {
                                            "replies": {
                                                "from": email_data["from"],
                                                "subject": email_data["subject"],
                                                "date": email_data["date"],
                                                "body": email_data["body"]
                                            }
                                        }
                                    }
                                )

                if idx % 10 == 0:
                    print(f"📨 Processed {idx}/{len(email_ids)} emails in {folder}...")

            except Exception as fetch_err:
                print(f"❌ Error fetching email ID {eid}: {fetch_err}")
                continue

    except Exception as folder_err:
        print(f"❌ Could not process folder '{folder}': {folder_err}")

mail.logout()
print("\n✅ All recent emails synced with reply data in MongoDB.")
