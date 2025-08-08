
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import os
import pymongo
from datetime import datetime
import pandas as pd

# ========== Configuration ==========
EMAIL = "your_email@outlook.com"            #replace your email(outlook)
APP_PASSWORD = "you_generated_app_pass"               #replace with your app password
IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993
MONGO_URI = "mongodb://localhost:27017/"
DB_NAME = "outlook_email_logs"
COLLECTION_NAME = "emails"
ATTACHMENTS_DIR = "attachments"
FOLDERS = ["INBOX"]  # Can add more folders
MAX_EMAILS_PER_FOLDER = 200

# Create attachments folder if not exists
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# ========== MongoDB Setup ==========
client = pymongo.MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db[COLLECTION_NAME]

# ========== Connect to IMAP ==========
print("🔌 Connecting to Outlook server...")
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, APP_PASSWORD)

def clean_text(text):
    return ''.join(BeautifulSoup(text, "html.parser").stripped_strings)

def decode_mime_words(s):
    decoded_words = []
    for word, encoding in decode_header(s):
        if isinstance(word, bytes):
            decoded_words.append(word.decode(encoding or "utf-8", errors="ignore"))
        else:
            decoded_words.append(word)
    return ''.join(decoded_words)

def save_attachment(part, email_uid):
    filename = part.get_filename()
    if filename:
        filename = decode_mime_words(filename)
        filepath = os.path.join(ATTACHMENTS_DIR, f"{email_uid}_{filename}")
        with open(filepath, "wb") as f:
            f.write(part.get_payload(decode=True))
        return filepath
    return None

def parse_email(msg):
    try:
        subject = decode_mime_words(msg.get("Subject", "No Subject"))
        date_str = msg.get("Date")
        from_email = msg.get("From", "")
        to_email = msg.get("To", "")
        cc_email = msg.get("Cc", "")
        message_id = msg.get("Message-ID", "no-id").strip()
        in_reply_to = msg.get("In-Reply-To", None)

        # Parse date
        try:
            date_obj = email.utils.parsedate_to_datetime(date_str)
        except:
            date_obj = None

        body = ""
        attachments = []

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if part.get_content_disposition() == "attachment":
                    filepath = save_attachment(part, message_id.replace("<", "").replace(">", ""))
                    if filepath:
                        attachments.append(filepath)
                elif content_type == "text/plain":
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
            "to": to_email,
            "cc": cc_email,
            "subject": subject,
            "date": date_obj,
            "body": body,
            "attachments": attachments
        }

    except Exception as e:
        print(f"⚠️ Error parsing email: {e}")
        return None

all_emails_data = []

# ========== Loop through folders ==========
for folder in FOLDERS:
    print(f"\n📂 Processing folder: {folder}")
    try:
        mail.select(folder)
        result, data = mail.search(None, "ALL")
        if result != "OK":
            print(f"⚠️ Could not access {folder}")
            continue

        email_ids = data[0].split()
        email_ids = email_ids[-MAX_EMAILS_PER_FOLDER:]  # Limit

        for idx, eid in enumerate(reversed(email_ids), start=1):
            try:
                res, msg_data = mail.fetch(eid, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        email_data = parse_email(msg)

                        if not email_data:
                            continue

                        message_id = email_data["message_id"]

                        # Check if already exists in MongoDB
                        if not collection.find_one({"message_id": message_id}):
                            collection.insert_one(email_data)

                        # Track replies
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

                        all_emails_data.append(email_data)

                if idx % 10 == 0:
                    print(f"📨 Processed {idx}/{len(email_ids)} emails in {folder}...")

            except Exception as fetch_err:
                print(f"❌ Error fetching email ID {eid}: {fetch_err}")
                continue

    except Exception as folder_err:
        print(f"❌ Could not process folder '{folder}': {folder_err}")

mail.logout()

# ========== Save to Excel ==========
df = pd.DataFrame(all_emails_data)
df.to_excel("outlook_emails.xlsx", index=False)

print("\n✅ All recent emails synced with attachments, reply data, MongoDB, and saved to Excel.")
