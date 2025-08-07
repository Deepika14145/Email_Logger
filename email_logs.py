import imaplib
import email
from email.header import decode_header
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from email.utils import parsedate_to_datetime

# --- Step 1: Gmail Login Details ---
imap_server = "imap.gmail.com"
email_user = "youremail@gmail.com"
email_pass = "your_app_pass"

# --- Step 2: Connect to Gmail ---
mail = imaplib.IMAP4_SSL(imap_server)
mail.login(email_user, email_pass)
mail.select("inbox")

# --- Step 3: Fetch Emails (last 200) ---
status, messages = mail.search(None, 'ALL')
email_ids = messages[0].split()
email_data = []
reply_map = {}

for e_id in email_ids[-400:]:  # Get last 200 emails
    _, msg_data = mail.fetch(e_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Decode Subject
            subject_raw = msg.get("Subject", "")
            subject, encoding = decode_header(subject_raw)[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8", errors="ignore")

            from_ = msg.get("From", "")
            date_ = msg.get("Date", "")
            msg_id = msg.get("Message-ID", "").strip()
            in_reply_to = msg.get("In-Reply-To", "").strip()

            if "<" in from_ and ">" in from_:
                name = from_.split("<")[0].strip('" ')
                email_id = from_.split("<")[1].strip(">")
            else:
                name = from_ or ""
                email_id = from_ or ""

            # Extract and Clean Email Body
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() in ["text/plain", "text/html"]:
                        try:
                            raw_body = part.get_payload(decode=True).decode(errors="ignore")
                            soup = BeautifulSoup(raw_body, "html.parser")
                            body = soup.get_text()
                            break
                        except:
                            body = ""
            else:
                try:
                    raw_body = msg.get_payload(decode=True).decode(errors="ignore")
                    soup = BeautifulSoup(raw_body, "html.parser")
                    body = soup.get_text()
                except:
                    body = ""

            # Parse datetime
            try:
                parsed_date = parsedate_to_datetime(date_)
                if parsed_date.tzinfo is not None:
                    parsed_date = parsed_date.replace(tzinfo=None)
            except:
                parsed_date = datetime.min

            # Store email data
            email_record = {
                "Message-ID": msg_id,
                "Email ID": email_id,
                "Name": name,
                "Subject": subject,
                "Timestamp": date_,
                "ParsedDate": parsed_date,
                "Content": body.strip(),
                "Reply Count": 0,
                "Replies": []
            }

            email_data.append(email_record)
            reply_map[msg_id] = email_record

            if in_reply_to and in_reply_to in reply_map:
                original = reply_map[in_reply_to]
                original["Reply Count"] += 1
                original["Replies"].append({
                    "Reply Timestamp": parsed_date,
                    "Reply Body": body.strip()
                })

mail.logout()

# --- Step 4: Sort by ParsedDate (Most Recent First) ---
email_data.sort(key=lambda x: x["ParsedDate"], reverse=True)

# --- Step 5: Flatten Data for Excel ---
flat_data = []
for record in email_data:
    base = {
        "Email ID": record["Email ID"],
        "Name": record["Name"],
        "Subject": record["Subject"],
        "Original Timestamp": record["Timestamp"],
        "Original Body": record["Content"],
        "Replied Times": record["Reply Count"]
    }

    if record["Reply Count"] > 0:
        for i, reply in enumerate(record["Replies"], start=1):
            flat = base.copy()
            flat[f"Reply {i} Timestamp"] = reply["Reply Timestamp"]
            flat[f"Reply {i} Body"] = reply["Reply Body"]
            flat_data.append(flat)
    else:
        flat_data.append(base)

# --- Step 6: Save to Excel ---
df = pd.DataFrame(flat_data)
df.to_excel("Email_Replies_Log.xlsx", index=False)

print("Emails and replies saved to 'Email_Log.xlsx'")
