import imaplib
import email
import os
import pandas as pd
import json


# === CONFIGURATION ===
EMAIL = "ritesh.pal@ariantechsolutions.com"
PASSWORD = "sgiu btoz pwfs xepx"  # App Password use karo
IMAP_SERVER = "imap.gmail.com"
SAVE_DIR = "attachments"  # Attachments save hone ka folder
SPECIFIC_SENDER = "reeda.s@ariantechsolutions.com" # Yahan us user ka email dalein


# === Column Rename Mapping ===
RENAME_MAP = {
    "First Name": "fname",
    "Last Name": "lname",
    "Email": "email",
    "Mobile": "mobile",
    "Lead Source": "lead_source",
    "Primary Model Interest": "PMI",
    "Lead Owner": "lead_owner",
    "Enquiry Type": "enquiry_type",
    "Purchase Type": "purchase_type",
    "Retailer Name": "retailer_name"
}

# === Create folder if not exists ===
os.makedirs(SAVE_DIR, exist_ok=True)

# === CONNECT TO GMAIL ===
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)
mail.select("inbox")

# === GET UNREAD EMAILS ===
status, messages = mail.search(None, '(UNSEEN)')
email_ids = messages[0].split()

for num in email_ids:
    status, data = mail.fetch(num, '(RFC822)')
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    # === NEW: Check if the email is from the specific sender ===
    sender_email = email.utils.parseaddr(msg.get("From"))[1]
    if sender_email.lower() == SPECIFIC_SENDER.lower():
        print(f"Processing email from {sender_email}")
        
        for part in msg.walk():
            content_disposition = part.get("Content-Disposition", "")
            if "attachment" in content_disposition:
                filename = part.get_filename()
                if filename and (filename.endswith(".csv") or filename.endswith(".xls") or filename.endswith(".xlsx")):
                    filepath = os.path.join(SAVE_DIR, filename)

                    # === Save the attachment ===
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))

                    # === Read the file into DataFrame ===
                    try:
                        if filename.endswith(".csv"):
                            df = pd.read_csv(filepath)

                        elif filename.endswith(".xlsx"):
                            df = pd.read_excel(filepath, engine="openpyxl")

                        elif filename.endswith(".xls"):
                            try:
                                df = pd.read_excel(filepath, engine="xlrd")
                            except Exception as e:
                                with open(filepath, 'r', encoding="utf-8", errors="ignore") as f:
                                    first_512_bytes = f.read(512)
                                    if first_512_bytes.lstrip().startswith('<'):
                                        # Might be HTML - try read_html
                                        df_list = pd.read_html(filepath)
                                        df = df_list[0]  # take the first table found
                                    else:
                                        raise e

                        
                        df.columns = df.columns.str.strip()
                        df.rename(columns=RENAME_MAP, inplace=True)
                        df = df.applymap(lambda x: str(x).replace("\n", " ").strip() if pd.notnull(x) else x)

                        # === Replace NaN with None (so JSON will have null)
                        df = df.where(pd.notnull(df), None)

                        # === Convert to JSON ===
                        json_data = df.to_dict(orient="records")
                        json_path = os.path.splitext(filepath)[0] + ".json"
                        with open(json_path, "w", encoding="utf-8") as jf:
                            json.dump(json_data, jf, indent=2, ensure_ascii=False)


                        print(f"✅ JSON saved: {json_path}")

                    except Exception as e:
                        print(f"Error processing file {filename}: {e}")

    # === Mark this email as read regardless of sender ===
    mail.store(num, '+FLAGS', '\\Seen')

mail.logout()