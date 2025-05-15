import openpyxl
import requests
import os
import imaplib
import email
from email.header import decode_header
from datetime import datetime, time, timedelta
import schedule
import time as time_module

def log_message(message):
    timestamp = datetime.now().strftime("[%A %Y-%m-%d %H:%M:%S]") 
    with open("log.txt", "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} {message}\n")
    print(f"{timestamp} {message}")  

def get_Token(url, username, password):
    payload = {
        "UserName": username,
        "Password": password
    }
    try:
        response = requests.post(url, json=payload)
        data = response.json()
        if len(data['token']):
            log_message("✅ Token fetched successfully.")
        return data['token']
    except requests.exceptions.RequestException as e:
        log_message(f"⚠️ Request error: {e}")
        return None

def post_downloaded_file(file_path, party_id, sheet_name, token):
    url = "https://sunrisediam.com:8223/api/party/create_update_manual_upload"
    headers = {
        'Authorization': f'Bearer {token}'
    }
    with open(file_path, 'rb') as f:
        files = {
            'File_Location': (os.path.basename(file_path), f, 'application/octet-stream')
        }
        data = {
            'File_Id': 0,
            'Party_Id': party_id,
            'Sheet_Name': sheet_name,
            'Validity_Days': 3,
            'API_Flag': 'false',
            'Exclude': 'false',
            'Is_Overwrite': 'true',
            'Priority': 'false',
            'Upload_Type': 'T',
            'Is_Same_Id': 'false',
            'Overseas_Same_Id': 'false'
        }

        response = requests.post(url, headers=headers, files=files, data=data)

    if response.status_code == 200:
        log_message("✅ File uploaded successfully.")
        log_message(f"Response: {response.text}")
    else:
        log_message("❌ Upload failed.")
        log_message(f"Status Code: {response.status_code}")
        log_message(f"Response: {response.text}")

def download_email_attachments(EMAIL, SENDER_EMAIL, APP_PASSWORD, DOWNLOAD_FOLDER, party_id, sheet_name, token):
    import email.utils
    print(SENDER_EMAIL)

    imap_server = "imap.gmail.com"
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(EMAIL, APP_PASSWORD)
    mail.select("inbox")

    status, messages = mail.search(None, f'FROM "{SENDER_EMAIL}"')

    if not os.path.isdir(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)

    email_ids = messages[0].split()
    if not email_ids:
        log_message("No emails found from this sender.")
        mail.logout()
        return

    latest_email_id = email_ids[-1]
    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8", errors="ignore")
            log_message(f"📨 Latest Email Subject: {subject}")

            msg_date = msg.get("Date")
            parsed_date = email.utils.parsedate_to_datetime(msg_date)
            log_message(f"📅 Received At: {parsed_date}")

            now = datetime.now(parsed_date.tzinfo)
            if (now - parsed_date) > timedelta(hours=24):
                log_message("⚠️ Email is older than 24 hours. Skipping.")
                mail.logout()
                return

            for part in msg.walk():
                if part.get_content_disposition() == "attachments":
                    filename = part.get_filename()
                    if filename:
                        decoded_filename, encoding = decode_header(filename)[0]
                        if isinstance(decoded_filename, bytes):
                            filename = decoded_filename.decode(encoding or "utf-8", errors="ignore")

                        filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        log_message(f"✅ Downloaded: {filename} → {filepath}")

                        post_downloaded_file(
                            file_path=filepath,
                            party_id=party_id,
                            sheet_name=sheet_name,
                            token=token
                        )

                        # Delete the file after posting
                        os.remove(filepath)
                        log_message(f"🗑️ Deleted: {filepath}")

    mail.logout()

def schedule_downloads():
    excel_file = 'credentials.xlsx'
    workbook = openpyxl.load_workbook(excel_file,data_only=True)
    sheet = workbook.active

    url = "https://sunrisediam.com:8223/api/employee/employee_login"
    username = sheet["E2"].value
    password = sheet["E3"].value
    token = get_Token(url=url, username=username, password=password)
    if not token:
        log_message("❌ Could not get token. Skipping this run.")
        return

    total_supp = sheet["B6"].value
    for i in range(total_supp):
        ind = -1
        upload_time_obj = sheet[f"F{8 + i}"].value
        if isinstance(upload_time_obj, time):
            now = datetime.now()
            upload_time = datetime.combine(now.date(), upload_time_obj)
            time_diff = abs((upload_time - now).total_seconds())

            if time_diff < 600:
                supp_name = sheet[f"A{8 + i}"].value
                log_message(f"⏰ Triggering for Supp Name: {supp_name}, Time match: {upload_time_obj}")
                ind = i
        else:
            log_message(f"⚠️ Invalid time format at F{8+i}: {upload_time_obj}")

        if ind != -1:
            email = sheet["B2"].value
            password = sheet["B3"].value
            sender_email = sheet[f"C{8 + ind}"].value
            party_id = sheet[f"B{8 + ind}"].value
            sheet_name = sheet[f"E{8 + ind}"].value
            if(party_id == 22375):
                yesterday = datetime.today() - timedelta(days=1)
                sheet_name = sheet_name + " " +yesterday.strftime("%d-%m-%Y")
            
            active = sheet[f"G{8 + ind}"].value

            if active == "Active":
                download_email_attachments(
                    EMAIL=email,
                    SENDER_EMAIL=sender_email,
                    APP_PASSWORD=password,
                    DOWNLOAD_FOLDER="attachments",
                    party_id=party_id,
                    sheet_name=sheet_name,
                    token=token
                )

# schedule_downloads()
# Set schedule
schedule.every().hour.at(":15").do(schedule_downloads)


log_message("🔁 Scheduler started. ")

while True:
    schedule.run_pending()
    time_module.sleep(1)
