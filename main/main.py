from dotenv import load_dotenv
load_dotenv()

import os
import base64
import datetime
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import requests

# load DeepSeek API key
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

# Gmail + Sheets API scopes
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Google Sheet IDs to store leads
LEADS_SHEET_ID = "YOUR_LEADS_SHEET_ID"
LOST_LEADS_SHEET_ID = "YOUR_LOST_LEADS_SHEET_ID"

def authenticate():
    # check for existing token
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # if no valid creds, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # refresh token if expired
            creds.refresh(Request())
        else:
            # run OAuth flow for new token
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save token for future use
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    # return Gmail service and Sheets service
    return build('gmail', 'v1', credentials=creds), build('sheets', 'v4', credentials=creds)

def send_email(service, to, subject, body):
    # create email message
    message = MIMEText(body)
    message['to'] = to
    message['subject'] = subject
    # encode email in base64
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {'raw': raw}
    # send email via Gmail API
    sent = service.users().messages().send(userId="me", body=body).execute()
    print(f"✅ Email sent to {to}. Message ID: {sent['id']}")

def ask_deepseek(prompt):
    # prepare headers for DeepSeek request
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
    }
    json_data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}]
    }
    # send request and return AI output
    response = requests.post("https://api.deepseek.com/v1/chat/completions", headers=headers, json=json_data)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]

def append_to_sheet(service, sheet_id, row):
    # append a row to the Google Sheet
    sheet = service.spreadsheets()
    body = {"values": [row]}
    sheet.values().append(
        spreadsheetId=sheet_id,
        range="Sheet1!A:D",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

def follow_up_logic(gmail_service, sheets_service):
    # current UTC time
    now = datetime.datetime.utcnow()
    # timestamps for 2 and 7 days ago
    two_days_ago = now - datetime.timedelta(days=2)
    seven_days_ago = now - datetime.timedelta(days=7)

    # Gmail search query for sent messages 2-7 days ago
    query = f"after:{int(seven_days_ago.timestamp())} before:{int(two_days_ago.timestamp())} in:sent"
    messages = gmail_service.users().messages().list(userId="me", q=query).execute().get('messages', [])

    print(f"Found {len(messages)} sent messages between 2–7 days ago.")

    for msg in messages:
        # fetch full message data
        msg_data = gmail_service.users().messages().get(userId="me", id=msg['id'], format='full').execute()
        headers = msg_data['payload']['headers']
        # get email subject and recipient
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
        to = next((h['value'] for h in headers if h['name'] == 'To'), '')
        # get message body
        parts = msg_data['payload'].get('parts', [])
        body = ""
        if parts:
            for part in parts:
                if part['mimeType'] == 'text/plain':
                    body = base64.urlsafe_b64decode(part['body']['data']).decode()
                    break
        else:
            body = base64.urlsafe_b64decode(msg_data['payload']['body']['data']).decode()

        # classify email as proposal using DeepSeek
        classify_prompt = f"Is the following email a business proposal or pitch? Reply only 'Yes' or 'No'.\n\n{body}"
        result = ask_deepseek(classify_prompt).strip().lower()

        if "yes" in result:
            # add as open lead in Google Sheet
            append_to_sheet(sheets_service, LEADS_SHEET_ID, [str(now.date()), to, subject, "Open Lead"])

            # check if client rejected
            if "we went with another company" in body.lower():
                # add to lost leads sheet
                append_to_sheet(sheets_service, LOST_LEADS_SHEET_ID, [str(now.date()), to, subject, "Lost Lead"])
                # send sympathetic email
                sympathetic_msg = "Thank you for considering us. We wish you success with your chosen provider."
                send_email(gmail_service, to, f"RE: {subject}", sympathetic_msg)
            else:
                # generate polite follow-up via DeepSeek
                followup_prompt = f"""Write a short, polite follow-up email paragraph to a client based on this previous message:
\"\"\"{body}\"\"\""""
                follow_up_body = ask_deepseek(followup_prompt).strip()
                # send follow-up email
                send_email(gmail_service, to, f"RE: {subject}", follow_up_body)

if __name__ == "__main__":
    # authenticate Gmail and Sheets services
    gmail_service, sheets_service = authenticate()
    # run follow-up logic
    follow_up_logic(gmail_service, sheets_service)
