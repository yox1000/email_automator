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

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly'
]

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def send_email(service, to, subject, body):
    message = MIMEText(body)
    message['to'] = to
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {'raw': raw}
    sent = service.users().messages().send(userId="me", body=body).execute()
    print(f"✅ Follow-up sent to {to}. Message ID: {sent['id']}")

def ask_deepseek(prompt):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
    }
    json_data = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "user", "content": prompt}
        ]
    }
    response = requests.post("https://api.deepseek.com/v1/chat/completions", headers=headers, json=json_data)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]

def follow_up_logic(service):
    user_id = "me"
    now = datetime.datetime.utcnow()
    two_days_ago = now - datetime.timedelta(days=2)
    seven_days_ago = now - datetime.timedelta(days=7)

    query = f"after:{int(seven_days_ago.timestamp())} before:{int(two_days_ago.timestamp())} in:sent"
    messages = service.users().messages().list(userId=user_id, q=query).execute().get('messages', [])

    print(f"Found {len(messages)} sent messages between 2–7 days ago.")

    for msg in messages:
        msg_data = service.users().messages().get(userId=user_id, id=msg['id'], format='full').execute()
        headers = msg_data['payload']['headers']
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
        to = next((h['value'] for h in headers if h['name'] == 'To'), '')
        parts = msg_data['payload'].get('parts', [])
        body = ""
        if parts:
            for part in parts:
                if part['mimeType'] == 'text/plain':
                    body = base64.urlsafe_b64decode(part['body']['data']).decode()
                    break
        else:
            body = base64.urlsafe_b64decode(msg_data['payload']['body']['data']).decode()
        classify_prompt = f"Is the following email a business proposal or pitch? Reply only 'Yes' or 'No'.\n\n{body}"
        result = ask_deepseek(classify_prompt).strip().lower()
        if "yes" in result:
            followup_prompt = f"""Write a short, polite follow-up email paragraph to a client based on this previous message:
\"\"\"{body}\"\"\"

Include a reminder of the main idea in the intial proposal and express interest in hearing their thoughts on this initial proposal. Do not directly copy the original message."""
            follow_up_body = ask_deepseek(followup_prompt).strip()

            send_email(service, to, f"RE: {subject}", follow_up_body)


#add sending to excel spreadsheet  
if __name__ == "__main__":
    service = authenticate()
    follow_up_logic(service)
