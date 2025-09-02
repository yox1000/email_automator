from dotenv import load_dotenv
load_dotenv()

import os
import datetime
import requests
from msal import ConfidentialClientApplication

CLIENT_ID = os.getenv("OUTLOOK_CLIENT_ID")
CLIENT_SECRET = os.getenv("OUTLOOK_CLIENT_SECRET")
TENANT_ID = os.getenv("OUTLOOK_TENANT_ID")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

def authenticate_outlook():
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_silent(SCOPES, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Authentication failed", result.get("error_description"))

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

def send_email_outlook(access_token, to, subject, body):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    email_msg = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": to}}
            ]
        }
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json=email_msg)
    response.raise_for_status()
    print(f"Follow-up sent to {to}.")

def create_draft_email_outlook(access_token, to, subject, body):
    url = "https://graph.microsoft.com/v1.0/me/messages"
    draft_msg = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body
        },
        "toRecipients": [
            {"emailAddress": {"address": to}}
        ],
        "isDraft": True
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json=draft_msg)
    response.raise_for_status()
    print(f"📝 Draft created for {to}.")

def follow_up_logic_outlook(access_token, send_email_flag=True):
    now = datetime.datetime.utcnow()
    two_days_ago = now - datetime.timedelta(days=2)
    seven_days_ago = now - datetime.timedelta(days=7)

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    query = (
        f"https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages"
        f"?$filter=sentDateTime ge {seven_days_ago.isoformat()}Z and sentDateTime le {two_days_ago.isoformat()}Z"
        f"&$top=50"
    )

    response = requests.get(query, headers=headers)
    response.raise_for_status()
    messages = response.json().get("value", [])

    print(f"Found {len(messages)} sent messages between 2–7 days ago.")

    for msg in messages:
        subject = msg.get("subject", "")
        to_recipients = msg.get("toRecipients", [])
        to = to_recipients[0]["emailAddress"]["address"] if to_recipients else ""
        body = msg.get("body", {}).get("content", "")

        classify_prompt = f"Is the following email a business proposal or pitch? Reply only 'Yes' or 'No'.\n\n{body}"
        result = ask_deepseek(classify_prompt).strip().lower()

        if "yes" in result:
            followup_prompt = f"""Write a short, polite follow-up email paragraph to a client based on this previous message:
\"\"\"{body}\"\"\"

Include a reminder of the main idea in the initial proposal and express interest in hearing their thoughts on this initial proposal. Do not directly copy the original message."""
            follow_up_body = ask_deepseek(followup_prompt).strip()

            if send_email_flag:
                send_email_outlook(access_token, to, f"RE: {subject}", follow_up_body)
            else:
                create_draft_email_outlook(access_token, to, f"RE: {subject}", follow_up_body)

if __name__ == "__main__":
    access_token = authenticate_outlook()
    
    # Set send_email_flag to False to create drafts instead of sending
    follow_up_logic_outlook(access_token, send_email_flag=False)
