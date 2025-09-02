from dotenv import load_dotenv
load_dotenv()

import os
import datetime
import requests
from msal import ConfidentialClientApplication
import openpyxl

# load Outlook and DeepSeek credentials from environment variables
CLIENT_ID = os.getenv("OUTLOOK_CLIENT_ID")
CLIENT_SECRET = os.getenv("OUTLOOK_CLIENT_SECRET")
TENANT_ID = os.getenv("OUTLOOK_TENANT_ID")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

# Microsoft Graph API scope and authority URL
SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Excel files to store leads and lost leads
LEADS_FILE = "leads.xlsx"
LOST_LEADS_FILE = "lost_leads.xlsx"

def authenticate_outlook():
    # create an MSAL confidential client
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    # try to get a cached token silently
    result = app.acquire_token_silent(SCOPES, account=None)
    # if no cached token, request a new one
    if not result:
        result = app.acquire_token_for_client(scopes=SCOPES)
    # return access token if successful
    if "access_token" in result:
        return result["access_token"]
    else:
        # raise exception if authentication failed
        raise Exception("Authentication failed", result.get("error_description"))

def ask_deepseek(prompt):
    # prepare headers for DeepSeek API request
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
    }
    # payload for DeepSeek chat model
    json_data = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}]
    }
    # send request and raise exception if error occurs
    response = requests.post("https://api.deepseek.com/v1/chat/completions", headers=headers, json=json_data)
    response.raise_for_status()
    # return the AI response content
    return response.json()["choices"][0]["message"]["content"]

def send_email_outlook(access_token, to, subject, body):
    # Outlook API endpoint for sending email
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    # construct email message payload
    email_msg = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to}}]
        }
    }
    # set authorization header with token
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    # send the email
    response = requests.post(url, headers=headers, json=email_msg)
    response.raise_for_status()
    print(f"✅ Email sent to {to}.")

def append_to_excel(filename, row):
    # check if Excel file exists
    if not os.path.exists(filename):
        # create workbook and worksheet if file doesn't exist
        wb = openpyxl.Workbook()
        ws = wb.active
        # add header row
        ws.append(["Date", "Client", "Subject", "Status"])
    else:
        # load existing workbook
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
    # append new row to worksheet
    ws.append(row)
    # save workbook
    wb.save(filename)

def follow_up_logic_outlook(access_token):
    # current UTC time
    now = datetime.datetime.utcnow()
    # two days ago
    two_days_ago = now - datetime.timedelta(days=2)
    # seven days ago
    seven_days_ago = now - datetime.timedelta(days=7)

    # authorization header
    headers = {"Authorization": f"Bearer {access_token}"}
    # query to fetch sent messages between 2-7 days ago
    query = (
        f"https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages"
        f"?$filter=sentDateTime ge {seven_days_ago.isoformat()}Z and sentDateTime le {two_days_ago.isoformat()}Z"
        f"&$top=50"
    )
    # get messages from Outlook API
    response = requests.get(query, headers=headers)
    response.raise_for_status()
    messages = response.json().get("value", [])

    print(f"Found {len(messages)} sent messages between 2–7 days ago.")

    for msg in messages:
        # get subject of email
        subject = msg.get("subject", "")
        # get recipient email address
        to_recipients = msg.get("toRecipients", [])
        to = to_recipients[0]["emailAddress"]["address"] if to_recipients else ""
        # get email body
        body = msg.get("body", {}).get("content", "")

        # classify if email is a business proposal
        classify_prompt = f"Is the following email a business proposal or pitch? Reply only 'Yes' or 'No'.\n\n{body}"
        result = ask_deepseek(classify_prompt).strip().lower()

        if "yes" in result:
            # save as open lead in Excel
            append_to_excel(LEADS_FILE, [now.date(), to, subject, "Open Lead"])

            # check if client rejected lead
            if "we went with another company" in body.lower():
                # save lost lead in Excel
                append_to_excel(LOST_LEADS_FILE, [now.date(), to, subject, "Lost Lead"])
                # send sympathetic email
                sympathetic_msg = "Thank you for considering us. We wish you success with your chosen provider."
                send_email_outlook(access_token, to, f"RE: {subject}", sympathetic_msg)
            else:
                # generate polite follow-up via DeepSeek
                followup_prompt = f"""Write a short, polite follow-up email paragraph to a client based on this previous message:
\"\"\"{body}\"\"\""""
                follow_up_body = ask_deepseek(followup_prompt).strip()
                # send follow-up email
                send_email_outlook(access_token, to, f"RE: {subject}", follow_up_body)

if __name__ == "__main__":
    # authenticate and get access token
    token = authenticate_outlook()
    # run follow-up logic
    follow_up_logic_outlook(token)
