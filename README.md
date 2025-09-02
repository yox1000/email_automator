# Email Automator

Automates follow-up emails to clients for business proposals and tracks leads using AI assistance. Works with **Outlook** (Microsoft Graph API) or **Gmail** (Gmail API + Google Sheets). Uses **DeepSeek AI** to classify emails and generate polite follow-ups.

Features:

* Fetch sent emails from the last 2–7 days.
* Identify business proposals using AI.
* Automatically send follow-up emails to clients if no response.
* Track **open leads** and **lost leads** in Excel (Outlook) or Google Sheets (Gmail).
* Send sympathetic messages to clients who went with a different company.
* Optional: Send reminders to admin if no client response within 3 days (future enhancement).

Requirements:

* Python 3.10+
* Packages:
  pip install requests python-dotenv msal openpyxl google-auth google-auth-oauthlib google-api-python-client
* DeepSeek API Key (stored in `.env`)
* Outlook credentials (for Outlook version) or Gmail OAuth credentials (for Gmail version)
* Excel file (for Outlook) or Google Sheets (for Gmail) for tracking leads.

Setup:

1. Clone this repository.
2. Create a `.env` file in the project root with the following variables:

Outlook version:
OUTLOOK\_CLIENT\_ID=your\_client\_id
OUTLOOK\_CLIENT\_SECRET=your\_client\_secret
OUTLOOK\_TENANT\_ID=your\_tenant\_id
DEEPSEEK\_API\_KEY=your\_deepseek\_api\_key

Gmail version:
DEEPSEEK\_API\_KEY=your\_deepseek\_api\_key

3. For Gmail, download `credentials.json` from Google Cloud Console and place it in the project root.
4. For Gmail, create two Google Sheets for leads and lost leads, and note their sheet IDs.

Usage:

Outlook Version:
python outlook\_email\_automator.py

Gmail Version:
python gmail\_email\_automator.py

How it Works:

1. Authenticate with Outlook or Gmail API.
2. Fetch sent emails in the last 2–7 days.
3. Classify each email using DeepSeek:

   * "Yes" → business proposal.
   * "No" → skip.
4. Follow-up logic:

   * If client hasn’t responded → generate polite follow-up.
   * If client went with another company → generate sympathetic message.
5. Track leads:

   * Open leads → Excel / Google Sheet.
   * Lost leads → Excel / Google Sheet.

Notes:

* Ensure Excel files or Google Sheets exist and have correct headers: `Date | Client | Subject | Status`.
* DeepSeek AI generates responses; review them if necessary before sending in production.
* Gmail API requires `token.json` for OAuth token caching.

Future Enhancements:

* Add 3-day no-response admin reminders.
* Enable automatic OneDrive or Google Drive syncing for Excel files.
* Add email templates and personalization options.
* Support attachments and HTML emails.

License:
MIT License
