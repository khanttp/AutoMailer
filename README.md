# AutoMailer

A Google Apps Script that helps you send personalized emails directly from a Google Sheet using a Gmail draft. It sends emails in timed batches so you stay within Gmail's daily limits. Great for an org. or a campaign that needs sending out mass emails without hitting Google's spam detection or rate limits.

---

## What It Does

- Sends personalized emails using a Gmail draft as a template
- Replaces placeholders like `{{First Name}}` with data from your sheet
- Sends emails in batches every 5 minutes
- Automatically stops when Gmail's daily limit is hit
- Logs errors and sending status in the sheet

---


### Step 1: Set Up Your Sheet

Make sure your sheet includes at least:
- `Recipient` (email address)
- `Email Sent` (leave blank to track sending status)

You can add more columns for personalization like `First Name`, `Company`, etc then reference them like `{{First Name}}` in the email.

[[example Sheet](https://docs.google.com/spreadsheets/d/1k4NYdMDsi2Oc-mIvRS8vkZoRFIWqtFVSVP0I4j6oEsA/copy)]

### Step 2: Create a Gmail Draft

- Compose a new email in Gmail
- Use placeholders like `{{First Name}}` in the subject or body.
- Give it a unique subject. Save it as a draft.
- This draft will be your email template

---

## Installation

1. In your Google Sheet, go to **Extensions > Apps Script**
2. Paste the AutoMailer script into the editor (or use the example sheet that already has the script)
3. Click **Save**, then **Run** the script once to trigger the authorization prompt
4. Grant the requested permissions

---

## Sending Emails

1. Refresh the sheet after setup
2. Use the new **AutoMailer > Start AutoMailer** menu
3. Enter the exact subject line of your Gmail draft when prompted

Emails will begin sending within 5 minutes and continue in batches. The sheet may say “script finished” — that just means the trigger is set, not that all emails are sent yet.

---

## Managing AutoMailer

**Change Batch Size**
Edit the `BATCH_SIZE` variable in the script to control how many emails send every 5 minutes.

**Stop or Reset**
Use **AutoMailer > Reset AutoMailer** to stop the script and clear its settings.

---
