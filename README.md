# ✉️ Sheet Draft - Google Sheets Email Automation Tool 🚀

**Streamline your email communication directly from Google Sheets with Sheet Draft\!** 🎉 This Google Sheets add-on is designed to **automate email sending and draft creation** using your spreadsheet data. 🗂️

Whether you need to:

*   📢  Send bulk announcements efficiently
*   💌  Manage email marketing campaigns
*   📝  Prepare drafts for review and sending

**Sheet Draft simplifies your email workflow and boosts your productivity\!** ⏱️ It's a practical solution for managing email tasks within Google Sheets.

## Key Features 🌟

*   ✅ **Send or Draft Emails:** Choose your preferred action: send emails directly via Gmail 📤 or save them as drafts for later review and sending 💾.
*   📅 **Scheduled Sending:** Plan ahead with scheduled email delivery. 🗓️ Set emails to send automatically at a specific date and time, ideal for planned communications and campaigns.
*   ✍️ **Personalized Emails:**  Enhance engagement with personalized emails. ✨ Use data from your Google Sheet columns to customize recipient details, subject lines, and email body content dynamically.
*   📧 **Sender Email Management:**  Control your sender identity. ⚙️ Configure and utilize authorized "send-as" email addresses to send emails from various aliases or organizational accounts.
*   ⏳ **Random Delay:**  Maintain email sending best practices. 😇 Add a randomized delay between emails to simulate natural sending patterns, helping to avoid spam filters and adhere to Gmail's sending policies.
*   📊 **Daily Sending Limit:**  Stay within Gmail's recommended sending volumes. 🛡️  Manage the maximum number of emails sent daily to prevent account restrictions and ensure reliable delivery.
*   🤹 **Flexible Row Processing:**  Customize your email sends. 🎯 Process specific row ranges within your sheet or handle all rows containing data for maximum flexibility.
*   🧑‍💻 **User-Friendly Interface:** Easy to use for all skill levels. 🙌 An intuitive dialog box within Google Sheets simplifies setup and operation, making email automation accessible to everyone.

## Setup Instructions 🛠️

Follow these steps to set up Sheet Draft in your Google Sheets:

1.  **Open your Google Sheet 📂:** Go to the Google Sheet that contains your email data. Ensure you have columns for:

    *   📧 Recipient Emails
    *   Subject Lines 💬
    *   Email Bodies 📝

2.  **Open Script Editor 👨‍💻:** In Google Sheets, navigate to **"Extensions" > "Apps Script"**. This will open the Google Apps Script editor in a new window.

3.  **Copy and Paste Code ✂️ + 📋:**  Copy **all** of the provided code (HTML, CSS, JavaScript, and Google Apps Script functions).  In the Apps Script editor, **replace all existing code** in the default `Code.gs` file with the copied code.

4.  **Save the Script ✅:** Click the **"Save"** icon (💾 - floppy disk) in the Apps Script editor. Name your project "Sheet Draft" or any name you prefer.

5.  **Refresh Google Sheet 🔄:** Return to your Google Sheet and **refresh** the page in your browser.  A new custom menu, **"Sheet Draft by Josh"** (or your project name), should now be visible in the Google Sheets menu bar.

6.  **Authorize Script 🔑:**

    *   Click on **"Sheet Draft by Josh" > "Open Email Dialog"**.
    *   A dialog box will appear. The first time you run the script, you will be prompted to authorize access. Click **"Continue"**.
    *   **Choose your Google Account.** 👤
    *   **Review Permissions:** Examine the permissions Sheet Draft requests (access to your spreadsheet and Gmail). Click **"Allow"** to grant permissions. Sheet Draft requires these permissions to operate correctly.

7.  **Email List Sheet (Optional, Recommended) 📝:**

    *   For streamlined management of "send-as" emails, consider creating a sheet named **"Email List"** in your Google Sheet workbook.
    *   In the "Email List" sheet, create a single column (Column A) and list each email address you want to authorize as a "send-as" sender within Sheet Draft.

## How to Use Sheet Draft 🚀

Step-by-step guide to using Sheet Draft:

1.  **Open the Dialog Box 💬:** In Google Sheets, click on **"Sheet Draft by Josh" > "Open Email Dialog"**. The Sheet Draft control panel will appear.

2.  **Sender Email Configuration 📧:**

    *   **Sender Email List:** View your list of authorized sender emails, sourced from the "Email List" sheet (if used) or previously added emails.
    *   **Add Sender Email:**
        *   Enter a new email address in the **"Add Sender Email"** field.
        *   Click **"Add Email"**. The email will be added to the sender list and to the "Email List" sheet if you are using one.
    *   **"Send As" Selection 🎭:**
        *   Choose the desired sender email address from the **"Send As"** dropdown. This includes your primary Google account email and any emails from your Sender Email List.
        *   **Important:** "Send-as" alias emails require prior configuration in your Gmail settings.

3.  **Column Configuration 🗂️:**

    *   **Email Column:**  Enter the **column letter** (e.g., `A`, `B`, `C`) for the column containing **recipient email addresses**.
    *   **Subject Column:** Enter the **column letter** for the **subject line column**.
    *   **Body Column:** Enter the **column letter** for the **email body column**.
    *   **Data Previews:** Previews of data from the **second row** of the specified columns will display to help verify your column selections.

4.  **Scheduling (Optional) ⏰:**

    *   **"Schedule Email Sending" Option:** Enable email scheduling by checking the **"Schedule Email Sending"** box.
    *   **Schedule Date:** Select the desired sending date using the date picker. Defaults to the next day.
    *   **Schedule Time:**  Enter the sending time in **24-hour format** (HH:mm), e.g., `14:30` for 2:30 PM.
    *   **"Schedule Emails" Button:** Click **"Schedule Emails"** to schedule the process. Confirm in the dialog box.
    *   **"Clear Schedule" Button:** Cancel any scheduled process by clicking **"Clear Schedule"**.

5.  **Delay Settings (Optional) 🐌:**

    *   **"Add Random Delay" Option:** Enable a random delay between emails by checking **"Add Random Delay"**.
    *   **Min Delay (seconds):** Enter the **minimum** delay duration in seconds.
    *   **Max Delay (seconds):** Enter the **maximum** delay duration in seconds. A random delay within this range will be applied between each email.

6.  **Daily Sending Limit 🚦:**

    *   **"Enable Daily Limit" - Default Enabled:** The **"Enable Daily Limit"** box is checked by default for safety. Keep this enabled to manage Gmail sending limits.
    *   **"Max Emails per Day" Setting:** The default limit is **`490`**, suitable for most free Gmail accounts. Adjust if you have a Google Workspace account with higher limits, but stay within your account's limits.

7.  **Action 🎯:**

    *   Select your desired action from the **"Action"** dropdown:
        *   **"Save as Drafts"**: Save emails as drafts in your Gmail "Drafts" folder.
        *   **"Send Emails"**: Send emails directly to recipients.

8.  **Start Row Configuration 🏁:**

    *   **"Start Row" Setting:**  Processing starts from **row `2` by default** (assuming row 1 is headers). Adjust if your data starts on a different row.
    *   **"Enable End Row" Option:** Check **"Enable End Row"** to process a specific row range.
    *   **"End Row" Setting:** If enabled, enter the **end row number**. If disabled, all rows from "Start Row" to the last data row will be processed.

9.  **Email Processing and Testing 🚀:**

    *   **"Run Email Process" Button:** Click **"Run Email Process"** to begin sending emails or saving drafts based on your "Action" selection. Confirm in the dialog box.
    *   **"Send Test Email" Button:** Click **"Send Test Email"** to send a test email to your Gmail address using the selected "Send As" email. Use this to verify your configuration.

10. **Email Status 📊:**

    *   **Sent Emails:**  Count of emails successfully sent in the current run.
    *   **Draft Emails:**  Count of emails saved as drafts in the current run.
    *   **Daily Emails Sent:**  Tracks emails sent within the current day's script executions. (Persistent daily tracking requires more advanced storage).
    *   **"Reset All Counters" Button:** Click **"Reset All Counters"** to reset the sent, draft, and daily sent email counters to zero.
    *   **Processing Status:**  "Processing Emails..." or "Scheduling Your Emails..." messages indicate script activity.

## ⚠️  Important Notes and Disclaimers ⚠️

*   **Gmail Sending Limits:**  Be mindful of Gmail's sending limits, especially for free accounts (around 500 emails daily). Exceeding limits may lead to temporary account restrictions. Sheet Draft's daily limit feature helps manage this.
*   **"Send-as" Permissions:**  "Send-as" email addresses require proper setup in Gmail settings. Verify "send-as" permissions for any alias emails you intend to use.
*   **Script Authorization:**  Google Apps Script authorization is required to access your Google Sheet and Gmail. Review permissions before authorizing.
*   **Error Handling:**  Basic error handling is included. Test with a small batch first to ensure correct setup. Check Apps Script execution logs ("Executions") for detailed error information if needed.
*   **Scheduling Reliability:**  Scheduled triggers in Google Apps Script are generally reliable, but slight timing variations can occur. Test scheduling for critical time-sensitive emails.
*   **Random Delay Effectiveness:**  The random delay is a basic measure to help avoid spam filters. Dedicated email marketing platforms offer more robust solutions.
*   **Data Privacy:**  Sheet Draft processes data within your Google Sheet. Ensure you are comfortable with script access and processing of your sheet's data, especially sensitive information.
*   **No Warranty:**  This script is provided as-is, without warranty. Use at your own risk.

📜 **Disclaimer:** This README is based on the provided code and description.  Always review and test the script before using it for important email communications. The author is not responsible for issues arising from script use.

-----

Let me know if you have any questions or need further assistance\! 😊
