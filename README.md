# Gmail Cleaner

A powerful Google Apps Script tool that helps you analyse and automatically clean up your Gmail inbox using Google Sheets. Perfect for managing free-tier Gmail accounts that are overflowing with emails!

## Why?
I made the classic mistake of signing up for Temu and was bombarded with emails constantly, I sat down one day and after spending 30mins+ and not even making a dint in the amount of emails, I decided there has to be a better way to handle this, so I have created this little appscript, I hope it helps you manage your inbox.


## 🚀 Quick Start (Easiest Method)

**The fastest way to get started** - no coding required!

1. **[Click here to make a copy of the template](https://docs.google.com/spreadsheets/d/15CXvM1upr2rm3ciRKn7hAuwiDHDaUDZEylxqT_zvwLM/copy)**
2. In your copy, click **Gmail Cleaner** → **Launch Sidebar**
3. Authorise permissions when prompted
4. Click **Setup Sheets** in the sidebar
5. Click **Analyse Inbox** to start!

That's it! The script is already installed and ready to use.

---

**Prefer to install manually?** See the [Manual Setup Guide](#manual-setup) below.

---

## Features

- **Inbox Analysis**: Scan your inbox with adjustable limits
  - Default: 1,000 emails per scan (adjustable in UI: 100 - 20,000)
  - Processes in batches for real-time progress
  - Shows live email count as it scans
  - Automatically resumes if interrupted
- **Flexible Rules**: Create custom cleanup rules based on:
  - Sender email address
  - Subject line patterns
  - Email content (supports full Gmail search syntax)
- **Batch Processing**: Clean unlimited emails with progress tracking
  - Processes 50 emails at a time for responsive updates
  - No limits on total emails cleaned
  - Automatically retries on errors
- **Automation**: Schedule automatic cleanup (daily or weekly)
- **Activity Logs**: Track all cleanup activities
- **Modern UI**: Beautiful sidebar interface with custom notifications and progress bars

## Manual Setup

Only follow these steps if you didn't use the template above and want to install from scratch.

### Step 1: Create a New Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it something like "Gmail Cleaner"

### Step 2: Add the Apps Script Code

1. In your Google Sheet, go to **Extensions** → **Apps Script**
2. Delete any existing code in the editor
3. Create the following files:

#### File 1: Code.gs
- In the Apps Script editor, paste the contents of `Code.gs` from this repository
- This file contains all the backend logic

#### File 2: Sidebar.html
- In the Apps Script editor, click the **+** button next to "Files"
- Select **HTML**
- Name it `Sidebar`
- Paste the contents of `Sidebar.html` from this repository

### Step 3: Save and Authorise

1. Click the **Save** icon (💾) in the Apps Script editor
2. Close the Apps Script tab and return to your Google Sheet
3. Refresh the page
4. You should see a new menu: **Gmail Cleaner**
5. Click **Gmail Cleaner** → **Launch Sidebar**
6. The first time you run this, Google will ask you to authorise the script:
   - Click "Review Permissions"
   - Select your Google account
   - Click "Advanced" → "Go to Gmail Cleaner (unsafe)"
   - Click "Allow"

### Step 4: Setup Your Sheets - IMPORTANT

1. In the sidebar, click **Setup Sheets**
2. This will create three sheets:
   - **Analysis**: Where email statistics will appear
   - **Rules**: Where you define cleanup rules
   - **Logs**: Activity history

## How to Use

### Analysing Your Inbox

The analysis scans your inbox to count emails by sender. You can control how many emails to scan each run.

**Gmail API Limits**: 
- Consumer accounts: ~20,000 email reads per day
- Don't run excessive scans to avoid hitting this limit

1. Open the sidebar: **Gmail Cleaner** → **Launch Sidebar**
2. Set your **Scan Limit** (default: 1,000 emails, range: 100-20,000)
3. Click **Analyse Inbox**
4. Watch the progress bar update in real-time
   - You'll see: "50 emails scanned... 100 emails scanned... 150 emails scanned..."
   - If interrupted, you can resume from where it left off
5. Check the **Analysis** sheet to see:
   - Top senders by email count
   - When you last received an email from them
   - Sample subject lines

**Tip**: Start with 1,000-2,000 emails for your first scan. You can always run it again with a higher limit!

### Creating Cleanup Rules

1. Go to the **Rules** sheet
2. Add a new row with the following columns:

| Rule Type | Value | Action | Status |
|-----------|-------|--------|--------|
| Sender | newsletter@example.com | Trash | Active |
| Subject | Unsubscribe | Archive | Active |
| Content | promotional | Mark Read | Paused |

**Rule Types:**
- **Sender**: Match emails from a specific sender
- **Subject**: Match emails with specific text in the subject
- **Content**: Match using Gmail search syntax (e.g., "has:attachment")

**Actions:**
- **Trash**: Move emails to trash
- **Archive**: Remove from inbox (but keep in archive)
- **Mark Read**: Mark as read (but don't move)

**Status:**
- **Active**: Rule will be applied
- **Paused**: Rule will be skipped

### Running Cleanup

1. After creating rules, click **Run Cleanup** in the sidebar
2. Watch the progress as the script processes each rule:
   - Shows which rule is currently running
   - Displays live count of emails cleaned
   - Processes in batches of 50 for better responsiveness
3. Check the **Logs** sheet for details about what was cleaned

**Note**: Each rule can clean unlimited emails. The script processes in batches to prevent timeouts and shows real-time progress!

### Automating Cleanup

1. In the sidebar, find the **Automation** section
2. Select a schedule:
   - **Manual only**: No automatic cleanup
   - **Daily (2:00 AM)**: Runs every day at 2:00 AM
   - **Weekly (Monday 2:00 AM)**: Runs every Monday at 2:00 AM
3. The script will automatically run cleanup based on your rules

## Important Notes

### Performance & Reliability

- **Batch Processing**: Both analysis and cleanup process emails in small batches (50 at a time)
- **No Limits**: Can handle inboxes of any size - 1,000, 10,000, or 100,000+ emails
- **Auto-Resume**: If a scan is interrupted (timeout/error), it automatically saves progress and can resume
- **Real-time Progress**: See exactly what's happening with live counters and progress bars
- **Gmail API Quotas**: Google has daily quotas for Gmail operations. If you hit limits, wait 24 hours

### Safety Tips

1. **Test First**: Always test rules manually before enabling automation
2. **Start Small**: Begin with "Archive" or "Mark Read" before using "Trash"
3. **Check Logs**: Review the Logs sheet regularly to ensure rules are working correctly
4. **Backup Important Emails**: Consider backing up important emails before running bulk cleanup
5. **Use Specific Rules**: The more specific your rules, the better results you'll get

### Permissions

>Just a heads up, the permission request for this is going to look horrible and dodgy, the reason for this is that it doesn't have a beautifully branded Oauth and to get that due to the nature of this permission required to run this, Google requires an annual security audit, which ya boy is not going to pay for, so look at the code and understand it, but realise that this is purely between you and the google servers and data is not getting passed to any third party, etc.

This script requires the following permissions:
- **Gmail**: To read and modify your emails
- **Spreadsheet**: To read/write data to the sheets
- **Triggers**: To run scheduled tasks

## Troubleshooting

### "Analysis sheet not found" Error
- Run **Gmail Cleaner** → **Setup Sheets** from the menu

### "No rules found" Error
- Add at least one rule to the **Rules** sheet with Status = "Active"

### Script Times Out
- Shouldn't happen with the new batch processing system
- If analysis is interrupted, simply click "Analyse Inbox" again - it will resume from where it stopped
- If cleanup is interrupted, run it again - it will continue processing remaining emails

### Automation Not Working
- Check that you've set a schedule in the sidebar
- Go to **Extensions** → **Apps Script** → **Triggers** (clock icon) to see if triggers are set up
- Make sure you have active rules in the Rules sheet

## Example Use Cases

### Unsubscribe from Newsletters
1. Run **Analyse Inbox**
2. Identify newsletter senders in the Analysis sheet
3. Add rules for each sender with Action = "Trash"
4. Run cleanup manually or schedule it

### Archive Old Promotional Emails
1. Add a rule: Type = "Content", Value = "category:promotions older_than:30d", Action = "Archive"
2. This will archive promotional emails older than 30 days

### Clean Up Social Media Notifications
1. Add rules for: facebook.com, twitter.com, linkedin.com, etc.
2. Set Action = "Archive" or "Trash"
3. Enable daily automation
4. Watch as hundreds or thousands of notifications get cleaned automatically!

> **Note**: I will probably add more features to this if there is enough requests for something, but I think what we have here is enough for 99% of users.

## Configuration

### Adjusting the Scan Limit

The scan limit controls how many emails are analysed in a single run:

1. Open the Gmail Cleaner sidebar
2. In the **Inbox Analysis** section, you'll see a "Scan Limit" input field
3. Adjust the value (minimum 100, maximum 20,000 emails)
4. Click **Analyse Inbox** to start scanning with your chosen limit

**Default**: 1,000 emails  
**Recommended**: Start with 1,000-2,000 for your first scan. Increase if needed.  
**Note**: Gmail API has a daily limit of ~20,000 email reads for consumer accounts.

## Thanks

- I have to thank Temu, if it wasn't for you sending 300 emails a  week, I would never have been frustrated enough to build this

## Licence

This project is free to use and modify. No warranty provided - use at your own risk!

## Support

If you find this tool helpful, consider:
- Sharing it with others who might benefit
- Suggesting improvements
- Reporting bugs or issues

---

**Happy Cleaning!**
