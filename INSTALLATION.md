# 📦 Installation Guide

Step-by-step instructions for setting up OPS Communicator.

## Method 1: Manual Installation (Recommended)

### Step 1: Create Google Sheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Click "Blank" to create a new spreadsheet
3. Name it something like "OPS Communications" or "Daily Announcements"

> **💡 Demo Available**: You can view a working example at our [demo sheet](https://docs.google.com/spreadsheets/d/10k25AIzFcYtmRGEo8ap_d46hF154p4-6O6oHugca5C8/edit?usp=sharing)

### Step 2: Open Apps Script
1. In your Google Sheet, click `Extensions` in the menu
2. Click `Apps Script` 
3. This opens the Google Apps Script editor in a new tab

### Step 3: Add the Code
1. In the Apps Script editor, you'll see a file called `Code.gs`
2. Delete any existing code in that file
3. Copy the entire contents of `Code.gs` from this repository
4. Paste it into the editor
5. Click the save button (💾) or press Ctrl+S

### Step 4: Run Initial Setup
1. Go back to your Google Sheet tab
2. Refresh the page (F5 or Ctrl+R)
3. Wait about 10-15 seconds for the menu to appear
4. You should see a new menu called "📧 Daily Announcements"
5. Click `📧 Daily Announcements > 🔑 Setup & Authorize`

### Step 5: Grant Permissions
When you run Setup & Authorize, Google will ask for permissions:

1. Click "Review permissions"
2. Choose your Google account
3. Click "Advanced" if you see a warning
4. Click "Go to [Your Project Name] (unsafe)" 
5. Click "Allow"

**Required Permissions:**
- Read/write access to your Google Sheets
- Send emails through Gmail
- Create time-based triggers

### Step 6: Verify Setup
After granting permissions:
1. The setup should complete automatically
2. Your spreadsheet will have new column headers
3. You should see a success message

### Step 7: Create Your First Announcement
1. Click `📧 Daily Announcements > ➕ Add New Communication`
2. Fill out the form:
   - **Name**: "Test Announcement"
   - **Time**: Choose any time
   - **Days**: Select today's day
   - **From/To**: Use your email address
   - **Subject**: "Test: [Date]"
   - **Message**: "This is a test message!"
3. Click "Add Communication"

### Step 8: Test the System
1. Click `📧 Daily Announcements > 🧪 Test Send Now`
2. You should receive a test email immediately
3. Try `📧 Daily Announcements > 🕐 Simulate Hourly Check` to preview automation

## Method 2: Using Google Apps Script Directly

If you prefer to start from Apps Script:

1. Go to [script.google.com](https://script.google.com)
2. Click "New project"
3. Replace the default code with the contents of `Code.gs`
4. Save the project
5. You'll need to create a Google Sheet manually and link it

## Troubleshooting Installation

### "Menu doesn't appear"
- Wait 10-15 seconds after refreshing
- Try hard refresh (Ctrl+Shift+R)
- Check that you saved the code in Apps Script

### "Permission denied" errors
- Make sure you granted all requested permissions
- Try running Setup & Authorize again
- Check that your Google account has access to create triggers

### "Script function not found"
- Ensure you copied the entire `Code.gs` file
- Check that there are no syntax errors (red underlines in editor)
- Try saving and refreshing again

### "Execution transcript shows errors"
- Check the error message in Apps Script > View > Execution transcript
- Common issues: email quota exceeded, invalid email addresses

## Security Notes

This script requires several permissions because it needs to:
- **Read/Write Sheets**: Store your announcements and settings
- **Send Email**: Deliver your announcements via Gmail
- **Create Triggers**: Set up automatic hourly checking

The script only accesses:
- The specific Google Sheet you created
- Your Gmail for sending (not reading) emails
- Google's trigger service for automation

## Next Steps

After successful installation:
1. Read the [Usage Guide](README.md#usage-guide) in the main README
2. Create a few test announcements
3. Set up your real daily announcements
4. Monitor with the status checking tools

## Getting Help

If you encounter issues during installation:
1. Check the troubleshooting section above
2. Look at the Apps Script execution logs
3. Open an issue on GitHub with:
   - What step failed
   - Any error messages you see
   - Screenshots if helpful
