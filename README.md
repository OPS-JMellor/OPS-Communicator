# 📧 OPS Communicator

An automated email announcement system built with Google Apps Script and Google Sheets. Send scheduled daily announcements with a Gmail-like rich text editor, automatic date placeholders, and comprehensive management features.

## ✨ Features

- **Rich Text Editor**: Gmail-like interface with formatting buttons, links, and lists
- **Flexible Scheduling**: Choose specific days and hourly time slots (6 AM - 11 PM)  
- **Date Placeholders**: Automatic date insertion with `[Date]`, `[Day]`, and `[ShortDate]`
- **Communication Management**: Edit, delete, and manage existing announcements
- **Test & Simulate**: Test emails and simulate hourly checks before going live
- **Comprehensive Logging**: Full audit trail and status monitoring
- **One-Click Setup**: Automated spreadsheet setup and trigger configuration

## 🚀 Quick Start

### Prerequisites
- Google Account with access to Google Sheets and Gmail
- Basic familiarity with Google Apps Script (optional)

### Installation

> **💡 Want to see it in action first?** Check out our [demo sheet](https://docs.google.com/spreadsheets/d/10k25AIzFcYtmRGEo8ap_d46hF154p4-6O6oHugca5C8/edit?usp=sharing) to see how it works before setting up your own.

1. **Create a new Google Sheet**
   - Go to [Google Sheets](https://sheets.google.com)
   - Create a new blank spreadsheet
   - Give it a name like "OPS Communications" or "Daily Announcements"

2. **Open Apps Script Editor**
   - In your Google Sheet, go to `Extensions > Apps Script`
   - Delete any existing code in the editor

3. **Add the Code**
   - Copy the entire contents of `Code.gs` from this repository
   - Paste it into the Apps Script editor
   - Save the project (Ctrl+S or Cmd+S)

4. **Run Setup**
   - Go back to your Google Sheet
   - Refresh the page (you may need to wait a moment)
   - Look for the "📧 Daily Announcements" menu in the menu bar
   - Click `📧 Daily Announcements > 🔑 Setup & Authorize`
   - Grant the necessary permissions when prompted

5. **Create Your First Announcement**
   - Click `📧 Daily Announcements > ➕ Add New Communication`
   - Fill out the form with your announcement details
   - Click "Add Communication"

6. **Manage Existing Announcements**
   - Click `📧 Daily Announcements > ✏️ Manage Communications`
   - Select the row number of the communication you want to edit
   - The edit dialog will open with all current values pre-populated
   - Make your changes and click "💾 Save Changes"
   - You can also delete communications using the "🗑️ Delete" button

## 📋 Usage Guide

### Creating Announcements

1. **Communication Name**: Give your announcement a descriptive name
2. **Send Time**: Choose from hourly slots (6:00 AM - 11:00 PM)
3. **Days to Send**: Select which days of the week to send
4. **Email Settings**: Set sender and recipient emails
5. **Subject Line**: Use date placeholders for dynamic subjects
6. **Message**: Use the rich text editor with formatting options

### Managing Existing Announcements

#### Editing Communications
1. **Access Management**: Click `📧 Daily Announcements > ✏️ Manage Communications`
2. **Select Communication**: Choose the row number from the list (e.g., "2", "3", "4")
3. **Edit Interface**: The dialog opens with all current values pre-populated:
   - Communication name, send time, and active status
   - All selected days remain checked
   - Email addresses and subject line preserved  
   - Message content loads in the rich text editor
4. **Save Changes**: Click "💾 Save Changes" to update the communication
5. **Delete Option**: Use "🗑️ Delete" to permanently remove the communication

#### Tips for Management
- **Active Toggle**: Uncheck "Active" to temporarily disable without deleting
- **Time Changes**: Easy to reschedule by changing the send time dropdown
- **Day Adjustments**: Add or remove days by checking/unchecking boxes
- **Message Updates**: Full rich text editing with all formatting preserved

### Date Placeholders

- `[Date]` → January 15, 2025
- `[Day]` → Wednesday  
- `[ShortDate]` → 1/15/2025

### Rich Text Features

- **Formatting**: Bold, italic, underline (buttons or Ctrl+B, Ctrl+I, Ctrl+U)
- **Lists**: Bullet points and numbered lists
- **Links**: Highlight text and click 🔗 Link button, or just type URLs
- **Keyboard Shortcuts**: Ctrl+K for links, standard formatting shortcuts

### Management Features

- **📊 Check Status**: View today's sent announcements
- **✏️ Manage Communications**: Edit existing announcements
- **🧪 Test Send Now**: Send immediate test emails
- **🕐 Simulate Hourly Check**: Preview what would be sent at specific times

## 🛠️ Technical Details

### Architecture
- **Google Apps Script**: Server-side JavaScript runtime
- **Google Sheets**: Data storage and user interface
- **Gmail API**: Email delivery through MailApp service
- **HTML Service**: Rich text editor and modal dialogs
- **Trigger Service**: Automated hourly execution

### Key Functions
- `sendDailyAnnouncements()`: Main execution function (runs hourly)
- `showAddCommunicationDialog()`: Rich text announcement creation
- `showManageCommunicationsDialog()`: Edit existing announcements
- `parseTimeString()`: Flexible time parsing (handles both strings and Date objects)
- `replacePlaceholders()`: Dynamic date substitution

### Data Structure
The system uses a Google Sheet with these columns:
- Communication Name
- Send Time  
- Active (boolean)
- From Email
- To Emails (comma-separated)
- Subject
- Message (HTML content)
- Send Days (comma-separated day abbreviations)
- Sent Today (date tracking)

## 🔧 Configuration

### Time Zones
The system automatically uses your Google account's time zone. No manual configuration needed.

### Email Limits
Google Apps Script has daily email quotas:
- Personal accounts: 100 emails/day
- Google Workspace: 1,500 emails/day

### Triggers
The system creates an hourly trigger that runs `sendDailyAnnouncements()`. You can view and manage triggers in the Apps Script editor under "Triggers" in the left sidebar.

## 📈 Monitoring & Troubleshooting

### Status Checking
- Use `📧 Daily Announcements > 📊 Check Status` to see today's activity
- Check `📧 Daily Announcements > ⏰ View Trigger Info` for automation status

### Common Issues

**"No times showing in dropdown"**
- Ensure you've run "Setup & Authorize" first
- Check that your spreadsheet has the correct headers

**"Test email not working"**
- Verify sender email is your Google account or authorized sending address
- Check recipient email addresses for typos

**"Emails not sending automatically"**  
- Verify the trigger is installed with "View Trigger Info"
- Check that communications are marked as "Active"
- Ensure current day is selected in "Send Days"

### Debugging
- Apps Script logs: Go to Apps Script editor > View > Logs
- Execution transcript: Apps Script editor > View > Execution transcript

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues, fork the repository, and create pull requests.

### Development Setup
1. Fork this repository
2. Create a new Google Apps Script project
3. Copy your changes to test them
4. Submit pull requests with detailed descriptions

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙋 Support

If you encounter issues:
1. Check the troubleshooting section above
2. Review the Apps Script logs for error messages
3. Open an issue on GitHub with detailed information

## 🎯 Roadmap

Future enhancements being considered:
- [ ] Email templates with variables
- [ ] Attachment support
- [ ] Advanced scheduling (specific dates, holidays)
- [ ] Email analytics and tracking
- [ ] Multi-language support
- [ ] Integration with external calendar systems

---

---

**OPS Communicator** - Made with ❤️ for automated communication management  
© 2025 - Licensed under MIT License
