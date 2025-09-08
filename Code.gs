function sendDailyAnnouncements() {
  // Get the active spreadsheet and first sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get current date and time
  const now = new Date();
  const currentHour = now.getHours();
  const currentDay = now.getDay(); // 0 = Sunday, 1 = Monday, etc.
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const currentDayName = dayNames[currentDay];
  
  // Get all data from the sheet (assuming headers in row 1)
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log('No announcements configured');
    return;
  }
  
  const headers = data[0];
  
  // Find column indices
  const nameCol = headers.indexOf('Communication Name');
  const timeCol = headers.indexOf('Send Time');
  const activeCol = headers.indexOf('Active');
  const fromCol = headers.indexOf('From Email');
  const toCol = headers.indexOf('To Emails');
  const subjectCol = headers.indexOf('Subject');
  const messageCol = headers.indexOf('Message');
  const daysCol = headers.indexOf('Send Days');
  const sentTodayCol = headers.indexOf('Sent Today');
  
  if (nameCol === -1 || timeCol === -1) {
    console.log('Required columns not found. Please check sheet headers.');
    return;
  }
  
  // Process each row (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip if not active
    if (!row[activeCol]) continue;
    
    // Check if today is a send day
    const sendDays = row[daysCol] ? row[daysCol].toString() : '';
    if (sendDays && !sendDays.includes(currentDayName)) {
      console.log(`Skipping "${row[nameCol]}" - not scheduled for ${currentDayName}`);
      continue;
    }
    
    // If no send days specified, use old weekend logic (Mon-Fri only)
    if (!sendDays && (currentDay === 0 || currentDay === 6)) {
      console.log(`Skipping "${row[nameCol]}" - weekend and no specific days configured`);
      continue;
    }
    
    // Parse send time
    const sendTimeStr = row[timeCol];
    if (!sendTimeStr) continue;
    
    const sendHour = parseTimeString(sendTimeStr);
    if (sendHour === -1) continue; // Invalid time format
    
    // Check if it's time to send
    if (currentHour !== sendHour) continue;
    
    // Check if already sent today
    const sentToday = row[sentTodayCol];
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    if (sentToday === today) {
      console.log(`Already sent "${row[nameCol]}" today`);
      continue;
    }
    
    // Send the email
    try {
      const fromEmail = row[fromCol];
      const toEmails = row[toCol];
      const subject = replacePlaceholders(row[subjectCol], now);
      const message = replacePlaceholders(row[messageCol], now);
      
      // Split multiple recipients if comma-separated
      const recipients = toEmails.split(',').map(email => email.trim()).join(',');
      
      MailApp.sendEmail({
        to: recipients,
        subject: subject,
        htmlBody: message
      });
      
      // Mark as sent today
      sheet.getRange(i + 1, sentTodayCol + 1).setValue(today);
      
      console.log(`Sent "${row[nameCol]}" successfully`);
      
    } catch (error) {
      console.error(`Error sending "${row[nameCol]}":`, error);
    }
  }
}

function parseTimeString(timeStr) {
  // Parse time strings like "8:00 AM", "2:30 PM", "14:30"
  if (!timeStr) return -1;
  
  // Check if input is a Date object
  if (timeStr instanceof Date) {
    console.log(`parseTimeString received Date object: ${timeStr}`);
    return timeStr.getHours();
  }
  
  const time = timeStr.toString().trim().toUpperCase();
  console.log(`parseTimeString input: "${timeStr}" -> cleaned: "${time}"`);
  
  // Handle 12-hour format
  if (time.includes('AM') || time.includes('PM')) {
    // Extract AM/PM
    const period = time.includes('AM') ? 'AM' : 'PM';
    
    // Extract time part (everything before AM/PM)
    const timePart = time.replace(/\s*(AM|PM).*/, '').trim();
    console.log(`12-hour format - period: ${period}, timePart: "${timePart}"`);
    
    if (!timePart.includes(':')) return -1;
    
    const parts = timePart.split(':');
    const hours = parseInt(parts[0]);
    
    if (isNaN(hours) || hours < 1 || hours > 12) return -1;
    
    let hour24 = hours;
    if (period === 'PM' && hours !== 12) hour24 += 12;
    if (period === 'AM' && hours === 12) hour24 = 0;
    
    console.log(`Parsed 12-hour: ${hours} ${period} -> ${hour24}`);
    return hour24;
  }
  
  // Handle 24-hour format
  if (time.includes(':')) {
    const parts = time.split(':');
    const hours = parseInt(parts[0]);
    
    if (isNaN(hours) || hours < 0 || hours > 23) return -1;
    
    console.log(`Parsed 24-hour: "${time}" -> ${hours}`);
    return hours;
  }
  
  // Try parsing as just a number (like "15")
  const numericHour = parseInt(time);
  if (!isNaN(numericHour) && numericHour >= 0 && numericHour <= 23) {
    console.log(`Parsed numeric: "${time}" -> ${numericHour}`);
    return numericHour;
  }
  
  console.log(`Failed to parse: "${time}"`);
  return -1; // Invalid format
}

function replacePlaceholders(text, date) {
  if (!text) return '';
  
  // The text is now HTML from the rich editor, so just replace placeholders
  let formatted = text.toString()
    .replace(/\[Date\]/g, Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM d, yyyy'))
    .replace(/\[Day\]/g, Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE'))
    .replace(/\[ShortDate\]/g, Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d/yyyy'));
  
  return formatted;
}

function setupTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyAnnouncements') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new hourly trigger
  ScriptApp.newTrigger('sendDailyAnnouncements')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Trigger setup complete - will run every hour');
}

function testAnnouncements() {
  // Test function - sends email for selected row, or first active if none selected
  console.log('Running test send...');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No communications found', 'Please add a communication first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const headers = data[0];
  const nameCol = headers.indexOf('Communication Name');
  const activeCol = headers.indexOf('Active');
  const fromCol = headers.indexOf('From Email');
  const toCol = headers.indexOf('To Emails');
  const subjectCol = headers.indexOf('Subject');
  const messageCol = headers.indexOf('Message');
  
  // Get selected row (if any)
  const activeRange = sheet.getActiveRange();
  const selectedRow = activeRange ? activeRange.getRow() : null;
  
  let testRowIndex = null;
  let testSource = '';
  
  // Check if a specific row is selected (not header row)
  if (selectedRow && selectedRow > 1 && selectedRow <= data.length) {
    const rowData = data[selectedRow - 1];
    if (rowData[nameCol]) { // Has a communication name
      testRowIndex = selectedRow - 1;
      testSource = `selected row (${rowData[nameCol]})`;
    }
  }
  
  // If no valid selection, find first active communication
  if (testRowIndex === null) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[activeCol] && row[nameCol]) {
        testRowIndex = i;
        testSource = `first active communication (${row[nameCol]})`;
        break;
      }
    }
  }
  
  if (testRowIndex === null) {
    SpreadsheetApp.getUi().alert('No communications to test', 'Please select a row with a communication, or make sure you have at least one active communication.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const row = data[testRowIndex];
  const now = new Date();
  
  // Check if communication has all required fields
  if (!row[fromCol] || !row[subjectCol] || !row[messageCol]) {
    SpreadsheetApp.getUi().alert('❌ Incomplete Communication', `The communication "${row[nameCol]}" is missing required fields (From Email, Subject, or Message).`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Ask for test recipient email(s)
  const ui = SpreadsheetApp.getUi();
  const recipientResponse = ui.prompt(
    '📧 Test Email Recipients',
    `Enter email address(es) to send the test to:\n(Separate multiple emails with commas)\n\n` +
    `Original recipients: ${row[toCol] || 'None specified'}\n\n` +
    `💡 This test will NOT send to the original recipients - only to the email(s) you specify below.`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (recipientResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const testRecipients = recipientResponse.getResponseText().trim();
  if (!testRecipients) {
    ui.alert('❌ No Recipients', 'You must enter at least one email address to send the test to.', ui.ButtonSet.OK);
    return;
  }
  
  // Validate email addresses
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const emails = testRecipients.split(',').map(email => email.trim());
  const invalidEmails = emails.filter(email => !emailRegex.test(email));
  
  if (invalidEmails.length > 0) {
    ui.alert('❌ Invalid Email Address', `Invalid email address(es): ${invalidEmails.join(', ')}\n\nPlease enter valid email addresses.`, ui.ButtonSet.OK);
    return;
  }
  
  try {
    const fromEmail = row[fromCol];
    
    // Use the EXACT same subject as would be sent (with placeholders replaced)
    const subject = replacePlaceholders(row[subjectCol], now);
    
    // Use the EXACT same message formatting as would be sent
    const originalMessage = replacePlaceholders(row[messageCol], now);
    
    // Add test notice at the bottom with proper HTML formatting
    const testMessage = originalMessage + 
      '<br><br><div style="border-top: 2px solid #ccc; margin-top: 20px; padding-top: 15px; color: #666; font-style: italic;">' +
      '🧪 <strong>This is a test email</strong> - This message was sent using the "Test Send Now" feature. ' +
      'The actual scheduled email will look exactly like this but without this notice.<br><br>' +
      `<strong>Original recipients:</strong> ${row[toCol] || 'None specified'}` +
      '</div>';
    
    const recipients = emails.join(',');
    
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: testMessage,
      replyTo: fromEmail
    });
    
    console.log(`Test sent: "${row[nameCol]}" from ${testSource} to ${recipients}`);
    SpreadsheetApp.getUi().alert('✅ Test Sent!', `Test email sent successfully for "${row[nameCol]}" (${testSource}).\n\nSent to: ${recipients}\n\nThe email was sent exactly as it will appear when scheduled, with all formatting and placeholders properly replaced.`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error(`Error sending test for "${row[nameCol]}":`, error);
    SpreadsheetApp.getUi().alert('❌ Test Failed', `Error sending test for "${row[nameCol]}": ${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function simulateHourlyCheck() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask user what time to simulate
  const response = ui.prompt(
    '🕐 Simulate Hourly Check',
    'Enter the time to simulate (e.g., "2:00 PM", "14:00", or just "14"):\n\n' +
    'This will run the exact same logic as the hourly trigger, pretending it\'s that time right now.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const timeInput = response.getResponseText().trim().toUpperCase();
  let simulatedHour = -1;
  
  // Parse different time formats
  if (timeInput.includes('AM') || timeInput.includes('PM')) {
    // Handle 12-hour format like "2:00 PM"
    simulatedHour = parseTimeString(timeInput);
  } else if (timeInput.includes(':')) {
    // Handle 24-hour format like "14:00"
    const parts = timeInput.split(':');
    simulatedHour = parseInt(parts[0]);
  } else {
    // Handle just hour like "14"
    simulatedHour = parseInt(timeInput);
  }
  
  if (simulatedHour < 0 || simulatedHour > 23) {
    ui.alert('❌ Invalid Time', 'Please enter a valid time between 0-23 (24-hour) or use AM/PM format.', ui.ButtonSet.OK);
    return;
  }
  
  // Run the simulation
  console.log(`Simulating hourly check for hour: ${simulatedHour}`);
  
  const results = simulateHourlyCheckAtTime(simulatedHour);
  
  // Show results
  let message = `🕐 Simulated hourly check for ${simulatedHour}:00\n\n`;
  
  if (results.sent.length > 0) {
    message += `✅ WOULD SEND (${results.sent.length}):\n`;
    results.sent.forEach(comm => {
      message += `• "${comm.name}" to ${comm.recipients}\n`;
    });
    message += '\n';
  }
  
  if (results.skipped.length > 0) {
    message += `⏭️ WOULD SKIP (${results.skipped.length}):\n`;
    results.skipped.forEach(item => {
      message += `• "${item.name}": ${item.reason}\n`;
    });
    message += '\n';
  }
  
  if (results.sent.length === 0 && results.skipped.length === 0) {
    message += 'No communications found or configured.\n\n';
  }
  
  message += '💡 This was a simulation - no actual emails were sent.';
  
  ui.alert('🕐 Hourly Check Simulation Results', message, ui.ButtonSet.OK);
}

function simulateHourlyCheckAtTime(simulatedHour) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    return { sent: [], skipped: [] };
  }
  
  const headers = data[0];
  const nameCol = headers.indexOf('Communication Name');
  const timeCol = headers.indexOf('Send Time');
  const activeCol = headers.indexOf('Active');
  const fromCol = headers.indexOf('From Email');
  const toCol = headers.indexOf('To Emails');
  const daysCol = headers.indexOf('Send Days');
  const sentTodayCol = headers.indexOf('Sent Today');
  
  // Get current date info but override the hour
  const now = new Date();
  const currentDay = now.getDay();
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const currentDayName = dayNames[currentDay];
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const sent = [];
  const skipped = [];
  
  // Process each row (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const commName = row[nameCol] || `Row ${i + 1}`;
    
    // Skip if not active
    if (!row[activeCol]) {
      skipped.push({ name: commName, reason: 'Not active' });
      continue;
    }
    
    // Check if today is a send day
    const sendDays = row[daysCol] ? row[daysCol].toString() : '';
    if (sendDays && !sendDays.includes(currentDayName)) {
      skipped.push({ name: commName, reason: `Not scheduled for ${currentDayName}` });
      continue;
    }
    
    // If no send days specified, use old weekend logic (Mon-Fri only)
    if (!sendDays && (currentDay === 0 || currentDay === 6)) {
      skipped.push({ name: commName, reason: 'Weekend (no specific days configured)' });
      continue;
    }
    
    // Parse send time
    const sendTimeStr = row[timeCol];
    if (!sendTimeStr) {
      skipped.push({ name: commName, reason: 'No send time specified' });
      continue;
    }
    
    const sendHour = parseTimeString(sendTimeStr);
    console.log(`Debug: "${commName}" - Time string: "${sendTimeStr}" -> Parsed hour: ${sendHour}`);
    
    if (sendHour === -1 || isNaN(sendHour)) {
      skipped.push({ name: commName, reason: `Invalid time format: "${sendTimeStr}" -> ${sendHour}` });
      continue;
    }
    
    // Check if it's time to send (using simulated hour)
    if (simulatedHour !== sendHour) {
      skipped.push({ name: commName, reason: `Wrong time (scheduled for ${sendHour}:00, simulating ${simulatedHour}:00)` });
      continue;
    }
    
    // Check if already sent today
    const sentToday = row[sentTodayCol];
    if (sentToday === today) {
      skipped.push({ name: commName, reason: 'Already sent today' });
      continue;
    }
    
    // Would send!
    const toEmails = row[toCol];
    const recipients = toEmails ? toEmails.split(',').map(email => email.trim()).join(', ') : 'No recipients';
    sent.push({ name: commName, recipients: recipients });
  }
  
  return { sent, skipped };
}

// Create custom menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Daily Announcements')
    .addItem('📖 Instructions', 'showInstructions')
    .addSeparator()
    .addItem('➕ Add New Communication', 'showAddCommunicationDialog')
    .addItem('✏️ Manage Communications', 'showManageCommunicationsDialog')
    .addSeparator()
    .addItem('🔑 Setup & Authorize', 'setupAndAuthorize')
    .addItem('🧪 Test Send Now', 'testAnnouncements')
    .addItem('📤 Manual Send', 'showManualSendDialog')
    .addItem('🕐 Simulate Hourly Check', 'simulateHourlyCheck')
    .addSeparator()
    .addItem('📊 Check Status', 'checkStatus')
    .addItem('⏰ View Trigger Info', 'viewTriggerInfo')
    .addToUi();
}

function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  
  const instructions = `📧 DAILY ANNOUNCEMENTS SYSTEM
How to Use This Automated Email System

🚀 GETTING STARTED:
1. Click "Setup & Authorize" to give permission and activate the system
2. Click "Add New Communication" to create your first announcement
3. The system will automatically send emails every day at your chosen time!

✍️ CREATING ANNOUNCEMENTS:
• Communication Name: Give it a descriptive name (e.g., "Daily Lunch Menu")
• Send Time: Choose from dropdown (6:00 AM - 11:00 PM, hourly intervals only)
• Days to Send: Select which days of the week (Mon-Sun)
• From/To Emails: Set sender and recipients (comma-separate multiple emails)
• Subject: Use [Date], [Day], or [ShortDate] for automatic date insertion
• Message: Use the Gmail-like rich text editor with formatting buttons!

✨ MESSAGE FORMATTING FEATURES:
• Bold, Italic, Underline with buttons or Ctrl+B, Ctrl+I, Ctrl+U
• Bullet lists and numbered lists
• Easy link insertion - highlight text and click 🔗 Link button
• Just type URLs and they become clickable automatically!

📅 DATE PLACEHOLDERS:
• [Date] → December 5, 2024
• [Day] → Thursday  
• [ShortDate] → 12/5/2024

🧪 TESTING & MONITORING:
• "Test Send Now" - Sends a perfect preview of your email immediately
• "Check Status" - See which announcements sent today
• "View Trigger Info" - Confirm the system is running automatically

⚙️ HOW IT WORKS:
The system checks every hour for scheduled emails. When it's the right time and day, it automatically sends your announcements with all formatting and date placeholders properly filled in.

Need help? The system will show error messages if something needs attention!`;

  ui.alert('📖 Instructions', instructions, ui.ButtonSet.OK);
}

function setupAndAuthorize() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // This will trigger the authorization prompt
    const testEmail = Session.getActiveUser().getEmail();
    
    // Set up the trigger
    setupTrigger();
    
    // Show success message
    ui.alert(
      '✅ Setup Complete!', 
      `Authorization successful!\n\n` +
      `• Trigger created (runs every hour)\n` +
      `• Emails will be sent from: ${testEmail}\n` +
      `• Weekend emails are automatically skipped\n\n` +
      `You can now test with "Test Send Now" or wait for the scheduled time.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert(
      '❌ Setup Error', 
      `There was an issue during setup:\n\n${error.toString()}\n\nPlease try again or contact your IT administrator.`,
      ui.ButtonSet.OK
    );
  }
}

function checkStatus() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    ui.alert('No announcements configured yet!', 'Please add your announcement details to the sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  const headers = data[0];
  const sentTodayCol = headers.indexOf('Sent Today');
  const activeCol = headers.indexOf('Active');
  const nameCol = headers.indexOf('Communication Name');
  
  let statusMsg = 'Current Status:\n\n';
  let activeCount = 0;
  let sentTodayCount = 0;
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol] || `Row ${i + 1}`;
    const active = row[activeCol];
    const sentToday = row[sentTodayCol];
    
    if (active) {
      activeCount++;
      const sent = (sentToday === today) ? '✅ Sent' : '⏳ Pending';
      statusMsg += `• ${name}: ${sent}\n`;
      if (sentToday === today) sentTodayCount++;
    } else {
      statusMsg += `• ${name}: ⏸️ Inactive\n`;
    }
  }
  
  statusMsg += `\n📊 Summary: ${sentTodayCount}/${activeCount} sent today`;
  
  ui.alert('📊 Announcement Status', statusMsg, ui.ButtonSet.OK);
}

function viewTriggerInfo() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  const announcementTriggers = triggers.filter(t => t.getHandlerFunction() === 'sendDailyAnnouncements');
  
  if (announcementTriggers.length === 0) {
    ui.alert('❌ No Trigger Found', 'No trigger is currently set up. Use "Setup & Authorize" to create one.', ui.ButtonSet.OK);
    return;
  }
  
  const trigger = announcementTriggers[0];
  const triggerSource = trigger.getTriggerSource();
  const eventType = trigger.getEventType();
  
  ui.alert(
    '⏰ Trigger Information', 
    `Trigger Status: ✅ Active\n` +
    `Type: Time-based trigger\n` +
    `Runs: Every hour\n` +
    `Function: sendDailyAnnouncements\n\n` +
    `The script automatically checks every hour for scheduled announcements.`,
    ui.ButtonSet.OK
  );
}

function showManualSendDialog() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    ui.alert('No Communications Found', 'Please add a communication first using "Add New Communication".', ui.ButtonSet.OK);
    return;
  }
  
  const headers = data[0];
  const nameCol = headers.indexOf('Communication Name');
  const activeCol = headers.indexOf('Active');
  const timeCol = headers.indexOf('Send Time');
  const daysCol = headers.indexOf('Send Days');
  
  if (nameCol === -1) {
    ui.alert('Invalid Sheet', 'Sheet headers not found. Please use "Setup & Authorize" first.', ui.ButtonSet.OK);
    return;
  }
  
  // Build list of communications
  let commList = '';
  let validComms = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol] || `Communication ${i}`;
    const active = row[activeCol] ? '✅ Active' : '⏸️ Inactive';
    const time = row[timeCol] || 'No time set';
    const days = row[daysCol] || 'No days set';
    const rowNum = i + 1; // 1-based row number
    
    commList += `${rowNum}. ${name}\n   Status: ${active} | Time: ${time} | Days: ${days}\n\n`;
    validComms.push(rowNum);
  }
  
  const response = ui.prompt(
    '📤 Manual Send',
    `Select a communication to send immediately:\n\n${commList}Enter the row number (e.g., 2, 3, 4...):\n\n⚠️ This will send the email RIGHT NOW to all configured recipients, regardless of the scheduled time or day settings.`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedRow = parseInt(response.getResponseText().trim());
  
  if (!selectedRow || !validComms.includes(selectedRow)) {
    ui.alert('❌ Invalid Selection', `Please enter a valid row number from: ${validComms.join(', ')}`, ui.ButtonSet.OK);
    return;
  }
  
  // Confirm before sending
  const rowData = data[selectedRow - 1];
  const commName = rowData[nameCol] || `Communication ${selectedRow}`;
  const confirmResponse = ui.alert(
    '🚨 Confirm Manual Send',
    `Are you sure you want to send "${commName}" RIGHT NOW?\n\nThis will send the email immediately to all configured recipients with today's date placeholders filled in.`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) return;
  
  // Send the communication
  sendCommunicationManually(selectedRow);
}

function showAddCommunicationDialog() {
  const htmlTemplate = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .form-group { margin-bottom: 15px; }
          label { display: block; margin-bottom: 5px; font-weight: bold; }
          input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
          textarea { height: 80px; resize: vertical; }
          .checkbox-group { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 5px; }
          .checkbox-item { display: flex; align-items: center; }
          .checkbox-item input { width: auto; margin-right: 5px; }
          .buttons { margin-top: 20px; text-align: right; }
          .btn { padding: 10px 20px; margin-left: 10px; border: none; border-radius: 4px; cursor: pointer; }
          .btn-primary { background-color: #4285f4; color: white; }
          .btn-secondary { background-color: #f8f9fa; color: #333; border: 1px solid #ddd; }
          .help-text { font-size: 12px; color: #666; margin-top: 3px; }
          .editor-toolbar { 
            display: flex; 
            align-items: center; 
            gap: 5px; 
            padding: 8px; 
            background: #f8f9fa; 
            border: 1px solid #ddd; 
            border-bottom: none; 
            border-radius: 4px 4px 0 0; 
            flex-wrap: wrap;
          }
          .toolbar-btn { 
            padding: 6px 10px; 
            border: 1px solid #ccc; 
            background: white; 
            border-radius: 3px; 
            cursor: pointer; 
            font-size: 12px;
            min-width: 32px;
            height: 28px;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .toolbar-btn:hover { background: #e8f0fe; }
          .toolbar-btn:active { background: #d2e3fc; }
          .toolbar-separator { color: #ccc; margin: 0 5px; }
          .rich-editor { 
            min-height: 120px; 
            max-height: 300px; 
            overflow-y: auto; 
            padding: 12px; 
            border: 1px solid #ddd; 
            border-radius: 0 0 4px 4px; 
            background: white;
            outline: none;
            font-family: Arial, sans-serif;
            line-height: 1.4;
          }
          .rich-editor:focus { border-color: #4285f4; }
          .rich-editor[contenteditable]:empty::before {
            content: attr(placeholder);
            color: #999;
            font-style: italic;
          }
          .rich-editor p { margin: 0 0 10px 0; }
          .rich-editor ul, .rich-editor ol { margin: 10px 0; padding-left: 20px; }
          .rich-editor a { color: #1a73e8; text-decoration: underline; }
        </style>
      </head>
      <body>
        <h3>📧 Add New Communication</h3>
        <form>
          <div class="form-group">
            <label for="commName">Communication Name:</label>
            <input type="text" id="commName" placeholder="e.g., Daily Announcements, Lunch Menu" required>
          </div>
          
          <div class="form-group">
            <label for="sendTime">Send Time:</label>
            <select id="sendTime" required>
              <option value="">Select a time...</option>
              <option value="6:00 AM">6:00 AM</option>
              <option value="7:00 AM">7:00 AM</option>
              <option value="8:00 AM" selected>8:00 AM</option>
              <option value="9:00 AM">9:00 AM</option>
              <option value="10:00 AM">10:00 AM</option>
              <option value="11:00 AM">11:00 AM</option>
              <option value="12:00 PM">12:00 PM</option>
              <option value="1:00 PM">1:00 PM</option>
              <option value="2:00 PM">2:00 PM</option>
              <option value="3:00 PM">3:00 PM</option>
              <option value="4:00 PM">4:00 PM</option>
              <option value="5:00 PM">5:00 PM</option>
              <option value="6:00 PM">6:00 PM</option>
              <option value="7:00 PM">7:00 PM</option>
              <option value="8:00 PM">8:00 PM</option>
              <option value="9:00 PM">9:00 PM</option>
              <option value="10:00 PM">10:00 PM</option>
              <option value="11:00 PM">11:00 PM</option>
            </select>
            <div class="help-text">Choose when to send the announcement each day (system checks hourly on the hour)</div>
          </div>
          
          <div class="form-group">
            <label for="sendDays">Days to Send:</label>
            <div class="checkbox-group">
              <div class="checkbox-item">
                <input type="checkbox" id="monday" checked>
                <label for="monday">Monday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="tuesday" checked>
                <label for="tuesday">Tuesday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="wednesday" checked>
                <label for="wednesday">Wednesday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="thursday" checked>
                <label for="thursday">Thursday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="friday" checked>
                <label for="friday">Friday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="saturday">
                <label for="saturday">Saturday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="sunday">
                <label for="sunday">Sunday</label>
              </div>
            </div>
          </div>
          
          <div class="form-group">
            <label for="fromEmail">From Email:</label>
            <input type="email" id="fromEmail" value="<?= userEmail ?>" required>
            <div class="help-text">This sets the reply-to address</div>
          </div>
          
          <div class="form-group">
            <label for="toEmails">To Emails:</label>
            <input type="text" id="toEmails" placeholder="email1@school.edu, email2@school.edu" required>
            <div class="help-text">Separate multiple emails with commas</div>
          </div>
          
          <div class="form-group">
            <label for="subject">Subject Line:</label>
            <input type="text" id="subject" placeholder="Daily Announcements - [Date]" required>
            <div class="help-text">Use [Date], [Day], or [ShortDate] for automatic dates</div>
          </div>
          
          <div class="form-group">
            <label for="message">Message:</label>
            <div class="editor-toolbar">
              <button type="button" class="toolbar-btn" onclick="formatText('bold')" title="Bold"><b>B</b></button>
              <button type="button" class="toolbar-btn" onclick="formatText('italic')" title="Italic"><i>I</i></button>
              <button type="button" class="toolbar-btn" onclick="formatText('underline')" title="Underline"><u>U</u></button>
              <span class="toolbar-separator">|</span>
              <button type="button" class="toolbar-btn" onclick="formatText('insertUnorderedList')" title="Bullet List">• List</button>
              <button type="button" class="toolbar-btn" onclick="formatText('insertOrderedList')" title="Numbered List">1. List</button>
              <span class="toolbar-separator">|</span>
              <button type="button" class="toolbar-btn" onclick="showLinkDialog()" title="Insert Link">🔗 Link</button>
              <button type="button" class="toolbar-btn" onclick="cleanUpMessage()" title="Clean Up Messy Content">🧹 Clean</button>
              <button type="button" class="toolbar-btn" onclick="formatText('removeFormat')" title="Remove Formatting">Clear</button>
            </div>
            <div id="message" class="rich-editor" contenteditable="true" 
                 placeholder="Good morning everyone! Here are today's announcements...">
            </div>
            <div class="help-text">Format your message just like in Gmail! Highlight text and use the buttons above, or just type URLs and they'll become clickable automatically.<br>
            💡 <strong>Keyboard shortcuts:</strong> Ctrl+B (bold), Ctrl+I (italic), Ctrl+U (underline), Ctrl+K (link)</div>
          </div>
          
          <div class="buttons">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="button" class="btn btn-primary" onclick="addCommunication(); return false;">Add Communication</button>
          </div>
        </form>
        
        <script>
          
          // Rich text editor functions
          function formatText(command, value) {
            document.execCommand(command, false, value);
            document.getElementById('message').focus();
          }
          
          function showLinkDialog() {
            var selection = window.getSelection();
            var selectedText = selection.toString().trim();
            
            var url = prompt('Enter the website address (e.g., lunch.school.edu):', '');
            if (!url) return;
            
            // Add https if not present
            if (url.indexOf('http://') !== 0 && url.indexOf('https://') !== 0) {
              url = 'https://' + url;
            }
            
            var linkText = selectedText || prompt('What should the link say?', 'click here');
            if (!linkText) return;
            
            if (selectedText) {
              // Replace selected text with link
              document.execCommand('createLink', false, url);
            } else {
              // Insert new link at cursor
              var link = '<a href="' + url + '">' + linkText + '</a>';
              document.execCommand('insertHTML', false, link);
            }
            
            document.getElementById('message').focus();
          }
          
          // Simplified - no auto URL conversion for now
          function setupAutoLinking() {
            // Disabled auto-linking to fix JavaScript issues
            console.log('Auto-linking disabled');
          }
          
          // Keyboard shortcuts
          function setupKeyboardShortcuts() {
            var editor = document.getElementById('message');
            
            editor.addEventListener('keydown', function(e) {
              // Ctrl+B for bold
              if (e.ctrlKey && e.key === 'b') {
                e.preventDefault();
                formatText('bold');
              }
              // Ctrl+I for italic
              else if (e.ctrlKey && e.key === 'i') {
                e.preventDefault();
                formatText('italic');
              }
              // Ctrl+U for underline
              else if (e.ctrlKey && e.key === 'u') {
                e.preventDefault();
                formatText('underline');
              }
              // Ctrl+K for link
              else if (e.ctrlKey && e.key === 'k') {
                e.preventDefault();
                showLinkDialog();
              }
              // Enter key - create proper line breaks
              else if (e.key === 'Enter') {
                if (!e.shiftKey) {
                  document.execCommand('insertHTML', false, '<br><br>');
                  e.preventDefault();
                }
              }
            });
          }
          
          // Clean up messy HTML content
          function cleanUpMessage() {
            var editor = document.getElementById('message');
            var content = editor.innerHTML;
            
            // Simple cleanup - remove common messy elements
            content = content.replace(/gmail_quote/g, '');
            content = content.replace(/style="[^"]*"/g, '');
            content = content.replace(/class="[^"]*"/g, '');
            content = content.replace(/dir="[^"]*"/g, '');
            
            // Clean up multiple line breaks
            content = content.replace(/<br><br><br>/g, '<br><br>');
            
            editor.innerHTML = content;
            editor.focus();
            
            alert('Message cleaned up! Removed messy formatting.');
          }
          
          // Initialize auto-linking and keyboard shortcuts when page loads
          window.onload = function() {
            setupAutoLinking();
            setupKeyboardShortcuts();
          };
          
          function addCommunication() {
            try {
              var messageEditor = document.getElementById('message');
              var messageContent = messageEditor ? messageEditor.innerHTML.trim() : '';
              
              // Simplified content check
              if (!messageContent || messageContent === '<br>' || messageContent === '<div><br></div>') {
                messageContent = '';
              }
              var hasRealContent = messageContent.length > 0;
            
            var data = {
              name: document.getElementById('commName').value,
              time: document.getElementById('sendTime').value,
              fromEmail: document.getElementById('fromEmail').value,
              toEmails: document.getElementById('toEmails').value,
              subject: document.getElementById('subject').value,
              message: messageContent,
              days: {
                monday: document.getElementById('monday').checked,
                tuesday: document.getElementById('tuesday').checked,
                wednesday: document.getElementById('wednesday').checked,
                thursday: document.getElementById('thursday').checked,
                friday: document.getElementById('friday').checked,
                saturday: document.getElementById('saturday').checked,
                sunday: document.getElementById('sunday').checked
              }
            };
            
            console.log('Form data:', data); // Debug
            
            if (!data.name || !data.time || !data.fromEmail || !data.toEmails || !data.subject || !hasRealContent) {
              alert('Please fill in all required fields.');
              return;
            }
            
            var selectedDays = false;
            for (var day in data.days) {
              if (data.days[day]) {
                selectedDays = true;
                break;
              }
            }
            
            if (!selectedDays) {
              alert('Please select at least one day to send the communication.');
              return;
            }
            
            var button = event.target;
            button.disabled = true;
            button.textContent = 'Adding...';
            
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .addNewCommunication(data);
              
          } catch(error) {
            alert('JavaScript Error: ' + error.toString());
            console.error('Add Communication Error:', error);
          }
        }
          
          function onSuccess(result) {
            if (result.success) {
              alert('Communication added successfully! The system is ready to send emails automatically.');
              google.script.host.close();
            } else {
              alert('Error: ' + result.error);
              var button = document.querySelector('.btn-primary');
              button.disabled = false;
              button.textContent = 'Add Communication';
            }
          }
          
          function onFailure(error) {
            alert('Error adding communication: ' + error.toString());
            var button = document.querySelector('.btn-primary');
            button.disabled = false;
            button.textContent = 'Add Communication';
          }
        </script>
      </body>
    </html>
  `);
  
  htmlTemplate.userEmail = Session.getActiveUser().getEmail();
  
  const html = htmlTemplate.evaluate()
    .setWidth(550)
    .setHeight(700);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Communication');
}

function showManageCommunicationsDialog() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    ui.alert('No Communications Found', 'Please add a communication first using "Add New Communication".', ui.ButtonSet.OK);
    return;
  }
  
  const headers = data[0];
  const nameCol = headers.indexOf('Communication Name');
  const activeCol = headers.indexOf('Active');
  
  if (nameCol === -1) {
    ui.alert('Invalid Sheet', 'Sheet headers not found. Please use "Setup & Authorize" first.', ui.ButtonSet.OK);
    return;
  }
  
  // Build list of communications
  let commList = '';
  let validComms = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol] || `Communication ${i}`;
    const active = row[activeCol] ? '✅' : '❌';
    const rowNum = i + 1; // 1-based row number
    
    commList += `${rowNum}. ${active} ${name}\n`;
    validComms.push(rowNum);
  }
  
  const response = ui.prompt(
    '✏️ Manage Communications',
    `Select a communication to edit:\n\n${commList}\nEnter the row number (e.g., 2, 3, 4...):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedRow = parseInt(response.getResponseText().trim());
  
  if (!selectedRow || !validComms.includes(selectedRow)) {
    ui.alert('❌ Invalid Selection', `Please enter a valid row number from: ${validComms.join(', ')}`, ui.ButtonSet.OK);
    return;
  }
  
  // Show edit dialog for selected communication
  console.log(`Calling showEditCommunicationDialog with selectedRow: ${selectedRow}`);
  showEditCommunicationDialog(selectedRow);
}

function showEditCommunicationDialog(rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowData = data[rowNumber - 1]; // Convert to 0-based index
  
  // Debug logging
  console.log(`Edit dialog debug - Row ${rowNumber}:`);
  console.log(`Headers: ${JSON.stringify(headers)}`);
  console.log(`Row data: ${JSON.stringify(rowData)}`);
  
  // Check if headers are correct - if not, show error and suggest running setup
  const expectedHeaders = ['Communication Name', 'Send Time', 'Active', 'From Email', 'To Emails', 'Subject', 'Message', 'Send Days', 'Sent Today'];
  const hasCorrectHeaders = expectedHeaders.every(header => headers.includes(header));
  
  if (!hasCorrectHeaders) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('❌ Invalid Headers', 
      `Your spreadsheet headers are incorrect or missing columns.\n\nExpected: ${expectedHeaders.join(', ')}\nFound: ${headers.join(', ')}\n\nPlease run "Setup & Authorize" to fix the headers.`, 
      ui.ButtonSet.OK);
    return;
  }
  
  // Get column indices
  const nameCol = headers.indexOf('Communication Name');
  const timeCol = headers.indexOf('Send Time');
  const activeCol = headers.indexOf('Active');
  const fromCol = headers.indexOf('From Email');
  const toCol = headers.indexOf('To Emails');
  const subjectCol = headers.indexOf('Subject');
  const messageCol = headers.indexOf('Message');
  const daysCol = headers.indexOf('Send Days');
  
  console.log(`Column indices: name=${nameCol}, time=${timeCol}, active=${activeCol}, from=${fromCol}, to=${toCol}, subject=${subjectCol}, message=${messageCol}, days=${daysCol}`);
  
  // Parse existing days
  const existingDays = rowData[daysCol] ? rowData[daysCol].toString() : '';
  const daysChecked = {
    monday: existingDays.includes('Mon'),
    tuesday: existingDays.includes('Tue'),
    wednesday: existingDays.includes('Wed'),
    thursday: existingDays.includes('Thu'),
    friday: existingDays.includes('Fri'),
    saturday: existingDays.includes('Sat'),
    sunday: existingDays.includes('Sun')
  };
  
  // Get actual values for template substitution
  const commName = rowData[nameCol] || '';
  const sendTime = rowData[timeCol] || '';
  const isActive = rowData[activeCol] || false;
  const fromEmail = rowData[fromCol] || '';
  const toEmails = rowData[toCol] || '';
  const subject = rowData[subjectCol] || '';
  const message = rowData[messageCol] || '';
  
  // Debug actual values
  console.log('Actual values from spreadsheet:');
  console.log(`- commName: "${commName}"`);
  console.log(`- sendTime: "${sendTime}"`);
  console.log(`- isActive: ${isActive}`);
  console.log(`- message: "${message}"`);
  console.log(`- days: ${JSON.stringify(daysChecked)}`);
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .form-group { margin-bottom: 15px; }
          label { display: block; margin-bottom: 5px; font-weight: bold; }
          input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
          .checkbox-group { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 5px; }
          .checkbox-item { display: flex; align-items: center; }
          .checkbox-item input { width: auto; margin-right: 5px; }
          .buttons { margin-top: 20px; text-align: right; }
          .btn { padding: 10px 20px; margin-left: 10px; border: none; border-radius: 4px; cursor: pointer; }
          .btn-primary { background-color: #4285f4; color: white; }
          .btn-danger { background-color: #ea4335; color: white; }
          .btn-secondary { background-color: #f8f9fa; color: #333; border: 1px solid #ddd; }
          .help-text { font-size: 12px; color: #666; margin-top: 3px; }
          .editor-toolbar { 
            display: flex; 
            align-items: center; 
            gap: 5px; 
            padding: 8px; 
            background: #f8f9fa; 
            border: 1px solid #ddd; 
            border-bottom: none; 
            border-radius: 4px 4px 0 0; 
            flex-wrap: wrap;
          }
          .toolbar-btn { 
            padding: 6px 10px; 
            border: 1px solid #ccc; 
            background: white; 
            border-radius: 3px; 
            cursor: pointer; 
            font-size: 12px;
            min-width: 32px;
            height: 28px;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .toolbar-btn:hover { background: #e8f0fe; }
          .toolbar-btn:active { background: #d2e3fc; }
          .toolbar-separator { color: #ccc; margin: 0 5px; }
          .rich-editor { 
            min-height: 120px; 
            max-height: 300px; 
            overflow-y: auto; 
            padding: 12px; 
            border: 1px solid #ddd; 
            border-radius: 0 0 4px 4px; 
            background: white;
            outline: none;
            font-family: Arial, sans-serif;
            line-height: 1.4;
          }
          .rich-editor:focus { border-color: #4285f4; }
          .rich-editor p { margin: 0 0 10px 0; }
          .rich-editor ul, .rich-editor ol { margin: 10px 0; padding-left: 20px; }
          .rich-editor a { color: #1a73e8; text-decoration: underline; }
        </style>
      </head>
      <body>
        <h3>✏️ Edit Communication: ${commName}</h3>
        <form>
          <div class="form-group">
            <label for="commName">Communication Name:</label>
            <input type="text" id="commName" value="${commName}" required>
          </div>
          
          <div class="form-group">
            <label for="sendTime">Send Time:</label>
            <select id="sendTime" required>
              <option value="">Select a time...</option>
              <option value="6:00 AM"${sendTime === '6:00 AM' ? ' selected' : ''}>6:00 AM</option>
              <option value="7:00 AM"${sendTime === '7:00 AM' ? ' selected' : ''}>7:00 AM</option>
              <option value="8:00 AM"${sendTime === '8:00 AM' ? ' selected' : ''}>8:00 AM</option>
              <option value="9:00 AM"${sendTime === '9:00 AM' ? ' selected' : ''}>9:00 AM</option>
              <option value="10:00 AM"${sendTime === '10:00 AM' ? ' selected' : ''}>10:00 AM</option>
              <option value="11:00 AM"${sendTime === '11:00 AM' ? ' selected' : ''}>11:00 AM</option>
              <option value="12:00 PM"${sendTime === '12:00 PM' ? ' selected' : ''}>12:00 PM</option>
              <option value="1:00 PM"${sendTime === '1:00 PM' ? ' selected' : ''}>1:00 PM</option>
              <option value="2:00 PM"${sendTime === '2:00 PM' ? ' selected' : ''}>2:00 PM</option>
              <option value="3:00 PM"${sendTime === '3:00 PM' ? ' selected' : ''}>3:00 PM</option>
              <option value="4:00 PM"${sendTime === '4:00 PM' ? ' selected' : ''}>4:00 PM</option>
              <option value="5:00 PM"${sendTime === '5:00 PM' ? ' selected' : ''}>5:00 PM</option>
              <option value="6:00 PM"${sendTime === '6:00 PM' ? ' selected' : ''}>6:00 PM</option>
              <option value="7:00 PM"${sendTime === '7:00 PM' ? ' selected' : ''}>7:00 PM</option>
              <option value="8:00 PM"${sendTime === '8:00 PM' ? ' selected' : ''}>8:00 PM</option>
              <option value="9:00 PM"${sendTime === '9:00 PM' ? ' selected' : ''}>9:00 PM</option>
              <option value="10:00 PM"${sendTime === '10:00 PM' ? ' selected' : ''}>10:00 PM</option>
              <option value="11:00 PM"${sendTime === '11:00 PM' ? ' selected' : ''}>11:00 PM</option>
            </select>
            <div class="help-text">Choose when to send the announcement each day (system checks hourly on the hour)</div>
          </div>
          
          <div class="form-group">
            <label>Active:</label>
            <div class="checkbox-item">
              <input type="checkbox" id="active"${isActive ? ' checked' : ''}>
              <label for="active">Communication is active</label>
            </div>
          </div>
          
          <div class="form-group">
            <label for="sendDays">Days to Send:</label>
            <div class="checkbox-group">
              <div class="checkbox-item">
                <input type="checkbox" id="monday"${daysChecked.monday ? ' checked' : ''}>
                <label for="monday">Monday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="tuesday"${daysChecked.tuesday ? ' checked' : ''}>
                <label for="tuesday">Tuesday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="wednesday"${daysChecked.wednesday ? ' checked' : ''}>
                <label for="wednesday">Wednesday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="thursday"${daysChecked.thursday ? ' checked' : ''}>
                <label for="thursday">Thursday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="friday"${daysChecked.friday ? ' checked' : ''}>
                <label for="friday">Friday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="saturday"${daysChecked.saturday ? ' checked' : ''}>
                <label for="saturday">Saturday</label>
              </div>
              <div class="checkbox-item">
                <input type="checkbox" id="sunday"${daysChecked.sunday ? ' checked' : ''}>
                <label for="sunday">Sunday</label>
              </div>
            </div>
          </div>
          
          <div class="form-group">
            <label for="fromEmail">From Email:</label>
            <input type="email" id="fromEmail" value="${fromEmail}" required>
            <div class="help-text">This sets the reply-to address</div>
          </div>
          
          <div class="form-group">
            <label for="toEmails">To Emails:</label>
            <input type="text" id="toEmails" value="${toEmails}" required>
            <div class="help-text">Separate multiple emails with commas</div>
          </div>
          
          <div class="form-group">
            <label for="subject">Subject Line:</label>
            <input type="text" id="subject" value="${subject}" required>
            <div class="help-text">Use [Date], [Day], or [ShortDate] for automatic dates</div>
          </div>
          
          <div class="form-group">
            <label for="message">Message:</label>
            <div class="editor-toolbar">
              <button type="button" class="toolbar-btn" onclick="formatText('bold')" title="Bold"><b>B</b></button>
              <button type="button" class="toolbar-btn" onclick="formatText('italic')" title="Italic"><i>I</i></button>
              <button type="button" class="toolbar-btn" onclick="formatText('underline')" title="Underline"><u>U</u></button>
              <span class="toolbar-separator">|</span>
              <button type="button" class="toolbar-btn" onclick="formatText('insertUnorderedList')" title="Bullet List">• List</button>
              <button type="button" class="toolbar-btn" onclick="formatText('insertOrderedList')" title="Numbered List">1. List</button>
              <span class="toolbar-separator">|</span>
              <button type="button" class="toolbar-btn" onclick="showLinkDialog()" title="Insert Link">🔗 Link</button>
              <button type="button" class="toolbar-btn" onclick="cleanUpMessage()" title="Clean Up Messy Content">🧹 Clean</button>
              <button type="button" class="toolbar-btn" onclick="formatText('removeFormat')" title="Remove Formatting">Clear</button>
            </div>
            <div id="message" class="rich-editor" contenteditable="true">${message}</div>
            <div class="help-text">Format your message just like in Gmail! Highlight text and use the buttons above, or just type URLs and they'll become clickable automatically.<br>
            💡 <strong>Keyboard shortcuts:</strong> Ctrl+B (bold), Ctrl+I (italic), Ctrl+U (underline), Ctrl+K (link)</div>
          </div>
          
          <div class="buttons">
            <button type="button" class="btn btn-danger" onclick="deleteCommunication()">🗑️ Delete</button>
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="button" class="btn btn-primary" onclick="updateCommunication()">💾 Save Changes</button>
          </div>
        </form>
        
        <script>
          var currentRowNumber = ${rowNumber};
          
          // Rich text editor functions (same as add dialog)
          function formatText(command, value) {
            document.execCommand(command, false, value);
            document.getElementById('message').focus();
          }
          
          function showLinkDialog() {
            var selection = window.getSelection();
            var selectedText = selection.toString().trim();
            
            var url = prompt('Enter the website address (e.g., lunch.school.edu):', '');
            if (!url) return;
            
            if (url.indexOf('http://') !== 0 && url.indexOf('https://') !== 0) {
              url = 'https://' + url;
            }
            
            var linkText = selectedText || prompt('What should the link say?', 'click here');
            if (!linkText) return;
            
            if (selectedText) {
              document.execCommand('createLink', false, url);
            } else {
              var link = '<a href="' + url + '">' + linkText + '</a>';
              document.execCommand('insertHTML', false, link);
            }
            
            document.getElementById('message').focus();
          }
          
          function cleanUpMessage() {
            var editor = document.getElementById('message');
            var content = editor.innerHTML;
            
            content = content.replace(/gmail_quote/g, '');
            content = content.replace(/style="[^"]*"/g, '');
            content = content.replace(/class="[^"]*"/g, '');
            content = content.replace(/dir="[^"]*"/g, '');
            content = content.replace(/<br><br><br>/g, '<br><br>');
            
            editor.innerHTML = content;
            editor.focus();
            
            alert('Message cleaned up! Removed messy formatting.');
          }
          
          function updateCommunication() {
            try {
              var messageEditor = document.getElementById('message');
              var messageContent = messageEditor ? messageEditor.innerHTML.trim() : '';
              
              if (!messageContent || messageContent === '<br>' || messageContent === '<div><br></div>') {
                messageContent = '';
              }
              var hasRealContent = messageContent.length > 0;
              
              var data = {
                rowNumber: currentRowNumber,
                name: document.getElementById('commName').value,
                time: document.getElementById('sendTime').value,
                active: document.getElementById('active').checked,
                fromEmail: document.getElementById('fromEmail').value,
                toEmails: document.getElementById('toEmails').value,
                subject: document.getElementById('subject').value,
                message: messageContent,
                days: {
                  monday: document.getElementById('monday').checked,
                  tuesday: document.getElementById('tuesday').checked,
                  wednesday: document.getElementById('wednesday').checked,
                  thursday: document.getElementById('thursday').checked,
                  friday: document.getElementById('friday').checked,
                  saturday: document.getElementById('saturday').checked,
                  sunday: document.getElementById('sunday').checked
                }
              };
              
              if (!data.name || !data.time || !data.fromEmail || !data.toEmails || !data.subject || !hasRealContent) {
                alert('Please fill in all required fields.');
                return;
              }
              
              var selectedDays = false;
              for (var day in data.days) {
                if (data.days[day]) {
                  selectedDays = true;
                  break;
                }
              }
              
              if (!selectedDays) {
                alert('Please select at least one day to send the communication.');
                return;
              }
              
              var button = event.target;
              button.disabled = true;
              button.textContent = 'Saving...';
              
              google.script.run
                .withSuccessHandler(onUpdateSuccess)
                .withFailureHandler(onUpdateFailure)
                .updateExistingCommunication(data);
                
            } catch(error) {
              alert('JavaScript Error: ' + error.toString());
              console.error('Update Communication Error:', error);
            }
          }
          
          function deleteCommunication() {
            if (!confirm('Are you sure you want to delete this communication? This cannot be undone.')) {
              return;
            }
            
            google.script.run
              .withSuccessHandler(onDeleteSuccess)
              .withFailureHandler(onDeleteFailure)
              .deleteExistingCommunication(currentRowNumber);
          }
          
          function onUpdateSuccess(result) {
            if (result.success) {
              alert('✅ Communication updated successfully!');
              google.script.host.close();
            } else {
              alert('Error: ' + result.error);
              var button = document.querySelector('.btn-primary');
              button.disabled = false;
              button.textContent = '💾 Save Changes';
            }
          }
          
          function onUpdateFailure(error) {
            alert('Error updating communication: ' + error.toString());
            var button = document.querySelector('.btn-primary');
            button.disabled = false;
            button.textContent = '💾 Save Changes';
          }
          
          function onDeleteSuccess(result) {
            if (result.success) {
              alert('✅ Communication deleted successfully!');
              google.script.host.close();
            } else {
              alert('Error: ' + result.error);
            }
          }
          
          function onDeleteFailure(error) {
            alert('Error deleting communication: ' + error.toString());
          }
          
          // Initialize when page loads
          window.onload = function() {
            // Auto-focus the name field
            document.getElementById('commName').focus();
          };
        </script>
      </body>
    </html>
  `);
  
  html.setWidth(600);
  html.setHeight(750);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Communication');
}

function updateExistingCommunication(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndex = data.rowNumber - 1; // Convert to 0-based index
    
    // Convert days selection to text
    const selectedDays = [];
    if (data.days.monday) selectedDays.push('Mon');
    if (data.days.tuesday) selectedDays.push('Tue');
    if (data.days.wednesday) selectedDays.push('Wed');
    if (data.days.thursday) selectedDays.push('Thu');
    if (data.days.friday) selectedDays.push('Fri');
    if (data.days.saturday) selectedDays.push('Sat');
    if (data.days.sunday) selectedDays.push('Sun');
    
    const daysText = selectedDays.join(', ');
    
    // Get current data and headers
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    
    // Find column indices
    const nameCol = headers.indexOf('Communication Name');
    const timeCol = headers.indexOf('Send Time');
    const activeCol = headers.indexOf('Active');
    const fromCol = headers.indexOf('From Email');
    const toCol = headers.indexOf('To Emails');
    const subjectCol = headers.indexOf('Subject');
    const messageCol = headers.indexOf('Message');
    const daysCol = headers.indexOf('Send Days');
    
    // Update the row
    const updatedRow = [...allData[rowIndex]]; // Copy existing row
    updatedRow[nameCol] = data.name;
    updatedRow[timeCol] = data.time;
    updatedRow[activeCol] = data.active;
    updatedRow[fromCol] = data.fromEmail;
    updatedRow[toCol] = data.toEmails;
    updatedRow[subjectCol] = data.subject;
    updatedRow[messageCol] = data.message;
    updatedRow[daysCol] = daysText;
    
    // Write the updated row back
    const range = sheet.getRange(data.rowNumber, 1, 1, updatedRow.length);
    range.setValues([updatedRow]);
    
    return { 
      success: true, 
      message: 'Communication updated successfully!' 
    };
    
  } catch (error) {
    console.error('Error updating communication:', error);
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

function deleteExistingCommunication(rowNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Delete the row
    sheet.deleteRow(rowNumber);
    
    return { 
      success: true, 
      message: 'Communication deleted successfully!' 
    };
    
  } catch (error) {
    console.error('Error deleting communication:', error);
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

function cleanUpHTML(html) {
  if (!html) return '';
  
  let cleaned = html.toString();
  
  // Remove Gmail quote divs and nested structures
  cleaned = cleaned.replace(/<div class="gmail_quote"[^>]*>/g, '');
  cleaned = cleaned.replace(/<\/div>/g, '');
  cleaned = cleaned.replace(/<div[^>]*>/g, '');
  
  // Remove excessive inline styles but preserve basic formatting
  cleaned = cleaned.replace(/style="[^"]*"/g, '');
  cleaned = cleaned.replace(/class="[^"]*"/g, '');
  cleaned = cleaned.replace(/dir="[^"]*"/g, '');
  cleaned = cleaned.replace(/data-[^=]*="[^"]*"/g, '');
  
  // Clean up multiple line breaks
  cleaned = cleaned.replace(/<br\s*\/?>\s*<br\s*\/?>/g, '<br><br>');
  cleaned = cleaned.replace(/(<br\s*\/?>){3,}/g, '<br><br>');
  
  // Remove empty elements - use a simpler approach
  cleaned = cleaned.replace(/<[^>]+>\s*<\/[^>]+>/g, '');
  
  // Clean up whitespace
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  
  return cleaned;
}

function addNewCommunication(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Ensure headers exist
    setupSheetHeaders();
    
    // Convert days selection to text
    const selectedDays = [];
    if (data.days.monday) selectedDays.push('Mon');
    if (data.days.tuesday) selectedDays.push('Tue');
    if (data.days.wednesday) selectedDays.push('Wed');
    if (data.days.thursday) selectedDays.push('Thu');
    if (data.days.friday) selectedDays.push('Fri');
    if (data.days.saturday) selectedDays.push('Sat');
    if (data.days.sunday) selectedDays.push('Sun');
    
    const daysText = selectedDays.join(', ');
    
    // Add new row
    const newRow = [
      data.name,
      data.time,
      true, // Active
      data.fromEmail,
      data.toEmails,
      data.subject,
      data.message,
      daysText,
      '' // Sent Today (empty initially)
    ];
    
    sheet.appendRow(newRow);
    
    // Auto-authorize and setup trigger
    try {
      setupTrigger();
      
      return { 
        success: true, 
        message: 'Communication added and system authorized successfully!' 
      };
    } catch (authError) {
      return { 
        success: true, 
        message: 'Communication added! Please use "Setup & Authorize" from the menu to complete setup.' 
      };
    }
    
  } catch (error) {
    console.error('Error adding communication:', error);
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

function setupSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // If sheet is empty or doesn't have proper headers, set them up
  if (data.length === 0 || !data[0].includes('Communication Name')) {
    const headers = [
      'Communication Name',
      'Send Time',
      'Active',
      'From Email',
      'To Emails',
      'Subject',
      'Message',
      'Send Days',
      'Sent Today'
    ];
    
    // Clear existing content and add headers
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f1f3f4');
    
    // Set column widths
    sheet.setColumnWidth(1, 200); // Communication Name
    sheet.setColumnWidth(2, 100); // Send Time
    sheet.setColumnWidth(3, 80);  // Active
    sheet.setColumnWidth(4, 200); // From Email
    sheet.setColumnWidth(5, 250); // To Emails
    sheet.setColumnWidth(6, 250); // Subject
    sheet.setColumnWidth(7, 300); // Message
    sheet.setColumnWidth(8, 150); // Send Days
    sheet.setColumnWidth(9, 100); // Sent Today
  }
}

function sendCommunicationManually(rowNumber) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < rowNumber) {
    ui.alert('❌ Error', 'Invalid row number - communication not found.', ui.ButtonSet.OK);
    return;
  }
  
  const headers = data[0];
  const rowData = data[rowNumber - 1]; // Convert to 0-based index
  
  // Find column indices
  const nameCol = headers.indexOf('Communication Name');
  const fromCol = headers.indexOf('From Email');
  const toCol = headers.indexOf('To Emails');
  const subjectCol = headers.indexOf('Subject');
  const messageCol = headers.indexOf('Message');
  const sentTodayCol = headers.indexOf('Sent Today');
  
  const commName = rowData[nameCol] || `Communication ${rowNumber}`;
  const now = new Date();
  
  // Check if communication has all required fields
  if (!rowData[fromCol] || !rowData[toCol] || !rowData[subjectCol] || !rowData[messageCol]) {
    ui.alert('❌ Incomplete Communication', `The communication "${commName}" is missing required fields (From Email, To Emails, Subject, or Message).`, ui.ButtonSet.OK);
    return;
  }
  
  try {
    const fromEmail = rowData[fromCol];
    const toEmails = rowData[toCol];
    const subject = replacePlaceholders(rowData[subjectCol], now);
    const message = replacePlaceholders(rowData[messageCol], now);
    
    // Split multiple recipients if comma-separated
    const recipients = toEmails.split(',').map(email => email.trim()).join(',');
    
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: message,
      replyTo: fromEmail
    });
    
    // Mark as sent today
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.getRange(rowNumber, sentTodayCol + 1).setValue(today);
    
    console.log(`Manual send: "${commName}" sent successfully to ${recipients}`);
    
    ui.alert(
      '✅ Email Sent Successfully!', 
      `"${commName}" has been sent manually!\n\n` +
      `To: ${recipients}\n` +
      `Subject: ${subject}\n\n` +
      `The communication has been marked as sent today and will not be sent again automatically until tomorrow.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error(`Error in manual send for "${commName}":`, error);
    ui.alert('❌ Send Failed', `Error sending "${commName}":\n\n${error.toString()}`, ui.ButtonSet.OK);
  }
}
