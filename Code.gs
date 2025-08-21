// Complete Sales Pipeline Tracker Code
function createLeadForm() {
  // Create or get the form
  let form = FormApp.create('Lead Capture Form');
  
  // Add questions
  form.addTextItem().setTitle('Company Name').setRequired(true);
  form.addTextItem().setTitle('Contact Person').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Phone');
  form.addTextItem().setTitle('Estimated Value ($)');
  form.addMultipleChoiceItem()
      .setTitle('How did you hear about us?')
      .setChoiceValues(['Website', 'Referral', 'Social Media', 'Other']);
  
  // Link form to your spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActive().getId());
  
  Logger.log('Form URL: ' + form.getPublishedUrl());
  SpreadsheetApp.getUi().alert('Form created! Check logs for the URL.');
}

function sendFollowupReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nextFollowupDate = row[10]; // Next Follow-up Date column (index 10)
    const status = row[5];
    const email = row[3];
    const company = row[1];
    const contactPerson = row[2];
    
    if (nextFollowupDate && isSameDay(new Date(nextFollowupDate), today) && 
        status !== 'Closed-Won' && status !== 'Closed-Lost' && email) {
      
      // Send email reminder
      const subject = `Follow-up reminder: ${company}`;
      const body = `Hi there,\n\nThis is a reminder to follow up with ${contactPerson} at ${company}.\n\nCurrent status: ${status}\n\nCheck the sales pipeline for details: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
      
      try {
        MailApp.sendEmail(email, subject, body);
        
        // Update last contact date
        sheet.getRange(i+1, 10).setValue(today); // Update Last Contact Date
        
        Logger.log(`Sent follow-up reminder for ${company}`);
      } catch (e) {
        Logger.log(`Failed to send email to ${email}: ${e.toString()}`);
      }
    }
  }
}

function isSameDay(date1, date2) {
  return date1.getDate() === date2.getDate() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getFullYear() === date2.getFullYear();
}

function updateLeadStatus(leadId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === leadId) { // Lead ID column
      sheet.getRange(i+1, 6).setValue(newStatus); // Status column
      
      // If status changed to "Proposal Sent", schedule follow-up
      if (newStatus === 'Proposal Sent') {
        const followUpDate = new Date();
        followUpDate.setDate(followUpDate.getDate() + 7); // Follow up in 7 days
        sheet.getRange(i+1, 11).setValue(followUpDate); // Next Follow-up Date
      }
      
      Logger.log(`Updated status for lead ${leadId} to ${newStatus}`);
      break;
    }
  }
}

function generatePipelineReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Count leads by status
  const statusCount = {};
  let totalValue = 0;
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][5];
    const value = parseFloat(data[i][6]) || 0;
    
    if (status) {
      statusCount[status] = (statusCount[status] || 0) + 1;
      
      if (status !== 'Closed-Lost') {
        totalValue += value;
      }
    }
  }
  
  // Create a simple report
  let report = 'SALES PIPELINE REPORT\n\n';
  report += `Total Pipeline Value: $${totalValue.toFixed(2)}\n\n`;
  report += 'Leads by Status:\n';
  
  for (const status in statusCount) {
    report += `${status}: ${statusCount[status]} leads\n`;
  }
  
  // Display the report
  SpreadsheetApp.getUi().alert(report);
  return report;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sales Pipeline')
    .addItem('Create Lead Form', 'createLeadForm')
    .addItem('Send Follow-up Reminders', 'sendFollowupReminders')
    .addItem('Generate Pipeline Report', 'generatePipelineReport')
    .addToUi();
}
