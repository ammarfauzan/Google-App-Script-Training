// This Function to add menu in navbar Names HR Dept

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('HR Dept') // You can name your custom menu anything
      .addItem('Header Formatting', 'ownHeaderFormatting')
      .addItem('Send Email Attandence Data')
      .addToUi();
}

function ownHeaderFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the range of your header row.
  const headerRange = sheet.getRange('A1:I1');

  // Apply formatting:
  headerRange.setBackground('#DDEBF7'); // Light Blue background (optional, for visual appeal)
  headerRange.setFontWeight('bold'); // Makes the text bold
  headerRange.setFontSize(12);       // Sets the font size to 12
  // headerRange.setHorizontalAlignment('center'); // Center aligns the text (optional)
  // headerRange.setVerticalAlignment('middle');   // Vertically aligns the text in the middle (optional)

  Logger.log("Header row A1:I1 formatted successfully: bold, font size 12, light blue background.");
  SpreadsheetApp.getUi().alert("Header formatted!"); // User-friendly alert
}

// Function to Send Attachment csv to Email
function exportSheetAsCsvAndEmail() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Export Data 1 - Email"); // Replace with your sheet name
  const folderId = "1nxlheOqbrQSk9HujPbpHmaxfzMfyLrjP"; // Replace with your Google Drive folder ID

  const recipientEmail = "maf.fauzan@gmail.com, datarangerindonesia@gmail.com"; // The specific email address
  const subject = "Google Sheets Export: HR Data " + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const body = "Dear Team,\n\nPlease find the attached HR data attendence export.\n\nRegards,\nYour Automated System";

  if (!sheet) {
    console.error("Sheet 'HR Data' not found!");
    // You might want to send an email notification about the failure here too
    return;
  }

  const fileName = sheet.getName() + "_" + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss") + ".csv";
  const csvData = convertSheetToCsv(sheet);

  // Create the file in Google Drive (optional, but good for backup)
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(fileName, csvData, MimeType.CSV);

  // Get the file content as a Blob for attachment
  const attachmentBlob = file.getAs(MimeType.CSV);

  // Send the email with the attachment
  Logger.log("Try send to Email");
  try {
    GmailApp.sendEmail(recipientEmail, subject, body, {
      attachments: [attachmentBlob],
      name: "Google Sheets Automated Export" // Sender's name in the email
    });
    console.log("Exported " + fileName + " successfully to Google Drive and sent to " + recipientEmail + "!");
  } catch (e) {
    console.error("Failed to send email: " + e.toString());
  }

  // You can optionally delete the file from Drive after sending if you only need it as an email attachment
  // folder.removeFile(file);
}

function convertSheetToCsv(sheet) {
  const data = sheet.getDataRange().getValues();
  let csv = "";
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (let j = 0; j < row.length; j++) {
      let cell = row[j];
      // Handle commas and quotes in data for proper CSV formatting
      if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\n'))) {
        cell = '"' + cell.replace(/"/g, '""') + '"';
      }
      csv += cell;
      if (j < row.length - 1) {
        csv += ",";
      }
    }
    csv += "\n";
  }
  Logger.log("Convert Data into CSV successful!");
  return csv;
}