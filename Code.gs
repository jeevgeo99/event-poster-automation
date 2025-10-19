function createPosterFromForm(e) {
  try {
    // Spreadsheet and sheet setup
    const ss = SpreadsheetApp.openById('1SJneYqRRXzojv0SOo1i5YoNbNRidqGLnuF86qqi_u9U');
    const sheet = ss.getSheetByName('Form responses 1');

    if (!sheet) {
      Logger.log('❌ ERROR: Sheet "Form responses 1" not found!');
      return;
    }

    // Get last submitted row
    const lastRow = sheet.getLastRow();
    const lastRowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Build placeholder-to-value map
    const data = {
      '<<Email>>': lastRowData[headers.indexOf('Email Address')],
      '<<District>>': lastRowData[headers.indexOf('District')],
      '<<Division>>': lastRowData[headers.indexOf('Division')],
      '<<Area>>': lastRowData[headers.indexOf('Area')],
      '<<Club>>': lastRowData[headers.indexOf('Club')],
      '<<EventType>>': lastRowData[headers.indexOf('Event Type')],
      '<<Time>>': lastRowData[headers.indexOf('Time')],
      '<<Date>>': lastRowData[headers.indexOf('Date')],
      '<<Venue>>': lastRowData[headers.indexOf('Venue')],
      '<<ContactName>>': lastRowData[headers.indexOf('Contact Name')],
      '<<ContactEmail>>': lastRowData[headers.indexOf('Contact Email')],
      '<<TemplateChoice>>': lastRowData[headers.indexOf('Template Choice')],
    };

    Logger.log('✅ Data: ' + JSON.stringify(data));

    const recipientEmail = data['<<Email>>'];
    if (!recipientEmail) {
      Logger.log('❌ ERROR: No Email Address!');
      return;
    }

    // Map template choices to Google Slides IDs
    const templateMap = {
      'Template 1': '1JwxwufsHEKNqQZOCJyjV6_-zQxwiT5pnYqJTZl8s52s',
      'Template 2': '194WFVyFRq8aSs3ywS3fq0le1NQavDpucCv9y4xlPwUI' 
    };

    const templateChoice = data['<<TemplateChoice>>'];
    const templateId = templateMap[templateChoice];

    if (!templateId) {
      Logger.log(`❌ ERROR: No template found for choice "${templateChoice}"`);
      return;
    }

    Logger.log('✅ Using template ID: ' + templateId);

    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(`Poster_${data['<<Club>>']}_${new Date().toISOString()}`);
    const presentationId = newFile.getId();

    if (newFile.getMimeType() !== MimeType.GOOGLE_SLIDES) {
      throw new Error('❌ Copied file is not a Google Slides presentation!');
    }

    const presentation = SlidesApp.openById(presentationId);
    const slides = presentation.getSlides();

    slides.forEach(slide => {
      for (let placeholder in data) {
        slide.replaceAllText(placeholder, data[placeholder] || '');
      }
    });

    presentation.saveAndClose();

    const pdf = DriveApp.getFileById(presentationId).getAs('application/pdf');

    const htmlBody = `
      <p>Hello,</p>

      <p>Thank you for using <b>District 91's Branded Event Poster Generator!</b><br>
      I’m pleased to inform you that your poster has been generated and is attached to this email. 
      The poster is designed to highlight the key details of your upcoming meeting and help engage your members.</p>

      <p>The poster is designed to be easily printed for display during your meeting. 
      If you prefer to share the poster digitally with your members, you can distribute it via email 
      or upload it to your club’s website.</p>

      <p>Thank you for your leadership and dedication in organizing your meeting. 
      If you need further assistance or have any questions regarding the poster, 
      please feel free to reply to this email.</p>

      <p>If you found this useful, please share this tool with fellow Toastmasters clubs 
      to make event planning even easier!<br>
      Your feedback is invaluable in helping us improve, so feel free to share your suggestions.</p>

      <p>Thank you for embracing technology and leading the way in District 91.</p>

      <p>Best regards,<br>
      <b>D91 Toastmasters Team,</b><br>
      Toastmasters International</p>
    `;

    MailApp.sendEmail({
      to: recipientEmail,
      subject: `Your Event Poster: ${data['<<EventType>>']} (${data['<<Club>>']})`,
      htmlBody: htmlBody,
      attachments: [pdf]
    });

    Logger.log('✅ Poster sent to: ' + recipientEmail);

  } catch (err) {
    Logger.log('❌ Error: ' + err.message);
    throw err;
  }
}

// Trigger creation
function createPosterTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('createPosterFromForm')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  Logger.log('✅ Trigger created: onFormSubmit -> createPosterFromForm');
}






















































































// function createPosterFromForm(e) {
//   try {
//     // Spreadsheet and sheet setup
//     const ss = SpreadsheetApp.openById('1SJneYqRRXzojv0SOo1i5YoNbNRidqGLnuF86qqi_u9U');
//     const sheet = ss.getSheetByName('Form responses 1');

//     if (!sheet) {
//       Logger.log('❌ ERROR: Sheet "Form responses 1" not found!');
//       return;
//     }

//     // Get last submitted row
//     const lastRow = sheet.getLastRow();
//     const lastRowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
//     const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

//     // Build placeholder-to-value map
//     const data = {
//       '<<Email>>': lastRowData[headers.indexOf('Email Address')],
//       '<<District>>': lastRowData[headers.indexOf('District')],
//       '<<Division>>': lastRowData[headers.indexOf('Division')],
//       '<<Area>>': lastRowData[headers.indexOf('Area')],
//       '<<Club>>': lastRowData[headers.indexOf('Club')],
//       '<<EventType>>': lastRowData[headers.indexOf('Event Type')],
//       '<<Time>>': lastRowData[headers.indexOf('Time')],
//       '<<Date>>': lastRowData[headers.indexOf('Date')],
//       '<<Venue>>': lastRowData[headers.indexOf('Venue')],
//       '<<ContactName>>': lastRowData[headers.indexOf('Contact Name')],
//       '<<ContactEmail>>': lastRowData[headers.indexOf('Contact Email')],
//     };

//     Logger.log('✅ Data: ' + JSON.stringify(data));

//     const recipientEmail = data['<<Email>>'];
//     if (!recipientEmail) {
//       Logger.log('❌ ERROR: No Email Address!');
//       return;
//     }

//     // Template ID
//     const templateId = '1JwxwufsHEKNqQZOCJyjV6_-zQxwiT5pnYqJTZl8s52s';
//     Logger.log('✅ Using template ID: ' + templateId);

//     const templateFile = DriveApp.getFileById(templateId);
//     const newFile = templateFile.makeCopy(`Poster_${data['<<Club>>']}_${new Date().toISOString()}`);
//     const presentationId = newFile.getId();

//     if (newFile.getMimeType() !== MimeType.GOOGLE_SLIDES) {
//       throw new Error('❌ Copied file is not a Google Slides presentation!');
//     }

//     const presentation = SlidesApp.openById(presentationId);
//     const slides = presentation.getSlides();

//     slides.forEach(slide => {
//       for (let placeholder in data) {
//         slide.replaceAllText(placeholder, data[placeholder] || '');
//       }
//     });

//     presentation.saveAndClose();

//     const pdf = DriveApp.getFileById(presentationId).getAs('application/pdf');

//     const htmlBody = `
//       <p>Hello,</p>

//       <p>Thank you for using <b>District 91's Branded Event Poster Generator!</b><br>
//       I’m pleased to inform you that your poster has been generated and is attached to this email. 
//       The poster is designed to highlight the key details of your upcoming meeting and help engage your members.</p>

//       <p>The poster is designed to be easily printed for display during your meeting. 
//       If you prefer to share the poster digitally with your members, you can distribute it via email 
//       or upload it to your club’s website.</p>

//       <p>Thank you for your leadership and dedication in organizing your meeting. 
//       If you need further assistance or have any questions regarding the poster, 
//       please feel free to reply to this email.</p>

//       <p>If you found this useful, please share this tool with fellow Toastmasters clubs 
//       to make event planning even easier!<br>
//       Your feedback is invaluable in helping us improve, so feel free to share your suggestions.</p>

//       <p>Thank you for embracing technology and leading the way in District 91.</p>

//       <p>Best regards,<br>
//       <b>D91 Toastmasters Team,</b><br>
//       Toastmasters International</p>
//     `;

//     MailApp.sendEmail({
//       to: recipientEmail,
//       subject: `Your Event Poster: ${data['<<EventType>>']} (${data['<<Club>>']})`,
//       htmlBody:htmlBody,
//       // body: `Hello,\n\nAttached is your automatically generated event poster.\n\nRegards,\nDistrict 91 Poster Bot`,
//       attachments: [pdf]
//     });

//     Logger.log('✅ Poster sent to: ' + recipientEmail);

//   } catch (err) {
//     Logger.log('❌ Error: ' + err.message);
//     throw err;
//   }
// }

// // Trigger creation
// function createPosterTrigger() {
//   const ss = SpreadsheetApp.getActive();
//   ScriptApp.newTrigger('createPosterFromForm')
//     .forSpreadsheet(ss)
//     .onFormSubmit()
//     .create();
//   Logger.log('✅ Trigger created: onFormSubmit -> createPosterFromForm');
// }
