const template_ID = "1jIO0t2OEHlTP0BcEhDC77Q_faV4Bx8hxMx-zjFWscF4";

const sheet = "Form responses 1";

function onFormSubmit(e){
  
  const name = e.values[1];
  const company = e.values[2];
  const service = e.values[3];
  const budget = e.values[4];
  const date = e.values[5];

  const copy = DriveApp.getFileById(template_ID).makeCopy(`Proposal - ${company}`);

  const doc = DocumentApp.openById(copy.getId());

  const body = doc.getBody();
      body.replaceText('{{Name}}', name);
      body.replaceText('{{Company}}', company);
      body.replaceText('{{Service}}', service);
      body.replaceText('{{Budge}}', budget);
      body.replaceText('{{Date}}', date);

  doc.saveAndClose();

  const pdfBlob = copy.getAs(MimeType.PDF);

  const email = "hamzuansari7@gmail.com";

  GmailApp.sendEmail(
    email, 'Proposal - ' +
    company, " Hi " +
    name + ", Please find attached our proposl.",
    {
      attachments: [pdfBlob]
    }
  );

}
