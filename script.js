const extractEmail = () => {
  let emailAddressContainer = [];
  let rows = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    let googleDriveUrl = rows[i].toString();
    let fileId = googleDriveUrl.match(/[-\w]{25,}/);
    // Read the PDF file in Google Drive
    const pdfDocument = DriveApp.getFileById(fileId);
    let validPdf = pdfDocument
      .toString()
      .match(/(?<![<>][\s]|\w)[\w-]*?.pdf\b/gm);
    if (validPdf) {
      let language = "en"; // English
      // Use OCR to convert PDF to a temporary Google Document
      // Restrict the response to include file Id and Title fields only

      const { id } = Drive.Files.insert(
        {
          title: pdfDocument.getName().replace(/\.pdf$/, ""),
          mimeType: pdfDocument.getMimeType() || "application/pdf",
        },
        pdfDocument.getBlob(),
        {
          ocr: true,
          ocrLanguage: language,
          fields: "id",
        }
      );

      // Use the Document API to extract text from the Google Document
      const textContent = DocumentApp.openById(id).getBody().getText();
      // Delete Google Document since it is no longer needed
      Drive.Files.remove(id);
      let emailAddress = textContent.match(
        /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi
      );

      if (emailAddress) {
        emailAddressContainer.push(emailAddress.toString());
      }
    } else {
      let docx = DriveApp.getFileById(fileId);
      let blob = docx.getBlob();
      let file = Drive.Files.insert({}, blob, { convert: true });
      let id = file["id"];
      let doc = DocumentApp.openById(id);
      // Delete Google Document since it is no longer needed
      Drive.Files.remove(id);
      let text = doc.getBody().getText();
      let emailAddress = text.match(
        /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi
      );
      if (emailAddress) {
        emailAddressContainer.push(emailAddress.toString());
      }
    }
  }
  const email = emailAddressContainer.join("\r\n");
  DriveApp.createFile("emailAddress.txt", email);
};
