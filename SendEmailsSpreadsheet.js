function sendArticleCountEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("A1:H1"); // Primeira célula de dados : última célula de dados (diagonal) //// apenas primeira linha
  var data = dataRange.getDisplayValues();
 for (i in data) {
    var rowData = data[i];
    var recipient = rowData[2]; // Nome
    var emailAddress = rowData[0]; // Email
    var subject = rowData[3]; // Assunto do Email
    var CNPJ = rowData[1]; // número do CNPJ
   //
    var arquivoOps = rowData[6]; // numero do boleto.pdf
    var file = DriveApp.getFilesByName(arquivoOps + ".pdf");
    var dataVencimento = rowData[5]; // 23/10/2023 etc
    var lorem = rowData[7];

    // Construct the message with the variable
    var messagecarta = '<div style="font-size: 13pt; font-family: Arial, sans-serif;">\
        <p><strong>Prezados(as) Colegas,</strong></p>\
        \
        <p>Paragraph</p>\
        <ul>
            <li style="margin-bottom: 10px">List item</li>\
        </ul>\
        <p> Paragraph <strong>' + dataVencimento + '</strong> lorem ipsum.</p>\
        <p>Paragraph <a href="mailto:EMAIL@EMAIL.com">EMAIL@EMAIL.COM</a></p>\
        <p style="color:#0b5394;"><strong>NAME OF SIGNATURE<br>Signature</strong></p>\
    </div>';

    //
   
    //
    //
    MailApp.sendEmail(
     emailAddress, 
     subject + " " + dataVencimento, 
     messagecarta, 
     {htmlBody:messagecarta, 
      attachments: [file.next().getBlob()]}); // envia email
      console.log(emailAddress);

  }
 
}

