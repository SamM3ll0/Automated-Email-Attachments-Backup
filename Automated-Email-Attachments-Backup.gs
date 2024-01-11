// Substitua "PUT_YOUR_FOLDER_ID_HERE" pelo ID da pasta no Google Drive
var destinationFolderId = "PUT_YOUR_FOLDER_ID_HERE";
// Substitua "PUT_YOUR_SPREADSHEET_URL_HERE" pelo URL da planilha no Google Sheets
var spreadsheetUrl = "PUT_YOUR_SPREADSHEET_URL_HERE";

// Modifique a consulta de pesquisa conforme necessário
var searchQuery = "seuemail@gmail.com has:attachment";

function saveNewAttachmentsToDriveAndSpreadsheet() {
  var folder = DriveApp.getFolderById(destinationFolderId);
  var threads = GmailApp.search(searchQuery);

  // Abre a planilha
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheet = spreadsheet.getSheetByName("Anexos");

  // Adiciona cabeçalhos se a planilha ou a folha estiverem vazias
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Anexos");
    sheet.appendRow(["Data de Envio", "Hora de Envio", "URL no Drive"]);
  }

  threads.forEach(function (thread) {
    var messages = thread.getMessages();

    messages.forEach(function (message) {
      var attachments = message.getAttachments();

      attachments.forEach(function (attachment) {
        // Adiciona a data e hora ao nome do arquivo
        var date = Utilities.formatDate(message.getDate(), "GMT", "yyyy-MM-dd");
        var time = Utilities.formatDate(message.getDate(), "GMT", "HH:mm:ss");
        var fileName = date + " " + time;

        // Verifica se o arquivo já existe na pasta
        var existingFiles = folder.getFilesByName(fileName);
        if (!existingFiles.hasNext()) {
          // Cria o arquivo no Google Drive apenas se não existir
          var file = folder.createFile(attachment.copyBlob().setName(fileName));

          // Adiciona informações na planilha
          sheet.appendRow([date, time, file.getUrl()]);
        }
      });
    });
  });
}
