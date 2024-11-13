function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{ name: "Criar aArquivo CSV", functionName: "saveCSV" }];
  ss.updateMenu("Exportar CSV", csvMenuEntries);
};

function saveCSV(){
  var userInterface = HtmlService.createHtmlOutput()
    .setHeight(10)
    .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Gerando CSV');

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = sheet.getActiveSheet();

  var lastLine = activeSheet.getLastRow();
  var lastColumn = activeSheet.getLastColumn();

  var data = activeSheet.getRange(1,1,lastLine, lastColumn).getValues();

  SpreadsheetApp.flush();

  var csvData = csvConverter(data);

  function csvConverter(arr){

    var lineData = [];

    for(var i of arr){
      var line = i.join(";");
      lineData.push(line);
    }

    return "\uFEFF" + lineData.join("\n");

  }

  SpreadsheetApp.flush();

  var fileName = sheet.getName();

  var folder = DriveApp.createFolder(fileName.toLowerCase());

  var file = folder.createFile(fileName + ".csv", csvData, "text/csv");

  data.length = 0;

  downloader();
 
  //Browser.msgBox('O arquivo foi salvo em uma pasta chamada ' + folder.getName());


  function downloader(){
    var url = "https://drive.google.com/uc?export=download&id=" + file.getId();
    var html = "<script>window.open('"+ url +"');google.script.host.close();</script>"
    var userInterface = HtmlService.createHtmlOutput(html)
      .setHeight(10)
      .setWidth(100);
    SpreadsheetApp.getUi().showModalDialog(userInterface, 'Baixando CSV');
    Utilities.sleep(2000);

    var folderToDelete = DriveApp.getFolderById(folder.getId());
    folderToDelete.setTrashed(true);
  }

}
