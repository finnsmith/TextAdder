function onInstall() {
  onOpen();
  showSidebar();
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem("Add text to selection", "showSidebar")
  .addToUi();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile("addTextSidebar")
    .evaluate()
    .setTitle("Add text to selection"); 
  SpreadsheetApp.getUi().showSidebar(html);
}

function addText(newWord, wordSelection, position) {
  Logger.log("addText running")
  var selection = SpreadsheetApp.getActiveRange();
  if (position == 'at the beginning of') position = 'front';
  if (wordSelection == 'just the first word') wordSelection = 'firstWord';
  var oldData = selection.getValues();
  var newData = [];
  
  var inappropriateRange = true;
  
  for (i in oldData) {
    var newRow = [];
    for (j in oldData[i]){
      if (typeof oldData[i][j] == 'string' && oldData[i][j].length > 0) {
        inappropriateRange = false;
        var oldText;
        if (wordSelection == 'firstWord') {
          if (position == 'front') { 
            newRow.push(newWord + oldData[i][j]);
          } else {
            oldText = oldData[i][j].split(' ');
            newRow.push(oldText[0] + newWord + " " + oldText.slice(1).join(' '));
          }
        } else {
          oldText = oldData[i][j].split(' ');
          var newText = [];
          for (w in oldText) {
            if (position == 'front') {
              newText.push(newWord + oldText[w]);
            } else {
              newText.push(oldText[w] + newWord);
            }
          }
          
          newText = newText.join(' ');
          newRow.push(newText);
        } 
      } else {
        newRow.push(oldData[i][j]);
      }
    }
    newData.push(newRow);
  }
  selection.setValues(newData);
  if (inappropriateRange) Browser.msgBox("Text Adder can only add to cells that have text in them. It won't work on numbers or blank cells.")
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getContent();
}
