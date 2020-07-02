// form是google表單
var form = FormApp.openById("1UAPVrsrUHhKWD9NfpBHdef3iFON-bwhxm-JLMHxTOFs");
// sheet是google試算表 "output"
var sheet = SpreadsheetApp.openById(
  "1LWvBu3oRkT1VrKmR78_0vQ1T8bKOFitHn5JIlKHPn9k"
);

function newSubmitTrigger() {
  var date = new Date();
  var month = date.getMonth() + 1;
  var timestamp =
    date.getFullYear() +
    "/" +
    month +
    "/" +
    date.getDate() +
    " " +
    date.getHours() +
    ":" +
    date.getMinutes() +
    ":" +
    date.getSeconds();
  var responses = [timestamp];
  var lastformResponse = form.getResponses().slice(-1)[0];
  var itemResponses = lastformResponse.getItemResponses();
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var question = itemResponse.getItem().getTitle();
    var response = itemResponse.getResponse();
    Logger.log(
      'Response #%s to the question "%s" was "%s"',
      (i + 1).toString(),
      question,
      response
    );
    responses.push(response);
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var array = sheet.getRange("工作表1!C1:C" + lastRow).getValues();

  Logger.log("search area: " + array);
  //檢查重覆學號
  Logger.log("search target: " + responses[2]);
  for (var row = 1; row <= lastRow; row++) {
    Logger.log(array[row]);
    if (array[row] == responses[2]) {
      range = sheet.getActiveSheet().getRange(parseInt(row + 1), 1, 1, 1);
      range.setValue(timestamp);
      for (var col = 1; col <= responses.length; col++) {
        range = sheet.getActiveSheet().getRange(parseInt(row + 1), col, 1, 1);
        range.setValue(responses[col - 1]);
      }
      return "duplicated student ID";
    }
  }
  //把form新的資料移到sheet上

  sheet.insertRowBefore(2);
  for (var col = 1; col <= responses.length; col++) {
    range = sheet.getActiveSheet().getRange(2, col, 1, 1);
    range.setValue(responses[col - 1]);
  }

  sortingByStudentId();
}

function setoutputTrigger() {
  ScriptApp.newTrigger("newSubmitTrigger")
    .forForm(form)
    .onFormSubmit()
    .create();
}
