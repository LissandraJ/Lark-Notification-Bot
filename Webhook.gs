function doPost(e) {
  let params = JSON.parse(e.postData.contents);
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tasks created");
  let lastRow = sheet.getLastRow();

  let task_id = params.task.id;
  let last_id = sheet.getRange(lastRow, 6).getValue();

  if ( last_id == task_id ) {
    return ContentService.createTextOutput("retry system");
  }

  let project_name = params.project.name
  let project_id = params.project.id;
  let task_title = params.task.title;
  let due_date = params.task.due_date;
  let date_time = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");

  sheet.insertRowsAfter(lastRow,1)

  sheet.getRange(lastRow + 1, 1).setValue(project_name);
  sheet.getRange(lastRow + 1, 2).setValue("master");
  sheet.getRange(lastRow + 1, 3).setValue(task_title);
  sheet.getRange(lastRow + 1, 4).setValue(due_date);
  sheet.getRange(lastRow + 1, 5).setValue(project_id);
  sheet.getRange(lastRow + 1, 6).setValue(task_id);
  sheet.getRange(lastRow + 1, 11).setValue(date_time);

  let now = new Date();
  now.setSeconds(now.getSeconds() + 5);

  ScriptApp.newTrigger('larkbot').timeBased().at(now).create();

  return ContentService.createTextOutput(params);
}
