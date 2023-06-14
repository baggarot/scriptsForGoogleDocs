function creareTimeDrivenTriggers() {
  ScriptApp.newTriggers('emailSend')
    .timeBased
    .everyDays(1)
    .atHour(8)
    .create();
}

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Предписания');

function emailSend() {
  var dvvGmal = "dvv@nso.ru" + "," + "gmal@nso.ru";
  var recipient = [dvvGmal, "grv@nso.ru", "aako@nso.ru", "shaa@nso.ru", "jds@nso.ru", "sti@nso.ru", "prias@nso.ru"];
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    if (dateComparison(row) > 0) {
      MailApp.sendEmail(recipient[inspectorsName(row)], "Оповещение о сроке исполнения предписания",
                        "Через " + dateComparison(row) + " дня истекает срок исполнения предписания по объекту: " +
                        alertTheme(row));
    }
  }
}

function inspectorsName(row) {
  var inspectors = ['Дмитриев+Гора', 'Желуницин', 'Колпакова', 'Шпильной', 'Жданов', 'Столбова', 'Приёмкин'];
  var inspector = sheet.getRange(`I${row}`).getValue();
  for (let i = 0; i < inspectors.length; i++) {
    if (inspectors[i] == inspector) return i;
  }
}

function alertTheme(row) {
  var objectName = sheet.getRange(`B${row}`).getValue();
  var address = sheet.getRange(`C${row}`).getValue();
  var organization = sheet.getRange(`D${row}`).getValue();
  Logger.log(objectName);
  return objectName + ", расположенный в " + address + ", организацией " + organization;
}

function reminder(counter) {
  var dates = new Date();
  var reminderDate = dates.getDate() + counter;
  var reminderMonth = dates.getMonth() + 1;
  var reminderYear = dates.getFullYear();
  return reminderDate + "." + reminderMonth + "." + reminderYear;
}

function dateComparison(row) {
  var cell = Utilities.formatDate(new Date(sheet.getRange(`G${row}`).getValue()), "GMT+7", "d.M.yyyy");
  for (let i = 7; i > 0; i--) {
    if (reminder(i) == cell) return i;
  }
  return 0;
}
