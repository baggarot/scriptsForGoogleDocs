function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('emailSend')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}

var months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
var date = new Date();

function emailSend() {
  var dvvGmal = "dvv@nso.ru" + "," + "gmal@nso.ru";
  var sheets = ['Дмитриев+Гора', 'Желуницин', 'Колпакова', 'Шпильной', 'Приёмкин'];
  var recipient = [dvvGmal, "grv@nso.ru", "aako@nso.ru", "shaa@nso.ru", "prias@nso.ru"];
  for (let i = 0; i < recipient.length; i++) {
    Logger.log(chekingValues(sheets[i]));
    Logger.log(date.getMonth());
    removingUnwantedColumns(sheets[i]);
    for (let row = 3; row <= valueLimit(sheets[i]); row++) {
      switch (chekingValues(sheets[i])) {
        case 0.0:
        case (date.getMonth() - 1):
        case 11.0:
          if (typeof searchCell(sheets[i], row) !== 'undefined') MailApp.sendEmail(recipient[i], "Оповещение о проверке",
                                                                  "У Вас в прошлом месяце на объекте: " + searchCell(sheets[i], row) +
                                                                  " была запланирована выездная проверка по утвержденной программе проверок!");
      }
    }
  }
}

function removingUnwantedColumns(nameSheet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  switch (date.getMonth()) {
    case 0.0:
      Logger.log('запуск case 0');
      if (sheet.getRange('D2').getValue() == months[10]) sheet.deleteColumn(4);
      break;
    case 1.0:
      Logger.log('запуск case 1');
      if (sheet.getRange('D2').getValue() == months[11]) sheet.deleteColumn(4);
      break;
    default:
      Logger.log('запуск default');
      if (chekingValues(nameSheet) == date.getMonth() - 2) sheet.deleteColumn(4);
  }
}

function searchCell(nameSheet, row) {
  var yellow = '#ffff00';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  var activeCell = sheet.getRange(row, 4);
  if (activeCell.getBackground() == yellow) {
    var objectName = sheet.getRange(`B${row}`).getValue();
    var address = sheet.getRange(`C${row}`).getValue();
    Logger.log(objectName);
    return objectName + ", расположенном в " + address;
  }
}

function chekingValues(nameSheet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  var cell = sheet.getRange('D2').getValue();
  for (let i = 0; i <= months.length; i++) {
    if (cell == months[i]) {
      date.setMonth(i);
      return date.getMonth();
    }
  }
}

function valueLimit(nameSheet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  return sheet.getLastRow();
}
