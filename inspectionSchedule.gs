function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('emailSend')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}

function emailSend() {
  var date = new Date();
  var dvvGmal = "dvv@nso.ru" + "," + "gmal@nso.ru";
  var sheets = ['Дмитриев+Гора', 'Желуницин', 'Колпакова', 'Шпильной', 'Жданов', 'Столбова', 'Приёмкин'];
  var recipient = [dvvGmal, "grv@nso.ru", "aako@nso.ru", "shaa@nso.ru", "jds@nso.ru", "sti@nso.ru", "prias@nso.ru"];
  for (let i = 0; i < recipient.length; i++) {
    removingUnwantedColumns(sheets[i]);
    for (let row = 3; row <= valueLimit(sheets[i]); row++) {
      if (typeof searchCell(sheets[i], row) !== 'undefined' &&
          chekingValues(sheets[i]) == date.getMonth() - 1) {
        MailApp.sendEmail(recipient[i], "Оповещение о проверке",
                          "У Вас в прошлом месяце на объекте: " + searchCell(sheets[i], row) +
                          " была запланирована выездная проверка по утвержденной программе проверок!");
      }
    }
  }
}

function removingUnwantedColumns(nameSheet) {
  var date = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  if (chekingValues(nameSheet) == date.getMonth() - 2) {
    sheet.deleteColumn(4);
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
  var date = new Date();
  var months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
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
