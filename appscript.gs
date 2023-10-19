// Variables globales
var timeZone = "America/Argentina/Catamarca";
var dateTimeFormat = "dd/MM/yyyy HH:mm:ss";
var logSpreadSheetId = "";
var attendanceLogSheetName = "attendance log";
var defaultTerminalName = "headquarter";
var mainTabName = "main tab";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Anyboards Menu')
    .addItem('Configuración Inicial', 'initialSetup')
    .addItem('Agregar Nuevas UID', 'addNewUIDsFromAttendanceLogUiHandler')
    .addItem('Agregar una UID Seleccionada', 'addOneSelectedUID')
    .addToUi();
}

function addOneSelectedUID() {
  var tabName = SpreadsheetApp.getActiveSheet().getName();
  if (tabName != attendanceLogSheetName) {
    SpreadsheetApp.getUi().alert('Debe estar en la hoja ' + attendanceLogSheetName);
  }
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  var col = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
  
  addNewUIDsFromAttendanceLog(row);
}

function addNewUIDsFromAttendanceLogUiHandler() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Todas las nuevas UID de ' + attendanceLogSheetName + ' se agregarán a la pestaña principal', '¿Estás seguro?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    addNewUIDsFromAttendanceLog();
  }
}

function addNewUIDsFromAttendanceLog(row) {
  var mainTab = getMainSheet();
  var data = mainTab.getRange(2, 1, mainTab.getLastRow(), 1).getValues();
  var registeredUIDs = [];
  data.forEach(x => registeredUIDs.push(x[0]));

  registeredUIDs = [...new Set(registeredUIDs)];

  var attendanceSheet = getAttendanceLogSheet();

  var data;
  if (row)
    data = attendanceSheet.getRange(row, 1, row, 2).getValues();
  else
    data = attendanceSheet.getRange(2, 1, attendanceSheet.getLastRow(), 2).getValues();
  var arr = [];

  for (var i = 0; i < data.length; i++) {
    var visit = [];
    var uid = data[i][1];
    if (!registeredUIDs.includes(uid)) {
      visit.date = data[i][0];
      visit.uid = uid;
      arr.push(visit)
      registeredUIDs.push(uid);
    }
  }

  var startRow = mainTab.getLastRow() + 1;
  data = [];
  for (var i = arr.length - 1; i >= 0; i--) {
    var row = [];
    row[0] = arr[i].uid;
    row[1] = 'Persona ' + (startRow - 2 + arr.length - i);
    row[2] = "salida";
    row[3] = "Estás registrado";
    row[4] = 0;
    row[5] = arr[i].date;
    data.push(row);
  }
  if (data.length > 0)
    mainTab.getRange(startRow, 1, data.length, data[0].length).setValues(data);
}

function initialSetup() {
  if (!getAttendanceLogSheet()) {
    var mainSheet = SpreadsheetApp.getActiveSheet().setName(mainTabName);
    var rowData = ['UID', 'Nombre', 'Acceso', 'Texto a Mostrar', 'Conteo de Visitas', 'Última Visita'];
    mainSheet.getRange(1, 1, 1, rowData.length).setValues([rowData]);
    mainSheet.setColumnWidths(1, rowData.length + 1, 150);

    rowData = ['Fecha y Hora', 'UID', 'Nombre', 'Resultado', 'Terminal'];
    var attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(attendanceLogSheetName);
    attendanceSheet.getRange(1, 1, 1, rowData.length).setValues([rowData]);
    attendanceSheet.setColumnWidths(1, rowData.length + 1, 150);

  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('El sistema de la hoja de cálculo ya ha sido inicializado');
  }
}

function doGet(e) {
  var access = "-1";
  var text = 'Ir a casa';
  var name = '¿Quien eres?';

  var dateTime = Utilities.formatDate(new Date(), timeZone, dateTimeFormat);
  var result = 'Ok';
  if (e.parameter == 'undefined') {
    result = 'Sin parámetros';
  } else {
    var uid = '';
    var terminal = defaultTerminalName;
    for (var param in e.parameter) {
      var value = stripQuotes(e.parameter[param]);

      switch (param) {
        case 'uid':
          uid = value;
          break;
        case 'terminal':
          terminal = value;
          break;

        default:
          result = "Parámetro no compatible";
      }
    }

    var mainSheet = getMainSheet();
    var data = mainSheet.getDataRange().getValues();
    if (data.length == 0)
      return;

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == uid) {
        name = data[i][1];
        access = data[i][2];
        var numOfVisits = mainSheet.getRange(i + 1, 5).getValue();

        if (access == "entrada") {
          access = "salida";
          text = "Estas saliendo";
        } else if (access == "salida") {
          access = "entrada";
          text = "Estas entrando";
        }

        mainSheet.getRange(i + 1, 5).setValue(numOfVisits + 1);
        mainSheet.getRange(i + 1, 6).setValue(dateTime + ' ' + terminal);
        mainSheet.getRange(i + 1, 3).setValue(access);
        break;
      }
    }

    var attendanceSheet = getAttendanceLogSheet();
    data = [dateTime, uid, name, access, terminal];
    attendanceSheet.getRange(attendanceSheet.getLastRow() + 1, 1, 1, data.length).setValues([data]);
  }

  result = access + ":" + name + ":" + text;
  return ContentService.createTextOutput(result);
}

function getAttendanceLogSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(attendanceLogSheetName);
}

function getMainSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainTabName);
}

function stripQuotes(value) {
  return value.replace(/^["']|['"]$/g, "");
}
