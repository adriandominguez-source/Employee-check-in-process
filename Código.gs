// appsscript.json ref: https://developers.google.com/apps-script/manifest
// timezone ids: https://gist.github.com/mhawksey/8673e904a03a91750c26c2754fe0977a

function doGet(e){
  var tpl = HtmlService.createTemplateFromFile("page.html");
  tpl.data = e.parameters;
  tpl.data.id = SpreadsheetApp.getActiveSpreadsheet().getId();
  return tpl.evaluate();
}

function checkin(userid, control){
  if(userid && control){
    var now = new Date();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro").appendRow([userid,control,now]);
    return now.toLocaleString();
  }
}
function generarReporteCompletoFichajes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values.shift(); 

  const idColIndex = headers.indexOf('ID');
  const tipoColIndex = headers.indexOf('TIPO');
  const fechaColIndex = headers.indexOf('FECHA');
  const horaColIndex = headers.indexOf('HORA');

  if (idColIndex === -1 || tipoColIndex === -1 || fechaColIndex === -1 || horaColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error: No se encontraron las columnas "ID", "TIPO", "FECHA" o "HORA".');
    return;
  }
  
  const dailyRecords = {};
  
  values.forEach(row => {
    const id = row[idColIndex];
    const tipo = row[tipoColIndex];
    const fecha = row[fechaColIndex];
    const hora = row[horaColIndex];

    if (!id || !fecha || !hora) return;

    const fechaHora = Utilities.formatDate(fecha, spreadsheet.getSpreadsheetTimeZone(), "yyyy/MM/dd") + ' ' + Utilities.formatDate(hora, spreadsheet.getSpreadsheetTimeZone(), "HH:mm:ss");
    const timestampDate = new Date(fechaHora);
    
    const dateKey = Utilities.formatDate(timestampDate, spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd");

    if (!dailyRecords[id]) {
      dailyRecords[id] = {};
    }
    if (!dailyRecords[id][dateKey]) {
      dailyRecords[id][dateKey] = { entradas: [], salidas: [] };
    }

    if (tipo.toLowerCase() === 'entrada') {
      dailyRecords[id][dateKey].entradas.push(timestampDate);
    } else if (tipo.toLowerCase() === 'salida') {
      dailyRecords[id][dateKey].salidas.push(timestampDate);
    }
  });

  const resumenDiario = [];
  const resumenMensual = {};
  const horasComplementarias = [];
  const alertasFichaje = [];
  const totalExtraHoursPerId = {};

  for (const id in dailyRecords) {
    for (const dateKey in dailyRecords[id]) {
      const dayData = dailyRecords[id][dateKey];

      if (dayData.entradas.length === 0 || dayData.salidas.length === 0) {
        alertasFichaje.push([id, new Date(dateKey), dayData.entradas.length === 0 ? "Falta entrada" : "Falta salida"]);
        continue;
      }
      
      const firstEntry = new Date(Math.min(...dayData.entradas));
      const lastExit = new Date(Math.max(...dayData.salidas));
      
      const durationMs = lastExit.getTime() - firstEntry.getTime();
      const hoursWorked = durationMs / (1000 * 60 * 60);

      resumenDiario.push([id, new Date(dateKey), firstEntry, lastExit, hoursWorked / 24]);

      const monthKey = Utilities.formatDate(firstEntry, spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM");
      if (!resumenMensual[id]) {
        resumenMensual[id] = {};
      }
      if (!resumenMensual[id][monthKey]) {
        resumenMensual[id][monthKey] = { totalHours: 0, totalDays: 0 };
      }
      resumenMensual[id][monthKey].totalHours += hoursWorked;
      resumenMensual[id][monthKey].totalDays += 1;

      const extraHours = hoursWorked - 8;
      if (extraHours > 0) {
        horasComplementarias.push([id, new Date(dateKey), extraHours / 24]);
        totalExtraHoursPerId[id] = (totalExtraHoursPerId[id] || 0) + extraHours;
      }
    }
  }

  // --- Generar la hoja "Resumen Diario" ---
  const resumenDiarioSheet = crearOlimpiarHoja("Resumen Diario", spreadsheet);
  resumenDiarioSheet.getRange(1, 1, 1, 5).setValues([['ID', 'Fecha', 'Primera Entrada', 'Última Salida', 'Horas Trabajadas']]);
  if (resumenDiario.length > 0) {
    resumenDiarioSheet.getRange(2, 1, resumenDiario.length, 5).setValues(resumenDiario);
    resumenDiarioSheet.getRange('B:B').setNumberFormat('dd/MM/yyyy');
    resumenDiarioSheet.getRange('C:C').setNumberFormat('HH:mm:ss');
    resumenDiarioSheet.getRange('D:D').setNumberFormat('HH:mm:ss');
    resumenDiarioSheet.getRange('E:E').setNumberFormat('[HH]:mm');
  }
  // Formato profesional
  const diarioHeader = resumenDiarioSheet.getRange(1, 1, 1, 5);
  diarioHeader.setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  resumenDiarioSheet.getRange(1, 1, resumenDiario.length > 0 ? resumenDiario.length + 1 : 1, 5).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');


  // --- Generar la hoja "Resumen Mensual" ---
  const resumenMensualSheet = crearOlimpiarHoja("Resumen Mensual", spreadsheet);
  const monthlyData = [];
  for (const id in resumenMensual) {
    for (const month in resumenMensual[id]) {
      const totalHours = resumenMensual[id][month].totalHours;
      const totalDays = resumenMensual[id][month].totalDays;
      const averageHours = totalHours / totalDays;
      monthlyData.push([id, month, totalHours / 24, averageHours / 24]);
    }
  }
  resumenMensualSheet.getRange(1, 1, 1, 4).setValues([['ID', 'Mes', 'Total de Horas Trabajadas', 'Promedio Diario']]);
  if (monthlyData.length > 0) {
    resumenMensualSheet.getRange(2, 1, monthlyData.length, 4).setValues(monthlyData);
    resumenMensualSheet.getRange('C:C').setNumberFormat('[HH]:mm');
    resumenMensualSheet.getRange('D:D').setNumberFormat('[HH]:mm');
  }
  // Formato profesional
  const mensualHeader = resumenMensualSheet.getRange(1, 1, 1, 4);
  mensualHeader.setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  resumenMensualSheet.getRange(1, 1, monthlyData.length > 0 ? monthlyData.length + 1 : 1, 4).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');


  // --- Generar la hoja "Horas Complementarias" ---
  const horasComplementariasSheet = crearOlimpiarHoja("Horas Complementarias", spreadsheet);
  horasComplementariasSheet.getRange(1, 1, 1, 3).setValues([['ID', 'Fecha', 'Horas Complementarias']]);
  if (horasComplementarias.length > 0) {
    horasComplementariasSheet.getRange(2, 1, horasComplementarias.length, 3).setValues(horasComplementarias);
    horasComplementariasSheet.getRange('B:B').setNumberFormat('dd/MM/yyyy');
    horasComplementariasSheet.getRange('C:C').setNumberFormat('[HH]:mm');
  }
  // Formato profesional
  const hcHeader = horasComplementariasSheet.getRange(1, 1, 1, 3);
  hcHeader.setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  horasComplementariasSheet.getRange(1, 1, horasComplementarias.length > 0 ? horasComplementarias.length + 1 : 1, 3).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');


  // --- Generar la hoja "Alertas de Fichaje" ---
  const alertasFichajeSheet = crearOlimpiarHoja("Alertas de Fichaje", spreadsheet);
  alertasFichajeSheet.getRange(1, 1, 1, 3).setValues([['ID', 'Fecha', 'Motivo de la Alerta']]);
  if (alertasFichaje.length > 0) {
    alertasFichajeSheet.getRange(2, 1, alertasFichaje.length, 3).setValues(alertasFichaje);
    alertasFichajeSheet.getRange('B:B').setNumberFormat('dd/MM/yyyy');
  }
  // Formato profesional
  const alertasHeader = alertasFichajeSheet.getRange(1, 1, 1, 3);
  alertasHeader.setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  alertasFichajeSheet.getRange(1, 1, alertasFichaje.length > 0 ? alertasFichaje.length + 1 : 1, 3).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');


  // --- Generar la hoja "Total" (con formato profesional) ---
  const totalSheet = crearOlimpiarHoja("Total", spreadsheet);
  
  totalSheet.getRange('A1').setValue('Resumen de Horas Complementarias Totales');
  totalSheet.getRange('A1:B1').merge();
  totalSheet.getRange('A1').setFontSize(14).setFontWeight('bold');

  const totalData = [['ID', 'Total de Horas Extras']];
  const uniqueIds = [];
  for (const id in totalExtraHoursPerId) {
    totalData.push([id, totalExtraHoursPerId[id] / 24]);
    uniqueIds.push(id);
  }
  
  totalSheet.getRange(3, 1, totalData.length, 2).setValues(totalData);

  const headerRange = totalSheet.getRange(3, 1, 1, 2);
  headerRange.setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  totalSheet.getRange(3, 1, totalData.length, 2).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');
  totalSheet.getRange('B:B').setNumberFormat('[HH]:mm');

  const idCell = totalSheet.getRange('A4');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(uniqueIds).build();
  idCell.setDataValidation(rule);
  
  const chart = totalSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(totalSheet.getRange(3, 1, totalData.length, 2))
    .setOption('title', 'Total de Horas Complementarias por Empleado')
    .setOption('hAxis', {title: 'ID del Empleado'})
    .setOption('vAxis', {title: 'Horas Extras'})
    .setOption('legend', {position: 'none'})
    .setPosition(3, 4, 0, 0)
    .build();
  
  totalSheet.insertChart(chart);

  SpreadsheetApp.getUi().alert('Análisis de fichajes completado. Revisa las nuevas hojas de resumen.');
}

function crearOlimpiarHoja(sheetName, spreadsheet) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}
