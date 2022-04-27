// inspirado de https://www.youtube.com/watch?v=iLALWX0_OYs&ab_channel=JeffEverhart
// @dbrisaro

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  cont = menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create New Docs', 'createNewGoogleDocs');
  menu.addToUi();
}

function createNewGoogleDocs() {
  const googleDocTemplate = DriveApp.getFileById('1oynt3dNOLPrEf9abNui5YHIgtrDVMsVr5Bi5yidxqW8');  // id del doc template
  const destinationFolder = DriveApp.getFolderById('1d_zgT7gJywzLBSIn3YiDOvzA0EmH6ruK');     // id de la carpeta de destino de los docs generados
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');    // planilla desde donde se leen los datos
  const rows = sheet.getDataRange().getValues();    // filas de la planilla
  
  rows.forEach(function(row, index){
    if (index === 0) return;  // en la fila 0 no se hace nada
    if (row[0]) return;    // si la primera columna tiene data, no se hace nada.
    
    const copy = googleDocTemplate.makeCopy(`${row[1]}_${row[3]}_Carta`, destinationFolder);    // title del doc a generar
    const doc = DocumentApp.openById(copy.getId());     // copia del archivo template
    const body = doc.getBody();     // toma el body
    const friendlyDate = new Date(row[5]).toLocaleDateString();   // reformateo de la fecha
    
    body.replaceText('{{Nombre}}', row[1]);
    body.replaceText('{{DNI}}', row[2]);
    body.replaceText('{{Convocatoria}}', row[3]);
    body.replaceText('{{Firma}}', row[4]);
    body.replaceText('{{Fecha}}', friendlyDate);   // pasos para reemplazar los tags 
    
    doc.saveAndClose();    // cierro el nuevo doc generado
    
    const url = doc.getUrl();
    sheet.getRange(index + 1, 1).setValue(url)   // se guarda el link al nuevo doc
                       
  })

}