/**
 * ArteDelCantar Agenda - Backend Autónomo
 */
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function setupInitialConfig() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('ArteDelCantar_Agenda');
    const spreadsheetId = ss.getId();
    SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', spreadsheetId);
    Logger.log('✅ Configuración completada. Spreadsheet ID: ' + spreadsheetId);
    return { ok: true, spreadsheetId };
  } catch (error) { return { ok: false, error: error.message }; }
}

function doGet(e) {
    return ContentService.createTextOutput(JSON.stringify({status: "ok"}))
      .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
    // Aquí iría tu lógica de negocio
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
}
