const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ArteDelCantar')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleCors(() => {
    let requestData;
    try {
      requestData = JSON.parse(e?.postData?.contents || '{}');
    } catch (err) { return respondJson({ ok: false, error: 'JSON inválido' }, 400); }
    
    const { action, payload = {}, token } = requestData;
    if (!action) return respondJson({ ok: false, error: 'Acción requerida' }, 400);

    if (action === 'createDirectBooking') return handleCreateDirectBooking(payload);
    if (action === 'loginAdmin') return handleLoginAdmin(payload);

    try {
      verifyToken(token);
    } catch (err) { return respondJson({ ok: false, error: 'No autorizado' }, 401); }

    switch (action) {
      case 'getSessions': return handleGetSessions();
      case 'createSession': return handleCreateSession(payload);
      case 'deleteSession': return handleDeleteSession(payload);
      case 'getStudentRequests': return handleGetStudentRequests(payload);
      case 'updateStudentStatus': return handleUpdateStudentStatus(payload);
      case 'updateStudentNote': return handleUpdateStudentNote(payload);
      case 'getAssignments': return handleGetAssignments();
      case 'deleteAssignment': return handleDeleteAssignment(payload);
      default: return respondJson({ ok: false, error: 'Acción desconocida' }, 400);
    }
  });
}

function handleCors(callback) {
  const result = callback();
  const output = (typeof result === 'object' && result.setMimeType) ? result : respondJson(result);
  return output.setHeader("Access-Control-Allow-Origin", "*")
               .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
               .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function respondJson(data, status = 200) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupInitialConfig() {
  let ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('ArteDelCantar_Agenda');
  SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', ss.getId());
  return { ok: true, spreadsheetId: ss.getId() };
}

function getSpreadsheet() { return SpreadsheetApp.openById(SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID')); }
function getAllRecords(sheetName) { const sheet = getSpreadsheet().getSheetByName(sheetName); return sheet ? sheet.getDataRange().getValues().slice(1).map((row, i, arr) => { /* lógica de parseo */ return {}; }) : []; }
function insertRecord(sheetName, record) { getSpreadsheet().getSheetByName(sheetName).appendRow(Object.values(record)); return record; }
function deleteRecord(sheetName, id) { /* lógica de borrado */ }
function getRecordById(sheetName, id) { /* lógica de búsqueda */ }
function hashPassword(pwd) { return "hashed"; } 
function createToken(payload) { return "token"; }
function verifyToken(token) { return true; }
function handleLoginAdmin(payload) { return { ok: true, data: { token: "fake-token" } }; }
function handleGetSessions() { return { ok: true, data: [] }; }
function handleCreateSession(data) { return { ok: true, data: {} }; }
function handleDeleteSession(data) { return { ok: true, data: {} }; }
function handleGetStudentRequests(data) { return { ok: true, data: [] }; }
function handleUpdateStudentStatus(data) { return { ok: true, data: {} }; }
function handleUpdateStudentNote(data) { return { ok: true, data: {} }; }
function handleGetAssignments() { return { ok: true, data: [] }; }
function handleDeleteAssignment(data) { return { ok: true, data: {} }; }
function handleCreateDirectBooking(data) { return { ok: true, data: {} }; }

