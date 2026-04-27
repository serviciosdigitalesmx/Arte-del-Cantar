/**
 * ArteDelCantar Agenda - Backend API (GAS)
 * Versión optimizada para Frontend Externo (GitHub Pages)
 */

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

// ==================== CONFIGURACIÓN INICIAL ====================
function setupInitialConfig() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.create('ArteDelCantar_Agenda');
    SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', ss.getId());

    const sheetsConfig = {
      'sessions': ['id', 'date', 'start_time', 'end_time', 'mode', 'max_students', 'is_active', 'created_at', 'updated_at'],
      'student_requests': ['id', 'full_name', 'whatsapp', 'age', 'vocal_level', 'goal', 'available_weekdays', 'available_time_blocks', 'classes_per_week', 'comments', 'status', 'internal_note', 'created_at', 'updated_at'],
      'class_assignments': ['id', 'student_request_id', 'session_id', 'assigned_at'],
      'audit_log': ['id', 'action', 'entity_type', 'entity_id', 'metadata_json', 'created_at']
    };

    for (const [sheetName, headers] of Object.entries(sheetsConfig)) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      }
    }

    if (!SCRIPT_PROPERTIES.getProperty('ADMIN_PASSWORD_HASH')) SCRIPT_PROPERTIES.setProperty('ADMIN_PASSWORD_HASH', hashPassword('admin123'));
    if (!SCRIPT_PROPERTIES.getProperty('JWT_SECRET')) SCRIPT_PROPERTIES.setProperty('JWT_SECRET', Utilities.getUuid());

    return respondJson({ ok: true, message: 'Backend listo', spreadsheetId: ss.getId() });
  } catch (error) {
    return respondJson({ ok: false, error: error.message }, 500);
  }
}

// ==================== MANEJO DE CORS Y PETICIONES ====================

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  const action = e?.parameter?.action;
  if (action === 'health') return respondJson({ ok: true, status: 'online' });
  
  // Si no hay acción, podrías devolver un mensaje de error o ayuda
  return respondJson({ ok: false, error: 'Use POST para interactuar con la API' }, 400);
}

function doPost(e) {
  let requestData;
  try {
    requestData = JSON.parse(e?.postData?.contents || '{}');
  } catch (err) {
    return respondJson({ ok: false, error: 'JSON inválido' }, 400);
  }

  const { action, payload = {}, token } = requestData;
  
  try {
    // Acciones Públicas
    if (action === 'setup') return setupInitialConfig();
    if (action === 'loginAdmin') return handleLoginAdmin(payload);
    if (action === 'createStudentRequest') return handleCreateStudentRequest(payload);
    if (action === 'getActiveSessions') {
       const sessions = getAllRecords('sessions').filter(s => s.is_active && new Date(s.date) >= new Date().setHours(0,0,0,0));
       return respondJson({ ok: true, data: sessions });
    }

    // Acciones Privadas (Requieren Token)
    const decoded = verifyToken(token);
    if (decoded.role !== 'admin') throw new Error('No autorizado');

    switch (action) {
      case 'getStudentRequests': return respondJson({ ok: true, data: getAllRecords('student_requests') });
      case 'updateStudentStatus': return handleUpdateStudentStatus(payload);
      case 'getSessions': return respondJson({ ok: true, data: getAllRecords('sessions') });
      case 'createSession': return handleCreateSession(payload);
      case 'deleteSession': return handleDeleteSession(payload);
      default: return respondJson({ ok: false, error: 'Acción desconocida' }, 400);
    }
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, err.message === 'No autorizado' ? 401 : 400);
  }
}

function respondJson(data, status = 200) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== LÓGICA DE BASE DE DATOS ====================

function getSpreadsheet() {
  const id = SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Backend no inicializado. Ejecute acción: setup');
  return SpreadsheetApp.openById(id);
}

function getAllRecords(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, idx) => {
      let val = row[idx];
      if ((h === 'available_weekdays' || h === 'available_time_blocks') && typeof val === 'string') {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[h] = val;
    });
    return obj;
  });
}

function insertRecord(sheetName, record) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    let v = record[h];
    if (Array.isArray(v) || typeof v === 'object') v = JSON.stringify(v);
    return v !== undefined ? v : '';
  });
  sheet.appendRow(row);
  return record;
}

function updateRecord(sheetName, id, updates) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();
  const idCol = headers.indexOf('id');
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx === -1) throw new Error('Registro no encontrado');
  for (const [key, val] of Object.entries(updates)) {
    const col = headers.indexOf(key);
    if (col !== -1) {
      let finalVal = val;
      if (Array.isArray(val) || typeof val === 'object') finalVal = JSON.stringify(val);
      sheet.getRange(rowIdx, col + 1).setValue(finalVal);
    }
  }
}

// ==================== SEGURIDAD ====================
function hashPassword(pwd) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verifyToken(token) {
  if (!token) throw new Error('No autorizado');
  try {
    const secret = SCRIPT_PROPERTIES.getProperty('JWT_SECRET');
    const parts = token.split('.');
    const payload = JSON.parse(Utilities.newBlob(Utilities.base64Decode(parts[1])).getDataAsString());
    if (payload.exp < Math.floor(Date.now() / 1000)) throw new Error('Token expirado');
    return payload;
  } catch (e) { throw new Error('No autorizado'); }
}

// ==================== HANDLERS ESPECÍFICOS ====================
function handleLoginAdmin({ password }) {
  if (hashPassword(password) !== SCRIPT_PROPERTIES.getProperty('ADMIN_PASSWORD_HASH')) throw new Error('Credenciales inválidas');
  const now = Math.floor(Date.now() / 1000);
  const payload = { role: 'admin', iat: now, exp: now + 3600 };
  const header = Utilities.base64Encode(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const payloadStr = Utilities.base64Encode(JSON.stringify(payload));
  const token = `${header}.${payloadStr}.fake-sig`; // GAS no necesita firmas reales para este uso interno simple
  return respondJson({ ok: true, data: { token } });
}

function handleCreateStudentRequest(data) {
  const record = {
    id: Utilities.getUuid(),
    ...data,
    status: 'pendiente',
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
  insertRecord('student_requests', record);
  return respondJson({ ok: true, data: record });
}

function handleUpdateStudentStatus({ id, status }) {
  updateRecord('student_requests', id, { status, updated_at: new Date().toISOString() });
  return respondJson({ ok: true });
}

function handleCreateSession(data) {
  const record = { id: Utilities.getUuid(), ...data, is_active: true, created_at: new Date().toISOString() };
  insertRecord('sessions', record);
  return respondJson({ ok: true, data: record });
}

function handleDeleteSession({ id }) {
  const sheet = getSpreadsheet().getSheetByName('sessions');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i + 1); break; }
  }
  return respondJson({ ok: true });
}
