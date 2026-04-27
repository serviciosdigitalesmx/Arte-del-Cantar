/**
 * ARTE DEL CANTAR - SISTEMA OPERATIVO PARA PROFESORES (V2)
 * Backend: Google Apps Script API
 * DB: Google Sheets
 */

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

// ==================== 1. CONFIGURACIÓN E INFRAESTRUCTURA ====================
function setup() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.create('ArteDelCantar_CRM_Agenda');
    SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', ss.getId());

    const sheetsConfig = {
      'sessions': ['id', 'day', 'start_time', 'end_time', 'mode', 'max_students', 'current_students', 'is_active'],
      'student_requests': [
        'id', 'full_name', 'whatsapp', 'age', 'vocal_level', 
        'commitment', 'goal', 'availability_json', 'classes_per_week', 
        'status', 'internal_note', 'created_at', 'updated_at'
      ],
      'class_assignments': ['id', 'student_id', 'session_id', 'assigned_at']
    };

    for (const [sheetName, headers] of Object.entries(sheetsConfig)) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      }
    }

    if (!SCRIPT_PROPERTIES.getProperty('ADMIN_PASS')) SCRIPT_PROPERTIES.setProperty('ADMIN_PASS', 'admin123');
    if (!SCRIPT_PROPERTIES.getProperty('JWT_SECRET')) SCRIPT_PROPERTIES.setProperty('JWT_SECRET', Utilities.getUuid());

    return respondJson({ ok: true, message: 'Sistema Operativo inicializado' });
  } catch (e) { return respondJson({ ok: false, error: e.message }, 500); }
}

// ==================== 2. MANEJO DE PETICIONES (API) ====================
function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }
function doGet(e) { return respondJson({ ok: false, error: 'Acceso solo por POST' }, 405); }

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const { action, payload = {}, token } = requestData;

    // Públicos
    if (action === 'setup') return setup();
    if (action === 'createStudentRequest') return handleCreateStudentRequest(payload);
    if (action === 'login') return handleLogin(payload);
    if (action === 'getPublicSessions') return respondJson({ ok: true, data: getAllRecords('sessions').filter(s => s.is_active) });

    // Privados (Auth)
    const decoded = verifyAuth(token);
    
    switch (action) {
      case 'getDashboardStats': return handleGetDashboardStats();
      case 'getStudents': return respondJson({ ok: true, data: getAllRecords('student_requests') });
      case 'getSessions': return respondJson({ ok: true, data: getAllRecords('sessions') });
      case 'updateStudent': return handleUpdateStudent(payload);
      case 'bulkImport': return handleBulkImport(payload);
      case 'createSession': return handleCreateSession(payload);
      case 'deleteSession': return handleDeleteSession(payload);
      case 'assignStudent': return handleAssignStudent(payload);
      case 'unassignStudent': return handleUnassignStudent(payload);
      default: throw new Error('Acción no reconocida');
    }
  } catch (e) { return respondJson({ ok: false, error: e.message }, 400); }
}

// ==================== 3. LÓGICA DE NEGOCIO ====================

function handleCreateStudentRequest(data) {
  const id = Utilities.getUuid();
  const record = {
    id,
    full_name: data.full_name,
    whatsapp: data.whatsapp.replace(/\D/g, ''),
    age: data.age || '',
    vocal_level: data.vocal_level,
    commitment: data.commitment || 'interesado',
    goal: data.goal,
    availability_json: JSON.stringify(data.availability || {}),
    classes_per_week: data.classes_per_week || 1,
    status: 'pendiente',
    internal_note: '',
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
  insertRecord('student_requests', record);
  return respondJson({ ok: true, data: { id } });
}

function handleGetDashboardStats() {
  const students = getAllRecords('student_requests');
  const sessions = getAllRecords('sessions');
  
  const stats = {
    total: students.length,
    pending: students.filter(s => s.status === 'pendiente').length,
    contacted: students.filter(s => s.status === 'contactado').length,
    assigned: students.filter(s => s.status === 'asignado').length,
    active: students.filter(s => s.status === 'activo').length,
    free_slots: sessions.reduce((acc, s) => acc + (s.max_students - s.current_students), 0)
  };
  return respondJson({ ok: true, data: stats });
}

function handleUpdateStudent({ id, updates }) {
  updates.updated_at = new Date().toISOString();
  updateRecord('student_requests', id, updates);
  return respondJson({ ok: true });
}

function handleAssignStudent({ student_id, session_id }) {
  const session = getRecordById('sessions', session_id);
  if (session.current_students >= session.max_students) throw new Error('Sesión llena');
  
  const id = Utilities.getUuid();
  insertRecord('class_assignments', { id, student_id, session_id, assigned_at: new Date().toISOString() });
  
  // Update counts and status
  updateRecord('sessions', session_id, { current_students: session.current_students + 1 });
  updateRecord('student_requests', student_id, { status: 'asignado', updated_at: new Date().toISOString() });
  
  return respondJson({ ok: true });
}

function handleCreateSession(data) {
  const record = {
    id: Utilities.getUuid(),
    day: data.day,
    start_time: data.start_time,
    end_time: data.end_time,
    mode: data.mode || 'presencial',
    max_students: parseInt(data.max_students) || 1,
    current_students: 0,
    is_active: true
  };
  insertRecord('sessions', record);
  return respondJson({ ok: true, data: record });
}

// ==================== 4. UTILIDADES CORE ====================

function respondJson(data, status = 200) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function handleLogin({ password }) {
  if (password !== SCRIPT_PROPERTIES.getProperty('ADMIN_PASS')) throw new Error('Password incorrecto');
  const token = Utilities.base64Encode(JSON.stringify({ role: 'admin', exp: Date.now() + 86400000 }));
  return respondJson({ ok: true, data: { token } });
}

function verifyAuth(token) {
  if (!token) throw new Error('No autorizado');
  return JSON.parse(Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString());
}

function getSpreadsheet() { return SpreadsheetApp.openById(SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID')); }

function getAllRecords(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (h.endsWith('_json')) { try { val = JSON.parse(val); } catch(e){} }
      obj[h] = val;
    });
    return obj;
  });
}

function insertRecord(sheetName, record) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => record[h] !== undefined ? record[h] : '');
  sheet.appendRow(row);
}

function updateRecord(sheetName, id, updates) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      for (const [key, val] of Object.entries(updates)) {
        const col = headers.indexOf(key);
        if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(typeof val === 'object' ? JSON.stringify(val) : val);
      }
      return;
    }
  }
}

function getRecordById(sheetName, id) { return getAllRecords(sheetName).find(r => r.id === id); }
function handleDeleteSession({ id }) {
  const sheet = getSpreadsheet().getSheetByName('sessions');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i + 1); break; }
  }
  return respondJson({ ok: true });
}
function handleBulkImport({ numbers }) {
  const existing = getAllRecords('student_requests').map(s => s.whatsapp);
  let imported = 0;
  
  numbers.forEach(num => {
    const clean = num.replace(/\D/g, '');
    if (clean.length >= 8 && !existing.includes(clean)) {
      const id = Utilities.getUuid();
      insertRecord('student_requests', {
        id,
        full_name: `[Importado] ${clean}`,
        whatsapp: clean,
        status: 'pendiente',
        commitment: 'desconocido',
        goal: 'Importado de WhatsApp',
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      });
      imported++;
    }
  });
  return respondJson({ ok: true, data: { imported } });
}
