/**
 * ArteDelCantar Agenda - Backend Autónomo (Versión Corregida)
 * - Crea la hoja de cálculo y las hojas si no existen
 * - Maneja CORS automáticamente
 * - Contraseña inicial: admin123
 */

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

// ==================== CONFIGURACIÓN INICIAL ====================
function setupInitialConfig() {
  try {
    // 1. Crear o usar un Spreadsheet
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.create('ArteDelCantar_Agenda');
    }
    const spreadsheetId = ss.getId();
    SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', spreadsheetId);

    // 2. Crear las hojas necesarias (con sus encabezados)
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
        Logger.log(`Hoja creada: ${sheetName}`);
      } else {
        // Verificar que los encabezados existan (opcional)
        const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (existingHeaders.length !== headers.length) {
          Logger.log(`Advertencia: La hoja ${sheetName} tiene estructura diferente.`);
        }
      }
    }

    // 3. Configurar propiedades de seguridad (solo si no existen)
    if (!SCRIPT_PROPERTIES.getProperty('ADMIN_PASSWORD_HASH')) {
      SCRIPT_PROPERTIES.setProperty('ADMIN_PASSWORD_HASH', hashPassword('admin123'));
    }
    if (!SCRIPT_PROPERTIES.getProperty('JWT_SECRET')) {
      SCRIPT_PROPERTIES.setProperty('JWT_SECRET', Utilities.getUuid() + Utilities.getUuid());
    }
    if (!SCRIPT_PROPERTIES.getProperty('TOKEN_TTL_SECONDS')) {
      SCRIPT_PROPERTIES.setProperty('TOKEN_TTL_SECONDS', '3600');
    }

    Logger.log('✅ Configuración inicial completada. Spreadsheet ID: ' + spreadsheetId);
    return { ok: true, message: 'Configuración completada', spreadsheetId };
  } catch (error) {
    Logger.log('❌ Error en setupInitialConfig: ' + error.message);
    return { ok: false, error: error.message };
  }
}

// ==================== FUNCIONES AUXILIARES ====================
function getSpreadsheetId() {
  const id = SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Ejecuta primero setupInitialConfig()');
  return id;
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
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
      if ((h === 'available_weekdays' || h === 'available_time_blocks' || h === 'metadata_json') &&
          typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))) {
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
    if (h === 'available_weekdays' || h === 'available_time_blocks' || h === 'metadata_json') {
      if (Array.isArray(v) || typeof v === 'object') v = JSON.stringify(v);
    }
    return v !== undefined ? v : '';
  });
  sheet.appendRow(row);
  return record;
}

function updateRecord(sheetName, id, updates) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('id')] === id) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx === -1) throw new Error('Registro no encontrado');
  for (const [key, val] of Object.entries(updates)) {
    const col = headers.indexOf(key);
    if (col !== -1) {
      let finalVal = val;
      if (key === 'available_weekdays' || key === 'available_time_blocks' || key === 'metadata_json') {
        if (Array.isArray(val) || typeof val === 'object') finalVal = JSON.stringify(val);
      }
      sheet.getRange(rowIdx, col + 1).setValue(finalVal);
    }
  }
  return getRecordById(sheetName, id);
}

function deleteRecord(sheetName, id) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('id')] === id) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx !== -1) sheet.deleteRow(rowIdx);
}

function getRecordById(sheetName, id) {
  return getAllRecords(sheetName).find(r => r.id === id);
}

// ==================== AUTENTICACIÓN ====================
function hashPassword(pwd) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function createToken(payload) {
  const secret = SCRIPT_PROPERTIES.getProperty('JWT_SECRET');
  const ttl = parseInt(SCRIPT_PROPERTIES.getProperty('TOKEN_TTL_SECONDS') || '3600');
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'HS256', typ: 'JWT' };
  const tokenPayload = { ...payload, exp: now + ttl, iat: now };
  const headerB64 = Utilities.base64Encode(JSON.stringify(header)).replace(/=/g, '');
  const payloadB64 = Utilities.base64Encode(JSON.stringify(tokenPayload)).replace(/=/g, '');
  const signature = Utilities.computeHmacSha256Signature(headerB64 + '.' + payloadB64, secret);
  const sigB64 = Utilities.base64Encode(signature).replace(/=/g, '');
  return `${headerB64}.${payloadB64}.${sigB64}`;
}

function verifyToken(token) {
  if (!token) throw new Error('Token no proporcionado');
  const secret = SCRIPT_PROPERTIES.getProperty('JWT_SECRET');
  const parts = token.split('.');
  if (parts.length !== 3) throw new Error('Token malformado');
  const [headerB64, payloadB64, sigB64] = parts;
  const expectedSig = Utilities.computeHmacSha256Signature(headerB64 + '.' + payloadB64, secret);
  const expectedSigB64 = Utilities.base64Encode(expectedSig).replace(/=/g, '');
  if (sigB64 !== expectedSigB64) throw new Error('Firma inválida');
  const payload = JSON.parse(Utilities.newBlob(Utilities.base64Decode(payloadB64)).getDataAsString());
  if (payload.exp < Math.floor(Date.now() / 1000)) throw new Error('Token expirado');
  return payload;
}

// ==================== MANEJO HTTP CON CORS ====================
function doGet(e) {
  return handleCors(() => {
    const action = e?.parameter?.action;
    if (action === 'health') {
      return respondJson({ ok: true, data: { status: 'ok', timestamp: new Date().toISOString() } });
    }
    if (action === 'getActiveSessions') {
      const sessions = getAllRecords('sessions').filter(s => s.is_active === true && new Date(s.date) >= new Date().setHours(0,0,0,0));
      return respondJson({ ok: true, data: sessions });
    }
    return respondJson({ ok: false, error: 'Acción no soportada' }, 400);
  });
}

function doPost(e) {
  return handleCors(() => {
    let requestData;
    try {
      requestData = JSON.parse(e?.postData?.contents || '{}');
    } catch (err) {
      return respondJson({ ok: false, error: 'JSON inválido' }, 400);
    }
    const { action, payload = {}, token } = requestData;
    if (!action) return respondJson({ ok: false, error: 'Action requerida' }, 400);

    // Acciones públicas
    if (action === 'createDirectBooking') return handleCreateDirectBooking(payload);
    if (action === 'loginAdmin') return handleLoginAdmin(payload);

    // Acciones privadas (requieren token)
    let decoded;
    try {
      decoded = verifyToken(token);
      if (!decoded || decoded.role !== 'admin') throw new Error('No autorizado');
    } catch (err) {
      return respondJson({ ok: false, error: 'Token inválido o expirado' }, 401);
    }

    switch (action) {
      case 'getSessions': return handleGetSessions();
      case 'createSession': return handleCreateSession(payload);
      case 'updateSession': return handleUpdateSession(payload);
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
  if (result && typeof result.getHeaders === 'function') {
    result.setHeader('Access-Control-Allow-Origin', '*');
    result.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    result.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return result;
}

function respondJson(data, status = 200) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}

// ==================== LOGIN ADMIN ====================
function handleLoginAdmin({ password }) {
  const storedHash = SCRIPT_PROPERTIES.getProperty('ADMIN_PASSWORD_HASH');
  if (!storedHash) {
    return respondJson({ ok: false, error: 'Configuración incompleta. Ejecuta setupInitialConfig()' }, 500);
  }
  if (hashPassword(password) !== storedHash) {
    return respondJson({ ok: false, error: 'Contraseña incorrecta' }, 401);
  }
  const token = createToken({ role: 'admin' });
  return respondJson({ ok: true, data: { token } });
}

// ==================== SESIONES (horarios con fecha real) ====================
function handleGetSessions() {
  return respondJson({ ok: true, data: getAllRecords('sessions') });
}

function handleCreateSession(data) {
  try {
    if (!data.date || !/^\d{4}-\d{2}-\d{2}$/.test(data.date)) throw new Error('Fecha inválida (YYYY-MM-DD)');
    if (new Date(data.date) < new Date().setHours(0,0,0,0)) throw new Error('No se puede crear sesión en fecha pasada');
    if (!data.start_time || !/^([01]?[0-9]|2[0-3]):[0-5][0-9]$/.test(data.start_time)) throw new Error('Hora inicio inválida');
    if (!data.end_time || !/^([01]?[0-9]|2[0-3]):[0-5][0-9]$/.test(data.end_time)) throw new Error('Hora fin inválida');
    if (data.start_time >= data.end_time) throw new Error('La hora inicio debe ser menor que la hora fin');
    if (!['presencial', 'online', 'ambas'].includes(data.mode)) throw new Error('Modalidad inválida');
    let maxStudents = data.max_students ? parseInt(data.max_students) : null;
    if (maxStudents !== null && (isNaN(maxStudents) || maxStudents < 1)) throw new Error('Cupo máximo debe ser número positivo');

    const record = {
      id: Utilities.getUuid(),
      date: data.date,
      start_time: data.start_time,
      end_time: data.end_time,
      mode: data.mode,
      max_students: maxStudents,
      is_active: data.is_active !== undefined ? data.is_active : true,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    };
    insertRecord('sessions', record);
    return respondJson({ ok: true, data: record });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

function handleUpdateSession(data) {
  try {
    if (!data.id) throw new Error('ID requerido');
    const existing = getRecordById('sessions', data.id);
    if (!existing) throw new Error('Sesión no encontrada');
    const updates = { ...data, updated_at: new Date().toISOString() };
    updateRecord('sessions', data.id, updates);
    return respondJson({ ok: true, data: { ...existing, ...updates } });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

function handleDeleteSession({ id }) {
  try {
    if (!id) throw new Error('ID requerido');
    const assignments = getAllRecords('class_assignments').filter(a => a.session_id === id);
    if (assignments.length) throw new Error('No se puede eliminar una sesión con alumnos asignados');
    deleteRecord('sessions', id);
    return respondJson({ ok: true, data: { deleted: true } });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

// ==================== SOLICITUDES DE ALUMNOS ====================
function handleCreateStudentRequest(data) {
  try {
    if (!data.full_name || data.full_name.trim().length < 2) throw new Error('Nombre completo requerido');
    const whatsapp = data.whatsapp.replace(/\D/g, '');
    if (whatsapp.length < 8) throw new Error('WhatsApp inválido');
    if (!['principiante', 'intermedio', 'avanzado'].includes(data.vocal_level)) throw new Error('Nivel inválido');
    if (!data.goal) throw new Error('Objetivo requerido');
    if (!data.available_weekdays || !data.available_weekdays.length) throw new Error('Seleccione al menos un día');
    if (!data.available_time_blocks || !data.available_time_blocks.length) throw new Error('Seleccione al menos un bloque');
    const classesPerWeek = parseInt(data.classes_per_week);
    if (isNaN(classesPerWeek) || classesPerWeek < 1 || classesPerWeek > 7) throw new Error('Clases por semana entre 1 y 7');

    const now = new Date().toISOString();
    const record = {
      id: Utilities.getUuid(),
      full_name: data.full_name.trim(),
      whatsapp: whatsapp,
      age: data.age || '',
      vocal_level: data.vocal_level,
      goal: data.goal,
      available_weekdays: JSON.stringify(data.available_weekdays),
      available_time_blocks: JSON.stringify(data.available_time_blocks),
      classes_per_week: classesPerWeek,
      comments: data.comments || '',
      status: 'pendiente',
      internal_note: '',
      created_at: now,
      updated_at: now,
    };
    insertRecord('student_requests', record);
    return respondJson({ ok: true, data: { id: record.id, message: 'Solicitud enviada' } });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

function handleGetStudentRequests(filters = {}) {
  let records = getAllRecords('student_requests');
  if (filters.status) records = records.filter(r => r.status === filters.status);
  if (filters.vocal_level) records = records.filter(r => r.vocal_level === filters.vocal_level);
  if (filters.day) records = records.filter(r => JSON.parse(r.available_weekdays).includes(filters.day));
  if (filters.block) records = records.filter(r => JSON.parse(r.available_time_blocks).includes(filters.block));
  return respondJson({ ok: true, data: records });
}

function handleUpdateStudentStatus({ id, status }) {
  try {
    const valid = ['pendiente', 'contactado', 'asignado', 'descartado'];
    if (!valid.includes(status)) throw new Error('Estado inválido');
    const existing = getRecordById('student_requests', id);
    if (!existing) throw new Error('Solicitud no encontrada');
    const updated = { ...existing, status, updated_at: new Date().toISOString() };
    updateRecord('student_requests', id, updated);
    return respondJson({ ok: true, data: updated });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

function handleUpdateStudentNote({ id, internal_note }) {
  try {
    const existing = getRecordById('student_requests', id);
    if (!existing) throw new Error('Solicitud no encontrada');
    const updated = { ...existing, internal_note: internal_note || '', updated_at: new Date().toISOString() };
    updateRecord('student_requests', id, updated);
    return respondJson({ ok: true, data: updated });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

// ==================== ASIGNACIONES ====================
function handleGetAssignments() {
  return respondJson({ ok: true, data: getAllRecords('class_assignments') });
}

function handleCreateAssignment({ student_request_id, session_id }) {
  try {
    const student = getRecordById('student_requests', student_request_id);
    const session = getRecordById('sessions', session_id);
    if (!student || !session) throw new Error('Alumno o sesión no existe');
    if (session.is_active !== true) throw new Error('La sesión no está activa');
    const assignments = getAllRecords('class_assignments');
    if (assignments.some(a => a.student_request_id === student_request_id)) throw new Error('Este alumno ya tiene una asignación');
    const assignedCount = assignments.filter(a => a.session_id === session_id).length;
    if (session.max_students && assignedCount >= session.max_students) throw new Error(`Cupo lleno (máximo ${session.max_students})`);

    const newAssign = {
      id: Utilities.getUuid(),
      student_request_id,
      session_id,
      assigned_at: new Date().toISOString(),
    };
    insertRecord('class_assignments', newAssign);
    // Cambiar estado del alumno a "asignado"
    const updatedStudent = { ...student, status: 'asignado', updated_at: new Date().toISOString() };
    updateRecord('student_requests', student_request_id, updatedStudent);
    return respondJson({ ok: true, data: newAssign });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

function handleDeleteAssignment({ id }) {
  try {
    const assign = getRecordById('class_assignments', id);
    if (!assign) throw new Error('Asignación no encontrada');
    const student = getRecordById('student_requests', assign.student_request_id);
    if (student && student.status === 'asignado') {
      const updated = { ...student, status: 'pendiente', updated_at: new Date().toISOString() };
      updateRecord('student_requests', student.id, updated);
    }
    deleteRecord('class_assignments', id);
    return respondJson({ ok: true, data: { deleted: true } });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
}

// ==================== RESERVA DIRECTA (para el widget) ====================
function handleCreateDirectBooking(payload) {
  try {
    // Crear la solicitud del alumno
    const studentResult = handleCreateStudentRequest(payload);
    const studentData = JSON.parse(studentResult.getContent());
    if (!studentData.ok) throw new Error(studentData.error);
    const studentId = studentData.data.id;

    // Asignar a la sesión elegida
    const assignResult = handleCreateAssignment({ student_request_id: studentId, session_id: payload.session_id });
    const assignData = JSON.parse(assignResult.getContent());
    if (!assignData.ok) throw new Error(assignData.error);

    return respondJson({ ok: true, data: { message: 'Clase reservada con éxito', studentId, assignment: assignData.data } });
  } catch (err) {
    return respondJson({ ok: false, error: err.message }, 400);
  }
  // 1. Esta función atiende la petición pre-flight del navegador
function doOptions(e) {
  return createCORSResponse(JSON.stringify({ "status": "preflight" }));
}
// 2. Integra tu lógica en doPost
function doPost(e) {
  try {
    // --- AQUÍ VA TU LÓGICA ACTUAL ---
    // (Ejemplo: procesar datos, guardar en Sheets, etc.)
    var data = JSON.parse(e.postData.contents);
    // ... tu código de procesamiento ...
    
    var result = { "status": "success", "message": "Datos recibidos" };
    // ---------------------------------
    
    return createCORSResponse(JSON.stringify(result));
    
  } catch (err) {
    return createCORSResponse(JSON.stringify({ "status": "error", "message": err.toString() }));
  }
}

// 3. Esta función es la que añade los headers mágicos
function createCORSResponse(content) {
  var output = ContentService.createTextOutput(content);
  output.setMimeType(ContentService.MimeType.JSON);
  
  output.addHeader("Access-Control-Allow-Origin", "*");
  output.addHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  output.addHeader("Access-Control-Allow-Headers", "Content-Type");
  
  return output;
}
}