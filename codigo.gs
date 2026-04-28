/**
 * ARTE DEL CANTAR - SISTEMA OPERATIVO V3
 * Backend con autenticación HMAC, soft delete, y gestión real de asignaciones
 */

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const SHEETS_CONFIG = {
  'sessions': ['id', 'day', 'start_time', 'end_time', 'mode', 'max_students', 'current_students', 'is_active', 'deleted'],
  'student_requests': [
    'id', 'full_name', 'whatsapp', 'age', 'vocal_level', 'commitment', 'goal',
    'availability_json', 'classes_per_week', 'status', 'internal_note',
    'created_at', 'updated_at', 'deleted'
  ],
  'class_assignments': ['id', 'student_id', 'session_id', 'assigned_at', 'deleted'],
  'notifications': ['id', 'type', 'title', 'message', 'student_id', 'is_read', 'created_at', 'read_at', 'deleted'],
  'audit_log': ['id', 'event_type', 'action', 'entity_type', 'entity_id', 'payload_json', 'user_agent', 'created_at', 'deleted']
};

// ==================== 1. CONFIGURACIÓN E INFRAESTRUCTURA ====================
function setup() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.create('ArteDelCantar_CRM_Agenda');
    SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', ss.getId());

    ensureInfrastructure(ss);

    if (!SCRIPT_PROPERTIES.getProperty('JWT_SECRET')) {
      SCRIPT_PROPERTIES.setProperty('JWT_SECRET', Utilities.getUuid());
    }

    let adminPass = SCRIPT_PROPERTIES.getProperty('ADMIN_PASS');
    let createdAdminPass = false;
    if (!adminPass) {
      adminPass = 'admin123';
      SCRIPT_PROPERTIES.setProperty('ADMIN_PASS', adminPass);
      createdAdminPass = true;
    }

    return respondJson({
      ok: true,
      message: 'Sistema Operativo inicializado correctamente',
      data: {
        createdAdminPass,
        adminPass: createdAdminPass ? adminPass : undefined
      }
    });
  } catch (e) {
    return respondJson({ ok: false, error: e.message }, 500);
  }
}

// ==================== 2. MANEJO DE PETICIONES (API) ====================
function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }
function doGet(e) { return respondJson({ ok: false, error: 'Acceso solo por POST' }, 405); }

function doPost(e) {
  try {
    ensureInfrastructure();
    const requestData = JSON.parse(e.postData.contents);
    const { action, payload = {}, token, event_type = 'api' } = requestData;
    const userAgent = (e && e.parameter && e.parameter.ua) ? String(e.parameter.ua) : '';

    // Acciones públicas (sin token)
    if (action === 'setup') return setup();
    if (action === 'createStudentRequest') return handleCreateStudentRequest(payload, { event_type, userAgent });
    if (action === 'login') return handleLogin(payload, { event_type, userAgent });
    if (action === 'setAdminPassword') return handleSetAdminPassword(payload, { event_type, userAgent });
    if (action === 'setTeacherContact') return handleSetTeacherContact(payload, { event_type, userAgent });
    if (action === 'getNotifications') return handleGetNotifications(payload);
    if (action === 'markNotificationRead') return handleMarkNotificationRead(payload, { event_type, userAgent });
    if (action === 'trackInteraction') return handleTrackInteraction(payload, { event_type, userAgent });
    if (action === 'getPublicSessions') {
      const sessions = getAllRecords('sessions').filter(s => s.is_active === true && s.deleted !== true);
      return respondJson({ ok: true, data: sessions });
    }

    // Acciones protegidas (requieren token válido)
    const decoded = verifyAuth(token);
    const auditContext = { event_type, userAgent, actor: decoded.role || 'admin' };

    switch (action) {
      case 'getDashboardStats': return handleGetDashboardStats(auditContext);
      case 'getStudents': return respondJson({ ok: true, data: getAllRecords('student_requests') });
      case 'getSessions': return respondJson({ ok: true, data: getAllRecords('sessions') });
      case 'getAvailableSessions': return handleGetAvailableSessions(auditContext);
      case 'getFullSchedule': return handleGetFullSchedule(auditContext);
      case 'updateStudent': return handleUpdateStudent(payload, auditContext);
      case 'deleteStudent': return handleDeleteStudent(payload, auditContext);
      case 'bulkImport': return handleBulkImport(payload, auditContext);
      case 'createSession': return handleCreateSession(payload, auditContext);
      case 'deleteSession': return handleDeleteSession(payload, auditContext);
      case 'assignStudent': return handleAssignStudent(payload, auditContext);
      case 'unassignStudent': return handleUnassignStudent(payload, auditContext);
      default: throw new Error('Acción no reconocida');
    }
  } catch (e) {
    console.error(e);
    return respondJson({ ok: false, error: e.message }, 400);
  }
}

// ==================== 3. LÓGICA DE NEGOCIO (PÚBLICA Y PRIVADA) ====================

function handleCreateStudentRequest(data, context = {}) {
  const cleanPhone = data.whatsapp ? data.whatsapp.replace(/\D/g, '') : '';
  const existing = getAllRecords('student_requests').find(s => s.whatsapp === cleanPhone && s.deleted !== true);

  if (existing) {
    // Actualizar disponibilidad y otros campos relevantes
    const updates = {
      availability_json: JSON.stringify(data.availability || {}),
      updated_at: new Date().toISOString(),
      status: 'pendiente'
    };
    if (data.full_name) updates.full_name = data.full_name;
    if (data.vocal_level) updates.vocal_level = data.vocal_level;
    if (data.goal) updates.goal = data.goal;
    if (data.classes_per_week) updates.classes_per_week = data.classes_per_week;
    updateRecord('student_requests', existing.id, updates);
    logAudit('student_request_updated', 'createStudentRequest', 'student_request', existing.id, data, context);
    return respondJson({ ok: true, message: 'Disponibilidad actualizada exitosamente' });
  } else {
    const id = Utilities.getUuid();
    const record = {
      id,
      full_name: data.full_name,
      whatsapp: cleanPhone,
      age: data.age || '',
      vocal_level: data.vocal_level,
      commitment: data.commitment || 'interesado',
      goal: data.goal,
      availability_json: JSON.stringify(data.availability || {}),
      classes_per_week: data.classes_per_week || 1,
      status: 'pendiente',
      internal_note: '',
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
      deleted: false
    };
    insertRecord('student_requests', record);
    logAudit('student_request_created', 'createStudentRequest', 'student_request', id, data, context);
    const teacherWhatsappUrl = createNotificationAndMaybeAlertTeacher({
      type: 'new_student_request',
      title: 'Nueva solicitud de alumno',
      message: `Se registró un nuevo alumno: ${data.full_name || 'Sin nombre'}`,
      student_id: id
    });
    return respondJson({ ok: true, message: 'Alumno registrado exitosamente', data: { id, teacher_whatsapp_url: teacherWhatsappUrl } });
  }
}

function handleGetDashboardStats(context = {}) {
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
  logAudit('dashboard_viewed', 'getDashboardStats', 'dashboard', '', { stats }, context);
  return respondJson({ ok: true, data: stats });
}

function handleGetAvailableSessions(context = {}) {
  const sessions = getAllRecords('sessions').filter(s => s.is_active === true && s.deleted !== true);
  const available = sessions.filter(s => s.current_students < s.max_students);
  logAudit('available_sessions_viewed', 'getAvailableSessions', 'sessions', '', { count: available.length }, context);
  return respondJson({ ok: true, data: available });
}

function handleGetFullSchedule(context = {}) {
  const sessions = getAllRecords('sessions').filter(s => s.is_active === true && s.deleted !== true);
  const assignments = getAllRecords('class_assignments').filter(a => a.deleted !== true);
  const students = getAllRecords('student_requests');

  const schedule = sessions.map(session => {
    const assignedStudents = assignments
      .filter(a => a.session_id === session.id)
      .map(a => students.find(s => s.id === a.student_id))
      .filter(s => s && s.deleted !== true);
    return {
      ...session,
      students: assignedStudents
    };
  });
  logAudit('full_schedule_viewed', 'getFullSchedule', 'schedule', '', { sessions: schedule.length }, context);
  return respondJson({ ok: true, data: schedule });
}

function handleAssignStudent({ student_id, session_id }, context = {}) {
  const session = getRecordById('sessions', session_id);
  if (!session || session.deleted) throw new Error('Sesión no encontrada');
  if (session.current_students >= session.max_students) throw new Error('Sesión llena');

  const existingAssignment = getAllRecords('class_assignments').find(a => a.student_id === student_id && a.session_id === session_id && a.deleted !== true);
  if (existingAssignment) throw new Error('El alumno ya está asignado a esta sesión');

  const id = Utilities.getUuid();
  insertRecord('class_assignments', {
    id,
    student_id,
    session_id,
    assigned_at: new Date().toISOString(),
    deleted: false
  });

  // Actualizar contador de la sesión
  updateRecord('sessions', session_id, { current_students: session.current_students + 1 });
  // Cambiar estado del alumno a 'activo'
  updateRecord('student_requests', student_id, { status: 'activo', updated_at: new Date().toISOString() });
  logAudit('student_assigned', 'assignStudent', 'class_assignment', id, { student_id, session_id }, context);

  return respondJson({ ok: true, message: 'Alumno asignado correctamente' });
}

function handleUnassignStudent({ student_id, session_id }, context = {}) {
  const assignment = getAllRecords('class_assignments').find(a => a.student_id === student_id && a.session_id === session_id && a.deleted !== true);
  if (!assignment) throw new Error('Asignación no encontrada');

  // Soft delete la asignación
  updateRecord('class_assignments', assignment.id, { deleted: true });

  // Disminuir contador de la sesión
  const session = getRecordById('sessions', session_id);
  if (session) {
    updateRecord('sessions', session_id, { current_students: Math.max(0, session.current_students - 1) });
  }
  // Cambiar estado del alumno a 'pendiente' o el que corresponda
  updateRecord('student_requests', student_id, { status: 'pendiente', updated_at: new Date().toISOString() });
  logAudit('student_unassigned', 'unassignStudent', 'class_assignment', assignment.id, { student_id, session_id }, context);

  return respondJson({ ok: true, message: 'Alumno desasignado correctamente' });
}

// Resto de funciones: handleUpdateStudent, handleDeleteStudent (soft delete), handleCreateSession, handleDeleteSession, bulkImport
// Las mantengo similares pero con soft delete.

function handleUpdateStudent({ id, updates }, context = {}) {
  updates.updated_at = new Date().toISOString();
  updateRecord('student_requests', id, updates);
  logAudit('student_updated', 'updateStudent', 'student_request', id, { updates }, context);
  return respondJson({ ok: true });
}

function handleDeleteStudent({ id }, context = {}) {
  // Soft delete
  updateRecord('student_requests', id, { deleted: true });
  // Opcional: también desasignar de todas las sesiones
  const assignments = getAllRecords('class_assignments').filter(a => a.student_id === id && a.deleted !== true);
  assignments.forEach(a => {
    handleUnassignStudent({ student_id: id, session_id: a.session_id }, context);
  });
  logAudit('student_deleted', 'deleteStudent', 'student_request', id, {}, context);
  return respondJson({ ok: true, message: 'Alumno eliminado (soft delete)' });
}

function handleCreateSession(data, context = {}) {
  const record = {
    id: Utilities.getUuid(),
    day: data.day,
    start_time: data.start_time,
    end_time: data.end_time,
    mode: data.mode || 'presencial',
    max_students: parseInt(data.max_students) || 1,
    current_students: 0,
    is_active: true,
    deleted: false
  };
  insertRecord('sessions', record);
  logAudit('session_created', 'createSession', 'session', record.id, data, context);
  return respondJson({ ok: true, data: record });
}

function handleDeleteSession({ id }, context = {}) {
  // Soft delete
  updateRecord('sessions', id, { deleted: true });
  logAudit('session_deleted', 'deleteSession', 'session', id, {}, context);
  return respondJson({ ok: true });
}

function handleBulkImport({ numbers }, context = {}) {
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
        updated_at: new Date().toISOString(),
        deleted: false
      });
      imported++;
    }
  });
  logAudit('bulk_import_completed', 'bulkImport', 'student_request', '', { imported, numbers: Array.isArray(numbers) ? numbers.length : 0 }, context);
  return respondJson({ ok: true, data: { imported } });
}

// ==================== 4. AUTENTICACIÓN (HMAC) ====================
function handleLogin({ password }, context = {}) {
  const correctPass = SCRIPT_PROPERTIES.getProperty('ADMIN_PASS');
  if (!correctPass) throw new Error('Sistema no configurado: falta ADMIN_PASS');
  if (password !== correctPass) throw new Error('Contraseña incorrecta');

  const payload = {
    role: 'admin',
    exp: Date.now() + 86400000 // 24h
  };
  const payloadStr = JSON.stringify(payload);
  const secret = SCRIPT_PROPERTIES.getProperty('JWT_SECRET');
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payloadStr, secret);
  const token = Utilities.base64Encode(payloadStr) + '.' + Utilities.base64Encode(signature);
  logAudit('login_success', 'login', 'auth', '', {}, context);
  return respondJson({ ok: true, data: { token } });
}

function handleSetAdminPassword({ currentPassword, newPassword }, context = {}) {
  const correctPass = SCRIPT_PROPERTIES.getProperty('ADMIN_PASS');
  if (!correctPass) throw new Error('Sistema no configurado: falta ADMIN_PASS');
  if (currentPassword !== correctPass) throw new Error('Contraseña actual incorrecta');
  if (!newPassword || String(newPassword).length < 8) throw new Error('La nueva contraseña debe tener al menos 8 caracteres');
  SCRIPT_PROPERTIES.setProperty('ADMIN_PASS', String(newPassword));
  logAudit('password_changed', 'setAdminPassword', 'auth', '', {}, context);
  return respondJson({ ok: true, message: 'Contraseña actualizada' });
}

function handleSetTeacherContact({ currentPassword, teacherEmail = '', teacherWhatsapp = '' }, context = {}) {
  const correctPass = SCRIPT_PROPERTIES.getProperty('ADMIN_PASS');
  if (!correctPass) throw new Error('Sistema no configurado: falta ADMIN_PASS');
  if (currentPassword !== correctPass) throw new Error('Contraseña actual incorrecta');

  if (teacherEmail) SCRIPT_PROPERTIES.setProperty('TEACHER_EMAIL', String(teacherEmail).trim());
  if (teacherWhatsapp) SCRIPT_PROPERTIES.setProperty('TEACHER_WHATSAPP', String(teacherWhatsapp).trim());
  logAudit('teacher_contact_updated', 'setTeacherContact', 'settings', '', { teacherEmail, teacherWhatsapp }, context);

  return respondJson({ ok: true, message: 'Datos de contacto actualizados' });
}

function handleResetAdminPassword({ currentPassword }) {
  const correctPass = SCRIPT_PROPERTIES.getProperty('ADMIN_PASS');
  if (!correctPass) throw new Error('Sistema no configurado: falta ADMIN_PASS');
  if (currentPassword !== correctPass) throw new Error('Contraseña actual incorrecta');
  SCRIPT_PROPERTIES.setProperty('ADMIN_PASS', 'admin123');
  return respondJson({ ok: true, message: 'Contraseña restablecida a admin123' });
}

function handleGetNotifications(payload) {
  const limit = Math.max(1, Math.min(parseInt(payload.limit, 10) || 20, 100));
  const items = getAllRecords('notifications')
    .filter(n => n.deleted !== true)
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at))
    .slice(0, limit);
  return respondJson({ ok: true, data: items });
}

function handleMarkNotificationRead({ id }, context = {}) {
  if (!id) throw new Error('Notification id requerido');
  updateRecord('notifications', id, { is_read: true, read_at: new Date().toISOString() });
  logAudit('notification_read', 'markNotificationRead', 'notification', id, {}, context);
  return respondJson({ ok: true });
}

function handleTrackInteraction({ event_name = '', target = '', route = '', details = {} }, context = {}) {
  const payload = {
    event_name: String(event_name || ''),
    target: String(target || ''),
    route: String(route || ''),
    details: details || {}
  };
  logAudit('ui_interaction', String(event_name || 'trackInteraction'), 'ui', String(target || ''), payload, context);
  return respondJson({ ok: true });
}

function createNotificationAndMaybeAlertTeacher({ type, title, message, student_id }) {
  const id = Utilities.getUuid();
  const record = {
    id,
    type,
    title,
    message,
    student_id: student_id || '',
    is_read: false,
    created_at: new Date().toISOString(),
    read_at: '',
    deleted: false
  };
  insertRecord('notifications', record);

  const teacherEmail = SCRIPT_PROPERTIES.getProperty('TEACHER_EMAIL');
  if (teacherEmail) {
    try {
      MailApp.sendEmail(teacherEmail, title, message);
    } catch (e) {
      console.error('Email notification failed', e);
    }
  }

  const teacherWhatsapp = SCRIPT_PROPERTIES.getProperty('TEACHER_WHATSAPP') || '';
  const normalized = teacherWhatsapp.replace(/\D/g, '');
  return normalized ? `https://wa.me/${normalized}?text=${encodeURIComponent(message)}` : '';
}

function logAudit(eventType, action, entityType, entityId, payload, context = {}) {
  try {
    insertRecord('audit_log', {
      id: Utilities.getUuid(),
      event_type: String(eventType || 'api'),
      action: String(action || ''),
      entity_type: String(entityType || ''),
      entity_id: String(entityId || ''),
      payload_json: payload || {},
      user_agent: String(context.userAgent || ''),
      created_at: new Date().toISOString(),
      deleted: false
    });
  } catch (e) {
    console.error('audit_log failed', e);
  }
}

function verifyAuth(token) {
  if (!token) throw new Error('No autorizado');
  const parts = token.split('.');
  if (parts.length !== 2) throw new Error('Token inválido');
  const payloadStr = Utilities.newBlob(Utilities.base64Decode(parts[0])).getDataAsString();
  const signature = Utilities.base64Decode(parts[1]);
  const secret = SCRIPT_PROPERTIES.getProperty('JWT_SECRET');
  const expectedSignature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payloadStr, secret);
  if (signature.join('') !== expectedSignature.join('')) throw new Error('Firma incorrecta');
  const payload = JSON.parse(payloadStr);
  if (payload.exp < Date.now()) throw new Error('Token expirado');
  return payload;
}

// ==================== 5. UTILIDADES CORE (CON SOFT DELETE) ====================
function respondJson(data, status = 200) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet() {
  ensureInfrastructure();
  return SpreadsheetApp.openById(SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID'));
}

function ensureInfrastructure(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID')) || SpreadsheetApp.create('ArteDelCantar_CRM_Agenda');
  if (!SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID')) {
    SCRIPT_PROPERTIES.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  }

  Object.entries(SHEETS_CONFIG).forEach(([sheetName, headers]) => {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      return;
    }
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach((header, idx) => {
      if (existingHeaders[idx] !== header && !existingHeaders.includes(header)) {
        sheet.getRange(1, existingHeaders.length + 1).setValue(header);
        existingHeaders.push(header);
      }
    });
  });
  return spreadsheet;
}

function getAllRecords(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (h.endsWith('_json') && typeof val === 'string') {
        try { val = JSON.parse(val); } catch(e) { val = {}; }
      }
      obj[h] = val;
    });
    return obj;
  }).filter(obj => obj.deleted !== true); // soft delete filter
}

function insertRecord(sheetName, record) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet no encontrada: ${sheetName}`);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    const val = record[h];
    if (val === undefined) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });
  sheet.appendRow(row);
}

function updateRecord(sheetName, id, updates) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet no encontrada: ${sheetName}`);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      for (const [key, val] of Object.entries(updates)) {
        const col = headers.indexOf(key);
        if (col !== -1) {
          const valueToSet = (typeof val === 'object') ? JSON.stringify(val) : val;
          sheet.getRange(i + 1, col + 1).setValue(valueToSet);
        }
      }
      return;
    }
  }
}

function getRecordById(sheetName, id) {
  return getAllRecords(sheetName).find(r => r.id === id);
}
