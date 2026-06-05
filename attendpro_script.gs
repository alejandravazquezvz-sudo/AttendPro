// ══════════════════════════════════════════════════════════════════
//  AttendPro — Google Apps Script v2.2
//  INSTRUCCIONES:
//  1. Pega este código en script.google.com
//  2. Implementar → Nueva implementación → Aplicación web
//     - Ejecutar como: Yo
//     - Quién tiene acceso: Cualquier persona
//  3. Copia la URL y pégala en attendpro.html donde dice SCRIPT_URL
// ══════════════════════════════════════════════════════════════════

// Nombres de las hojas (cámbialos si los tuyos son diferentes)
const SHEET_USUARIOS  = 'Usuarios';
const SHEET_REGISTROS = 'Registros';

// ── ENTRY POINTS ────────────────────────────────────────────────
function doGet(e) {
  const action   = (e.parameter.action || '').trim();
  const callback = e.parameter.callback;
  let result;

  try {
    if      (action === 'login')     result = handleLogin(e.parameter);
    else if (action === 'register')  result = handleRegister(e.parameter);
    else if (action === 'getUsers')  result = handleGetUsers();
    else if (action === 'addRecord') result = handleAddRecord(e.parameter);
    else result = { status: 'AttendPro API activa ✓', version: '2.2' };
  } catch(err) {
    result = { error: err.message };
  }

  const json = JSON.stringify(result);

  // JSONP para evitar CORS desde el navegador
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${json})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = (body.action || '').trim();
    if      (action === 'addRecord') result = handleAddRecord(body);
    else if (action === 'register')  result = handleRegister(body);
    else result = { error: 'Acción no reconocida' };
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ─────────────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    // Formato de encabezado
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a6fe8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
  return sheet;
}

// ── LOGIN ────────────────────────────────────────────────────────
function handleLogin(params) {
  const user = (params.user || '').trim().toLowerCase();
  const pass = (params.pass || '').trim();
  if (!user || !pass) return { success: false, error: 'Faltan credenciales' };

  const sheet = getOrCreateSheet(SHEET_USUARIOS, ['Nombre', 'Usuario', 'Contraseña', 'Área/Rol', 'Fecha de registro']);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowUser = String(data[i][1]).trim().toLowerCase();
    const rowPass = String(data[i][2]).trim();
    if (rowUser === user && rowPass === pass) {
      return {
        success: true,
        user: {
          nombre: String(data[i][0]),
          user:   String(data[i][1]),
          pass:   String(data[i][2]),
          role:   String(data[i][3]),
        }
      };
    }
  }
  return { success: false, error: 'Usuario o contraseña incorrectos' };
}

// ── REGISTER ─────────────────────────────────────────────────────
function handleRegister(params) {
  const nombre = (params.nombre || '').trim();
  const user   = (params.user   || '').trim().toLowerCase();
  const pass   = (params.pass   || '').trim();
  const role   = (params.role   || '').trim();

  if (!nombre || !user || !pass) {
    return { success: false, error: 'Faltan datos (nombre, usuario o contraseña)' };
  }
  if (pass.length < 4) {
    return { success: false, error: 'La contraseña debe tener al menos 4 caracteres' };
  }

  const sheet = getOrCreateSheet(SHEET_USUARIOS, ['Nombre', 'Usuario', 'Contraseña', 'Área/Rol', 'Fecha de registro']);
  const data  = sheet.getDataRange().getValues();

  // Verificar si ya existe
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === user) {
      return { success: false, error: 'El usuario "' + user + '" ya existe' };
    }
  }

  // Agregar nuevo usuario
  sheet.appendRow([nombre, user, pass, role, new Date()]);
  return { success: true, user };
}

// ── GET USERS ────────────────────────────────────────────────────
function handleGetUsers() {
  const sheet = getOrCreateSheet(SHEET_USUARIOS, ['Nombre', 'Usuario', 'Contraseña', 'Área/Rol', 'Fecha de registro']);
  const data  = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue; // saltar filas vacías
    users.push({
      nombre: String(data[i][0]),
      user:   String(data[i][1]).trim(),
      pass:   String(data[i][2]).trim(),
      role:   String(data[i][3]),
    });
  }
  return { users };
}

// ── ADD RECORD ───────────────────────────────────────────────────
function handleAddRecord(params) {
  const sheet = getOrCreateSheet(SHEET_REGISTROS, ['Nombre', 'Tipo', 'Hora', 'Fecha', 'Área/Rol', 'Manual/Online', 'Timestamp ISO']);

  const name      = String(params.name      || '').trim();
  const type      = String(params.type      || '').trim();
  const role      = String(params.role      || '').trim();
  const tsRaw     = params.timestamp;
  const isManual  = params.manual ? 'Manual' : 'Online';

  if (!name || !type) return { success: false, error: 'Faltan datos del registro' };

  let ts;
  try {
    ts = tsRaw ? new Date(tsRaw) : new Date();
  } catch(e) {
    ts = new Date();
  }

  const hora  = Utilities.formatDate(ts, 'America/Mexico_City', 'HH:mm:ss');
  const fecha = Utilities.formatDate(ts, 'America/Mexico_City', 'dd/MM/yyyy');

  sheet.appendRow([name, type, hora, fecha, role, isManual, ts.toISOString()]);
  return { success: true };
}
