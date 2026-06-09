// ══════════════════════════════════════════════════════════════════
//  AttendPro — Google Apps Script v2.3
//  INSTRUCCIONES:
//  1. Pega este código en script.google.com
//  2. Implementar → Nueva implementación → Aplicación web
//     - Ejecutar como: Yo
//     - Quién tiene acceso: Cualquier persona
//  3. Copia la URL y pégala en index.html donde dice SCRIPT_URL
// ══════════════════════════════════════════════════════════════════

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
    else result = { status: 'AttendPro API activa ✓', version: '2.3' };
  } catch(err) {
    result = { error: err.message };
  }

  const json = JSON.stringify(result);
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
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a6fe8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getFecha(ts) {
  return Utilities.formatDate(ts, 'America/Mexico_City', 'dd/MM/yyyy');
}
function getHora(ts) {
  return Utilities.formatDate(ts, 'America/Mexico_City', 'HH:mm:ss');
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

  if (!nombre || !user || !pass)
    return { success: false, error: 'Faltan datos (nombre, usuario o contraseña)' };
  if (pass.length < 4)
    return { success: false, error: 'La contraseña debe tener al menos 4 caracteres' };

  const sheet = getOrCreateSheet(SHEET_USUARIOS, ['Nombre', 'Usuario', 'Contraseña', 'Área/Rol', 'Fecha de registro']);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === user)
      return { success: false, error: 'El usuario "' + user + '" ya existe' };
  }

  sheet.appendRow([nombre, user, pass, role, new Date()]);
  return { success: true, user };
}

// ── GET USERS ────────────────────────────────────────────────────
function handleGetUsers() {
  const sheet = getOrCreateSheet(SHEET_USUARIOS, ['Nombre', 'Usuario', 'Contraseña', 'Área/Rol', 'Fecha de registro']);
  const data  = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
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
// Estructura de la hoja Registros:
// Col A: Nombre
// Col B: Área/Rol
// Col C: Fecha (dd/MM/yyyy)
// Col D: Hora Entrada
// Col E: Hora Salida
// Col F: Horas trabajadas
// ─────────────────────────────────────────────────────────────────
function handleAddRecord(params) {
  const sheet = getOrCreateSheet(SHEET_REGISTROS,
    ['Nombre', 'Área/Rol', 'Fecha', 'Hora Entrada', 'Hora Salida', 'Horas trabajadas']);

  const name  = String(params.name || '').trim();
  const type  = String(params.type || '').trim(); // 'entrada' o 'salida'
  const role  = String(params.role || '').trim();

  if (!name || !type) return { success: false, error: 'Faltan datos del registro' };

  let ts;
  try { ts = params.timestamp ? new Date(params.timestamp) : new Date(); }
  catch(e) { ts = new Date(); }

  const fecha = getFecha(ts);
  const hora  = getHora(ts);

  const data = sheet.getDataRange().getValues();

  // Buscar si ya existe una fila para este nombre y fecha
  for (let i = 1; i < data.length; i++) {
    const rowName  = String(data[i][0]).trim().toLowerCase();
    const rowFecha = String(data[i][2]).trim();

    if (rowName === name.toLowerCase() && rowFecha === fecha) {
      // Fila encontrada — actualizar entrada o salida
      if (type === 'entrada') {
        // Solo escribe si la celda de entrada está vacía
        if (!data[i][3]) {
          sheet.getRange(i + 1, 4).setValue(hora);
        }
      } else if (type === 'salida') {
        sheet.getRange(i + 1, 5).setValue(hora);
        // Calcular horas trabajadas si hay entrada
        const entradaStr = data[i][3] || sheet.getRange(i + 1, 4).getValue();
        if (entradaStr) {
          const horasTrabajadas = calcularHoras(String(entradaStr), hora);
          sheet.getRange(i + 1, 6).setValue(horasTrabajadas);
        }
      }
      return { success: true };
    }
  }

  // No existe fila para hoy — crear nueva
  if (type === 'entrada') {
    sheet.appendRow([name, role, fecha, hora, '', '']);
  } else {
    // Salida sin entrada previa — registrar de todos modos
    sheet.appendRow([name, role, fecha, '', hora, '']);
  }

  // Formato zebra para legibilidad
  const lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, 6).setBackground('#f3f4f6');
  }

  return { success: true };
}

// ── CALCULAR HORAS TRABAJADAS ────────────────────────────────────
function calcularHoras(entradaStr, salidaStr) {
  try {
    const [eh, em, es] = entradaStr.split(':').map(Number);
    const [sh, sm, ss] = salidaStr.split(':').map(Number);
    const entradaMins = eh * 60 + em;
    const salidaMins  = sh * 60 + sm;
    let diffMins = salidaMins - entradaMins;
    if (diffMins < 0) diffMins += 24 * 60; // cruce de medianoche
    const horas   = Math.floor(diffMins / 60);
    const minutos = diffMins % 60;
    return horas + 'h ' + (minutos < 10 ? '0' : '') + minutos + 'min';
  } catch(e) {
    return '';
  }
}
