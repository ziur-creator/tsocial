// ============================================================
//  Turnos Trabajo Social - Hospital Piñero GCBA
//  1. Creá una Google Spreadsheet nueva y copiá su ID aquí:
var SPREADSHEET_ID = 'TU_SPREADSHEET_ID_AQUI';
//     (la ID está en la URL: docs.google.com/spreadsheets/d/ID/edit)
//  2. Pegá este código en script.google.com → Nuevo proyecto
//  3. Ejecutá setupSheets() una vez para crear las hojas
//  4. Implementar → Nueva implementación → App web
//     Ejecutar como: Yo  |  Acceso: Cualquier persona
// ============================================================

function getSpreadsheet() {
  if (SPREADSHEET_ID === 'TU_SPREADSHEET_ID_AQUI' || !SPREADSHEET_ID) {
    throw new Error('Configurá SPREADSHEET_ID en la línea 4 del script antes de continuar.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  e = e || {};  // guard: cuando se ejecuta desde el editor e es undefined
  var p = e.parameter || {};
  // Apps Script parte valores con comas en arrays — reconstruir fechas si es necesario
  try {
    if (e.parameters && e.parameters.fechas) {
      // e.parameters.fechas es array cuando hay múltiples valores separados por coma en la URL
      var fechasArr = e.parameters.fechas;
      p.fechas = Array.isArray(fechasArr) ? fechasArr.join(',') : String(fechasArr);
    }
  } catch(x) {}
  try {
    if (e.postData && e.postData.contents && e.postData.contents.length > 0) {
      try {
        // Intentar parsear como JSON (enviado con Content-Type text/plain)
        var parsed = JSON.parse(e.postData.contents);
        for (var key in parsed) p[key] = parsed[key];
      } catch(je) {
        // Fallback form-encoded
        var pairs = e.postData.contents.split('&');
        for (var i=0; i<pairs.length; i++) {
          var kv = pairs[i].split('=');
          if (kv.length >= 2) {
            var val = kv.slice(1).join('=');
            try { p[decodeURIComponent(kv[0])] = decodeURIComponent(val); } catch(de){}
          }
        }
      }
    }
  } catch(x) {}  var result;
  try {
    switch (p.action) {
      case 'register':               result = register(p);               break;
      case 'login':                  result = login(p);                  break;
      case 'getReservations':        result = getReservations(p);        break;
      case 'addReservation':         result = addReservation(p);         break;
      case 'addReservationBatch':    result = addReservationBatch(p);    break;
      case 'cancelReservation':      result = cancelReservation(p);      break;
      case 'getUsers':               result = getUsers(p);               break;
      case 'updateUser':             result = updateUser(p);             break;
      case 'deleteUser':             result = deleteUser(p);             break;
      case 'adminCancelReservation': result = adminCancelReservation(p); break;
      case 'getConsultorios':        result = getConsultorios(p);        break;
      case 'updateConsultorio':      result = updateConsultorio(p);      break;
      case 'updateProfile':          result = updateProfile(p);          break;
      case 'recoverPassword':        result = recoverPassword(p);        break;
      case 'verifyEmail':            result = verifyEmail(p);            break;
      case 'resendVerification':     result = resendVerification(p);     break;
      case 'setupDemoUsers':         result = setupDemoUsers();          break;
      case 'getAvisos':              result = getAvisos(p);              break;
      case 'addAviso':               result = addAviso(p);               break;
      case 'deleteAviso':            result = deleteAviso(p);            break;
      default: result = {ok:false, error:'Accion desconocida: '+p.action};
    }
  } catch(err) { result = {ok:false, error:err.message}; }
  var out = ContentService.createTextOutput(JSON.stringify(result));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// ---------- SETUP ----------
// Ejecutá setupSheets() UNA SOLA VEZ para crear todas las hojas y usuarios iniciales.
// Los usuarios de prueba tienen contraseña: 123456
// El administrador tiene contraseña: admin123
function setupSheets() {
  var ss = getSpreadsheet();

  // USUARIOS
  var su = ss.getSheetByName('Usuarios') || ss.insertSheet('Usuarios');
  if (su.getLastRow() === 0) {
    su.appendRow(['id','nombre','email','password','rol','telefono','foto','createdAt','verified','verifyCode']);
    su.appendRow([genId(),'Administrador','admin@pineiro.gob.ar',hashPw('admin123'),'admin','','',new Date().toISOString(),'true','']);
  }
  // Siempre verificar que los usuarios de prueba existan (se puede correr setupSheets más de una vez)
  var demos = [
    ['Dr. García',      'garcia@demo.com',      '+5491111111111'],
    ['Dra. Martínez',   'martinez@demo.com',    '+5491122222222'],
    ['Dr. Rodríguez',   'rodriguez@demo.com',   '+5491133333333'],
    ['Dra. López',      'lopez@demo.com',        '+5491144444444'],
    ['Dr. Fernández',   'fernandez@demo.com',   '+5491155555555'],
    ['Dra. Sánchez',    'sanchez@demo.com',      '+5491166666666']
  ];
  var existingUsers = sheetToObjects(su);
  var existingEmails = existingUsers.map(function(u){ return String(u.email).toLowerCase(); });
  var createdCount = 0;
  for (var i=0; i<demos.length; i++) {
    if (existingEmails.indexOf(demos[i][1]) === -1) {
      su.appendRow([genId(), demos[i][0], demos[i][1], hashPw('123456'), 'user', demos[i][2], '', new Date().toISOString(), 'true', '']);
      createdCount++;
    }
  }
  Logger.log('Usuarios de prueba creados: ' + createdCount);

  // RESERVAS
  var sr = ss.getSheetByName('Reservas') || ss.insertSheet('Reservas');
  if (sr.getLastRow() === 0)
    sr.appendRow(['id','consulId','fecha','startH','endH','userId','userName','createdAt']);

  // CONSULTORIOS
  var sc = ss.getSheetByName('Consultorios') || ss.insertSheet('Consultorios');
  if (sc.getLastRow() === 0) {
    sc.appendRow(['id','nombre','color','foto','descripcion']);
    var defs = [
      [1,'Consultorio A','#1a6abf','','Planta baja - Servicio de Trabajo Social'],
      [2,'Consultorio B','#26874A','','Planta baja - Servicio de Trabajo Social']
    ];
    for (var i=0;i<defs.length;i++) sc.appendRow(defs[i]);
  }
  // AVISOS
  var sa = ss.getSheetByName('Avisos') || ss.insertSheet('Avisos');
  if (sa.getLastRow() === 0) {
    sa.appendRow(['id','tipo','titulo','texto','fecha','createdAt']);
    var now = new Date().toISOString();
    var defaultAvisos = [
      ['ok','Sistema de turnos en línea','Bienvenidos al sistema de reservas del Servicio de Trabajo Social del Hospital Piñero. Desde esta app podés reservar cualquiera de los dos consultorios disponibles.','1 mar 2026'],
      ['general','Uso del espacio','Recordamos que al finalizar cada turno es necesario dejar el consultorio ordenado y en condiciones para el siguiente usuario.','5 mar 2026'],
      ['info','Horario de atención','El servicio atiende de lunes a viernes de 8:00 a 20:00 hs. Los consultorios se pueden reservar con hasta 30 días de anticipación.','10 mar 2026'],
      ['general','Horarios disponibles','Los consultorios están disponibles para reserva de lunes a viernes de 8:00 a 20:00 hs.','10 mar 2026']
    ];
    for (var i=0; i<defaultAvisos.length; i++) {
      sa.appendRow([genId(), defaultAvisos[i][0], defaultAvisos[i][1], defaultAvisos[i][2], defaultAvisos[i][3], now]);
    }
  }

  return {ok:true, msg:'Setup completo. Hojas y usuarios creados.'};
}
function getAvisos(p) {
  var sheet = getSheet('Avisos');
  if (!sheet) return {ok:true, avisos:[]};
  var rows = sheetToObjects(sheet);
  var avisos = rows.map(function(r){
    return {id:String(r.id),tipo:r.tipo||'general',titulo:r.titulo||'',texto:r.texto||'',fecha:r.fecha||'',userName:r.userName||'',userId:r.userId||''};
  });
  return {ok:true, avisos:avisos};
}

function addAviso(p) {
  // En T. Social Piñero todos los usuarios pueden publicar avisos
  if (!p.userId && !p.adminId) return {ok:false, error:'Sin identificación'};
  if (!p.titulo||!p.texto) return {ok:false, error:'Faltan datos'};
  var sheet = getSheet('Avisos') || getSpreadsheet().insertSheet('Avisos');
  if (sheet.getLastRow() === 0) sheet.appendRow(['id','tipo','titulo','texto','fecha','createdAt','userId','userName']);
  var id = genId();
  // Insertar en fila 2 (debajo del header) para que quede al tope
  if (sheet.getLastRow() > 1) sheet.insertRowAfter(1);
  sheet.getRange(2,1,1,8).setValues([[id, p.tipo||'general', p.titulo, p.texto, p.fecha||'', new Date().toISOString(), p.userId||'', p.userName||'']]);
  return {ok:true, id:id};
}

function deleteAviso(p) {
  // Admin puede eliminar cualquier aviso; el autor puede eliminar el suyo
  var sheet = getSheet('Avisos');
  if (!sheet) return {ok:false, error:'Sin avisos'};
  var rowIdx = findRowById(sheet, p.id);
  if (rowIdx < 0) return {ok:false, error:'No encontrado'};

  var isAdm = isAdmin(p.adminId);

  if (!isAdm) {
    if (!p.userId) return {ok:false, error:'Sin permiso'};
    // Columna G (índice 6, columna 7) = userId
    var data = sheet.getDataRange().getValues();
    var avisoUserId = String(data[rowIdx - 1][6] || '');
    // Si el aviso no tiene userId (aviso viejo sin autor) → permitir
    // Si tiene userId y no coincide → rechazar
    if (avisoUserId && avisoUserId !== String(p.userId)) {
      return {ok:false, error:'Sin permiso'};
    }
  }

  sheet.deleteRow(rowIdx);
  return {ok:true};
}

// Función alternativa si el sheet ya existe y solo querés agregar los usuarios de prueba.
// Ejecutala desde el editor de Apps Script si ya corriste setupSheets() antes.
function setupDemoUsers() {
  var sheet = getSheet('Usuarios');
  if (!sheet) return {ok:false, error:'Primero ejecutá setupSheets()'};
  var existing = sheetToObjects(sheet);
  var existingEmails = existing.map(function(u){ return String(u.email).toLowerCase(); });
  var demos = [
    ['Dr. García',      'garcia@demo.com',      '+5491111111111'],
    ['Dra. Martínez',   'martinez@demo.com',    '+5491122222222'],
    ['Dr. Rodríguez',   'rodriguez@demo.com',   '+5491133333333'],
    ['Dra. López',      'lopez@demo.com',        '+5491144444444'],
    ['Dr. Fernández',   'fernandez@demo.com',   '+5491155555555'],
    ['Dra. Sánchez',    'sanchez@demo.com',      '+5491166666666']
  ];
  var created = 0;
  for (var i=0; i<demos.length; i++) {
    if (existingEmails.indexOf(demos[i][1]) !== -1) { Logger.log('Ya existe: '+demos[i][1]); continue; }
    sheet.appendRow([genId(), demos[i][0], demos[i][1], hashPw('123456'), 'user', demos[i][2], '', new Date().toISOString(), 'true', '']);
    created++;
  }
  Logger.log('Usuarios creados: ' + created);
  return {ok:true, msg:'Usuarios creados: ' + created};
}
function genId() { return 'id_'+Date.now()+'_'+Math.floor(Math.random()*9999); }
function hashPw(pw) {
  var h=0;
  for (var i=0;i<pw.length;i++) { h=((h<<5)-h)+pw.charCodeAt(i); h|=0; }
  return 'h_'+Math.abs(h).toString(36)+'_'+pw.length;
}
function getSheet(n) { return getSpreadsheet().getSheetByName(n); }

// Google Sheets auto-convierte celdas con formato fecha a objetos Date.
// Esta función normaliza cualquier valor al tipo correcto según el nombre de columna.
function normalizeCell(key, val) {
  if (val instanceof Date) {
    // Columnas de fecha de reserva: convertir a YYYY-MM-DD
    if (key === 'fecha') {
      var y = val.getFullYear();
      var m = val.getMonth() + 1;
      var d = val.getDate();
      return y + '-' + (m<10?'0'+m:m) + '-' + (d<10?'0'+d:d);
    }
    // Otras fechas (createdAt) dejarlas como ISO string
    return val.toISOString();
  }
  return val;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0], rows = [];
  for (var i=1;i<data.length;i++) {
    var obj={};
    for (var j=0;j<headers.length;j++) obj[headers[j]] = normalizeCell(headers[j], data[i][j]);
    rows.push(obj);
  }
  return rows;
}
function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i=1;i<data.length;i++) if (String(data[i][0])===String(id)) return i+1;
  return -1;
}
function isAdmin(adminId) {
  if (!adminId) return false;
  var users = sheetToObjects(getSheet('Usuarios'));
  for (var i=0;i<users.length;i++) if (String(users[i].id)===String(adminId)&&users[i].rol==='admin') return true;
  return false;
}

// ---------- AUTH ----------
function register(p) {
  if (!p.nombre||!p.email||!p.password) return {ok:false,error:'Faltan campos obligatorios'};
  var sheet = getSheet('Usuarios');
  var users = sheetToObjects(sheet);
  for (var i=0;i<users.length;i++)
    if (users[i].email.toLowerCase()===p.email.toLowerCase()) return {ok:false,error:'El email ya esta registrado'};
  var id = genId();
  var code = String(Math.floor(100000 + Math.random()*900000));
  sheet.appendRow([id,p.nombre,p.email.toLowerCase(),hashPw(p.password),'user',p.telefono||'',p.foto||'',new Date().toISOString(),'false',code]);
  sendVerificationEmail(p.email.toLowerCase(), p.nombre, code);
  return {ok:true, pendingVerification:true, email:p.email.toLowerCase()};
}

function login(p) {
  if (!p.email||!p.password) return {ok:false,error:'Faltan campos'};
  var users = sheetToObjects(getSheet('Usuarios'));
  for (var i=0;i<users.length;i++) {
    var u=users[i];
    if (u.email.toLowerCase()===p.email.toLowerCase()&&u.password===hashPw(p.password)) {
      // Chequear verificacion (admins y usuarios pre-existentes sin campo = verificados)
      if (u.verified === 'false') {
        return {ok:false, error:'Verifica tu email antes de ingresar', pendingVerification:true, email:u.email};
      }
      return {ok:true, user:{id:u.id,nombre:u.nombre,email:u.email,rol:u.rol,telefono:u.telefono||'',foto:u.foto||''}};
    }
  }
  return {ok:false,error:'Email o contrasena incorrectos'};
}

function updateProfile(p) {
  if (!p.userId) return {ok:false,error:'Sin identificacion'};
  // Solo el propio usuario o un admin puede editar el perfil de otro
  var callerIsAdmin = p.adminId && isAdmin(p.adminId);
  var callerIsSelf  = p.callerId && String(p.callerId) === String(p.userId);
  if (!callerIsAdmin && !callerIsSelf && p.adminId !== p.userId) {
    // compatibilidad: si no se pasa adminId ni callerId, asumir que es el propio usuario
  }
  var sheet = getSheet('Usuarios');
  var rowIdx = findRowById(sheet, p.userId);
  if (rowIdx<0) return {ok:false,error:'Usuario no encontrado'};
  var data = sheet.getDataRange().getValues();
  var headers = data[0], row = data[rowIdx-1], obj={};
  for (var j=0;j<headers.length;j++) obj[headers[j]]=row[j];
  if (p.nombre)    obj.nombre    = p.nombre;
  if (p.telefono !== undefined) obj.telefono = p.telefono;
  // Permitir foto vacía (borrado): aceptar '' explícitamente
  if (p.foto !== undefined && p.foto !== 'undefined') obj.foto = p.foto;
  if (p.password)  obj.password  = hashPw(p.password);
  var newRow = headers.map(function(h){return obj[h];});
  sheet.getRange(rowIdx,1,1,newRow.length).setValues([newRow]);
  return {ok:true, user:{id:obj.id,nombre:obj.nombre,email:obj.email,rol:obj.rol,telefono:obj.telefono,foto:obj.foto}};
}

// ---------- RESERVATIONS ----------
function getReservations(p) {
  var reservas      = sheetToObjects(getSheet('Reservas'));
  var consultorios  = sheetToObjects(getSheet('Consultorios'));
  return {ok:true, reservations:reservas, consultorios:consultorios};
}

function addReservation(p) {
  if (!p.userId||!p.consulId||!p.fecha||p.startH===undefined||p.endH===undefined)
    return {ok:false,error:'Faltan datos de reserva'};
  var reservas = sheetToObjects(getSheet('Reservas'));
  var sH=parseInt(p.startH), eH=parseInt(p.endH), cId=String(p.consulId);
  var fecha=String(p.fecha).trim();
  for (var i=0;i<reservas.length;i++) {
    var r=reservas[i];
    var rs=parseInt(r.startH), re=parseInt(r.endH);
    var rFecha=String(r.fecha).trim();
    var solapan = !(eH<=rs||sH>=re);
    // Conflicto de consultorio
    if (String(r.consulId)===cId && rFecha===fecha && solapan)
      return {ok:false, error:'El consultorio ya esta reservado en ese horario'};
    // Mismo usuario, mismo día, mismo horario en cualquier consultorio
    if (String(r.userId)===String(p.userId) && rFecha===fecha && solapan)
      return {ok:false, error:'Ya tenes una reserva en ese horario. No podes reservar dos consultorios al mismo tiempo'};
  }
  var id=genId();
  getSheet('Reservas').appendRow([id,cId,fecha,sH,eH,String(p.userId),p.userName||'',new Date().toISOString()]);
  if (!p.skipEmail) sendConfirmationEmail(p, id);
  return {ok:true, id:id};
}

function cancelReservation(p) {
  if (!p.id||!p.userId) return {ok:false,error:'Faltan datos'};
  var sheet=getSheet('Reservas'), reservas=sheetToObjects(sheet);
  for (var i=0;i<reservas.length;i++) {
    if (String(reservas[i].id)===String(p.id)) {
      if (String(reservas[i].userId)!==String(p.userId)) return {ok:false,error:'Sin permiso'};
      var reservaCancelada = reservas[i];
      sheet.deleteRow(i+2);
      sendCancellationEmail(reservaCancelada);
      return {ok:true};
    }
  }
  return {ok:false,error:'Reserva no encontrada'};
}

function adminCancelReservation(p) {
  if (!isAdmin(p.adminId)) return {ok:false,error:'Sin permiso'};
  var sheet=getSheet('Reservas'), rowIdx=findRowById(sheet,p.id);
  if (rowIdx<0) return {ok:false,error:'Reserva no encontrada'};
  var data=sheet.getDataRange().getValues();
  var headers=data[0], row=data[rowIdx-1], rObj={};
  for (var j=0;j<headers.length;j++) rObj[headers[j]]=normalizeCell(headers[j],row[j]);
  sheet.deleteRow(rowIdx);
  sendCancellationEmail(rObj);
  return {ok:true};
}

// ---------- USERS (ADMIN) ----------
function getUsers(p) {
  var users = sheetToObjects(getSheet('Usuarios'));
  var safe  = users.map(function(u){
    return {
      id:       String(u.id||''),
      nombre:   String(u.nombre||''),
      email:    String(u.email||''),
      rol:      String(u.rol||'user'),
      telefono: String(u.telefono||''),
      foto:     String(u.foto||'')
    };
  });
  return {ok:true, users:safe};
}

function updateUser(p) {
  if (!isAdmin(p.adminId)) return {ok:false,error:'Sin permiso'};
  var sheet=getSheet('Usuarios'), rowIdx=findRowById(sheet,p.id);
  if (rowIdx<0) return {ok:false,error:'Usuario no encontrado'};
  var data=sheet.getDataRange().getValues(), headers=data[0], row=data[rowIdx-1], obj={};
  for (var j=0;j<headers.length;j++) obj[headers[j]]=row[j];
  if (p.nombre)    obj.nombre    = p.nombre;
  if (p.email)     obj.email     = p.email;
  if (p.rol)       obj.rol       = p.rol;
  if (p.telefono !== undefined) obj.telefono = p.telefono;
  if (p.password)  obj.password  = hashPw(p.password);
  sheet.getRange(rowIdx,1,1,headers.length).setValues([headers.map(function(h){return obj[h];})]);
  return {ok:true};
}

function deleteUser(p) {
  if (!isAdmin(p.adminId)) return {ok:false,error:'Sin permiso'};
  var sheet=getSheet('Usuarios'), rowIdx=findRowById(sheet,p.id);
  if (rowIdx<0) return {ok:false,error:'No encontrado'};
  sheet.deleteRow(rowIdx);
  // Eliminar todas las reservas del usuario
  var resSheet = getSheet('Reservas');
  var reservas = sheetToObjects(resSheet);
  // Recorrer de abajo hacia arriba para no alterar índices al borrar
  for (var i = reservas.length - 1; i >= 0; i--) {
    if (String(reservas[i].userId) === String(p.id)) {
      resSheet.deleteRow(i + 2); // +2 porque fila 1 es header y array es 0-based
    }
  }
  return {ok:true};
}

// ---------- CONSULTORIOS ----------
function getConsultorios(p) {
  return {ok:true, consultorios:sheetToObjects(getSheet('Consultorios'))};
}

function updateConsultorio(p) {
  if (!isAdmin(p.adminId)) return {ok:false,error:'Sin permiso'};
  var sheet=getSheet('Consultorios'), rowIdx=findRowById(sheet,p.id);
  if (rowIdx<0) return {ok:false,error:'Consultorio no encontrado'};
  var data=sheet.getDataRange().getValues(), headers=data[0], row=data[rowIdx-1], obj={};
  for (var j=0;j<headers.length;j++) obj[headers[j]]=row[j];
  if (p.nombre)      obj.nombre      = p.nombre;
  if (p.descripcion !== undefined) obj.descripcion = p.descripcion;
  // Permitir foto vacía (borrado)
  if (p.foto !== undefined && p.foto !== 'undefined') obj.foto = p.foto;
  sheet.getRange(rowIdx,1,1,headers.length).setValues([headers.map(function(h){return obj[h]||'';})]);
  return {ok:true, foto:String(obj.foto||'')};
}

// ---------- EMAIL CONFIRMACIÓN ----------
function sendConfirmationEmail(p, reservationId) {
  try {
    // Obtener datos del consultorio
    var consultorios = sheetToObjects(getSheet('Consultorios'));
    var consultorio = 'Consultorio';
    for (var i=0; i<consultorios.length; i++) {
      if (String(consultorios[i].id) === String(p.consulId)) {
        consultorio = consultorios[i].nombre;
        break;
      }
    }

    // Obtener email del usuario
    var users = sheetToObjects(getSheet('Usuarios'));
    var userEmail = p.email || '';
    var userName  = p.userName || '';
    for (var i=0; i<users.length; i++) {
      if (String(users[i].id) === String(p.userId)) {
        userEmail = users[i].email;
        userName  = users[i].nombre;
        break;
      }
    }
    if (!userEmail) return;

    // Formatear fecha y horario
    var parts = String(p.fecha).split('-');
    var fechaDisplay = parts.length === 3
      ? parts[2]+'/'+parts[1]+'/'+parts[0]
      : p.fecha;
    var startH = parseInt(p.startH);
    var endH   = parseInt(p.endH);
    var horario = startH+':00 - '+endH+':00 hs';

    var asunto = '[T. Social Piñero] Reserva confirmada: ' + consultorio;

    var cuerpoHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body><div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;background:#fdf6f0;border-radius:12px;overflow:hidden">'
      + '<div style="background:#1a6abf;padding:24px 28px;text-align:center">'
      + '<p style="color:rgba(255,255,255,.7);font-size:12px;margin:0;letter-spacing:2px">TRABAJO SOCIAL · PIÑERO</p>'
      + '<p style="color:rgba(255,255,255,.6);font-size:11px;margin:4px 0 0">Ampliando Posibilidades</p>'
      + '</div>'
      + '<div style="padding:28px">'
      + '<p style="color:#1a1410;font-size:16px;font-weight:700;margin:0 0 20px">Hola '+userName+', tu reserva est&aacute; confirmada &#10003;</p>'
      + '<table style="width:100%;border-collapse:collapse">'
      + '<tr><td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#6b5a50;font-size:13px">Consultorio</td>'
      + '<td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#1a1410;font-size:13px;font-weight:600;text-align:right">'+consultorio+'</td></tr>'
      + '<tr><td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#6b5a50;font-size:13px">Fecha</td>'
      + '<td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#1a1410;font-size:13px;font-weight:600;text-align:right">'+fechaDisplay+'</td></tr>'
      + '<tr><td style="padding:10px 0;color:#6b5a50;font-size:13px">Horario</td>'
      + '<td style="padding:10px 0;color:#1a1410;font-size:13px;font-weight:600;text-align:right">'+horario+'</td></tr>'
      + '</table>'
      + '<div style="background:#e8f0fb;border-left:3px solid #1a6abf;border-radius:4px;padding:12px 16px;margin-top:20px">'
      + '<p style="color:#6b5a50;font-size:12px;margin:0">Para cancelar tu reserva ingres&aacute; a la app en <strong>Mis Turnos</strong>.</p>'
      + '</div>'
      + '</div>'
      + '</div></body></html>';

    GmailApp.sendEmail(userEmail, asunto, 
      'Reserva confirmada: '+consultorio+' el '+fechaDisplay+' de '+horario,
      { htmlBody: cuerpoHtml, name: 'T. Social Piñero', charset: 'UTF-8' }
    );
  } catch(err) {
    // No interrumpir la reserva si falla el email
    Logger.log('Error enviando email: '+err.message);
  }
}

function sendCancellationEmail(reserva) {
  try {
    var users = sheetToObjects(getSheet('Usuarios'));
    var userEmail = '', userName = '';
    for (var i=0; i<users.length; i++) {
      if (String(users[i].id) === String(reserva.userId)) {
        userEmail = users[i].email;
        userName  = users[i].nombre;
        break;
      }
    }
    if (!userEmail) return;

    var consultorios = sheetToObjects(getSheet('Consultorios'));
    var consultorio = 'Consultorio';
    for (var i=0; i<consultorios.length; i++) {
      if (String(consultorios[i].id) === String(reserva.consulId)) {
        consultorio = consultorios[i].nombre;
        break;
      }
    }

    var parts = String(reserva.fecha).split('-');
    var fechaDisplay = parts.length === 3 ? parts[2]+'/'+parts[1]+'/'+parts[0] : reserva.fecha;
    var horario = parseInt(reserva.startH)+':00 - '+parseInt(reserva.endH)+':00 hs';

    var asunto = '[T. Social Piñero] Reserva cancelada: ' + consultorio;
    var cuerpoHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body><div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;background:#fdf6f0;border-radius:12px;overflow:hidden">'
      + '<div style="background:#1a6abf;padding:24px 28px;text-align:center">'
      + '<p style="color:rgba(255,255,255,.7);font-size:12px;margin:0;letter-spacing:2px">TRABAJO SOCIAL · PIÑERO</p>'
      + '</div>'
      + '<div style="padding:28px">'
      + '<p style="color:#1a1410;font-size:16px;font-weight:700;margin:0 0 20px">Hola '+userName+', tu reserva fue cancelada.</p>'
      + '<table style="width:100%;border-collapse:collapse">'
      + '<tr><td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#6b5a50;font-size:13px">Consultorio</td>'
      + '<td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#1a1410;font-size:13px;font-weight:600;text-align:right">'+consultorio+'</td></tr>'
      + '<tr><td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#6b5a50;font-size:13px">Fecha</td>'
      + '<td style="padding:10px 0;border-bottom:1px solid #f0e0d0;color:#1a1410;font-size:13px;font-weight:600;text-align:right">'+fechaDisplay+'</td></tr>'
      + '<tr><td style="padding:10px 0;color:#6b5a50;font-size:13px">Horario</td>'
      + '<td style="padding:10px 0;color:#1a1410;font-size:13px;font-weight:600;text-align:right">'+horario+'</td></tr>'
      + '</table>'
      + '</div></div></body></html>';

    GmailApp.sendEmail(userEmail, asunto,
      'Tu reserva en '+consultorio+' el '+fechaDisplay+' fue cancelada.',
      { htmlBody: cuerpoHtml, name: 'T. Social Piñero', charset: 'UTF-8' }
    );
  } catch(err) {
    Logger.log('Error email cancelacion: '+err.message);
  }
}

// ---------- RECUPERAR CONTRASEÑA ----------
function recoverPassword(p) {
  if (!p.email) return {ok:false, error:'Falta el email'};
  var sheet = getSheet('Usuarios');
  var users = sheetToObjects(sheet);
  var found = null;
  for (var i=0; i<users.length; i++) {
    if (users[i].email.toLowerCase() === p.email.toLowerCase()) {
      found = users[i];
      break;
    }
  }
  if (!found) return {ok:false, error:'No encontramos una cuenta con ese email'};

  // Generar contrase&ntilde;a temporal legible
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  var tempPw = '';
  for (var j=0; j<8; j++) tempPw += chars[Math.floor(Math.random()*chars.length)];

  // Guardar nueva contrase&ntilde;a hasheada
  var rowIdx = findRowById(sheet, found.id);
  var data = sheet.getDataRange().getValues();
  var headers = data[0], row = data[rowIdx-1], obj={};
  for (var k=0; k<headers.length; k++) obj[headers[k]] = row[k];
  obj.password = hashPw(tempPw);
  sheet.getRange(rowIdx, 1, 1, headers.length).setValues([headers.map(function(h){return obj[h];})]);

  // Enviar email
  try {
    var asunto = '[T. Social Piñero] Tu nueva contrasena temporal';
    var cuerpoHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body><div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;background:#fdf6f0;border-radius:12px;overflow:hidden">'
      + '<div style="background:#1a6abf;padding:24px 28px;text-align:center">'
      + '<p style="color:rgba(255,255,255,.7);font-size:12px;margin:0;letter-spacing:2px">TRABAJO SOCIAL · PIÑERO</p>'
      + '</div>'
      + '<div style="padding:28px">'
      + '<p style="color:#1a1410;font-size:16px;font-weight:700;margin:0 0 16px">Hola '+found.nombre+', aqu&iacute; est&aacute; tu contrase&ntilde;a temporal.</p>'
      + '<div style="background:#e8f0fb;border:2px solid #c25a00;border-radius:10px;padding:18px;text-align:center;margin:20px 0">'
      + '<p style="color:#6b5a50;font-size:12px;margin:0 0 8px;text-transform:uppercase;letter-spacing:1px">Tu nueva contrase&ntilde;a</p>'
      + '<p style="color:#c25a00;font-size:28px;font-weight:800;margin:0;letter-spacing:4px;font-family:monospace">'+tempPw+'</p>'
      + '</div>'
      + '<p style="color:#6b5a50;font-size:13px;margin:0">Una vez que ingreses, te recomendamos cambiarla desde <strong>Perfil → Mis datos</strong>.</p>'
      + '</div>'
      + '</div>';

    GmailApp.sendEmail(found.email, asunto,
      'Tu nueva contrasena temporal es: ' + tempPw,
      { htmlBody: cuerpoHtml, name: 'T. Social Piñero', charset: 'UTF-8' }
    );
  } catch(err) {
    Logger.log('Error email recuperacion: ' + err.message);
    return {ok:false, error:'Error al enviar el email. Contacta al administrador.'};
  }

  return {ok:true};
}

// ---------- RESERVAS EN LOTE (repetir) ----------
function addReservationBatch(p) {
  if (!p.userId || !p.consulId || !p.fechas || !p.startH || !p.endH)
    return {ok:false, error:'Faltan datos'};

  var fechas = String(p.fechas).replace(/%7C/gi,'|').replace(/%2C/gi,',').replace(/,/g,'|').split('|');
  var creadas = [], errores = [];

  for (var f=0; f<fechas.length; f++) {
    var fecha = fechas[f].trim();
    if (!fecha) continue;
    // Reusar addReservation con skipEmail para no mandar email por cada una
    var res = addReservation({
      userId:  p.userId,
      userName: p.userName || '',
      consulId: p.consulId,
      fecha:    fecha,
      startH:   p.startH,
      endH:     p.endH,
      skipEmail: true
    });
    if (res.ok) creadas.push(fecha);
    else        errores.push(fecha);
  }

  // Un solo email con todas las fechas creadas
  if (creadas.length > 0) sendBatchConfirmationEmail(p, creadas);

  return {ok:true, creadas:creadas, errores:errores};
}

// ---------- EMAIL CONFIRMACIÓN EN LOTE ----------
function sendBatchConfirmationEmail(p, fechas) {
  try {
    var consultorios = sheetToObjects(getSheet('Consultorios'));
    var consultorio = 'Consultorio';
    for (var i=0; i<consultorios.length; i++) {
      if (String(consultorios[i].id) === String(p.consulId)) {
        consultorio = consultorios[i].nombre; break;
      }
    }
    var users = sheetToObjects(getSheet('Usuarios'));
    var userEmail = '', userName = p.userName || '';
    for (var i=0; i<users.length; i++) {
      if (String(users[i].id) === String(p.userId)) {
        userEmail = users[i].email;
        userName  = users[i].nombre;
        break;
      }
    }
    if (!userEmail) return;

    var sH = parseInt(p.startH), eH = parseInt(p.endH);
    var horario = sH + ':00 - ' + eH + ':00 hs';

    var asunto = '[T. Social Piñero] ' + fechas.length + ' reservas confirmadas: ' + consultorio;

    var filas = '';
    for (var i=0; i<fechas.length; i++) {
      var parts = String(fechas[i]).split('-');
      var fd = parts.length===3 ? parts[2]+'/'+parts[1]+'/'+parts[0] : fechas[i];
      var bg = i%2===0 ? '#ffffff' : '#fdf6f0';
      filas += '<tr style="background:'+bg+'">'
        + '<td style="padding:9px 12px;font-size:13px;color:#1a1410;border-bottom:1px solid #f0e0d0">'+fd+'</td>'
        + '<td style="padding:9px 12px;font-size:13px;color:#1a1410;border-bottom:1px solid #f0e0d0;text-align:right">'+horario+'</td>'
        + '</tr>';
    }

    var cuerpoHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>'
      + '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#fdf6f0;border-radius:12px;overflow:hidden">'
      + '<div style="background:#1a6abf;padding:24px 28px;text-align:center">'
      + '<p style="color:rgba(255,255,255,.7);font-size:12px;margin:0;letter-spacing:2px">TRABAJO SOCIAL · PIÑERO</p>'
      + '</div>'
      + '<div style="padding:28px">'
      + '<p style="color:#1a1410;font-size:16px;font-weight:700;margin:0 0 6px">Hola '+userName+'</p>'
      + '<p style="color:#6b5a50;font-size:13px;margin:0 0 20px">Se crearon <strong>'+fechas.length+' reservas</strong> para <strong>'+consultorio+'</strong>.</p>'
      + '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;border:1px solid #f0e0d0">'
      + '<thead><tr style="background:#f5ece3">'
      + '<th style="padding:9px 12px;font-size:11px;color:#6b5a50;text-align:left;text-transform:uppercase;letter-spacing:1px">Fecha</th>'
      + '<th style="padding:9px 12px;font-size:11px;color:#6b5a50;text-align:right;text-transform:uppercase;letter-spacing:1px">Horario</th>'
      + '</tr></thead>'
      + '<tbody>'+filas+'</tbody>'
      + '</table>'
      + '<div style="background:#e8f0fb;border-left:3px solid #1a6abf;border-radius:4px;padding:12px 16px;margin-top:20px">'
      + '<p style="color:#6b5a50;font-size:12px;margin:0">Para cancelar cualquier reserva ingres&aacute; a la app en <strong>Mis Turnos</strong>.</p>'
      + '</div>'
      + '</div></div></body></html>';

    GmailApp.sendEmail(userEmail, asunto,
      'Se crearon ' + fechas.length + ' reservas en ' + consultorio + ' de ' + horario,
      { htmlBody: cuerpoHtml, name: 'T. Social Piñero', charset: 'UTF-8' });
  } catch(err) {
    Logger.log('Error email batch: ' + err.message);
  }
}

// ---------- VERIFICACION DE EMAIL ----------
function verifyEmail(p) {
  if (!p.email||!p.code) return {ok:false,error:'Faltan datos'};
  var sheet = getSheet('Usuarios');
  var users = sheetToObjects(sheet);
  for (var i=0;i<users.length;i++) {
    var u = users[i];
    if (u.email.toLowerCase() === p.email.toLowerCase()) {
      if (String(u.verifyCode) !== String(p.code).trim())
        return {ok:false, error:'Codigo incorrecto. Revisá tu email.'};
      // Marcar como verificado
      var rowIdx = findRowById(sheet, u.id);
      var data = sheet.getDataRange().getValues();
      var headers = data[0], row = data[rowIdx-1], obj={};
      for (var j=0;j<headers.length;j++) obj[headers[j]]=row[j];
      obj.verified = 'true';
      obj.verifyCode = '';
      sheet.getRange(rowIdx,1,1,headers.length).setValues([headers.map(function(h){return obj[h]||'';})]);
      return {ok:true, user:{id:u.id,nombre:u.nombre,email:u.email,rol:u.rol,telefono:u.telefono||'',foto:u.foto||''}};
    }
  }
  return {ok:false, error:'Email no encontrado'};
}

function resendVerification(p) {
  if (!p.email) return {ok:false,error:'Falta el email'};
  var sheet = getSheet('Usuarios');
  var users = sheetToObjects(sheet);
  for (var i=0;i<users.length;i++) {
    var u = users[i];
    if (u.email.toLowerCase() === p.email.toLowerCase()) {
      if (u.verified === 'true') return {ok:false, error:'Esta cuenta ya esta verificada'};
      var code = String(Math.floor(100000 + Math.random()*900000));
      var rowIdx = findRowById(sheet, u.id);
      var data = sheet.getDataRange().getValues();
      var headers = data[0], row = data[rowIdx-1], obj={};
      for (var j=0;j<headers.length;j++) obj[headers[j]]=row[j];
      obj.verifyCode = code;
      sheet.getRange(rowIdx,1,1,headers.length).setValues([headers.map(function(h){return obj[h]||'';})]);
      sendVerificationEmail(u.email, u.nombre, code);
      return {ok:true};
    }
  }
  return {ok:false, error:'Email no encontrado'};
}

function sendVerificationEmail(email, nombre, code) {
  try {
    var asunto = '[T. Social Piñero] Tu codigo de verificacion: ' + code;
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>'
      + '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;background:#fdf6f0;border-radius:12px;overflow:hidden">'
      + '<div style="background:#1a6abf;padding:24px 28px;text-align:center">'
      + '<p style="color:rgba(255,255,255,.7);font-size:12px;margin:0;letter-spacing:2px">TRABAJO SOCIAL · PIÑERO</p>'
      + '</div>'
      + '<div style="padding:28px">'
      + '<p style="color:#1a1410;font-size:16px;font-weight:700;margin:0 0 8px">Hola ' + nombre + '</p>'
      + '<p style="color:#6b5a50;font-size:13px;margin:0 0 24px">Ingresa este codigo en la app para verificar tu cuenta:</p>'
      + '<div style="background:#e8f0fb;border:2px solid #c25a00;border-radius:10px;padding:20px;text-align:center;margin:0 0 20px">'
      + '<p style="color:#6b5a50;font-size:11px;margin:0 0 8px;text-transform:uppercase;letter-spacing:2px">Codigo de verificacion</p>'
      + '<p style="color:#c25a00;font-size:40px;font-weight:800;margin:0;letter-spacing:10px;font-family:monospace">' + code + '</p>'
      + '</div>'
      + '<p style="color:#a08878;font-size:12px;margin:0">Si no creaste esta cuenta, ignora este mensaje.</p>'
      + '</div></div></body></html>';
    GmailApp.sendEmail(email, asunto,
      'Tu codigo de verificacion es: ' + code,
      {htmlBody: html, name: 'T. Social Piñero', charset: 'UTF-8'});
  } catch(e) {
    Logger.log('Error email verificacion: ' + e.message);
  }
}
