// ============================================================
//  CANAL 10 — Apps Script v4.0
//  Sistema completo: Pedidos + Editores + Superadmin + Drive
//
//  CONFIGURACIÓN:
//  1. Pegá este código en Extensions > Apps Script del Sheets
//  2. Cambiá DRIVE_FOLDER_ID por el ID de tu carpeta compartida
//  3. Desplegá como Web App:
//       Ejecutar como: Yo
//       Acceso: Cualquier persona
// ============================================================

var DRIVE_FOLDER_ID = "1JslwxmipJQ9t88PZeIkozaoXit6wVp3p";  // ← carpeta de adjuntos

var SHEET_PEDIDOS  = "Pedidos";
var SHEET_USUARIOS = "Usuarios";
var SHEET_ARCHIVOS = "Archivos";

var ENCABEZADOS_PEDIDOS = [
  "N° Ticket","Fecha y hora","Área","Tipo de pieza","Prioridad",
  "Duración","Deadline","Material / Ubicación","Descripción",
  "Solicitante","Email","Estado","Editor asignado","Inicio edición",
  "Fin edición","Archivos adjuntos"
];
var ENCABEZADOS_USUARIOS = ["Email","Nombre","Área","Rol","Fecha registro","Activo"];
var ENCABEZADOS_ARCHIVOS = ["Ticket","Nombre archivo","Tipo","URL Drive","Tamaño","Fecha subida"];

var COL = {
  TICKET:1,FECHA:2,AREA:3,TIPO:4,PRIO:5,DUR:6,DEADLINE:7,
  UBIC:8,DESC:9,NOMBRE:10,EMAIL:11,ESTADO:12,EDITOR:13,
  INICIO:14,FIN:15,ARCHIVOS:16
};

var ROLES = { SOLICITANTE:"solicitante", EDITOR:"editor", ADMIN:"admin" };

// ─── Helpers ─────────────────────────────────────────────
function ss()  { return SpreadsheetApp.getActiveSpreadsheet(); }
function jsonOut(obj) {
  var out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
function ahoraAR() {
  return Utilities.formatDate(new Date(),"America/Argentina/Buenos_Aires","dd/MM/yyyy HH:mm");
}

function getSheet(name) {
  var sh = ss().getSheetByName(name);
  if (!sh) {
    sh = ss().insertSheet(name);
    if (name===SHEET_PEDIDOS)  { sh.appendRow(ENCABEZADOS_PEDIDOS);  fmtH(sh,ENCABEZADOS_PEDIDOS.length);  setWidths(sh); }
    if (name===SHEET_USUARIOS) { sh.appendRow(ENCABEZADOS_USUARIOS); fmtH(sh,ENCABEZADOS_USUARIOS.length); }
    if (name===SHEET_ARCHIVOS) { sh.appendRow(ENCABEZADOS_ARCHIVOS); fmtH(sh,ENCABEZADOS_ARCHIVOS.length); }
  }
  return sh;
}

function fmtH(sh, len) {
  var r = sh.getRange(1,1,1,len);
  r.setBackground("#111827"); r.setFontColor("#FFF");
  r.setFontWeight("bold"); r.setFontSize(10);
  sh.setFrozenRows(1);
}

function setWidths(sh) {
  [120,140,110,140,140,100,140,180,320,140,200,120,160,140,140,220]
    .forEach(function(w,i){ sh.setColumnWidth(i+1,w); });
}

function colorEstado(sh, fila, estado) {
  var map = { PENDIENTE:["#FEF9C3","#854D0E"], "EN CURSO":["#DBEAFE","#1E3A8A"], REALIZADO:["#DCFCE7","#14532D"] };
  var c = map[estado] || map.PENDIENTE;
  var cel = sh.getRange(fila,COL.ESTADO);
  cel.setBackground(c[0]); cel.setFontColor(c[1]); cel.setFontWeight("bold");
}

// ─── Validar usuario ─────────────────────────────────────
function validarUsuario(email) {
  if (!email) return {ok:false,error:"Email requerido"};
  var sh = getSheet(SHEET_USUARIOS);
  var data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim()===email.toLowerCase().trim()) {
      if (data[i][5]===false||data[i][5]==="FALSE"||data[i][5]==="")
        return {ok:false,error:"Usuario desactivado"};
      return {ok:true, nombre:data[i][1], area:data[i][2], rol:data[i][3]};
    }
  }
  return {ok:false,error:"not_registered"};
}

function esAdmin(email) {
  var v = validarUsuario(email);
  return v.ok && v.rol===ROLES.ADMIN;
}

function esEditor(email) {
  var v = validarUsuario(email);
  return v.ok && (v.rol===ROLES.EDITOR || v.rol===ROLES.ADMIN);
}

// ════════════════════════════════════════════════════════
//  doPost — router
// ════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var a = datos.accion || "nuevo_pedido";
    if (a==="registrar_usuario")  return registrarUsuario(datos);
    if (a==="nuevo_pedido")       return nuevoPedido(datos);
    if (a==="subir_archivo")      return subirArchivoDrive(datos);
    if (a==="asignar_editor")     return asignarEditor(datos);
    if (a==="cambiar_estado")     return cambiarEstado(datos);
    if (a==="marcar_inicio")      return marcarInicio(datos);
    if (a==="marcar_fin")         return marcarFin(datos);
    if (a==="listar_editores")    return listarEditores(datos);
    if (a==="activar_usuario")    return activarDesactivar(datos, true);
    if (a==="desactivar_usuario") return activarDesactivar(datos, false);
    if (a==="cambiar_rol")        return cambiarRol(datos);
    return jsonOut({ok:false,error:"Acción desconocida"});
  } catch(err) {
    return jsonOut({ok:false,error:err.toString()});
  }
}

// ════════════════════════════════════════════════════════
//  doGet — router
// ════════════════════════════════════════════════════════
function doGet(e) {
  if (!e.parameter||!e.parameter.action)
    return ContentService.createTextOutput("Canal 10 API v4.0 ✓").setMimeType(ContentService.MimeType.TEXT);
  var a = e.parameter.action;
  if (a==="check_user")       return checkUser(e.parameter.email);
  if (a==="dashboard")        return getDashboard(e.parameter.email);
  if (a==="mis_tareas")       return getMisTareas(e.parameter.email);
  if (a==="archivos")         return getArchivos(e.parameter.ticket);
  if (a==="admin_dashboard")  return getAdminDashboard(e.parameter.email);
  if (a==="admin_usuarios")   return getAdminUsuarios(e.parameter.email);
  return jsonOut({ok:false,error:"Acción desconocida"});
}

// ════════════════════════════════════════════════════════
//  AUTH
// ════════════════════════════════════════════════════════
function checkUser(email) { return jsonOut(validarUsuario(email)); }

function registrarUsuario(datos) {
  var email  = (datos.email||"").trim().toLowerCase();
  var nombre = (datos.nombre||"").trim();
  var area   = (datos.area||"").trim();
  var rolPedido = (datos.rol||ROLES.SOLICITANTE).trim().toLowerCase();

  if (!email||!nombre) return jsonOut({ok:false,error:"Faltan datos"});

  // Solo admin y editor son roles válidos para auto-registro (editores piden rol editor, resto = solicitante)
  var rolFinal = ROLES.SOLICITANTE;
  if (rolPedido===ROLES.EDITOR) rolFinal = ROLES.EDITOR; // se registra como editor (pendiente de aprobación)

  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0].toString().toLowerCase()===email) {
      if (rows[i][5]===false||rows[i][5]==="FALSE"||rows[i][5]==="")
        return jsonOut({ok:false,error:"Usuario desactivado"});
      return jsonOut({ok:true,accion:"login",nombre:rows[i][1],area:rows[i][2],rol:rows[i][3]});
    }
  }
  sh.appendRow([email, nombre, area, rolFinal, ahoraAR(), true]);
  return jsonOut({ok:true,accion:"registro",nombre:nombre,area:area,rol:rolFinal});
}

// ════════════════════════════════════════════════════════
//  PEDIDOS
// ════════════════════════════════════════════════════════
function nuevoPedido(datos) {
  var v = validarUsuario(datos.email);
  if (!v.ok) return jsonOut({ok:false,error:v.error});

  var sh = getSheet(SHEET_PEDIDOS);
  var anio = new Date().getFullYear();
  var ticket = "C10-"+anio+"-"+String(sh.getLastRow()).padStart(3,"0");

  sh.appendRow([
    ticket, ahoraAR(),
    datos.area||"", datos.tipo||"", datos.prio||"",
    datos.dur||"", datos.deadline||"", datos.ubic||"",
    datos.desc||"", datos.nombre||"", datos.email||"",
    "PENDIENTE","","","",""
  ]);

  var nf = sh.getLastRow();
  var bgMap = {P1:"#FADBD8",P2:"#FDEBD0",P3:"#EAFAF1",P4:"#EBF5FB"};
  var prio = datos.prio||"";
  var bg = "#FFFFFF";
  Object.keys(bgMap).forEach(function(k){ if(prio.includes(k)) bg=bgMap[k]; });
  sh.getRange(nf,1,1,ENCABEZADOS_PEDIDOS.length).setBackground(bg);
  colorEstado(sh,nf,"PENDIENTE");

  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(["PENDIENTE","EN CURSO","REALIZADO"],true)
    .setAllowInvalid(false).build();
  sh.getRange(nf,COL.ESTADO).setDataValidation(regla);
  crearTrigger();

  return jsonOut({ok:true,ticket:ticket});
}

// ════════════════════════════════════════════════════════
//  ARCHIVOS — subida REAL a Drive via base64
// ════════════════════════════════════════════════════════
function subirArchivoDrive(datos) {
  var v = validarUsuario(datos.email);
  if (!v.ok) return jsonOut({ok:false,error:v.error});
  if (!datos.ticket||!datos.nombre||!datos.base64||!datos.mimeType)
    return jsonOut({ok:false,error:"Faltan datos del archivo"});

  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    // Crear subcarpeta por ticket si no existe
    var subFolderName = datos.ticket;
    var subFolders = folder.getFoldersByName(subFolderName);
    var subFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(subFolderName);

    // Decodificar base64 y crear archivo
    var decoded = Utilities.base64Decode(datos.base64);
    var blob = Utilities.newBlob(decoded, datos.mimeType, datos.nombre);
    var file = subFolder.createFile(blob);

    // Hacer el archivo accesible con link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = file.getUrl();
    var fileId = file.getId();
    var dlUrl = "https://drive.google.com/uc?export=download&id="+fileId;
    var size = datos.size || 0;

    // Guardar en hoja Archivos
    getSheet(SHEET_ARCHIVOS).appendRow([
      datos.ticket, datos.nombre, datos.mimeType, url, size, ahoraAR()
    ]);

    // Actualizar columna archivos en Pedidos
    var shP = getSheet(SHEET_PEDIDOS);
    var rows = shP.getDataRange().getValues();
    for (var i=1; i<rows.length; i++) {
      if (rows[i][COL.TICKET-1]===datos.ticket) {
        var actual = rows[i][COL.ARCHIVOS-1]||"";
        var nuevoVal = actual ? actual+" | "+datos.nombre+" ("+url+")" : datos.nombre+" ("+url+")";
        shP.getRange(i+1,COL.ARCHIVOS).setValue(nuevoVal);
        break;
      }
    }

    return jsonOut({ok:true, url:url, dlUrl:dlUrl, fileId:fileId, nombre:datos.nombre});

  } catch(err) {
    return jsonOut({ok:false,error:"Error al subir a Drive: "+err.toString()});
  }
}

function getArchivos(ticketId) {
  var sh = getSheet(SHEET_ARCHIVOS);
  var rows = sh.getDataRange().getValues();
  var arch = [];
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0]===ticketId) {
      var fileId = "";
      var url = rows[i][3]||"";
      var m = url.match(/\/d\/([^\/]+)/);
      if (m) fileId = m[1];
      arch.push({
        nombre: rows[i][1]||"",
        tipo:   rows[i][2]||"",
        url:    url,
        dlUrl:  fileId?"https://drive.google.com/uc?export=download&id="+fileId:"",
        size:   rows[i][4]||0,
        fecha:  rows[i][5]?rows[i][5].toString():""
      });
    }
  }
  return jsonOut({ok:true,archivos:arch});
}

// ════════════════════════════════════════════════════════
//  ADMIN — asignar editor
// ════════════════════════════════════════════════════════
function asignarEditor(datos) {
  if (!esAdmin(datos.emailAdmin))
    return jsonOut({ok:false,error:"Sin permisos"});
  if (!datos.ticket||!datos.editor)
    return jsonOut({ok:false,error:"Faltan datos"});

  var sh = getSheet(SHEET_PEDIDOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][COL.TICKET-1]===datos.ticket) {
      sh.getRange(i+1,COL.EDITOR).setValue(datos.editor);
      // Si estaba PENDIENTE, pasa a EN CURSO automáticamente
      if ((rows[i][COL.ESTADO-1]||"PENDIENTE")==="PENDIENTE") {
        sh.getRange(i+1,COL.ESTADO).setValue("EN CURSO");
        sh.getRange(i+1,COL.INICIO).setValue(ahoraAR());
        colorEstado(sh,i+1,"EN CURSO");
      }
      return jsonOut({ok:true});
    }
  }
  return jsonOut({ok:false,error:"Ticket no encontrado"});
}

// ════════════════════════════════════════════════════════
//  ESTADOS
// ════════════════════════════════════════════════════════
function cambiarEstado(datos) {
  // Admin puede cambiar cualquier estado; editor solo los suyos
  var v = validarUsuario(datos.email);
  if (!v.ok) return jsonOut({ok:false,error:v.error});
  var validos = ["PENDIENTE","EN CURSO","REALIZADO"];
  if (!validos.includes(datos.estado)) return jsonOut({ok:false,error:"Estado inválido"});

  var sh = getSheet(SHEET_PEDIDOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][COL.TICKET-1]===datos.ticket) {
      // Editor solo puede cambiar si está asignado a él
      if (v.rol===ROLES.EDITOR) {
        var editorAsig = (rows[i][COL.EDITOR-1]||"").toString().toLowerCase().trim();
        if (editorAsig && editorAsig!==datos.email.toLowerCase().trim())
          return jsonOut({ok:false,error:"No tenés permiso sobre este pedido"});
      }
      sh.getRange(i+1,COL.ESTADO).setValue(datos.estado);
      var ahora = ahoraAR();
      if (datos.estado==="EN CURSO")  { sh.getRange(i+1,COL.INICIO).setValue(ahora); sh.getRange(i+1,COL.FIN).setValue(""); }
      if (datos.estado==="REALIZADO") { if(!rows[i][COL.INICIO-1])sh.getRange(i+1,COL.INICIO).setValue(ahora); sh.getRange(i+1,COL.FIN).setValue(ahora); }
      if (datos.estado==="PENDIENTE") { sh.getRange(i+1,COL.INICIO).setValue(""); sh.getRange(i+1,COL.FIN).setValue(""); }
      colorEstado(sh,i+1,datos.estado);
      return jsonOut({ok:true});
    }
  }
  return jsonOut({ok:false,error:"Ticket no encontrado"});
}

function marcarInicio(datos) {
  return cambiarEstado({email:datos.email,ticket:datos.ticket,estado:"EN CURSO"});
}
function marcarFin(datos) {
  return cambiarEstado({email:datos.email,ticket:datos.ticket,estado:"REALIZADO"});
}

// ════════════════════════════════════════════════════════
//  QUERIES
// ════════════════════════════════════════════════════════
function getDashboard(email) {
  var v = validarUsuario(email);
  if (!v.ok) return jsonOut({ok:false,error:v.error});
  return jsonOut({ok:true, pedidos:todosLosPedidos(), usuario:v});
}

function getMisTareas(email) {
  if (!esEditor(email)) return jsonOut({ok:false,error:"Sin permisos"});
  var todos = todosLosPedidos();
  var emailLow = email.toLowerCase().trim();
  var mias = todos.filter(function(p){
    return (p.editor||"").toLowerCase().trim()===emailLow;
  });
  return jsonOut({ok:true,pedidos:mias});
}

function getAdminDashboard(email) {
  if (!esAdmin(email)) return jsonOut({ok:false,error:"Sin permisos admin"});
  var pedidos = todosLosPedidos();
  var editores = getEditoresList();
  return jsonOut({ok:true,pedidos:pedidos,editores:editores});
}

function getAdminUsuarios(email) {
  if (!esAdmin(email)) return jsonOut({ok:false,error:"Sin permisos admin"});
  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  var users = [];
  for (var i=1; i<rows.length; i++) {
    if (!rows[i][0]) continue;
    users.push({
      email:rows[i][0]||"", nombre:rows[i][1]||"", area:rows[i][2]||"",
      rol:rows[i][3]||"", fecha:rows[i][4]?rows[i][4].toString():"",
      activo:rows[i][5]!==false && rows[i][5]!=="FALSE" && rows[i][5]!==""
    });
  }
  return jsonOut({ok:true,usuarios:users});
}

function listarEditores(datos) {
  var v = validarUsuario(datos.email);
  if (!v.ok || (v.rol!==ROLES.ADMIN && v.rol!==ROLES.EDITOR))
    return jsonOut({ok:false,error:"Sin permisos"});
  return jsonOut({ok:true,editores:getEditoresList()});
}

function getEditoresList() {
  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  var eds = [];
  for (var i=1; i<rows.length; i++) {
    if (!rows[i][0]) continue;
    var rol = (rows[i][3]||"").toString().toLowerCase();
    var activo = rows[i][5]!==false && rows[i][5]!=="FALSE" && rows[i][5]!=="";
    if ((rol===ROLES.EDITOR||rol===ROLES.ADMIN) && activo) {
      eds.push({email:rows[i][0],nombre:rows[i][1],area:rows[i][2],rol:rows[i][3]});
    }
  }
  return eds;
}

function todosLosPedidos() {
  var sh = getSheet(SHEET_PEDIDOS);
  var rows = sh.getDataRange().getValues();
  var pedidos = [];
  for (var i=1; i<rows.length; i++) {
    var r = rows[i];
    if (!r[COL.TICKET-1]) continue;
    pedidos.push({
      ticket:  r[COL.TICKET-1]||"",
      fecha:   r[COL.FECHA-1]?r[COL.FECHA-1].toString():"",
      area:    r[COL.AREA-1]||"",
      tipo:    r[COL.TIPO-1]||"",
      prio:    r[COL.PRIO-1]||"",
      dur:     r[COL.DUR-1]||"",
      deadline:r[COL.DEADLINE-1]||"",
      ubic:    r[COL.UBIC-1]||"",
      desc:    r[COL.DESC-1]||"",
      nombre:  r[COL.NOMBRE-1]||"",
      email:   r[COL.EMAIL-1]||"",
      estado:  r[COL.ESTADO-1]||"PENDIENTE",
      editor:  r[COL.EDITOR-1]||"",
      inicio:  r[COL.INICIO-1]?r[COL.INICIO-1].toString():"",
      fin:     r[COL.FIN-1]?r[COL.FIN-1].toString():"",
      archivos:r[COL.ARCHIVOS-1]||""
    });
  }
  var po = {P1:1,P2:2,P3:3,P4:4};
  pedidos.sort(function(a,b){
    var pa=5,pb=5;
    Object.keys(po).forEach(function(k){
      if(a.prio.includes(k))pa=po[k];
      if(b.prio.includes(k))pb=po[k];
    });
    return pa-pb;
  });
  return pedidos;
}

// ════════════════════════════════════════════════════════
//  ADMIN — gestión de usuarios
// ════════════════════════════════════════════════════════
function activarDesactivar(datos, activar) {
  if (!esAdmin(datos.emailAdmin)) return jsonOut({ok:false,error:"Sin permisos"});
  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0].toString().toLowerCase()===datos.email.toLowerCase()) {
      sh.getRange(i+1,6).setValue(activar);
      return jsonOut({ok:true});
    }
  }
  return jsonOut({ok:false,error:"Usuario no encontrado"});
}

function cambiarRol(datos) {
  if (!esAdmin(datos.emailAdmin)) return jsonOut({ok:false,error:"Sin permisos"});
  var rolesValidos = [ROLES.SOLICITANTE,ROLES.EDITOR,ROLES.ADMIN];
  if (!rolesValidos.includes(datos.rol)) return jsonOut({ok:false,error:"Rol inválido"});

  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0].toString().toLowerCase()===datos.email.toLowerCase()) {
      sh.getRange(i+1,4).setValue(datos.rol);
      return jsonOut({ok:true});
    }
  }
  return jsonOut({ok:false,error:"Usuario no encontrado"});
}

// ─── Trigger onEdit ─────────────────────────────────────
function onEditTrigger(e) {
  var sh = e.source.getActiveSheet();
  if (sh.getName()!==SHEET_PEDIDOS) return;
  var col=e.range.getColumn(), fila=e.range.getRow();
  if (col!==COL.ESTADO||fila<2) return;
  var nuevo = sh.getRange(fila,COL.ESTADO).getValue().toString().trim().toUpperCase();
  var validos = ["PENDIENTE","EN CURSO","REALIZADO"];
  if (!validos.includes(nuevo)) return;
  var ahora = ahoraAR();
  if (nuevo==="EN CURSO")  { sh.getRange(fila,COL.INICIO).setValue(ahora); sh.getRange(fila,COL.FIN).setValue(""); }
  if (nuevo==="REALIZADO") { if(!sh.getRange(fila,COL.INICIO).getValue())sh.getRange(fila,COL.INICIO).setValue(ahora); sh.getRange(fila,COL.FIN).setValue(ahora); }
  if (nuevo==="PENDIENTE") { sh.getRange(fila,COL.INICIO).setValue(""); sh.getRange(fila,COL.FIN).setValue(""); }
  colorEstado(sh,fila,nuevo);
}

function crearTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  for (var t=0; t<ts.length; t++) { if(ts[t].getHandlerFunction()==="onEditTrigger") return; }
  ScriptApp.newTrigger("onEditTrigger").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

function instalarTrigger() {
  crearTrigger();
  SpreadsheetApp.getUi().alert("✅ Trigger instalado correctamente.");
}

// ─── Función para crear el primer admin ─────────────────
// Corré esta función UNA VEZ manualmente desde Apps Script para crear tu cuenta admin
function crearPrimerAdmin() {
  var email  = "posproduccion.canal10@gmail.com"; // ← tu email
  var nombre = "Superadmin Canal 10";
  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0].toString().toLowerCase()===email.toLowerCase()) {
      sh.getRange(i+1,4).setValue(ROLES.ADMIN);
      sh.getRange(i+1,6).setValue(true);
      Logger.log("✅ " + email + " ahora es ADMIN.");
      return;
    }
  }
  sh.appendRow([email, nombre, "Posproducción", ROLES.ADMIN, ahoraAR(), true]);
  SpreadsheetApp.getUi().alert("✅ Admin creado: "+email);
}
