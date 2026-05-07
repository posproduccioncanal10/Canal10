// ============================================================
//  CANAL 10 — Apps Script v3.0
//  Sistema de Pedidos con autenticación, archivos y estados
// ============================================================

var SHEET_PEDIDOS  = "Pedidos";
var SHEET_USUARIOS = "Usuarios";
var SHEET_ARCHIVOS = "Archivos";

var ENCABEZADOS_PEDIDOS = [
  "N° Ticket", "Fecha y hora", "Área", "Tipo de pieza", "Prioridad",
  "Duración", "Deadline", "Material / Ubicación", "Descripción",
  "Solicitante", "Email", "Estado", "Editor asignado", "Inicio edición",
  "Fin edición", "Archivos adjuntos"
];

var ENCABEZADOS_USUARIOS = [
  "Email", "Nombre", "Área", "Rol", "Fecha registro", "Activo"
];

var ENCABEZADOS_ARCHIVOS = [
  "Ticket", "Nombre archivo", "Tipo", "URL Drive", "Fecha subida"
];

var COL = {
  TICKET:1, FECHA:2, AREA:3, TIPO:4, PRIO:5,
  DUR:6, DEADLINE:7, UBIC:8, DESC:9, NOMBRE:10,
  EMAIL:11, ESTADO:12, EDITOR:13, INICIO:14, FIN:15, ARCHIVOS:16
};

var ESTADOS = {
  "PENDIENTE": { pill:"#F9E79F", color:"#7D6608" },
  "EN CURSO":  { pill:"#AED6F1", color:"#1A5276" },
  "REALIZADO": { pill:"#A9DFBF", color:"#1E8449" }
};

// ─── Helpers ─────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === SHEET_PEDIDOS)  { sh.appendRow(ENCABEZADOS_PEDIDOS);  formatH(sh, ENCABEZADOS_PEDIDOS.length);  setWidths(sh); }
    if (name === SHEET_USUARIOS) { sh.appendRow(ENCABEZADOS_USUARIOS); formatH(sh, ENCABEZADOS_USUARIOS.length); }
    if (name === SHEET_ARCHIVOS) { sh.appendRow(ENCABEZADOS_ARCHIVOS); formatH(sh, ENCABEZADOS_ARCHIVOS.length); }
  }
  return sh;
}

function formatH(sh, len) {
  var r = sh.getRange(1,1,1,len);
  r.setBackground("#111827"); r.setFontColor("#FFF");
  r.setFontWeight("bold"); r.setFontSize(10);
  sh.setFrozenRows(1);
}

function setWidths(sh) {
  [120,140,110,140,140,100,140,180,320,140,200,120,160,140,140,200].forEach(function(w,i){ sh.setColumnWidth(i+1,w); });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ahoraAR() {
  return Utilities.formatDate(new Date(),"America/Argentina/Buenos_Aires","dd/MM/yyyy HH:mm");
}

function aplicarEstilo(sh, fila, estado) {
  var e = ESTADOS[estado] || ESTADOS["PENDIENTE"];
  var c = sh.getRange(fila, COL.ESTADO);
  c.setBackground(e.pill); c.setFontColor(e.color); c.setFontWeight("bold");
}

// ─── Validar usuario ─────────────────────────────────────
function validarUsuario(email) {
  if (!email) return { ok:false, error:"Email requerido" };
  var sh = getSheet(SHEET_USUARIOS);
  var data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      if (data[i][5]===false || data[i][5]==="FALSE" || data[i][5]==="")
        return { ok:false, error:"Tu usuario está desactivado. Contactá a Posproducción." };
      return { ok:true, nombre:data[i][1], area:data[i][2], rol:data[i][3] };
    }
  }
  return { ok:false, error:"not_registered" };
}

// ════════════════════════════════════════════════════════
//  doPost
// ════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var accion = datos.accion || "nuevo_pedido";
    if (accion==="registrar_usuario") return registrarUsuario(datos);
    if (accion==="nuevo_pedido")      return nuevoPedido(datos);
    if (accion==="adjuntar_archivo")  return adjuntarArchivo(datos);
    return jsonOut({ ok:false, error:"Acción desconocida" });
  } catch(err) {
    return jsonOut({ ok:false, error:err.toString() });
  }
}

function registrarUsuario(datos) {
  var email  = (datos.email  ||"").trim().toLowerCase();
  var nombre = (datos.nombre ||"").trim();
  var area   = (datos.area   ||"").trim();
  if (!email||!nombre) return jsonOut({ ok:false, error:"Faltan datos" });

  var sh = getSheet(SHEET_USUARIOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0].toString().toLowerCase()===email) {
      if (rows[i][5]===false||rows[i][5]==="FALSE"||rows[i][5]==="")
        return jsonOut({ ok:false, error:"Usuario desactivado" });
      return jsonOut({ ok:true, accion:"login", nombre:rows[i][1], area:rows[i][2], rol:rows[i][3] });
    }
  }
  sh.appendRow([email, nombre, area, "solicitante", ahoraAR(), true]);
  return jsonOut({ ok:true, accion:"registro", nombre:nombre, area:area, rol:"solicitante" });
}

function nuevoPedido(datos) {
  var v = validarUsuario(datos.email);
  if (!v.ok) return jsonOut({ ok:false, error:v.error });

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
  var col = "#FFF";
  var p = datos.prio||"";
  if (p.includes("P1")) col="#FADBD8";
  else if (p.includes("P2")) col="#FDEBD0";
  else if (p.includes("P3")) col="#EAFAF1";
  else if (p.includes("P4")) col="#EBF5FB";
  sh.getRange(nf,1,1,ENCABEZADOS_PEDIDOS.length).setBackground(col);
  aplicarEstilo(sh, nf, "PENDIENTE");

  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(["PENDIENTE","EN CURSO","REALIZADO"],true)
    .setAllowInvalid(false).build();
  sh.getRange(nf, COL.ESTADO).setDataValidation(regla);

  crearTriggerSiNoExiste();
  return jsonOut({ ok:true, ticket:ticket });
}

function adjuntarArchivo(datos) {
  var v = validarUsuario(datos.email);
  if (!v.ok) return jsonOut({ ok:false, error:v.error });

  getSheet(SHEET_ARCHIVOS).appendRow([
    datos.ticket||"", datos.nombre||"", datos.tipo||"", datos.url||"", ahoraAR()
  ]);

  var shP = getSheet(SHEET_PEDIDOS);
  var rows = shP.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][COL.TICKET-1]===datos.ticket) {
      var actual = rows[i][COL.ARCHIVOS-1]||"";
      shP.getRange(i+1,COL.ARCHIVOS).setValue(actual ? actual+" | "+datos.nombre : datos.nombre);
      break;
    }
  }
  return jsonOut({ ok:true });
}

// ════════════════════════════════════════════════════════
//  doGet
// ════════════════════════════════════════════════════════
function doGet(e) {
  if (!e.parameter||!e.parameter.action)
    return ContentService.createTextOutput("Canal 10 API v3.0 ✓").setMimeType(ContentService.MimeType.TEXT);

  var a = e.parameter.action;
  if (a==="estado")     return getEstado(e.parameter.ticket);
  if (a==="dashboard")  return getDashboard(e.parameter.email);
  if (a==="check_user") return checkUser(e.parameter.email);
  if (a==="archivos")   return getArchivos(e.parameter.ticket);
  return jsonOut({ ok:false, error:"Acción desconocida" });
}

function checkUser(email) {
  return jsonOut(validarUsuario(email));
}

function getEstado(ticketId) {
  if (!ticketId) return jsonOut({ ok:false, error:"Falta ticket" });
  var sh = getSheet(SHEET_PEDIDOS);
  var rows = sh.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][COL.TICKET-1]===ticketId) {
      var r=rows[i];
      return jsonOut({ ok:true, ticket:ticketId,
        estado:r[COL.ESTADO-1]||"PENDIENTE", editor:r[COL.EDITOR-1]||"",
        inicio:r[COL.INICIO-1]?r[COL.INICIO-1].toString():"",
        fin:r[COL.FIN-1]?r[COL.FIN-1].toString():"",
        area:r[COL.AREA-1]||"", tipo:r[COL.TIPO-1]||"",
        prio:r[COL.PRIO-1]||"", nombre:r[COL.NOMBRE-1]||"",
        desc:r[COL.DESC-1]||"", deadline:r[COL.DEADLINE-1]||"",
        archivos:r[COL.ARCHIVOS-1]||""
      });
    }
  }
  return jsonOut({ ok:false, error:"Ticket no encontrado" });
}

function getDashboard(email) {
  var v = validarUsuario(email);
  if (!v.ok) return jsonOut({ ok:false, error:v.error });

  var sh = getSheet(SHEET_PEDIDOS);
  var rows = sh.getDataRange().getValues();
  var pedidos = [];
  for (var i=1; i<rows.length; i++) {
    var r=rows[i];
    if (!r[COL.TICKET-1]) continue;
    pedidos.push({
      ticket:r[COL.TICKET-1]||"", fecha:r[COL.FECHA-1]?r[COL.FECHA-1].toString():"",
      area:r[COL.AREA-1]||"", tipo:r[COL.TIPO-1]||"", prio:r[COL.PRIO-1]||"",
      dur:r[COL.DUR-1]||"", deadline:r[COL.DEADLINE-1]||"", ubic:r[COL.UBIC-1]||"",
      desc:r[COL.DESC-1]||"", nombre:r[COL.NOMBRE-1]||"", email:r[COL.EMAIL-1]||"",
      estado:r[COL.ESTADO-1]||"PENDIENTE", editor:r[COL.EDITOR-1]||"",
      inicio:r[COL.INICIO-1]?r[COL.INICIO-1].toString():"",
      fin:r[COL.FIN-1]?r[COL.FIN-1].toString():"",
      archivos:r[COL.ARCHIVOS-1]||""
    });
  }

  var po={"P1":1,"P2":2,"P3":3,"P4":4};
  pedidos.sort(function(a,b){
    var pa=5,pb=5;
    Object.keys(po).forEach(function(k){ if(a.prio.includes(k))pa=po[k]; if(b.prio.includes(k))pb=po[k]; });
    return pa-pb;
  });

  return jsonOut({ ok:true, pedidos:pedidos, usuario:v });
}

function getArchivos(ticketId) {
  var sh = getSheet(SHEET_ARCHIVOS);
  var rows = sh.getDataRange().getValues();
  var arch=[];
  for(var i=1;i<rows.length;i++){
    if(rows[i][0]===ticketId)
      arch.push({nombre:rows[i][1],tipo:rows[i][2],url:rows[i][3],fecha:rows[i][4]?rows[i][4].toString():""});
  }
  return jsonOut({ ok:true, archivos:arch });
}

// ─── onEdit trigger ───────────────────────────────────────
function onEditTrigger(e) {
  var sh=e.source.getActiveSheet();
  if(sh.getName()!==SHEET_PEDIDOS) return;
  var col=e.range.getColumn(), fila=e.range.getRow();
  if(col!==COL.ESTADO||fila<2) return;
  var nuevo=sh.getRange(fila,COL.ESTADO).getValue().toString().trim().toUpperCase();
  if(!ESTADOS[nuevo]) return;
  var ahora=ahoraAR();
  if(nuevo==="EN CURSO"){sh.getRange(fila,COL.INICIO).setValue(ahora);sh.getRange(fila,COL.FIN).setValue("");}
  else if(nuevo==="REALIZADO"){if(!sh.getRange(fila,COL.INICIO).getValue())sh.getRange(fila,COL.INICIO).setValue(ahora);sh.getRange(fila,COL.FIN).setValue(ahora);}
  else if(nuevo==="PENDIENTE"){sh.getRange(fila,COL.INICIO).setValue("");sh.getRange(fila,COL.FIN).setValue("");}
  aplicarEstilo(sh,fila,nuevo);
}

function crearTriggerSiNoExiste() {
  var ts=ScriptApp.getProjectTriggers();
  for(var t=0;t<ts.length;t++){if(ts[t].getHandlerFunction()==="onEditTrigger")return;}
  ScriptApp.newTrigger("onEditTrigger").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

function instalarTrigger() {
  crearTriggerSiNoExiste();
  SpreadsheetApp.getUi().alert("✅ Trigger instalado.");
}
