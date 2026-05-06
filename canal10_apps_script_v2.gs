// ============================================================
//  CANAL 10 — Apps Script para Pedidos de Edición v2
//  Pegá este código en: Extensions > Apps Script del Sheets
//  Luego desplegalo como Web App:
//    - Ejecutar como: Yo
//    - Acceso: Cualquier persona
// ============================================================

var ENCABEZADOS = [
  "N° Ticket", "Fecha y hora", "Área", "Tipo de pieza",
  "Prioridad", "Duración", "Deadline", "Material / Ubicación",
  "Descripción", "Solicitante", "Estado"
];

// ── Router principal ──────────────────────────────────────
function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);

    // Si trae "accion", es una operación del Kanban
    if (datos.accion === 'updateEstado') {
      return updateEstado(datos.ticket, datos.estado);
    }

    // Si no, es un pedido nuevo
    return guardarPedido(datos);

  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  // El Kanban pide todos los pedidos via GET
  try {
    return leerPedidos();
  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }
}

// ── Guardar pedido nuevo ──────────────────────────────────
function guardarPedido(datos) {
  var sheet = getSheet();

  // Encabezados si la hoja está vacía
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ENCABEZADOS);
    formatearEncabezados(sheet);
  }

  // Número de pedido
  var num = sheet.getLastRow(); // fila 1 = encabezado, num 1 = fila 2
  var anio = new Date().getFullYear();
  var ticket = datos.ticket || ("C10-" + anio + "-" + String(num).padStart(3, "0"));

  var ahora = Utilities.formatDate(
    new Date(), "America/Argentina/Buenos_Aires", "dd/MM/yyyy HH:mm"
  );

  var fila = [
    ticket,
    ahora,
    datos.area     || "",
    datos.tipo     || "",
    datos.prio     || "",
    datos.dur      || "",
    datos.deadline || "",
    datos.ubic     || "",
    datos.desc     || "",
    datos.nombre   || "",
    "PENDIENTE"
  ];

  sheet.appendRow(fila);

  // Color de fila según prioridad
  var fNum = sheet.getLastRow();
  var rango = sheet.getRange(fNum, 1, 1, ENCABEZADOS.length);
  var color = colorPrio(datos.prio);
  rango.setBackground(color);
  rango.setFontSize(10);
  sheet.getRange(fNum, 11).setFontWeight("bold").setBackground("#F9E79F");

  return json({ ok: true, ticket: ticket });
}

// ── Actualizar estado de un pedido ────────────────────────
function updateEstado(ticket, nuevoEstado) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === ticket) {
      var celdaEstado = sheet.getRange(i + 1, 11);
      celdaEstado.setValue(nuevoEstado);

      // Colores según estado
      if (nuevoEstado === "EN PROCESO") {
        celdaEstado.setBackground("#AED6F1").setFontColor("#1A5276");
      } else if (nuevoEstado === "LISTO") {
        celdaEstado.setBackground("#A9DFBF").setFontColor("#1E8449");
      } else {
        celdaEstado.setBackground("#F9E79F").setFontColor("#7D6608");
      }
      celdaEstado.setFontWeight("bold");
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: "Ticket no encontrado" });
}

// ── Leer todos los pedidos (para el Kanban) ───────────────
function leerPedidos() {
  var sheet = getSheet();
  if (sheet.getLastRow() <= 1) {
    return json({ ok: true, pedidos: [] });
  }

  var data = sheet.getDataRange().getValues();
  var pedidos = [];

  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    if (!fila[0]) continue; // saltar filas vacías
    pedidos.push({
      ticket:  fila[0]  + "",
      fecha:   fila[1]  + "",
      area:    fila[2]  + "",
      tipo:    fila[3]  + "",
      prio:    fila[4]  + "",
      dur:     fila[5]  + "",
      deadline:fila[6]  + "",
      ubic:    fila[7]  + "",
      desc:    fila[8]  + "",
      nombre:  fila[9]  + "",
      estado:  fila[10] + "",
      ts:      new Date().getTime() - (data.length - i) * 1000
    });
  }

  return json({ ok: true, pedidos: pedidos });
}

// ── Helpers ───────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function colorPrio(p) {
  var map = { P1: "#FADBD8", P2: "#FDEBD0", P3: "#EAFAF1", P4: "#EBF5FB" };
  return map[p] || "#FFFFFF";
}

function formatearEncabezados(sheet) {
  var h = sheet.getRange(1, 1, 1, ENCABEZADOS.length);
  h.setBackground("#1a1a2e");
  h.setFontColor("#FFFFFF");
  h.setFontWeight("bold");
  h.setFontSize(11);
  sheet.setFrozenRows(1);
  var anchos = [110, 130, 100, 130, 80, 90, 130, 160, 300, 130, 100];
  anchos.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
}
