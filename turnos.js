/**
 * ASIGNADOR DE TURNOS WAVE - CEUP  ·  "ONE WORLD, ONE WAVE"
 * Web app de Apps Script vinculada al Google Sheet.
 *
 * Hojas que TÚ llenas:
 *   - "notas" : padrón. Encabezados en fila 1: Código | Nombre | Nota
 *               (A = código UP, B = nombre de la persona, C = nota /20)
 *   - "buses" : respuestas del Google Form (código UP, ¿va a Wave?, hora llega, hora se retira)
 *
 * Hojas que crea/usa la herramienta (botón "Configurar hojas"):
 *   - "CONFIG"      : parámetros (tiempo de viaje, margen)
 *   - "CAT_TURNOS"  : Puesto | Inicio | Fin | Duración(min) | Personas | Encargado
 *                     Cada puesto se parte en sub-turnos del largo "Duración".
 *   - "CAT_BUSES"   : Bus | Tipo(ida/regreso) | Hora | Capacidad
 *   - "ASIGNACION"  : resultado (Código | Nombre | Nota | Llega | Sale | Turno | Bus ida | Bus regreso | Alertas | Manual)
 */

var SH = {
  NOTAS: 'notas', FORM: 'buses', CONFIG: 'CONFIG',
  TURNOS: 'CAT_TURNOS', BUSES: 'CAT_BUSES', ASIG: 'ASIGNACION'
};

// ---------- Menú ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Turnos WAVE')
    .addItem('Configurar hojas', 'setupSheets')
    .addItem('Abrir asignador', 'abrirAsignador')
    .addToUi();
}
function abrirAsignador() {
  var html = HtmlService.createHtmlOutput(INDEX_HTML)
    .setWidth(1280).setHeight(780).setTitle('Asignador de Turnos WAVE');
  SpreadsheetApp.getUi().showModalDialog(html, 'Asignador de Turnos WAVE');
}
function doGet() {
  return HtmlService.createHtmlOutput(INDEX_HTML).setTitle('Asignador de Turnos WAVE');
}

// ---------- Crear hojas ----------
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cfg = ss.getSheetByName(SH.CONFIG) || ss.insertSheet(SH.CONFIG);
  if (cfg.getLastRow() === 0) {
    cfg.getRange(1, 1, 3, 2).setValues([
      ['Parámetro', 'Valor'],
      ['viaje_minutos', 90],
      ['margen_minutos', 30]
    ]);
    cfg.getRange(1, 1, 1, 2).setFontWeight('bold');
    cfg.setColumnWidth(1, 180); cfg.setColumnWidth(2, 100);
  }

  // notas (solo crea encabezados si la hoja no existe; si ya la tienes, no la toca)
  if (!ss.getSheetByName(SH.NOTAS)) {
    var n = ss.insertSheet(SH.NOTAS);
    n.getRange(1, 1, 1, 3).setValues([['Código', 'Nombre', 'Nota']]);
    n.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  var t = ss.getSheetByName(SH.TURNOS) || ss.insertSheet(SH.TURNOS);
  if (t.getLastRow() === 0) {
    t.getRange(1, 1, 1, 6).setValues([['Puesto', 'Inicio', 'Fin', 'Duración (min)', 'Personas', 'Encargado']]);
    t.getRange(2, 1, 6, 6).setValues([
      ['Escaneo Puerta', '9:30 PM', '11:00 PM', 30, 6, ''],
      ['Mozo boxes',     '9:30 PM', '11:30 PM', 30, 2, 'Angela Navarrete'],
      ['Guardarropa',    '10:30 PM','12:00 AM', 30, 2, 'Alejandra Florez'],
      ['Foodtruck',      '1:00 AM', '3:30 AM',  30, 2, ''],
      ['Tópico',         '12:00 AM','5:00 AM',  60, 2, 'Rossana Alvarado'],
      ['Turno baño',     '12:00 AM','4:00 AM',  60, 2, 'Alejandra Florez']
    ]);
    t.getRange(1, 1, 1, 6).setFontWeight('bold');
    for (var ci = 1; ci <= 6; ci++) t.setColumnWidth(ci, 130);
  }

  var b = ss.getSheetByName(SH.BUSES) || ss.insertSheet(SH.BUSES);
  if (b.getLastRow() === 0) {
    b.getRange(1, 1, 1, 4).setValues([['Bus', 'Tipo', 'Hora', 'Capacidad']]);
    b.getRange(2, 1, 7, 4).setValues([
      ['Bus 1', 'ida', '8:00 PM', 49],
      ['Bus 2', 'ida', '8:30 PM', 49],
      ['Bus 3', 'ida', '9:00 PM', 55],
      ['Bus 4', 'ida', '9:30 PM', 55],
      ['Bus R1', 'regreso', '2:30 AM', 49],
      ['Bus R2', 'regreso', '3:30 AM', 49],
      ['Bus R3', 'regreso', '5:00 AM', 49]
    ]);
    b.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  var a = ss.getSheetByName(SH.ASIG) || ss.insertSheet(SH.ASIG);
  if (a.getLastRow() === 0) {
    a.getRange(1, 1, 1, 10).setValues([[
      'Código', 'Nombre', 'Nota', 'Hora llega', 'Hora sale', 'Turno', 'Bus ida', 'Bus regreso', 'Alertas', 'Manual'
    ]]);
    a.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  ss.toast('Hojas listas. Abre "Turnos WAVE > Abrir asignador".', 'Turnos WAVE', 5);
}

// ---------- Helpers de lectura ----------
function _sheetObjects(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = {}, hasData = false;
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
      if (values[i][j] !== '' && values[i][j] != null) hasData = true;
    }
    if (hasData) out.push(row);
  }
  return out;
}
function _pick(obj, keys) {
  for (var k in obj) {
    var kl = String(k).toLowerCase();
    for (var i = 0; i < keys.length; i++) if (kl.indexOf(keys[i]) !== -1) return obj[k];
  }
  return '';
}
function _config() {
  var c = {};
  _sheetObjects(SH.CONFIG).forEach(function (r) { c[String(r['Parámetro']).trim()] = r['Valor']; });
  return { viaje: Number(c['viaje_minutos']) || 90, margen: Number(c['margen_minutos']) || 30 };
}

// Padrón: código -> {nombre, nota}. Lee por encabezado; si no hay, usa A=cód,B=nombre,C=nota.
function _roster() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTAS);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var values = sh.getDataRange().getValues();
  var hdr = values[0].map(function (h) { return String(h).toLowerCase().trim(); });
  var iCod = -1, iNom = -1, iNota = -1;
  hdr.forEach(function (h, i) {
    if (iCod < 0 && (h.indexOf('cód') !== -1 || h.indexOf('cod') !== -1)) iCod = i;
    if (iNom < 0 && h.indexOf('nombre') !== -1) iNom = i;
    if (iNota < 0 && h.indexOf('nota') !== -1) iNota = i;
  });
  var hasHeader = (iCod !== -1);
  if (!hasHeader) { iCod = 0; iNom = 1; iNota = 2; }
  for (var r = hasHeader ? 1 : 0; r < values.length; r++) {
    var cod = String(values[r][iCod]).trim();
    if (!cod) continue;
    var nota = (iNota >= 0 && values[r][iNota] !== '' && values[r][iNota] != null) ? Number(values[r][iNota]) : null;
    var nom = (iNom >= 0) ? String(values[r][iNom]).trim() : '';
    map[cod] = { nombre: nom, nota: nota };
  }
  return map;
}

// Participantes = respuestas del Form (hoja "buses") que van a Wave, cruzadas con el padrón
function _participantes() {
  var roster = _roster();
  var out = [];
  _sheetObjects(SH.FORM).forEach(function (r) {
    var cod = String(_pick(r, ['código', 'codigo'])).trim();
    if (!cod) return;
    var va = String(_pick(r, ['vas a ir', 'va a wave', 'ir a wave'])).toLowerCase();
    var vaWave = va.indexOf('s') === 0 || va.indexOf('sí') !== -1 || va.indexOf('si') !== -1 || va === 'true';
    var info = roster[cod] || {};
    out.push({
      codigo: cod,
      nombre: info.nombre || '',
      nota: (info.nota == null) ? null : info.nota,
      vaWave: vaWave,
      llega: String(_pick(r, ['planeado ir', 'hora tienes planeado ir', 'llega', 'llegar'])).trim(),
      sale: String(_pick(r, ['retirarte', 'retirar', 'sale', 'salir'])).trim()
    });
  });
  return out;
}

// ---------- Horas ----------
function parseMin(v) {
  if (v === '' || v == null) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    var h0 = v.getHours(), m0 = v.getMinutes(); if (h0 < 8) h0 += 24; return h0 * 60 + m0;
  }
  var s = String(v).trim().toUpperCase();
  var m = s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/);
  if (!m) { var h = s.match(/(\d{1,2})\s*(AM|PM)/); if (!h) return null; m = [null, h[1], '00', h[2]]; }
  var hh = +m[1], mm = +m[2], ap = m[3];
  if (ap) { if (ap === 'PM' && hh !== 12) hh += 12; if (ap === 'AM' && hh === 12) hh = 0; }
  if (hh < 8) hh += 24;
  return hh * 60 + mm;
}
function fmtMin(m) {
  if (m == null) return '';
  var hh = Math.floor(m / 60) % 24, mm = m % 60;
  var ap = hh >= 12 ? 'PM' : 'AM'; var h12 = hh % 12; if (h12 === 0) h12 = 12;
  return h12 + ':' + (mm < 10 ? '0' : '') + mm + ' ' + ap;
}

// Genera los sub-turnos (slots) a partir de CAT_TURNOS según la Duración
function _slots() {
  var slots = [];
  _sheetObjects(SH.TURNOS).forEach(function (t) {
    var puesto = String(t['Puesto'] || '').trim(); if (!puesto) return;
    var ini = parseMin(t['Inicio']), fin = parseMin(t['Fin']);
    var dur = Number(t['Duración (min)']) || Number(_pick(t, ['duración', 'duracion'])) || 0;
    var personas = Number(t['Personas']) || 0;
    var enc = t['Encargado'] || '';
    if (ini == null || fin == null) {
      slots.push({ id: puesto, puesto: puesto, iniMin: ini, finMin: fin, inicio: t['Inicio'], fin: t['Fin'], necesarios: personas, encargado: enc });
      return;
    }
    if (!dur || dur >= (fin - ini)) {
      slots.push(_slot(puesto, ini, fin, personas, enc));
    } else {
      for (var s = ini; s < fin; s += dur) {
        var e = Math.min(s + dur, fin);
        slots.push(_slot(puesto, s, e, personas, enc));
      }
    }
  });
  return slots;
}
function _slot(puesto, ini, fin, personas, enc) {
  return {
    id: puesto + ' ' + fmtMin(ini) + '-' + fmtMin(fin),
    puesto: puesto, iniMin: ini, finMin: fin,
    inicio: fmtMin(ini), fin: fmtMin(fin),
    necesarios: personas, encargado: enc
  };
}

// ---------- API ----------
function getData() {
  var buses = _sheetObjects(SH.BUSES).map(function (b) {
    return { bus: b['Bus'], tipo: String(b['Tipo']).toLowerCase().trim(), hora: b['Hora'], capacidad: Number(b['Capacidad']) || 0, horaMin: parseMin(b['Hora']) };
  });
  return {
    config: _config(),
    participantes: _participantes(),
    turnos: _slots(),
    buses: buses,
    asignacion: _sheetObjects(SH.ASIG)
  };
}

// ---------- Auto-sugerir ----------
function autoAsignar() {
  var d = getData();
  var cfg = d.config;
  var personas = d.participantes.filter(function (p) { return p.vaWave; });
  personas.sort(function (a, b) {
    var na = (a.nota == null) ? -1 : a.nota, nb = (b.nota == null) ? -1 : b.nota; return nb - na;
  });

  var turnos = d.turnos.map(function (t) { return { t: t, asignados: 0 }; });
  var busIda = d.buses.filter(function (b) { return b.tipo === 'ida'; }).sort(function (a, b) { return a.horaMin - b.horaMin; });
  var busReg = d.buses.filter(function (b) { return b.tipo === 'regreso'; }).sort(function (a, b) { return a.horaMin - b.horaMin; });
  var capIda = {}, capReg = {};
  busIda.forEach(function (b) { capIda[b.bus] = 0; });
  busReg.forEach(function (b) { capReg[b.bus] = 0; });

  var res = [];
  personas.forEach(function (p) {
    var llega = parseMin(p.llega), sale = parseMin(p.sale);
    var elegido = null;
    turnos.forEach(function (slot) {
      if (slot.asignados >= slot.t.necesarios) return;
      if (llega != null && slot.t.iniMin != null && llega > slot.t.iniMin) return;
      if (sale != null && slot.t.finMin != null && sale < slot.t.finMin) return;
      if (!elegido || (slot.t.necesarios - slot.asignados) > (elegido.t.necesarios - elegido.asignados)) elegido = slot;
    });
    if (!elegido) {
      turnos.forEach(function (slot) {
        if (slot.asignados >= slot.t.necesarios) return;
        if (!elegido || (slot.t.necesarios - slot.asignados) > (elegido.t.necesarios - elegido.asignados)) elegido = slot;
      });
    }
    var turnoId = '', iniMin = null, finMin = null;
    if (elegido) { elegido.asignados++; turnoId = elegido.t.id; iniMin = elegido.t.iniMin; finMin = elegido.t.finMin; }

    var busI = '';
    if (iniMin != null) {
      var limite = iniMin - cfg.viaje - cfg.margen;
      for (var i = busIda.length - 1; i >= 0; i--) {
        if (busIda[i].horaMin <= limite && capIda[busIda[i].bus] < busIda[i].capacidad) { busI = busIda[i].bus; capIda[busI]++; break; }
      }
      if (!busI) for (var j = 0; j < busIda.length; j++) {
        if (capIda[busIda[j].bus] < busIda[j].capacidad) { busI = busIda[j].bus; capIda[busI]++; break; }
      }
    }
    var busR = '';
    var refSalida = (sale != null) ? sale : finMin;
    if (refSalida != null) {
      for (var k = 0; k < busReg.length; k++) {
        if (busReg[k].horaMin >= refSalida && capReg[busReg[k].bus] < busReg[k].capacidad) { busR = busReg[k].bus; capReg[busR]++; break; }
      }
      if (!busR) for (var l = busReg.length - 1; l >= 0; l--) {
        if (capReg[busReg[l].bus] < busReg[l].capacidad) { busR = busReg[l].bus; capReg[busR]++; break; }
      }
    }
    res.push([p.codigo, p.nombre, p.nota, p.llega, p.sale, turnoId, busI, busR, '', '']);
  });

  _escribir(res);
  return getData();
}

function _escribir(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.ASIG) || ss.insertSheet(SH.ASIG);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 10).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 10).setValues(rows);
}

// Guardar ediciones. rows = [{codigo,nombre,nota,llega,sale,turno,busIda,busReg,manual}]
function guardarAsignacion(rows) {
  var out = rows.map(function (r) {
    return [r.codigo, r.nombre, r.nota, r.llega, r.sale, r.turno, r.busIda, r.busReg, '', r.manual ? 'sí' : ''];
  });
  _escribir(out);
  return getData();
}


// ====== INTERFAZ (HTML incrustado; NO necesitas crear archivo aparte) ======
var INDEX_HTML = `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800;900&display=swap" rel="stylesheet">
  <style>
    @font-face { font-family:'Kaio'; src:url('kaio.otf') format('opentype'); font-weight:400 900; }
    :root{ --amarillo:#FAC012; --azul:#0844A7; --teal:#022B3A; --naranja:#F17121; --verde:#11A154; --rojo:#E61E14; --crema:#EDE0BE; --azulclaro:#2568C8; }
    *{ box-sizing:border-box; }
    body{ font-family:'Poppins',sans-serif; margin:0; color:var(--teal); background:var(--crema); }
    h1,h2,h3{ font-family:'Kaio',sans-serif; text-transform:uppercase; letter-spacing:.02em; margin:0; }
    .banner{ width:100%; display:block; }
    .bunting2{ display:none; }

    .bar{ display:flex; gap:10px; align-items:center; padding:12px 18px; background:#ad0a0a; backdrop-filter:blur(12px); -webkit-backdrop-filter:blur(12px); box-shadow:0 2px 12px rgba(15,45,107,0.15); border-bottom:none; flex-wrap:wrap; }
    .bar button{ background:rgba(255,255,255,0.15) !important; border:1px solid rgba(255,255,255,0.35) !important; color:#fff !important; box-shadow:none; }
    .bar button:hover{ background:rgba(255,255,255,0.25) !important; }
    .bar *{ color:#fff !important; }
    .bar .chip{ color:#fff !important; background:rgba(255,255,255,0.15); border-color:rgba(255,255,255,0.35); }
    button{ font-family:'Poppins',sans-serif; font-weight:800; color:#fff; border:0; padding:10px 16px; border-radius:30px; cursor:pointer; font-size:13px; box-shadow:0 3px 0 rgba(0,0,0,.18); transition:transform .05s, filter .15s; text-transform:uppercase; letter-spacing:.03em; }
    button:active{ transform:translateY(2px); box-shadow:0 1px 0 rgba(0,0,0,.18); }
    button:hover{ filter:brightness(1.06); }
    .b-auto{ background:linear-gradient(135deg,var(--naranja),var(--amarillo)); color:var(--teal); }
    .b-save{ background:linear-gradient(135deg,var(--verde),#0e8a45); }
    .b-reload{ background:linear-gradient(135deg,var(--azul),var(--azulclaro)); }
    button:disabled{ opacity:.5; }
    .chip{ background:var(--crema); border:2px solid var(--teal); border-radius:20px; padding:5px 12px; font-size:12px; font-weight:700; }
    #status{ font-weight:800; color:var(--verde); font-size:13px; }
    .sep{ width:2px; height:28px; background:#e2dcc4; margin:0 4px; }

    /* Tabs de vista */
    .tabs{ display:flex; gap:8px; }
    .tab{ background:#fff; color:var(--teal); border:2px solid var(--teal); border-radius:30px; box-shadow:none; }
    .tab.active{ background:var(--teal); color:#fff; }

    .wrap{ display:flex; gap:14px; padding:16px 18px; align-items:flex-start; }
    .main{ flex:1; min-width:0; }
    .side{ width:300px; flex-shrink:0; }
    .card{ background:#fff; border-radius:14px; box-shadow:0 4px 14px rgba(2,43,58,.12); padding:14px; margin-bottom:14px; border-top:6px solid var(--azul); }
    .card.t-yel{ border-top-color:var(--amarillo); } .card.t-grn{ border-top-color:var(--verde); } .card.t-red{ border-top-color:var(--rojo); } .card.t-org{ border-top-color:var(--naranja); }
    .card h3{ font-size:16px; color:var(--teal); margin-bottom:10px; display:flex; align-items:center; gap:8px; }
    .card h3 small{ font-family:'Poppins'; font-weight:700; text-transform:none; color:var(--azul); font-size:12px; margin-left:auto; }

    /* Slots estilo Excel */
    .slot{ display:flex; align-items:flex-start; gap:10px; padding:8px 6px; border-bottom:1px dashed #e6e0c8; }
    .slot:last-child{ border-bottom:0; }
    .slot .hora{ width:140px; flex-shrink:0; font-weight:700; font-size:12.5px; color:var(--teal); }
    .slot .badge{ flex-shrink:0; }
    .slot .people{ flex:1; display:flex; flex-wrap:wrap; gap:6px; align-items:center; }
    .pchip{ display:inline-flex; align-items:center; gap:6px; background:#eef4ff; border:1.5px solid var(--azul); color:var(--teal); border-radius:16px; padding:3px 6px 3px 10px; font-size:12px; font-weight:700; }
    .pchip.warn{ background:#fff4d6; border-color:var(--amarillo); }
    .pchip.err{ background:#ffe0dd; border-color:var(--rojo); }
    .pchip .x{ cursor:pointer; background:rgba(0,0,0,.12); border-radius:50%; width:16px; height:16px; display:inline-flex; align-items:center; justify-content:center; font-size:11px; line-height:1; }
    .pchip .x:hover{ background:var(--rojo); color:#fff; }
    .addsel{ border:1.5px dashed #b9c6e6; border-radius:16px; padding:3px 8px; font-size:12px; color:var(--azul); background:#fff; font-family:'Poppins'; font-weight:700; max-width:190px; }
    .badge{ display:inline-block; font-size:11px; font-weight:800; padding:2px 9px; border-radius:12px; }
    .b-ok{ background:#cdeed4; color:#0a7a2b; } .b-warn{ background:#ffe9a8; color:#8a5b00; } .b-err{ background:#ffd0cc; color:#a30c04; }
    .section-title{ font-family:'Kaio'; text-transform:uppercase; color:var(--azul); font-size:14px; margin:6px 0 4px; letter-spacing:.04em; }

    .pill{ display:inline-block; font-size:11px; font-weight:700; padding:2px 9px; border-radius:12px; margin:1px 0; }
    .err{ background:#ffd0cc; color:#a30c04; } .warn{ background:#ffe9a8; color:#8a5b00; } .ok{ background:#cdeed4; color:#0a7a2b; }
    .muted{ color:#6b7a82; font-size:12px; }
    .cup-line{ display:flex; align-items:center; gap:8px; margin-bottom:6px; font-size:12.5px; }
    .alert-item{ font-size:12px; padding:8px 10px; border-radius:10px; margin-bottom:7px; }
    .emoji{ font-size:18px; }
  </style>
<script src="https://unpkg.com/lucide@latest"></script>
</head>
<body>
  <img class="banner" src="banner.png" alt="WAVE">
  <div class="bunting2"></div>

  <div class="bar">
    <div class="tabs">
      <button id="tabTurnos" class="tab active" onclick="setView('turnos')"><i data-lucide="folder-open" style="width:15px;height:15px;vertical-align:middle"></i> Turnos</button>
      <button id="tabBuses" class="tab" onclick="setView('buses')"><i data-lucide="bus" style="width:15px;height:15px;vertical-align:middle"></i> Buses</button>
    </div>
    <div class="sep"></div>
    <button class="b-auto" onclick="auto()"><i data-lucide="zap" style="width:15px;height:15px;vertical-align:middle"></i> Auto-sugerir</button>
    <button class="b-save" onclick="save()"><i data-lucide="save" style="width:15px;height:15px;vertical-align:middle"></i> Guardar</button>
    <button class="b-reload" onclick="load()">↻ Recargar</button>
    <span id="status" style="margin-left:auto"></span>
  </div>

  <div class="wrap">
    <div class="main" id="mainView"></div>
    <div class="side">
      <div class="card t-grn"><h3><span class="emoji"><i data-lucide="bus" style="width:15px;height:15px;vertical-align:middle"></i></span> Cupos por bus <small style="font-family:'Poppins';font-weight:700;font-size:11px;margin-left:auto;color:var(--teal);background:var(--crema);border:1.5px solid var(--teal);border-radius:20px;padding:2px 8px;"><i data-lucide="bus" style="width:11px;height:11px;vertical-align:middle"></i> <b id="cfgViaje">–</b>m · <i data-lucide="timer" style="width:11px;height:11px;vertical-align:middle"></i> <b id="cfgMargen">–</b>m</small></h3><div id="cupBuses" class="muted">–</div></div>
      <div class="card t-red"><h3><span class="emoji"><i data-lucide="alert-triangle" style="width:15px;height:15px;vertical-align:middle"></i></span> Alertas (<span id="nalerts">0</span>)</h3><div id="alerts" class="muted">Sin alertas.</div></div>
    </div>
  </div>

<script>
var DATA=null, ROWS=[], VIEW='turnos';

function parseMin(v){ if(v===''||v==null)return null; var s=String(v).trim().toUpperCase();
  var m=s.match(/(\\d{1,2}):(\\d{2})\\s*(AM|PM)?/); if(!m){var h=s.match(/(\\d{1,2})\\s*(AM|PM)/); if(!h)return null; m=[null,h[1],'00',h[2]];}
  var hh=+m[1],mm=+m[2],ap=m[3]; if(ap){if(ap==='PM'&&hh!==12)hh+=12; if(ap==='AM'&&hh===12)hh=0;} if(hh<8)hh+=24; return hh*60+mm; }
function slotById(id){ return (DATA.turnos||[]).filter(function(t){return t.id===id;})[0]; }
function busBy(n){ return (DATA.buses||[]).filter(function(b){return b.bus===n;})[0]; }
function rowBy(cod){ return ROWS.filter(function(r){return r.codigo===cod;})[0]; }
function nameOf(r){ return r.nombre ? r.nombre : r.codigo; }

function load(){ setStatus('Cargando…');
  google.script.run.withSuccessHandler(function(d){ DATA=d;
    document.getElementById('cfgViaje').textContent=d.config.viaje;
    document.getElementById('cfgMargen').textContent=d.config.margen;
    buildRows(); render(); setStatus('Datos cargados <i data-lucide="check" style="width:15px;height:15px;vertical-align:middle"></i>');
  }).withFailureHandler(fail).getData(); }

// Base = quienes van a Wave; overlay = lo guardado en ASIGNACION (por código)
function buildRows(){ ROWS=[];
  var saved={}; (DATA.asignacion||[]).forEach(function(r){ saved[String(r['Código']).trim()]=r; });
  (DATA.participantes||[]).filter(function(p){return p.vaWave;}).forEach(function(p){
    var s=saved[p.codigo];
    ROWS.push({ codigo:p.codigo, nombre:p.nombre, nota:p.nota, llega:p.llega, sale:p.sale,
      turno:s?(s['Turno']||''):'', busIda:s?(s['Bus ida']||''):'', busReg:s?(s['Bus regreso']||''):'',
      manual:s?String(s['Manual']||'').toLowerCase()==='sí':false }); });
}

function setView(v){ VIEW=v;
  document.getElementById('tabTurnos').classList.toggle('active',v==='turnos');
  document.getElementById('tabBuses').classList.toggle('active',v==='buses'); render(); }

// ---------- Validación ----------
function validate(){ var byP={},list=[],turnoCount={},busCount={};
  ROWS.forEach(function(r){ if(r.turno)turnoCount[r.turno]=(turnoCount[r.turno]||0)+1;
    if(r.busIda)busCount[r.busIda]=(busCount[r.busIda]||0)+1; if(r.busReg)busCount[r.busReg]=(busCount[r.busReg]||0)+1; });
  ROWS.forEach(function(r){ var a=[]; var t=r.turno?slotById(r.turno):null; var bi=r.busIda?busBy(r.busIda):null; var br=r.busReg?busBy(r.busReg):null;
    if(!r.turno)a.push({t:'warn',m:'sin turno'});
    if(r.turno&&!r.busIda)a.push({t:'warn',m:'sin bus ida'});
    if(r.nota==null||r.nota==='')a.push({t:'warn',m:'sin nota'});
    if(t&&bi&&t.iniMin!=null&&bi.horaMin!=null){ if(bi.horaMin+DATA.config.viaje+DATA.config.margen>t.iniMin)a.push({t:'err',m:'<i data-lucide="bus" style="width:15px;height:15px;vertical-align:middle"></i> llega tarde al turno'}); }
    if(t&&br&&t.finMin!=null&&br.horaMin!=null&&br.horaMin<t.finMin)a.push({t:'err',m:'regreso antes de terminar'});
    var lp=parseMin(r.llega),sp=parseMin(r.sale);
    if(t&&lp!=null&&t.iniMin!=null&&lp>t.iniMin)a.push({t:'warn',m:'llega tras inicio'});
    if(t&&sp!=null&&t.finMin!=null&&sp<t.finMin)a.push({t:'warn',m:'se va antes de fin'});
    byP[r.codigo]=a; a.forEach(function(x){list.push({codigo:r.codigo,nombre:nameOf(r),t:x.t,m:x.m});}); });
  (DATA.turnos||[]).forEach(function(t){ var c=turnoCount[t.id]||0; if(c>t.necesarios)list.push({t:'err',m:'"'+t.id+'" sobre cupo ('+c+'/'+t.necesarios+')'}); });
  (DATA.buses||[]).forEach(function(b){ var c=busCount[b.bus]||0; if(c>b.capacidad)list.push({t:'err',m:'Bus "'+b.bus+'" sobre cupo ('+c+'/'+b.capacidad+')'}); });
  return {byP:byP,list:list,turnoCount:turnoCount,busCount:busCount}; }

function sev(byP,cod){ var a=byP[cod]||[]; if(a.some(function(x){return x.t==='err';}))return 'err'; if(a.length)return 'warn'; return ''; }
function tip(byP,cod){ var a=byP[cod]||[]; return a.map(function(x){return x.m;}).join(' · '); }

// ---------- Render principal ----------
function render(){ var al=validate();
  document.getElementById('mainView').innerHTML = (VIEW==='turnos') ? renderTurnos(al) : renderBuses(al);
  renderCupos(al); renderAlerts(al); }

function availSelect(filterFn, onchangeCall){
  var opts='<option value="">+ asignar…</option>';
  ROWS.filter(filterFn).sort(function(a,b){var na=a.nota==null?-1:a.nota,nb=b.nota==null?-1:b.nota;return nb-na;})
    .forEach(function(r){ opts+='<option value="'+r.codigo+'">'+nameOf(r)+(r.nota!=null&&r.nota!==''?' ('+r.nota+')':'')+'</option>'; });
  return '<select class="addsel" onchange="'+onchangeCall+'">'+opts+'</select>';
}
function chip(r,al,onx){ var s=sev(al.byP,r.codigo); return '<span class="pchip '+s+'" title="'+tip(al.byP,r.codigo)+'">'+nameOf(r)+'<span class="x" onclick="'+onx+'">✕</span></span>'; }

function renderTurnos(al){
  // agrupar slots por puesto, manteniendo orden
  var order=[], groups={};
  (DATA.turnos||[]).forEach(function(t){ if(!groups[t.puesto]){groups[t.puesto]=[];order.push(t.puesto);} groups[t.puesto].push(t); });
  if(!order.length) return '<div class="card t-yel"><h3><i data-lucide="folder-open" style="width:15px;height:15px;vertical-align:middle"></i> Turnos</h3><div class="muted">No hay turnos. Llena la hoja <b>CAT_TURNOS</b> y recarga.</div></div>';
  var html='';
  order.forEach(function(puesto){
    var enc=groups[puesto][0].encargado;
    html+='<div class="card t-yel"><h3><span class="emoji"><i data-lucide="clipboard-list" style="width:15px;height:15px;vertical-align:middle"></i></span> '+puesto+(enc?' <small><i data-lucide="user" style="width:15px;height:15px;vertical-align:middle"></i> '+enc+'</small>':'')+'</h3>';
    groups[puesto].forEach(function(t){
      var asign=ROWS.filter(function(r){return r.turno===t.id;});
      var c=asign.length, col=c>t.necesarios?'b-err':(c===t.necesarios?'b-ok':'b-warn');
      html+='<div class="slot"><div class="hora">'+t.inicio+'–'+t.fin+'</div>'+
        '<div class="badge '+col+'">'+c+'/'+t.necesarios+'</div><div class="people">';
      asign.forEach(function(r){ html+=chip(r,al,"removeTurno('"+r.codigo+"')"); });
      html+=availSelect(function(r){return !r.turno;}, "assignTurno('"+t.id.replace(/'/g,"\\\\'")+"',this.value)");
      html+='</div></div>';
    });
    html+='</div>';
  });
  return html;
}

function renderBuses(al){
  var ida=(DATA.buses||[]).filter(function(b){return b.tipo==='ida';});
  var reg=(DATA.buses||[]).filter(function(b){return b.tipo==='regreso';});
  function block(list,field){ var h='';
    list.forEach(function(b){ var asign=ROWS.filter(function(r){return r[field]===b.bus;});
      var c=asign.length, col=c>b.capacidad?'b-err':'b-ok';
      h+='<div class="slot"><div class="hora">'+b.bus+' · '+b.hora+'</div><div class="badge '+col+'">'+c+'/'+b.capacidad+'</div><div class="people">';
      asign.forEach(function(r){ h+=chip(r,al,"removeBus('"+r.codigo+"','"+field+"')"); });
      h+=availSelect(function(r){return !r[field];}, "assignBus('"+b.bus.replace(/'/g,"\\\\'")+"','"+field+"',this.value)");
      h+='</div></div>'; });
    return h; }
  var html='<div class="card t-org"><h3><span class="emoji"><i data-lucide="arrow-right" style="width:15px;height:15px;vertical-align:middle"></i></span> Buses de ida</h3>'+
    (ida.length?block(ida,'busIda'):'<div class="muted">Sin buses de ida en CAT_BUSES.</div>')+'</div>';
  html+='<div class="card t-grn"><h3><span class="emoji"><i data-lucide="arrow-left" style="width:15px;height:15px;vertical-align:middle"></i></span> Buses de regreso</h3>'+
    (reg.length?block(reg,'busReg'):'<div class="muted">Sin buses de regreso en CAT_BUSES.</div>')+'</div>';
  return html;
}

// ---------- Acciones ----------
function assignTurno(id,cod){ if(!cod)return; var r=rowBy(cod); if(r){r.turno=id;r.manual=true;render();} }
function removeTurno(cod){ var r=rowBy(cod); if(r){r.turno='';r.manual=true;render();} }
function assignBus(bus,field,cod){ if(!cod)return; var r=rowBy(cod); if(r){r[field]=bus;r.manual=true;render();} }
function removeBus(cod,field){ var r=rowBy(cod); if(r){r[field]='';r.manual=true;render();} }

// ---------- Paneles ----------
function renderCupos(al){
  var hb=(DATA.buses||[]).map(function(b){ var c=al.busCount[b.bus]||0; var col=c>b.capacidad?'err':'ok';
    return '<div class="cup-line"><span class="pill '+col+'">'+c+'/'+b.capacidad+'</span> <span>'+b.bus+'</span> <span class="muted">'+b.tipo+' '+b.hora+'</span></div>'; }).join('');
  document.getElementById('cupBuses').innerHTML=hb||'Sin buses en CAT_BUSES.'; }
function renderAlerts(al){ document.getElementById('nalerts').textContent=al.list.length; var box=document.getElementById('alerts');
  if(!al.list.length){ box.innerHTML='<span class="ok pill">¡Todo en orden! <i data-lucide="check" style="width:15px;height:15px;vertical-align:middle"></i></span>'; return; }
  box.innerHTML=al.list.map(function(x){ var who=x.nombre?('<b>'+x.nombre+'</b>: '):''; return '<div class="alert-item '+x.t+'">'+who+x.m+'</div>'; }).join(''); }

// ---------- Servidor ----------
function save(){ setStatus('Guardando…'); google.script.run.withSuccessHandler(function(d){ DATA=d; setStatus('Guardado <i data-lucide="check" style="width:15px;height:15px;vertical-align:middle"></i>'); }).withFailureHandler(fail).guardarAsignacion(ROWS); }
function auto(){ if(!confirm('Reemplaza la asignación actual con una sugerencia automática (orden: nota → preferencia). ¿Continuar?'))return;
  setStatus('Calculando…'); google.script.run.withSuccessHandler(function(d){ DATA=d; buildRows(); render(); setStatus('¡Sugerencia lista! Revisa y guarda.'); }).withFailureHandler(fail).autoAsignar(); }
function setStatus(s){ document.getElementById('status').textContent=s; }
function fail(e){ setStatus(''); alert('Error: '+(e&&e.message?e.message:e)); }

load();
</script>
<script>lucide.createIcons();</script>
</body>
</html>`;
