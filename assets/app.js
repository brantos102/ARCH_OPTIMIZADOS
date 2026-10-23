/* Interfaz del tablero. Toda la extracción ocurre en assets/parser.worker.js. */
(function () {
  'use strict';

  var MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var historico = {};   // periodo -> resultado
  var actual = null;    // resultado mostrado
  var worker = null;

  var $ = function (id) { return document.getElementById(id); };
  var fmtNum = function (n) { return Number(n || 0).toLocaleString('es-EC'); };
  var fmtPct = function (v) { return v == null ? 'n/d' : (v * 100).toFixed(1).replace('.', ',') + '%'; };
  function fmtDia(isoFecha) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoFecha || ''));
    if (!m) return String(isoFecha || '');
    return Number(m[3]) + ' ' + MESES[+m[2] - 1].slice(0, 3);
  }
  function fmtPeriodo(p) {
    var m = /^(\d{4})-(\d{2})$/.exec(String(p || ''));
    if (!m) return String(p || '');
    var n = MESES[+m[2] - 1] || m[2];
    return n.charAt(0).toUpperCase() + n.slice(1) + ' ' + m[1];
  }
  function corto(proceso) {
    var w = String(proceso || '').split(/\s+/)[0] || '';
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  }
  function el(tag, cls, texto) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (texto != null) n.textContent = texto;
    return n;
  }
  function claseEstado(e) { return e === 'medible' ? 'c-ok' : e === 'parcial' ? 'c-warn' : 'c-bad'; }
  function textoEstado(e) { return e === 'medible' ? 'Medible' : e === 'parcial' ? 'Parcial' : 'No medible'; }
  function colorValor(v) {
    if (v == null) return 'var(--fg-muted)';
    if (v >= 0.999) return 'var(--ok)';
    if (v === 0) return 'var(--bad)';
    return 'var(--warn)';
  }

  /* ------------------------------- tema ------------------------------- */
  try {
    var guardado = localStorage.getItem('tema');
    if (guardado) document.documentElement.setAttribute('data-theme', guardado);
  } catch (e) { /* almacenamiento bloqueado */ }
  $('tema').addEventListener('click', function () {
    var ahora = document.documentElement.getAttribute('data-theme');
    var oscuroPorDefecto = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    var nuevo = ahora ? (ahora === 'dark' ? 'light' : 'dark') : (oscuroPorDefecto ? 'light' : 'dark');
    document.documentElement.setAttribute('data-theme', nuevo);
    try { localStorage.setItem('tema', nuevo); } catch (e) { /* ignorado */ }
    if (actual) pintar(actual);
  });

  /* ------------------------------ carga ------------------------------ */
  function log(texto) {
    var caja = $('log');
    caja.hidden = false;
    caja.appendChild(el('div', null, texto));
    caja.scrollTop = caja.scrollHeight;
  }
  function progreso(pct, texto) {
    var barra = $('progress');
    barra.hidden = false;
    barra.firstElementChild.style.width = Math.max(0, Math.min(100, pct)) + '%';
    if (texto) log(texto);
  }
  /* Los navegadores no permiten Web Workers ni fetch bajo file://, así que
     la página necesita servirse por http. Avisamos antes de que falle. */
  var sinServidor = location.protocol === 'file:';
  if (sinServidor) {
    var aviso = el('p', 'cap');
    aviso.style.color = 'var(--bad)';
    aviso.textContent = 'Esta página está abierta como archivo local y el navegador bloquea la lectura de Excel en ese modo. '
      + 'Sírvela por http: abre una terminal en esta carpeta, ejecuta  python3 -m http.server 8000  y entra a http://localhost:8000';
    $('drop').appendChild(aviso);
    $('elegir').disabled = true;
    $('files').disabled = true;
  }

  function procesar(archivos) {
    if (sinServidor || !archivos || !archivos.length) return;
    $('log').replaceChildren();
    progreso(2, 'Iniciando lectura de ' + archivos.length + ' archivo(s)…');
    $('elegir').disabled = true;
    if (worker) worker.terminate();
    worker = new Worker('assets/parser.worker.js');
    worker.onmessage = function (ev) {
      var m = ev.data || {};
      if (m.tipo === 'progreso') { progreso(m.pct, m.texto); return; }
      if (m.tipo === 'error') {
        progreso(100, 'Error: ' + m.mensaje);
        $('elegir').disabled = false;
        return;
      }
      if (m.tipo === 'listo') {
        progreso(100, 'Listo. ' + m.resultado.hojas.length + ' pestañas leídas.');
        $('elegir').disabled = false;
        if (m.resultado.periodo) historico[m.resultado.periodo] = m.resultado;
        (m.resultado.avisos || []).forEach(log);
        llenarPeriodos(m.resultado.periodo);
        pintar(m.resultado);
        $('descargar').disabled = false;
        $('descargarCsv').disabled = false;
        $('descargarDiag').disabled = false;
      }
    };
    worker.onerror = function (e) {
      progreso(100, 'Error al iniciar el lector: ' + (e.message || 'revisa la consola'));
      $('elegir').disabled = false;
    };
    worker.postMessage({ archivos: Array.prototype.slice.call(archivos) });
  }

  $('elegir').addEventListener('click', function () { $('files').click(); });
  $('files').addEventListener('change', function (e) { procesar(e.target.files); });
  var zona = $('drop');
  ['dragenter', 'dragover'].forEach(function (t) {
    zona.addEventListener(t, function (e) { e.preventDefault(); zona.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (t) {
    zona.addEventListener(t, function (e) { e.preventDefault(); zona.classList.remove('over'); });
  });
  zona.addEventListener('drop', function (e) {
    if (e.dataTransfer && e.dataTransfer.files) procesar(e.dataTransfer.files);
  });

  $('periodo').addEventListener('change', function () {
    var r = historico[$('periodo').value];
    if (r) pintar(r);
  });

  function llenarPeriodos(seleccionado) {
    var sel = $('periodo');
    var claves = Object.keys(historico).sort().reverse();
    sel.replaceChildren.apply(sel, claves.map(function (k) {
      var o = document.createElement('option');
      o.value = k; o.textContent = fmtPeriodo(k);
      return o;
    }));
    sel.disabled = claves.length === 0;
    if (seleccionado && historico[seleccionado]) sel.value = seleccionado;
  }

  /* ----------------------------- descargas ----------------------------- */
  function bajar(nombre, contenido, tipo) {
    var blob = new Blob([contenido], { type: tipo });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }
  $('descargar').addEventListener('click', function () {
    bajar('indicadores.json', JSON.stringify({ ciclos: historico }, null, 1), 'application/json');
  });
  $('descargarDiag').addEventListener('click', function () {
    if (!actual) return;
    bajar('diagnostico-' + (actual.periodo || 'ciclo') + '.json',
      JSON.stringify({ generado: actual.generado, archivos: actual.archivos, hojas: actual.diagnostico || [] }, null, 1),
      'application/json');
  });
  $('descargarCsv').addEventListener('click', function () {
    if (!actual) return;
    var cab = ['periodo','ciclo_inicio','ciclo_fin','proceso','actividad_id','actividad','objetivo_pct','estado_medicion','indicador_sustituto','numerador','denominador','valor_pct','volumen','unidad_volumen','campo_faltante'];
    var esc = function (v) {
      var s = String(v == null ? '' : v);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    var filas = actual.actividades.map(function (a) {
      return [actual.periodo, actual.ciclo ? actual.ciclo.inicio : '', actual.ciclo ? actual.ciclo.fin : '',
        a.proceso, a.actividad_id, a.actividad, a.objetivo_pct, a.estado_medicion, a.indicador_sustituto,
        a.numerador, a.denominador, a.valor == null ? '' : Math.round(a.valor * 1000) / 10,
        a.volumen, a.unidad_volumen, a.campo_faltante].map(esc).join(',');
    });
    bajar('indicadores-' + (actual.periodo || 'ciclo') + '.csv', cab.join(',') + '\n' + filas.join('\n') + '\n', 'text/csv;charset=utf-8');
  });

  /* ------------------------------ pintado ------------------------------ */
  function pintar(r) {
    actual = r;
    var d = r.detalle || {};
    var rec = d.recepcion || {}, alm = d.almacenaje || {}, ali = d.alistamiento || {};

    var lede = $('lede');
    lede.replaceChildren();
    if (r.ciclo) {
      lede.appendChild(document.createTextNode('Ciclo '));
      lede.appendChild(el('b', null, fmtDia(r.ciclo.inicio) + ' – ' + fmtDia(r.ciclo.fin)));
      lede.appendChild(document.createTextNode(': de las '));
      lede.appendChild(el('b', null, String(r.resumen.total)));
      lede.appendChild(document.createTextNode(' actividades del contrato, '));
      lede.appendChild(el('b', null, String(r.resumen.medibles)));
      lede.appendChild(document.createTextNode(' son medibles con los archivos cargados. Faltan '));
      lede.appendChild(el('b', null, String(r.resumen.campos_ausentes)));
      lede.appendChild(document.createTextNode(' campos críticos en el origen, por lo que el cumplimiento no puede acreditarse ni rebatirse.'));
    }

    var tiles = [
      ['Actividades medibles', r.resumen.medibles + ' de ' + r.resumen.total, 'Con todos los campos que exige su fórmula'],
      ['Series recibidas', fmtNum(rec.filas), (rec.dias || 0) + ' días de arribo · ' + fmtNum(rec.duplicadas) + ' duplicadas'],
      ['Pedidos atendidos', fmtNum(ali.pedidos_ciclo), fmtNum(ali.pedidos) + ' en el archivo · ' + fmtNum(ali.filas_vacias) + ' filas vacías'],
      ['Stock al cierre', fmtNum(alm.stock), alm.movimiento
        ? fmtNum(alm.movimiento.ingresos) + ' ingresos · ' + fmtNum(alm.movimiento.salidas) + ' salidas'
        : fmtNum(alm.duplicados) + ' series duplicadas'],
    ];
    $('tiles').replaceChildren.apply($('tiles'), tiles.map(function (t) {
      var c = el('div', 'card');
      c.appendChild(el('div', 'stat-k', t[0]));
      c.appendChild(el('div', 'stat-v', t[1]));
      c.appendChild(el('div', 'stat-n', t[2]));
      return c;
    }));

    var cuerpo = $('tabla').querySelector('tbody');
    cuerpo.replaceChildren.apply(cuerpo, r.actividades.map(function (a) {
      var tr = el('tr');
      tr.appendChild(el('td', null, corto(a.proceso)));
      tr.appendChild(el('td', 'w', a.actividad));
      tr.appendChild(el('td', 'n', fmtPct(a.objetivo_pct / 100)));
      var tdE = el('td');
      tdE.appendChild(el('span', 'pill ' + claseEstado(a.estado_medicion), textoEstado(a.estado_medicion)));
      tr.appendChild(tdE);
      tr.appendChild(el('td', 'w', a.indicador_sustituto));
      var tdV = el('td', 'n', fmtPct(a.valor));
      tdV.style.color = colorValor(a.valor);
      tr.appendChild(tdV);
      tr.appendChild(el('td', 'ww', a.campo_faltante || '—'));
      return tr;
    }));

    var orden = r.actividades.slice().sort(function (a, b) {
      var va = a.valor == null ? 2 : a.valor, vb = b.valor == null ? 2 : b.valor;
      return va - vb;
    });
    $('barras').replaceChildren.apply($('barras'), orden.map(function (a) {
      return filaBarra(a.indicador_sustituto, a.actividad_id + ' ' + corto(a.proceso), a.valor, fmtPct(a.valor));
    }));
    $('eje').replaceChildren.apply($('eje'), ['0,0%', '50,0%', '100,0%'].map(function (t) { return el('span', null, t); }));
    $('notaBarras').textContent = 'Escala completa de 0% a 100%. «n/d» marca la actividad cuyo origen no aporta ninguna medición.';

    var criticos = r.campos.filter(function (c) { return c.critico_para_kpi === 'si'; })
      .sort(function (a, b) { return a.pct_poblado - b.pct_poblado || (a.proceso < b.proceso ? -1 : 1); });
    $('barrasCampos').replaceChildren.apply($('barrasCampos'), criticos.map(function (c) {
      return filaBarra(c.campo, corto(c.proceso), c.pct_poblado / 100, String(c.pct_poblado).replace('.', ',') + '%');
    }));

    barras($('chartPed'), r.diaria.filter(function (x) { return x.metrica === 'Pedidos creados'; }), 'var(--accent)');
    barras($('chartRec'), r.diaria.filter(function (x) { return x.metrica === 'Series recibidas'; }), 'var(--ok)');

    var cl = (ali.clientes || []);
    $('clientes').replaceChildren.apply($('clientes'), cl.map(function (c) {
      return filaBarra(c.cliente, fmtNum(c.items) + ' ítems', c.pct / 100, String(c.pct).replace('.', ',') + '%', 'var(--accent)');
    }));

    var th = $('hojas').querySelector('tbody');
    th.replaceChildren.apply(th, (r.hojas || []).map(function (h) {
      var tr = el('tr');
      tr.appendChild(el('td', 'w', h.nombre));
      var tdT = el('td');
      var pastilla = el('span', 'pill ' + (h.tipo === 'sin reconocer' ? 'c-warn' : 'c-ok'), h.tipo);
      tdT.appendChild(pastilla);
      tr.appendChild(tdT);
      tr.appendChild(el('td', 'n', fmtNum(h.filas)));
      tr.appendChild(el('td', 'ww', h.aporte || '—'));
      return tr;
    }));

    $('pie').textContent = 'Calculado el ' + new Date(r.generado).toLocaleString('es-EC') +
      ' a partir de ' + (r.archivos || []).map(function (a) { return a.nombre; }).join(', ') +
      '. Para conservar el histórico, descarga indicadores.json y súbelo a data/historico.json en el repositorio.';
  }

  function filaBarra(titulo, etiqueta, valor, texto, colorFijo) {
    var row = el('div', 'row');
    var top = el('div', 'row-top');
    var nombre = el('div', 'row-name', titulo + ' ');
    nombre.appendChild(el('span', null, '· ' + etiqueta));
    top.appendChild(nombre);
    var v = el('div', 'row-val', texto);
    v.style.color = colorFijo || colorValor(valor);
    top.appendChild(v);
    row.appendChild(top);
    var track = el('div', 'track');
    var fill = el('i');
    fill.style.width = ((valor == null ? 0 : Math.max(0, Math.min(1, valor))) * 100).toFixed(2) + '%';
    fill.style.background = colorFijo || (valor == null ? 'var(--grid)' : colorValor(valor));
    track.appendChild(fill);
    row.appendChild(track);
    return row;
  }

  /* Gráfico de barras en SVG, sin dependencias. */
  function barras(host, datos, color) {
    host.replaceChildren();
    if (!datos.length) {
      host.appendChild(el('p', 'empty', 'Sin datos para este ciclo.'));
      return;
    }
    var W = 560, H = 210, mT = 10, mB = 30, mL = 46, mR = 8;
    var w = W - mL - mR, h = H - mT - mB;
    var max = Math.max.apply(null, datos.map(function (d) { return d.valor; }));
    var paso = Math.pow(10, Math.floor(Math.log10(max || 1)));
    var tope = Math.ceil(max / paso) * paso || 1;
    var ticks = [0, tope / 2, tope];
    var ancho = w / datos.length;
    var svgNS = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', 'Barras por día');
    function nodo(t, attrs, texto) {
      var n = document.createElementNS(svgNS, t);
      Object.keys(attrs).forEach(function (k) { n.setAttribute(k, attrs[k]); });
      if (texto != null) n.textContent = texto;
      return n;
    }
    ticks.forEach(function (t) {
      var y = mT + h - (t / tope) * h;
      svg.appendChild(nodo('line', { x1: mL, x2: mL + w, y1: y, y2: y, stroke: 'var(--grid)', 'stroke-width': 1 }));
      svg.appendChild(nodo('text', { x: mL - 7, y: y + 4, 'text-anchor': 'end', 'font-size': 11, fill: 'var(--fg-muted)' }, fmtNum(Math.round(t))));
    });
    svg.appendChild(nodo('line', { x1: mL, x2: mL + w, y1: mT + h, y2: mT + h, stroke: 'var(--line)', 'stroke-width': 1 }));
    var saltos = Math.max(1, Math.ceil(datos.length / 8));
    datos.forEach(function (d, i) {
      var alto = Math.max(1, (d.valor / tope) * h);
      var x = mL + i * ancho + ancho * 0.15;
      var bw = Math.max(2, Math.min(28, ancho * 0.7));
      var rect = nodo('rect', { x: x, y: mT + h - alto, width: bw, height: alto, rx: 3, fill: color });
      rect.appendChild(nodo('title', {}, fmtDia(d.fecha) + ': ' + fmtNum(d.valor)));
      svg.appendChild(rect);
      if (i % saltos === 0) {
        svg.appendChild(nodo('text', { x: x + bw / 2, y: mT + h + 16, 'text-anchor': 'middle', 'font-size': 11, fill: 'var(--fg-muted)' }, fmtDia(d.fecha)));
      }
    });
    host.appendChild(svg);
  }

  /* ---------------------- histórico publicado ---------------------- */
  if (sinServidor) return;
  fetch('data/historico.json', { cache: 'no-store' })
    .then(function (r) { return r.ok ? r.json() : null; })
    .then(function (j) {
      if (!j || !j.ciclos) return;
      Object.keys(j.ciclos).forEach(function (k) { if (!historico[k]) historico[k] = j.ciclos[k]; });
      var claves = Object.keys(historico).sort();
      if (!claves.length) return;
      llenarPeriodos(claves[claves.length - 1]);
      pintar(historico[claves[claves.length - 1]]);
      $('descargar').disabled = false;
      $('descargarCsv').disabled = false;
    })
    .catch(function () { /* sin histórico publicado todavía */ });
})();
