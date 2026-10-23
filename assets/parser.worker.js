/* Extrae indicadores de Activos Fijos desde los libros de Excel del proveedor.
   Corre dentro de un Web Worker: el archivo nunca sale del navegador. */
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');

var MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

function norm(v) {
  return String(v == null ? '' : v)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}
function vacio(v) { return v == null || (typeof v === 'string' && v.trim() === ''); }
function aFecha(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && isFinite(v)) {
    var d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
    return isNaN(d) ? null : d;
  }
  if (typeof v === 'string') {
    var s = v.trim();
    var m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
    m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
  }
  return null;
}
function tieneHora(v, d) {
  if (typeof v === 'number') return (v % 1) !== 0;
  if (!d) return false;
  return d.getUTCHours() !== 0 || d.getUTCMinutes() !== 0 || d.getUTCSeconds() !== 0;
}
function iso(d) { return d.toISOString().slice(0, 10); }
function aNumero(v) {
  if (typeof v === 'number') return v;
  if (vacio(v)) return null;
  var n = Number(String(v).replace(/\s/g, '').replace(/,/g, ''));
  return isFinite(n) ? n : null;
}
/* El ciclo de reporte corre del 24 de un mes al 23 del siguiente, tal como lo
   nombra la proforma: «DEL 24 DE AGOSTO AL 23 DE SEPTIEMBRE». */
var CORTE_DIA = 24;
function cicloDe(fecha) {
  var y = fecha.getUTCFullYear(), m = fecha.getUTCMonth(), d = fecha.getUTCDate();
  if (d >= CORTE_DIA) m += 1;
  var p = new Date(Date.UTC(y, m, 1));
  return p.getUTCFullYear() + '-' + String(p.getUTCMonth() + 1).padStart(2, '0');
}
function rangoCiclo(periodo) {
  var m = /^(\d{4})-(\d{2})$/.exec(periodo);
  if (!m) return null;
  var y = +m[1], mes = +m[2] - 1;
  return {
    inicio: iso(new Date(Date.UTC(y, mes - 1, CORTE_DIA))),
    fin: iso(new Date(Date.UTC(y, mes, CORTE_DIA - 1))),
  };
}
function etiquetaPeriodo(p) {
  var m = /^(\d{4})-(\d{2})$/.exec(p || '');
  if (!m) return p || '';
  var n = MESES[+m[2] - 1] || m[2];
  return n.charAt(0).toUpperCase() + n.slice(1) + ' ' + m[1];
}
function diaHabilSiguiente(d) {
  var x = new Date(d.getTime() + 86400000);
  while (x.getUTCDay() === 0 || x.getUTCDay() === 6) x = new Date(x.getTime() + 86400000);
  return x;
}

/* ---- Detección del tipo de pestaña por su contenido, no por su nombre ---- */
/* La firma de cada formato. Se exige más de una columna para que una pestaña
   ajena no entre por coincidir en un solo nombre genérico y contamine los
   totales en silencio, que es peor que quedar fuera. */
function cabecera(filas, limite) {
  for (var i = 0; i < Math.min(filas.length, limite || 12); i++) {
    var f = filas[i] || [];
    var h = f.map(norm);
    var llenas = h.filter(function (c) { return c !== ''; }).length;
    var tiene = function (n) { return h.indexOf(n) >= 0; };

    if (tiene('DOC_EXT') && tiene('FEC_CREAC')) return { fila: i, cols: h, tipo: 'alistamiento' };

    var recep = ['STATUS', 'TIPO', 'SERIE', 'ARTICULO', 'QTY', 'FECHA INICIAL']
      .filter(tiene).length;
    if (tiene('SERIE') && tiene('FECHA INICIAL') && recep >= 4) return { fila: i, cols: h, tipo: 'recepcion' };

    if (tiene('ALMACEN_DESC') && tiene('SERIE')) return { fila: i, cols: h, tipo: 'alm_stock' };

    if (h.some(function (c) { return c.indexOf('CANTIDAD DE UBICACION') >= 0; })) {
      return { fila: i, cols: h, tipo: 'alm_ubicacion' };
    }
    /* Resumen de dos columnas. DESCRIPCION + CANTIDAD por sí solas aparecen en
       muchas pestañas de la proforma, así que solo vale si la tabla no tiene
       más columnas que esas. */
    if (tiene('DESCRIPCION') && tiene('CANTIDAD') && llenas <= 3) {
      return { fila: i, cols: h, tipo: 'alm_cantidad' };
    }
  }
  return null;
}
function buscaCol(cols, re, desde) {
  for (var i = desde || 0; i < cols.length; i++) if (re.test(cols[i])) return i;
  return -1;
}
function todasCols(cols, re) {
  var r = [];
  for (var i = 0; i < cols.length; i++) if (re.test(cols[i])) r.push(i);
  return r;
}

var RE_DISPONIBLE = /DISPONIBILIZ|FECHA.?FIN|FECHA.?DISPONIBLE|HORA.?FIN/;
var RE_DESPACHO = /DESPACH/;
var RE_FRANJA = /FRANJA|JORNADA|HORARIO/;
var RE_CONTEO = /CONTEO|FISICO/;
var RE_ETIQUETA = /ETIQUETA|PROYECTO|ADMINISTRADOR|\bOC\b/;

/* ---------------------------- Extracción ---------------------------- */
function leerRecepcion(filas, cab, ctx) {
  var c = cab.cols;
  var iSerie = buscaCol(c, /^SERIE$/), iFecha = buscaCol(c, /^FECHA INICIAL$/);
  var iArt = buscaCol(c, /^ARTICULO$/), iQty = buscaCol(c, /^QTY$/), iStatus = buscaCol(c, /^STATUS$/);
  var iDisp = buscaCol(c, RE_DISPONIBLE);
  var o = ctx.recepcion;
  o.pestanas.push(cab.nombre);
  if (iDisp >= 0) o.tieneDisponibilizacion = true;
  for (var r = cab.fila + 1; r < filas.length; r++) {
    var f = filas[r] || [];
    if (vacio(f[iSerie]) && vacio(f[iStatus])) continue;
    var serie = String(f[iSerie] == null ? '' : f[iSerie]).trim();
    if (!serie) continue;
    o.filas++;
    o.series[serie] = (o.series[serie] || 0) + 1;
    var bruto = f[iFecha], d = aFecha(bruto);
    if (d) {
      var k = iso(d);
      o.porDia[k] = (o.porDia[k] || 0) + 1;
      o.ciclos[cicloDe(d)] = (o.ciclos[cicloDe(d)] || 0) + 1;
      if (tieneHora(bruto, d)) o.conHora++;
      if (iDisp >= 0) {
        var dd = aFecha(f[iDisp]);
        if (dd) {
          var limite = new Date(d.getTime() + 86400000);
          if (d.getUTCHours() >= 13) limite = diaHabilSiguiente(limite);
          o.sla.total++;
          if (dd <= limite) o.sla.ok++;
        }
      }
    }
    o.campos.SERIE = (o.campos.SERIE || 0) + (serie ? 1 : 0);
    if (!vacio(f[iArt])) o.campos.ARTICULO = (o.campos.ARTICULO || 0) + 1;
    if (!vacio(f[iQty])) o.campos.QTY = (o.campos.QTY || 0) + 1;
    if (!vacio(bruto)) o.campos['FECHA INICIAL'] = (o.campos['FECHA INICIAL'] || 0) + 1;
  }
}

function leerAlmStock(filas, cab, ctx) {
  var c = cab.cols;
  var inicios = todasCols(c, /^SERIE$/);
  var o = ctx.almacenaje;
  o.pestanas.push(cab.nombre);
  inicios.forEach(function (i0, idx) {
    var cols = {
      serie: i0,
      desc: buscaCol(c, /^DESCRIPCION$/, i0),
      almacen: buscaCol(c, /^ALMACEN_DESC$/, i0),
      ubic: buscaCol(c, /^UBICACION$/, i0),
      fing: buscaCol(c, /^FECHA_INGRESO$/, i0),
      cod: buscaCol(c, /^COD ARTICULO$/, i0),
      conteo: buscaCol(c, RE_CONTEO, i0),
      etiq: buscaCol(c, RE_ETIQUETA, i0),
    };
    var lim = inicios[idx + 1] === undefined ? c.length : inicios[idx + 1];
    Object.keys(cols).forEach(function (k) { if (cols[k] >= lim) cols[k] = -1; });
    var bloque = { indice: idx, series: {}, filas: 0, ubic: 0, fing: 0, conteo: 0, etiq: 0, clases: {}, coincide: 0, contados: 0 };
    for (var r = cab.fila + 1; r < filas.length; r++) {
      var f = filas[r] || [];
      if (vacio(f[cols.serie])) continue;
      var serie = String(f[cols.serie]).trim();
      bloque.filas++;
      bloque.series[serie] = (bloque.series[serie] || 0) + 1;
      if (cols.almacen >= 0 && !vacio(f[cols.almacen])) {
        var cl = String(f[cols.almacen]).trim();
        bloque.clases[cl] = (bloque.clases[cl] || 0) + 1;
      }
      if (cols.ubic >= 0 && !vacio(f[cols.ubic])) bloque.ubic++;
      if (cols.fing >= 0 && !vacio(f[cols.fing])) bloque.fing++;
      if (cols.etiq >= 0 && !vacio(f[cols.etiq])) bloque.etiq++;
      if (cols.conteo >= 0 && !vacio(f[cols.conteo])) {
        bloque.conteo++;
        bloque.contados++;
        var sis = aNumero(f[cols.cod >= 0 ? cols.cod : cols.serie]);
        var fis = aNumero(f[cols.conteo]);
        if (fis != null && sis != null && fis === sis) bloque.coincide++;
      }
    }
    if (bloque.filas) o.bloques.push(bloque);
    if (cols.conteo >= 0) o.tieneConteo = true;
    if (cols.etiq >= 0) o.tieneEtiqueta = true;
  });
}

function leerAlmCantidad(filas, cab, ctx) {
  var c = cab.cols;
  var iD = buscaCol(c, /^DESCRIPCION$/), iC = buscaCol(c, /^CANTIDAD$/);
  var o = ctx.almacenaje;
  o.pestanas.push(cab.nombre);
  for (var r = cab.fila + 1; r < filas.length; r++) {
    var f = filas[r] || [];
    if (vacio(f[iD])) continue;
    if (/TOTAL/.test(norm(f[iD]))) continue;
    var n = aNumero(f[iC]);
    if (n == null) continue;
    o.simcards += n;
    o.lineasSimcard++;
  }
}

/* La pestaña apila un bloque por mes, cada uno encabezado por el nombre del
   mes en la columna A con la columna B vacía. Vale el último bloque, que es el
   cierre del ciclo; sumarlos todos mezclaría dos meses. */
function leerAlmUbicacion(filas, cab, ctx) {
  var c = cab.cols;
  var iT = 0, iC = buscaCol(c, /CANTIDAD DE UBICACION/);
  var o = ctx.almacenaje;
  o.pestanas.push(cab.nombre);
  var bloques = [], actual = null;
  for (var r = 0; r < filas.length; r++) {
    var f = filas[r] || [];
    if (vacio(f[iT])) continue;
    var etiqueta = String(f[iT]).trim();
    if (/CANTIDAD|TIPO DE PRODUCTO/.test(norm(etiqueta))) continue;
    var n = aNumero(f[iC]);
    if (n == null) { actual = { mes: etiqueta, total: 0, detalle: [] }; bloques.push(actual); continue; }
    if (!actual) { actual = { mes: '', total: 0, detalle: [] }; bloques.push(actual); }
    actual.total += n;
    actual.detalle.push({ tipo: etiqueta, cantidad: n });
  }
  var ultimo = bloques.length ? bloques[bloques.length - 1] : null;
  if (ultimo) {
    o.ubicaciones += ultimo.total;
    o.detalleUbicaciones = o.detalleUbicaciones.concat(ultimo.detalle);
    if (bloques.length > 1) {
      ctx.avisos.push('«' + cab.nombre + '» trae ' + bloques.length + ' bloques de mes; se tomó el último ('
        + (ultimo.mes || 'sin nombre') + ') como cierre del ciclo.');
    }
  }
}

function leerAlistamiento(filas, cab, ctx) {
  var c = cab.cols;
  var iDoc = buscaCol(c, /^DOC_EXT$/), iFec = buscaCol(c, /^FEC_CREAC$/);
  var iCant = buscaCol(c, /^CANTIDAD$/), iCli = buscaCol(c, /^CLIENTE$/);
  var iTipo = buscaCol(c, /^TIPO ARTICULO$/), iCanal = buscaCol(c, /^CANAL$/);
  var iImei = buscaCol(c, /^IMEI$/), iIcc = buscaCol(c, /^ICC$/), iRasca = buscaCol(c, /^RASCA$/), iAcc = buscaCol(c, /^ACCESORIO$/);
  var iDesp = buscaCol(c, RE_DESPACHO), iFranja = buscaCol(c, RE_FRANJA);
  /* Columnas de evento logístico. Estaban en el archivo desde el principio y el
     lector no llegaba a ellas: la exploración inicial truncó la cabecera en 16
     columnas y el resto del parser se construyó sobre esa suposición. */
  var iEv = buscaCol(c, /^CODIGO_EVENTO$/), iFecEv = buscaCol(c, /^FECHA_EV$/);
  var iServ = buscaCol(c, /^TIPO_SERVICIO$/), iCiudad = buscaCol(c, /^CIUDAD$/);
  var iExp = buscaCol(c, /^EXPRESO$/);
  var o = ctx.alistamiento;
  o.pestanas.push(cab.nombre);
  if (iDesp >= 0) o.tieneDespacho = true;
  if (iFranja >= 0) o.tieneFranja = true;
  for (var r = cab.fila + 1; r < filas.length; r++) {
    var f = filas[r] || [];
    if (vacio(f[iDoc])) { o.filasVacias++; continue; }
    o.filas++;
    var doc = String(f[iDoc]).trim();
    var bruto = f[iFec], d = aFecha(bruto);
    var p = o.pedidos[doc];
    if (!p) {
      p = o.pedidos[doc] = {
        fecha: null, hora: false, items: 0, cli: {}, desp: null, franja: null, pest: cab.nombre,
        eventos: {}, transito: null, servicio: null, ciudad: null, expreso: null,
      };
    }
    p.items++;
    var cli = vacio(f[iCli]) ? '(sin cliente)' : String(f[iCli]).trim();
    p.cli[cli] = (p.cli[cli] || 0) + 1;
    o.itemsCliente[cli] = (o.itemsCliente[cli] || 0) + 1;
    if (!vacio(f[iTipo])) { var t = String(f[iTipo]).trim(); o.tipos[t] = (o.tipos[t] || 0) + 1; }
    if (!vacio(f[iCanal])) { var ca = String(f[iCanal]).trim(); o.canales[ca] = (o.canales[ca] || 0) + 1; }
    if (d && (!p.fecha || d < p.fecha)) p.fecha = d;
    if (tieneHora(bruto, d)) p.hora = true;
    if (iDesp >= 0) { var dd = aFecha(f[iDesp]); if (dd && (!p.desp || dd > p.desp)) p.desp = dd; }
    if (iFranja >= 0 && !vacio(f[iFranja])) p.franja = String(f[iFranja]).trim();
    if (iEv >= 0 && !vacio(f[iEv])) {
      var ev = norm(f[iEv]);
      p.eventos[ev] = true;
      o.eventos[ev] = (o.eventos[ev] || 0) + 1;
    }
    if (iFecEv >= 0) {
      var de = aFecha(f[iFecEv]);
      if (de && d) {
        var horas = (de - d) / 3600000;
        if (horas > 0 && (p.transito === null || horas > p.transito)) p.transito = horas;
      }
    }
    if (iServ >= 0 && !vacio(f[iServ])) p.servicio = String(f[iServ]).trim();
    if (iCiudad >= 0 && !vacio(f[iCiudad])) p.ciudad = String(f[iCiudad]).trim();
    if (iExp >= 0 && !vacio(f[iExp])) p.expreso = String(f[iExp]).trim();
    o.campos.doc_ext = (o.campos.doc_ext || 0) + 1;
    if (!vacio(bruto)) o.campos.FEC_CREAC = (o.campos.FEC_CREAC || 0) + 1;
    if (tieneHora(bruto, d)) o.campos['FEC_CREAC (con hora)'] = (o.campos['FEC_CREAC (con hora)'] || 0) + 1;
    if (!vacio(f[iCant])) o.campos.CANTIDAD = (o.campos.CANTIDAD || 0) + 1;
    if (!vacio(f[iCli])) o.campos.CLIENTE = (o.campos.CLIENTE || 0) + 1;
    if (!vacio(f[iImei]) || !vacio(f[iIcc]) || !vacio(f[iRasca]) || !vacio(f[iAcc])) {
      o.campos['IDENTIFICADOR (IMEI/ICC/RASCA)'] = (o.campos['IDENTIFICADOR (IMEI/ICC/RASCA)'] || 0) + 1;
    }
  }
}

/* Perfil de una pestaña sin exponer valores: solo nombres de columna, tipo y
   cuántas filas la traen poblada. Sirve para extender el lector a formatos
   nuevos sin tener que mover el libro, que es lo que no cabe en ningún lado. */
function perfilar(filas, filaCab) {
  var cab = filas[filaCab] || [];
  var muestra = Math.min(filas.length, filaCab + 1 + 400);
  return cab.slice(0, 40).map(function (h, j) {
    var pobladas = 0, tipos = {};
    for (var r = filaCab + 1; r < muestra; r++) {
      var v = (filas[r] || [])[j];
      if (vacio(v)) continue;
      pobladas++;
      var t = v instanceof Date ? 'fecha' : typeof v === 'number' ? 'numero' : 'texto';
      tipos[t] = (tipos[t] || 0) + 1;
    }
    return {
      col: j,
      nombre: String(h == null ? '' : h).slice(0, 60),
      pobladas: pobladas,
      de: Math.max(0, muestra - filaCab - 1),
      tipos: tipos,
    };
  }).filter(function (c) { return c.nombre !== '' || c.pobladas > 0; });
}

function contexto() {
  return {
    archivos: [], hojas: [], avisos: [], diagnostico: [],
    recepcion: { filas: 0, series: {}, porDia: {}, ciclos: {}, conHora: 0, campos: {}, pestanas: [], tieneDisponibilizacion: false, sla: { ok: 0, total: 0 } },
    almacenaje: { bloques: [], simcards: 0, lineasSimcard: 0, ubicaciones: 0, detalleUbicaciones: [], pestanas: [], tieneConteo: false, tieneEtiqueta: false },
    alistamiento: { filas: 0, filasVacias: 0, pedidos: {}, itemsCliente: {}, tipos: {}, canales: {}, campos: {}, pestanas: [], tieneDespacho: false, tieneFranja: false, eventos: {} },
  };
}

/* ------------------------- Cálculo de indicadores ------------------------- */
function calcular(ctx, periodoPedido) {
  var ali = ctx.alistamiento, rec = ctx.recepcion, alm = ctx.almacenaje;

  var votos = {};
  Object.keys(ali.pedidos).forEach(function (k) {
    var p = ali.pedidos[k];
    if (p.fecha) { var c = cicloDe(p.fecha); votos[c] = (votos[c] || 0) + 1; }
  });
  Object.keys(rec.ciclos).forEach(function (c) { votos[c] = (votos[c] || 0) + rec.ciclos[c] / 100; });
  var periodos = Object.keys(votos).sort();
  var periodo = periodoPedido && votos[periodoPedido] !== undefined ? periodoPedido
    : periodos.sort(function (a, b) { return votos[b] - votos[a] || (a < b ? 1 : -1); })[0] || null;
  periodos = Object.keys(votos).sort();
  var rango = periodo ? rangoCiclo(periodo) : null;
  var dentro = function (isoFecha) { return rango && isoFecha >= rango.inicio && isoFecha <= rango.fin; };

  var seriesRec = Object.keys(rec.series);
  var dupRec = seriesRec.filter(function (s) { return rec.series[s] > 1; }).length;
  var recEnCiclo = 0, recFuera = 0;
  Object.keys(rec.porDia).forEach(function (d) {
    if (dentro(d)) recEnCiclo += rec.porDia[d]; else recFuera += rec.porDia[d];
  });

  var bloqueActual = alm.bloques.length ? alm.bloques[alm.bloques.length - 1] : null;
  var bloquePrevio = alm.bloques.length > 1 ? alm.bloques[alm.bloques.length - 2] : null;
  var stock = bloqueActual ? bloqueActual.filas : 0;
  var stockUnicos = bloqueActual ? Object.keys(bloqueActual.series).length : 0;
  var movimiento = null;
  if (bloqueActual && bloquePrevio) {
    var a = Object.keys(bloquePrevio.series), b = Object.keys(bloqueActual.series);
    var setA = {}, setB = {};
    a.forEach(function (s) { setA[s] = 1; }); b.forEach(function (s) { setB[s] = 1; });
    movimiento = {
      ingresos: b.filter(function (s) { return !setA[s]; }).length,
      salidas: a.filter(function (s) { return !setB[s]; }).length,
      permanecen: b.filter(function (s) { return setA[s]; }).length,
    };
  }

  var pedidos = Object.keys(ali.pedidos).map(function (k) {
    var p = ali.pedidos[k]; p.doc = k; return p;
  });
  var pedidosCiclo = pedidos.filter(function (p) { return p.fecha && dentro(iso(p.fecha)); });
  var conHora = pedidos.filter(function (p) { return p.hora; }).length;
  var porDiaPedidos = {};
  pedidosCiclo.forEach(function (p) { var k = iso(p.fecha); porDiaPedidos[k] = (porDiaPedidos[k] || 0) + 1; });

  /* Indicadores que la línea de eventos sí permite medir. No son los SLA de
     Activos Fijos: miden la entrega al destino, no el picking en bodega. */
  var ENTREGA = { ENTREGADO: 1, PEDIDO_ENTREGADO: 1 };
  var INCIDENCIA = { INCIDENCIA: 1, PEDIDO_CON_INCIDENCIA: 1 };
  var tieneAlguno = function (p, mapa) {
    return Object.keys(p.eventos || {}).some(function (e) { return mapa[e]; });
  };
  var entregados = pedidosCiclo.filter(function (p) { return tieneAlguno(p, ENTREGA); }).length;
  var conIncidencia = pedidosCiclo.filter(function (p) { return tieneAlguno(p, INCIDENCIA); }).length;
  var transitos = pedidosCiclo.map(function (p) { return p.transito; })
    .filter(function (t) { return t !== null && t > 0; }).sort(function (a, b) { return a - b; });
  var percentil = function (lista, q) {
    return lista.length ? lista[Math.min(lista.length - 1, Math.floor(lista.length * q))] : null;
  };
  var dentro72 = transitos.filter(function (t) { return t <= 72; }).length;
  var operativos = pedidosCiclo.length ? [
    {
      id: 'OP-1', nombre: 'Pedidos con evento de entrega registrado',
      detalle: 'El transportista reportó entrega al destino',
      numerador: entregados, denominador: pedidosCiclo.length,
      valor: entregados / pedidosCiclo.length,
    },
    {
      id: 'OP-2', nombre: 'Pedidos sin incidencia',
      detalle: 'Ningún evento de incidencia en la ruta',
      numerador: pedidosCiclo.length - conIncidencia, denominador: pedidosCiclo.length,
      valor: (pedidosCiclo.length - conIncidencia) / pedidosCiclo.length,
    },
  ] : [];
  if (transitos.length) {
    operativos.push({
      id: 'OP-3', nombre: 'Pedidos entregados dentro de 72 h',
      detalle: 'Desde la creación del pedido hasta el último evento registrado',
      numerador: dentro72, denominador: transitos.length, valor: dentro72 / transitos.length,
    });
  }

  var slaAli = { ok: 0, total: 0, oficina: { ok: 0, total: 0 }, extendido: { ok: 0, total: 0 } };
  if (ali.tieneDespacho) {
    pedidosCiclo.forEach(function (p) {
      if (!p.fecha || !p.desp) return;
      var min = (p.desp - p.fecha) / 60000;
      var ok = min <= 120;
      slaAli.total++; if (ok) slaAli.ok++;
      var esOficina = p.franja ? !/EXTEND|FUERA/.test(norm(p.franja))
        : (p.fecha.getUTCHours() >= 8 && p.fecha.getUTCHours() < 18);
      var caja = esOficina ? slaAli.oficina : slaAli.extendido;
      caja.total++; if (ok) caja.ok++;
    });
  }

  function act(id, proceso, actividad, objetivo, medible, campoFaltante, sustituto, num, den, volumen, unidad, fuente) {
    return {
      actividad_id: id, proceso: proceso, actividad: actividad, objetivo_pct: objetivo,
      estado_medicion: medible, campo_faltante: campoFaltante, indicador_sustituto: sustituto,
      numerador: num, denominador: den,
      valor: den > 0 ? num / den : null,
      volumen: volumen, unidad_volumen: unidad, fuente: fuente,
    };
  }

  var P_REC = 'RECEPCIÓN DE EQUIPOS MATERIALES Y/O REPUESTOS';
  var P_ALI = 'ALISTAMIENTO DE PEDIDOS Y DESPACHO';
  var P_ALM = 'ALMACENAJE';
  var P_REP = 'REPORTERÍA';
  var totalRec = rec.filas;
  var ubicPobl = bloqueActual ? bloqueActual.ubic : 0;
  var procesosConArchivo = [rec.filas > 0, ali.filas > 0, stock > 0].filter(Boolean).length;

  var actividades = [
    rec.tieneDisponibilizacion && rec.sla.total > 0
      ? act('AF-1', P_REC, 'Recepción y procesamiento de materiales en 24 h', 100, 'medible', '',
            'Series procesadas dentro del SLA de 24 h', rec.sla.ok, rec.sla.total, totalRec, 'series recibidas', rec.pestanas.join(' + '))
      : act('AF-1', P_REC, 'Recepción y procesamiento de materiales en 24 h', 100, 'no_medible',
            'Fecha y hora de disponibilización en inventario. FECHA INICIAL no trae hora',
            'Unicidad de series recibidas', seriesRec.length - dupRec, seriesRec.length || 0, totalRec, 'series recibidas', rec.pestanas.join(' + ')),

    ali.tieneDespacho && slaAli.oficina.total > 0
      ? act('AF-2', P_ALI, 'Picking y embalaje en 2 h - horario de oficina', 100, 'medible', '',
            'Pedidos despachados dentro de 2 h en horario de oficina', slaAli.oficina.ok, slaAli.oficina.total, pedidosCiclo.length, 'pedidos del ciclo', ali.pestanas.length + ' pestañas')
      : act('AF-2', P_ALI, 'Picking y embalaje en 2 h - horario de oficina', 100, 'no_medible',
            'Qué codigo_evento marca la salida de bodega. Los eventos registrados son de transporte y entrega, no de picking',
            'Pedidos con hora de creación registrada', conHora, pedidos.length, pedidosCiclo.length, 'pedidos del ciclo', ali.pestanas.length + ' pestañas'),

    ali.tieneDespacho && ali.tieneFranja && slaAli.extendido.total > 0
      ? act('AF-3', P_ALI, 'Picking y embalaje en 2 h - horario extendido', 100, 'medible', '',
            'Pedidos despachados dentro de 2 h fuera de horario', slaAli.extendido.ok, slaAli.extendido.total, slaAli.extendido.total, 'pedidos fuera de horario', ali.pestanas.length + ' pestañas')
      : act('AF-3', P_ALI, 'Picking y embalaje en 2 h - horario extendido', 100, 'no_medible',
            'No hay columna de franja horaria. TIPO_SERVICIO distingue NORMAL de EXCLUSIVO, que es nivel de servicio y no jornada',
            'Pedidos clasificables por franja horaria', ali.tieneFranja ? pedidos.length : 0, pedidos.length, 0, 'pedidos fuera de horario identificados', ali.pestanas.length + ' pestañas'),

    ubicPobl > 0
      ? act('AF-4', P_ALM, 'Verificación de etiquetado y categorización', 100, 'medible', '',
            'Registros con ubicación informada', ubicPobl, stock, stock, 'series en stock', alm.pestanas.join(' + '))
      : act('AF-4', P_ALM, 'Verificación de etiquetado y categorización', 100, 'no_medible',
            'UBICACION vacía en el 100% de los registros y las clases del almacén no mapean al contrato',
            'Registros con ubicación informada', 0, stock, stock, 'series en stock', alm.pestanas.join(' + ')),

    (bloqueActual && bloqueActual.contados > 0)
      ? act('AF-5', P_ALM, 'Tomas físicas - coincidencia sistema vs físico', 99, 'medible', '',
            'Coincidencia entre conteo físico y sistema', bloqueActual.coincide, bloqueActual.contados, stock, 'series en stock', alm.pestanas.join(' + '))
      : act('AF-5', P_ALM, 'Tomas físicas - coincidencia sistema vs físico', 99, 'no_medible',
            'No existe columna de conteo físico para contrastar contra el sistema',
            'Unicidad de series en stock', stockUnicos, stock, stock, 'series en stock', alm.pestanas.join(' + ')),

    act('AF-6', P_REP, 'Disponibilidad 7x24 del sistema logístico', 100, 'no_medible',
        'No se entregó registro de disponibilidad ni de caídas del sistema',
        'Sin sustituto disponible', 0, 0, 0, 'horas monitoreadas', 'no entregado'),

    act('AF-7', P_REP, 'Reporte de movimientos e inventario', 100, 'parcial',
        'Fecha y hora de entrega de cada reporte diario',
        'Procesos con archivo del ciclo entregado', procesosConArchivo, 4, procesosConArchivo, 'archivos del ciclo', ctx.archivos.length + ' archivos'),

    act('AF-8', P_REP, 'Reporte de cierre mensual', 100, 'parcial',
        'Fecha y hora de entrega del cierre',
        'Procesos con archivo del ciclo entregado', procesosConArchivo, 4, procesosConArchivo ? 1 : 0, 'cierre del ciclo', ctx.archivos.length + ' archivos'),
  ];

  var campos = [];
  function campo(proceso, pestana, nombre, filas, pobladas, critico) {
    campos.push({
      proceso: proceso, pestana: pestana, campo: nombre, filas: filas, pobladas: pobladas,
      pct_poblado: filas > 0 ? Math.round((pobladas / filas) * 1000) / 10 : 0, critico_para_kpi: critico ? 'si' : 'no',
    });
  }
  var pestRec = rec.pestanas.join(' + ') || 'sin pestaña';
  if (totalRec) {
    campo(P_REC, pestRec, 'SERIE', totalRec, rec.campos.SERIE || 0, false);
    campo(P_REC, pestRec, 'ARTICULO', totalRec, rec.campos.ARTICULO || 0, false);
    campo(P_REC, pestRec, 'FECHA INICIAL', totalRec, rec.campos['FECHA INICIAL'] || 0, true);
    campo(P_REC, pestRec, 'HORA DE ARRIBO', totalRec, rec.conHora, true);
    campo(P_REC, pestRec, 'FECHA/HORA DE DISPONIBILIZACION', totalRec, rec.tieneDisponibilizacion ? rec.sla.total : 0, true);
  }
  var pestAlm = alm.pestanas.join(' + ') || 'sin pestaña';
  if (stock) {
    campo(P_ALM, pestAlm, 'SERIE', stock, bloqueActual.filas, false);
    campo(P_ALM, pestAlm, 'UBICACION', stock, bloqueActual.ubic, true);
    campo(P_ALM, pestAlm, 'FECHA_INGRESO', stock, bloqueActual.fing, true);
    campo(P_ALM, pestAlm, 'CONTEO FISICO', stock, bloqueActual.conteo, true);
    campo(P_ALM, pestAlm, 'ETIQUETA PROYECTO/ADMINISTRADOR/OC', stock, bloqueActual.etiq, true);
  }
  var pestAli = ali.pestanas.length + ' pestañas PICK&PACK';
  if (ali.filas) {
    campo(P_ALI, pestAli, 'doc_ext', ali.filas, ali.campos.doc_ext || 0, false);
    campo(P_ALI, pestAli, 'CLIENTE', ali.filas, ali.campos.CLIENTE || 0, false);
    campo(P_ALI, pestAli, 'IDENTIFICADOR (IMEI/ICC/RASCA)', ali.filas, ali.campos['IDENTIFICADOR (IMEI/ICC/RASCA)'] || 0, false);
    campo(P_ALI, pestAli, 'FEC_CREAC (fecha)', ali.filas, ali.campos.FEC_CREAC || 0, true);
    campo(P_ALI, pestAli, 'FEC_CREAC (con hora)', ali.filas, ali.campos['FEC_CREAC (con hora)'] || 0, true);
    campo(P_ALI, pestAli, 'FECHA/HORA DE DESPACHO', ali.filas, ali.tieneDespacho ? slaAli.total : 0, true);
    campo(P_ALI, pestAli, 'FRANJA HORARIA (oficina/extendido)', ali.filas, ali.tieneFranja ? ali.filas : 0, true);
  }
  campo(P_REP, 'no entregado', 'LOG DE DISPONIBILIDAD DEL SISTEMA', 0, 0, true);
  campo(P_REP, 'no entregado', 'FECHA/HORA DE ENTREGA DE REPORTES', 0, 0, true);

  var diaria = [];
  Object.keys(rec.porDia).sort().forEach(function (d) {
    if (dentro(d)) diaria.push({ fecha: d, proceso: P_REC, metrica: 'Series recibidas', valor: rec.porDia[d] });
  });
  Object.keys(porDiaPedidos).sort().forEach(function (d) {
    diaria.push({ fecha: d, proceso: P_ALI, metrica: 'Pedidos creados', valor: porDiaPedidos[d] });
  });

  var criticos = campos.filter(function (c) { return c.critico_para_kpi === 'si'; });
  var resumen = {
    total: actividades.length,
    medibles: actividades.filter(function (a) { return a.estado_medicion === 'medible'; }).length,
    parciales: actividades.filter(function (a) { return a.estado_medicion === 'parcial'; }).length,
    no_medibles: actividades.filter(function (a) { return a.estado_medicion === 'no_medible'; }).length,
    campos_criticos: criticos.length,
    campos_ausentes: criticos.filter(function (c) { return c.pct_poblado === 0; }).length,
  };

  var clientes = Object.keys(ali.itemsCliente).map(function (k) {
    return { cliente: k, items: ali.itemsCliente[k] };
  }).sort(function (a, b) { return b.items - a.items; });
  var totalItems = clientes.reduce(function (s, c) { return s + c.items; }, 0);
  clientes.forEach(function (c) { c.pct = totalItems ? Math.round((c.items / totalItems) * 1000) / 10 : 0; });

  return {
    generado: new Date().toISOString(),
    periodo: periodo, periodos: periodos, etiqueta: etiquetaPeriodo(periodo), ciclo: rango,
    archivos: ctx.archivos, hojas: ctx.hojas, avisos: ctx.avisos, diagnostico: ctx.diagnostico,
    resumen: resumen, actividades: actividades, operativos: operativos, campos: campos, diaria: diaria,
    detalle: {
      recepcion: {
        filas: totalRec, series_unicas: seriesRec.length, duplicadas: dupRec,
        dias: Object.keys(rec.porDia).length, en_ciclo: recEnCiclo, fuera_ciclo: recFuera, con_hora: rec.conHora,
      },
      almacenaje: {
        stock: stock, unicos: stockUnicos, duplicados: stock - stockUnicos,
        ubicaciones: alm.ubicaciones, simcards: alm.simcards, lineas_simcard: alm.lineasSimcard,
        movimiento: movimiento, clases: bloqueActual ? bloqueActual.clases : {}, bloques: alm.bloques.length,
      },
      alistamiento: {
        filas: ali.filas, filas_vacias: ali.filasVacias, pedidos: pedidos.length,
        pedidos_ciclo: pedidosCiclo.length, con_hora: conHora, clientes: clientes,
        tipos: ali.tipos, canales: ali.canales, eventos: ali.eventos,
        entregados: entregados, con_incidencia: conIncidencia,
        transito_p50: percentil(transitos, 0.5), transito_p90: percentil(transitos, 0.9),
      },
    },
  };
}

/* Cuánto aportó cada pestaña a cada total, para que ninguna suma sea opaca. */
function instantanea(ctx) {
  return {
    recepcion: ctx.recepcion.filas,
    alistamiento: ctx.alistamiento.filas,
    stock: ctx.almacenaje.bloques.reduce(function (a, b) { return a + b.filas; }, 0),
    simcards: ctx.almacenaje.simcards,
    ubicaciones: ctx.almacenaje.ubicaciones,
  };
}
function aporte(antes, despues) {
  var partes = [];
  var etiquetas = {
    recepcion: 'series recibidas', alistamiento: 'filas de pedido',
    stock: 'filas de inventario', simcards: 'unidades', ubicaciones: 'ubicaciones',
  };
  Object.keys(etiquetas).forEach(function (k) {
    var d = despues[k] - antes[k];
    if (d > 0) partes.push(d.toLocaleString('es-EC') + ' ' + etiquetas[k]);
  });
  return partes.join(' · ');
}
/* Para una pestaña sin reconocer, la fila con más celdas de texto corto es la
   mejor candidata a cabecera; es lo que se perfila. */
function buscaFilaMasLlena(filas) {
  var mejor = 0, puntaje = -1;
  for (var i = 0; i < Math.min(filas.length, 15); i++) {
    var n = (filas[i] || []).filter(function (v) {
      return typeof v === 'string' && v.trim() !== '' && v.length < 40;
    }).length;
    if (n > puntaje) { puntaje = n; mejor = i; }
  }
  return mejor;
}

/* ------------------------------- Worker ------------------------------- */
self.onmessage = function (ev) {
  var msg = ev.data || {};
  var archivos = msg.archivos || [];
  var ctx = contexto();
  try {
    for (var i = 0; i < archivos.length; i++) {
      var file = archivos[i];
      self.postMessage({ tipo: 'progreso', pct: Math.round((i / archivos.length) * 90), texto: 'Leyendo ' + file.name + '…' });
      var buf = new Uint8Array(new FileReaderSync().readAsArrayBuffer(file));
      var wb = XLSX.read(buf, { type: 'array', cellDates: true, dense: true });
      ctx.archivos.push({ nombre: file.name, bytes: file.size, hojas: wb.SheetNames.length });
      for (var j = 0; j < wb.SheetNames.length; j++) {
        var nombre = wb.SheetNames[j];
        self.postMessage({
          tipo: 'progreso',
          pct: Math.round(((i + (j + 1) / wb.SheetNames.length) / archivos.length) * 90),
          texto: file.name + ' · ' + nombre,
        });
        var ws = wb.Sheets[nombre];
        if (!ws) continue;
        var filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
        var cab = cabecera(filas);
        if (!cab) {
          ctx.hojas.push({ archivo: file.name, nombre: nombre, tipo: 'sin reconocer', filas: filas.length, aporte: '' });
          ctx.diagnostico.push({
            archivo: file.name, hoja: nombre, tipo: 'sin reconocer', filaCabecera: null,
            filas: filas.length, columnas: perfilar(filas, buscaFilaMasLlena(filas)),
          });
          if (filas.length > 1) ctx.avisos.push('La pestaña «' + nombre + '» no coincide con ningún formato conocido y se omitió.');
          continue;
        }
        cab.nombre = nombre;
        ctx.hojas.push({ archivo: file.name, nombre: nombre, tipo: cab.tipo, filas: Math.max(0, filas.length - cab.fila - 1) });
        ctx.diagnostico.push({
          archivo: file.name, hoja: nombre, tipo: cab.tipo, filaCabecera: cab.fila + 1,
          filas: Math.max(0, filas.length - cab.fila - 1), columnas: perfilar(filas, cab.fila),
        });
        var antes = instantanea(ctx);
        if (cab.tipo === 'recepcion') leerRecepcion(filas, cab, ctx);
        else if (cab.tipo === 'alm_stock') leerAlmStock(filas, cab, ctx);
        else if (cab.tipo === 'alm_cantidad') leerAlmCantidad(filas, cab, ctx);
        else if (cab.tipo === 'alm_ubicacion') leerAlmUbicacion(filas, cab, ctx);
        else if (cab.tipo === 'alistamiento') leerAlistamiento(filas, cab, ctx);
        ctx.hojas[ctx.hojas.length - 1].aporte = aporte(antes, instantanea(ctx));
        filas = null;
      }
      wb = null; buf = null;
    }
    self.postMessage({ tipo: 'progreso', pct: 95, texto: 'Calculando indicadores…' });
    self.postMessage({ tipo: 'listo', resultado: calcular(ctx, msg.periodo) });
  } catch (e) {
    self.postMessage({ tipo: 'error', mensaje: (e && e.message) || String(e) });
  }
};
