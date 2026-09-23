/* ================================================================
   etl.js  ·  Motor de datos del tablero Pintuco
   ----------------------------------------------------------------
   Lee los Excel de la carpeta fuentes/ y arma las estructuras que
   consume el tablero. Corre igual en el navegador (boton Actualizar)
   y en Node (script actualizar.bat).

   API:
     PINTUCO_ETL.compilar(archivos, datosPrevios) -> { datos, avisos, resumen }
       archivos      : { "<nombre.xlsx>": ArrayBuffer|Uint8Array|Buffer }
       datosPrevios  : objeto { D_TIEMPOS, D_VENTAS, D_COMPET, D_RESTO }
                       (lo ya cifrado en datos.js; los modulos que aun
                        no tienen fuente se conservan tal cual)

   FASE 2: reconstruye  tiempos · efectividad · tomas · sell out · actividades
   FASE 4: agregara     ventas · competencia · rotacion · formacion · inventario
   ================================================================ */
(function (root) {
  'use strict';

  var XLSX =
    (typeof root !== 'undefined' && root.XLSX) ? root.XLSX :
    (typeof require === 'function') ? require('xlsx') : null;

  // ---------------------------------------------------------------- utilidades
  var MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO',
               'SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

  function norm(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
  function up(s) { return norm(s).toUpperCase(); }
  function ascii(s) {
    return String(s == null ? '' : s).normalize('NFD')
      .replace(/[̀-ͯ]/g, '').replace(/[^\x20-\x7e]/g, '').toLowerCase().trim();
  }
  function num(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v).replace(/\s/g, '').replace(/%/g, '').replace(/\$/g, '');
    if (s.indexOf(',') > -1 && s.indexOf('.') > -1) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.indexOf(',') > -1) s = s.replace(',', '.');
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  // reubica un valor por orden de magnitud dentro de [lo,hi] (celdas con la coma
  // decimal corrida un numero variable de posiciones en algunos exports)
  function arreglarMagnitud(v, lo, hi) {
    v = +v || 0;
    if (v === 0) return { v: 0, corregido: false };
    var signo = v < 0 ? -1 : 1, a = Math.abs(v), corregido = false, it = 0;
    while (a > hi && it < 30) { a /= 10; it++; corregido = true; }
    while (a > 0 && a < lo && it < 30) { a *= 10; it++; corregido = true; }
    return { v: signo * a, corregido: corregido };
  }
  function segDesdeHora(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number' && v > 0 && v < 2) return Math.round(v * 86400); // fraccion de dia
    var m = String(v).trim().match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?$/);
    if (!m) return 0;
    return (+m[1]) * 3600 + (+m[2]) * 60 + (+(m[3] || 0));
  }
  function fechaISO(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) {
      return v.getFullYear() + '-' + p2(v.getMonth() + 1) + '-' + p2(v.getDate());
    }
    var s = String(v).trim();
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);        // ISO
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);       // dd/mm/yyyy
    if (m) return m[3] + '-' + p2(m[2]) + '-' + p2(m[1]);
    if (typeof v === 'number' && v > 20000 && v < 80000) { // serial Excel
      var d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
      return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate());
    }
    return null;
  }
  function p2(n) { n = +n; return (n < 10 ? '0' : '') + n; }
  function mesDesde(v) {
    var u = up(v);
    for (var i = 0; i < MESES.length; i++) if (u.indexOf(MESES[i]) === 0 || u === MESES[i]) return MESES[i];
    var iso = fechaISO(v);
    if (iso) return MESES[+iso.slice(5, 7) - 1];
    return null;
  }
  function regCorta(s) {
    var r = up(s).replace(/^REGIONAL\s+/, '').trim();
    return r === 'BOGOTA' ? 'BOGOTÁ' : r;
  }

  // interner: array unico + mapa nombre->indice
  function Interner() {
    this.arr = []; this.map = Object.create(null);
  }
  Interner.prototype.id = function (v) {
    var k = v == null ? '' : String(v);
    var i = this.map[k];
    if (i === undefined) { i = this.arr.length; this.arr.push(k); this.map[k] = i; }
    return i;
  };

  // ---------------------------------------------------------------- lectura de hojas
  function parseLibro(bin) {
    return XLSX.read(bin, { type: 'array', cellDates: true, cellNF: false, cellText: false });
  }
  function hojaAOA(wb, nombre) {
    return XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null, blankrows: false });
  }
  // variante que CONSERVA filas en blanco (para hojas-matriz con bloques en posiciones fijas)
  function hojaAOACompleta(wb, nombre) {
    return XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null, blankrows: true });
  }
  function buscarHoja(wb, candidatos) {
    var nombres = wb.SheetNames;
    for (var c = 0; c < candidatos.length; c++) {
      var want = ascii(candidatos[c]);
      for (var i = 0; i < nombres.length; i++) if (ascii(nombres[i]) === want) return nombres[i];
    }
    for (var c2 = 0; c2 < candidatos.length; c2++) {
      var w2 = ascii(candidatos[c2]);
      for (var j = 0; j < nombres.length; j++) if (ascii(nombres[j]).indexOf(w2) > -1) return nombres[j];
    }
    return null;
  }
  // localiza la fila de encabezado (la que contiene TODAS las etiquetas dadas) y mapea columnas
  function mapaColumnas(aoa, requeridas, opcionales, avisos, ctx) {
    opcionales = opcionales || {};
    var maxScan = Math.min(aoa.length, 15);
    for (var r = 0; r < maxScan; r++) {
      var fila = (aoa[r] || []).map(ascii);
      var idx = {}, faltan = [];
      Object.keys(requeridas).forEach(function (clave) {
        var alts = requeridas[clave];
        var found = -1;
        for (var a = 0; a < alts.length; a++) {
          var t = ascii(alts[a]);
          for (var c = 0; c < fila.length; c++) {
            if (fila[c] === t || (t.length > 3 && fila[c].indexOf(t) > -1)) { found = c; break; }
          }
          if (found > -1) break;
        }
        if (found > -1) idx[clave] = found; else faltan.push(clave);
      });
      if (faltan.length === 0) {
        Object.keys(opcionales).forEach(function (clave) {
          var alts = opcionales[clave], found = -1;
          for (var a = 0; a < alts.length; a++) {
            var t = ascii(alts[a]);
            for (var c = 0; c < fila.length; c++) {
              if (fila[c] === t || (t.length > 3 && fila[c].indexOf(t) > -1)) { found = c; break; }
            }
            if (found > -1) break;
          }
          if (found > -1) idx[clave] = found;
        });
        return { headerRow: r, idx: idx };
      }
    }
    avisos.push('[' + ctx + '] no encontre la fila de encabezado con las columnas: ' +
      Object.keys(requeridas).join(', '));
    return null;
  }

  // ---------------------------------------------------------------- selector de archivos
  function elegirArchivo(archivos, patrones) {
    var nombres = Object.keys(archivos);
    for (var p = 0; p < patrones.length; p++) {
      var rx = patrones[p];
      for (var i = 0; i < nombres.length; i++) if (rx.test(ascii(nombres[i]))) return nombres[i];
    }
    return null;
  }

  // ================================================================ MODULOS
  // ---- 1. TIEMPOS DE VISITA  (+ base para EFECTIVIDAD) -------------
  function construirVisitas(wb, avisos) {
    var hoja = buscarHoja(wb, ['Sheet0', 'VISITAS', 'DETALLE', 'BASE']);
    if (!hoja) { avisos.push('[visitas] no hallé la hoja de detalle (Sheet0)'); return null; }
    var aoa = hojaAOA(wb, hoja);
    var req = {
      emp: ['Empleado'], reg: ['Regional'], red: ['Red'], pdv: ['Punto de venta'],
      fecha: ['Fecha de la visita'], mes: ['MES']
    };
    var opc = {
      sup: ['Superior'], perf: ['Perfil del PDV'], cad: ['Cadena'], ciu: ['Ciudad'],
      status: ['Status'], fuera: ['Fuera del itinerario', 'Fuera del itinerar'],
      frec: ['Frecuencia del itinerario'], motivo: ['Motivo para no realizar la vis'],
      chkin: ['Primer check-in manual', 'Check-in automatico', 'Check-in automÃ¡tico'],
      tiempo: ['Tiempo ejecutado en el PDV', 'Tiempo en el PDV automatico']
    };
    var mc = mapaColumnas(aoa, req, opc, avisos, 'visitas');
    if (!mc) return null;
    var I = mc.idx, filas = aoa.slice(mc.headerRow + 1);

    // supervisores = quienes aparecen como "Superior" de alguien
    var sups = Object.create(null);
    if (I.sup != null) filas.forEach(function (f) { var s = up(f[I.sup]); if (s) sups[s] = 1; });

    // ---- 1a. pasada de lectura: objetos intermedios ----
    var brutos = [], fechasSet = Object.create(null), descartadas = 0;
    filas.forEach(function (f) {
      var emp = norm(f[I.emp]), pdv = norm(f[I.pdv]);
      if (!emp || !pdv) { descartadas++; return; }
      var iso = fechaISO(f[I.fecha]);
      var mes = mesDesde(f[I.mes]) || (iso ? MESES[+iso.slice(5, 7) - 1] : null);
      if (!iso && !mes) { descartadas++; return; }
      if (!iso) iso = '2026-' + p2(MESES.indexOf(mes) + 1) + '-01';
      fechasSet[iso] = 1;
      var seg = segDesdeHora(I.tiempo != null ? f[I.tiempo] : 0);
      var status = norm(I.status != null ? f[I.status] : '') || 'Realizada';
      brutos.push({
        iso: iso, mes: mes || MESES[+iso.slice(5, 7) - 1],
        emp: emp, pdv: pdv,
        red: norm(f[I.red]) || '(sin dato)',
        perf: norm(I.perf != null ? f[I.perf] : '') || '(sin dato)',
        reg: norm(f[I.reg]) || '(sin dato)',
        cad: norm(I.cad != null ? f[I.cad] : '') || '(sin dato)',
        rol: sups[up(emp)] ? 'SUPERVISOR' : 'PROMOTOR',
        st: status,
        ciu: norm(I.ciu != null ? f[I.ciu] : '') || '(sin dato)',
        sc: seg,
        realizada:
          (I.chkin != null && norm(f[I.chkin])) ? true :
          seg > 0 ? true :
          (I.motivo != null && norm(f[I.motivo])) ? false :
          up(status) !== 'RECHAZADA',
        fuera: I.fuera != null && /^S/i.test(norm(f[I.fuera]))
      });
    });

    // ---- 1b. dimension fechas ordenada + interners para el resto ----
    var fechas = Object.keys(fechasSet).sort();
    var fIdx = Object.create(null);
    fechas.forEach(function (d, i) { fIdx[d] = i; });
    var IN = {
      emp: new Interner(), red: new Interner(), perf: new Interner(), pdv: new Interner(),
      reg: new Interner(), cad: new Interner(), rol: new Interner(), st: new Interner(), ciu: new Interner()
    };
    brutos.sort(function (a, b) { return a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : 0; });

    var recs = brutos.map(function (r) {
      return {
        f: fIdx[r.iso], e: IN.emp.id(r.emp), rd: IN.red.id(r.red), pf: IN.perf.id(r.perf),
        pv: IN.pdv.id(r.pdv), rg: IN.reg.id(r.reg), cd: IN.cad.id(r.cad), rl: IN.rol.id(r.rol),
        st: IN.st.id(r.st), ci: IN.ciu.id(r.ciu), sc: r.sc,
        _mes: r.mes, _red: r.red, _reg: r.reg, _realizada: r.realizada, _fuera: r.fuera
      };
    });

    var out = {
      fechas: fechas, emp: IN.emp.arr, red: IN.red.arr, perf: IN.perf.arr, pdv: IN.pdv.arr,
      reg: IN.reg.arr, cad: IN.cad.arr, rol: IN.rol.arr, st: IN.st.arr, ciu: IN.ciu.arr,
      rows: recs.map(function (r) {
        return [r.f, r.e, r.rd, r.pf, r.pv, r.rg, r.cd, r.rl, r.st, r.ci, r.sc].join(',');
      }).join(';'),
      meta: {
        registros: recs.length,
        outliers12h: recs.filter(function (r) { return r.sc > 43200; }).length
      }
    };
    if (descartadas) avisos.push('[visitas] ' + descartadas + ' filas sin empleado/PDV/fecha se omitieron');
    return { D_TIEMPOS: out, recs: recs };
  }

  // ---- 2. EFECTIVIDAD  (deriva de las visitas) --------------------
  function construirEfectividad(recs, avisos) {
    // canal:  GRANDES CADENAS -> CADENAS ;  resto -> TRADICIONAL
    function canal(red) { return up(red).indexOf('CADENA') > -1 ? 'CADENAS' : 'TRADICIONAL'; }
    var acc = Object.create(null);       // mes|canal|regcorta -> [prog, vis]
    recs.forEach(function (r) {
      if (!r._mes || !r._reg) return;
      var k = r._mes + '|' + canal(r._red) + '|' + regCorta(r._reg);
      var o = acc[k] || (acc[k] = [0, 0]);
      if (!r._fuera) o[0]++;                 // programada = dentro de itinerario
      if (r._realizada) o[1]++;              // visitada = ejecutada (incluye extrarutas)
    });
    var rows = [], nac = Object.create(null);
    Object.keys(acc).forEach(function (k) {
      var p = k.split('|'), o = acc[k];
      rows.push([p[0], p[1], p[2], o[0], o[1]]);
      var nk = p[0] + '|' + p[1];
      var n = nac[nk] || (nac[nk] = [0, 0]); n[0] += o[0]; n[1] += o[1];
    });
    Object.keys(nac).forEach(function (nk) {
      var p = nk.split('|'), o = nac[nk];
      rows.push([p[0], p[1], 'NACIONAL', o[0], o[1]]);
    });
    // orden: mes cronologico, luego canal, luego regional (NACIONAL primero)
    rows.sort(function (a, b) {
      return MESES.indexOf(a[0]) - MESES.indexOf(b[0]) ||
        a[1].localeCompare(b[1]) ||
        (a[2] === 'NACIONAL' ? -1 : b[2] === 'NACIONAL' ? 1 : a[2].localeCompare(b[2]));
    });
    if (!rows.length) avisos.push('[efectividad] no se generaron filas');
    return { rows: rows };
  }

  // ---- 3. TOMAS DE ALMACEN --------------------------------------
  function construirTomas(archivos, wbs, avisos) {
    var nom = elegirArchivo(archivos, [/ejecucion.*toma/, /toma.*ejecucion/, /^ejecucion toma/]);
    if (!nom) { avisos.push('[tomas] no encontré "EJECUCION TOMAS ####.xlsx"'); return null; }
    var wb = wbs[nom];
    var rows = [];
    [['CADENAS', 'Cadenas'], ['TRADICIONAL', 'Tradicional']].forEach(function (par) {
      var hoja = buscarHoja(wb, [par[0]]);
      if (!hoja) { avisos.push('[tomas] falta la hoja ' + par[0] + ' en ' + nom); return; }
      var aoa = hojaAOA(wb, hoja);
      var mc = mapaColumnas(aoa,
        { mes: ['MES'], reg: ['Regional'], obj: ['Objetivo'], efe: ['Efectivas', 'Efectivos'] },
        {}, avisos, 'tomas/' + par[0]);
      if (!mc) return;
      var I = mc.idx;
      aoa.slice(mc.headerRow + 1).forEach(function (f) {
        var mes = mesDesde(f[I.mes]); if (!mes) return;
        var reg = norm(f[I.reg]); if (!reg) return;
        rows.push([mes, reg, par[1], num(f[I.obj]), num(f[I.efe]), 2026]);
      });
    });
    rows.sort(function (a, b) {
      return MESES.indexOf(a[0]) - MESES.indexOf(b[0]) || a[2].localeCompare(b[2]) || a[1].localeCompare(b[1]);
    });
    if (!rows.length) { avisos.push('[tomas] sin filas'); return null; }
    return { rows: rows };
  }

  // ---- 4. SELL OUT CADENAS ------------------------------------
  function construirSellout(archivos, wbs, avisos) {
    var nom = elegirArchivo(archivos, [/cuota.*cadena/, /cadena.*cuota/, /sell.?out/]);
    if (!nom) { avisos.push('[sellout] no encontré "CUOTA NACIONAL CADENAS ####.xlsx"'); return null; }
    var wb = wbs[nom];
    var hoja = buscarHoja(wb, ['Cuota Cadenas', 'Cuota', 'Detalle']);
    if (!hoja) { avisos.push('[sellout] falta la hoja "Cuota Cadenas"'); return null; }
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      mes: ['MES'], reg: ['Regional'], tie: ['Tienda'],
      cd: ['Cuota Dinero'], cl: ['Cuota Litros'],
      ad: ['Cumplimiento Dinero'], al: ['Cumplimiento Litros']
    }, {}, avisos, 'sellout');
    if (!mc) return null;
    var I = mc.idx, rows = [];
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var mes = mesDesde(f[I.mes]); if (!mes) return;
      var reg = norm(f[I.reg]), tie = norm(f[I.tie]); if (!reg || !tie) return;
      rows.push([mes, 2026, reg, tie, num(f[I.cd]), num(f[I.cl]), num(f[I.ad]), num(f[I.al])]);
    });
    rows.sort(function (a, b) {
      return MESES.indexOf(a[0]) - MESES.indexOf(b[0]) || a[2].localeCompare(b[2]) || a[3].localeCompare(b[3]);
    });
    if (!rows.length) { avisos.push('[sellout] sin filas'); return null; }
    return { rows: rows };
  }

  // ---- 5. ACTIVIDADES ---------------------------------------------
  function construirActividades(archivos, wbs, avisos) {
    var nom = elegirArchivo(archivos, [/actividad/]);
    if (!nom) { avisos.push('[actividades] no encontré "ACTIVIDADES ####.xlsx"'); return null; }
    var wb = wbs[nom];
    var out = { gen: [], pdv: [] };

    var hG = buscarHoja(wb, ['GENERAL', 'ACTIVIDADES GENERAL']);
    if (hG) {
      var aoaG = hojaAOA(wb, hG);
      var mcG = mapaColumnas(aoaG, {
        gc: ['Nombre general campaña', 'Nombre general campa', 'Nombre general campana'],
        cp: ['Nombre campaña', 'Nombre campa', 'Nombre campana'],
        reg: ['Regional'], obj: ['Objetivo'], eje: ['Ejecutado']
      }, {
        ciu: ['Ciudad'], sup: ['Supervisor'],
        mes: ['Mes ejecucion', 'Mes ejecución', 'Mes vigencia']
      }, avisos, 'actividades/GENERAL');
      if (mcG) {
        var G = mcG.idx;
        aoaG.slice(mcG.headerRow + 1).forEach(function (f) {
          var cp = norm(f[G.cp]); if (!cp && !norm(f[G.gc])) return;
          var mes = mesDesde(G.mes != null ? f[G.mes] : '') || null;
          out.gen.push([
            norm(f[G.gc]) || cp, cp || norm(f[G.gc]),
            norm(f[G.reg]), norm(G.ciu != null ? f[G.ciu] : ''),
            norm(G.sup != null ? f[G.sup] : ''), mes || '',
            num(f[G.obj]), num(f[G.eje])
          ]);
        });
      }
    } else avisos.push('[actividades] falta la hoja GENERAL');

    var hP = buscarHoja(wb, ['POR PDV', 'ACTIVIDADES POR PDV', 'PDV']);
    if (hP) {
      var aoaP = hojaAOA(wb, hP);
      var mcP = mapaColumnas(aoaP, {
        pdv: ['Nombre pdv', 'Nombre pdv '], reg: ['Regional'],
        cp: ['Nombre campaña', 'Nombre campa', 'Nombre campana'],
        eje: ['Ejecutado']
      }, { mes: ['Mes'], sup: ['Supervisor'] }, avisos, 'actividades/POR PDV');
      if (mcP) {
        var P = mcP.idx;
        aoaP.slice(mcP.headerRow + 1).forEach(function (f) {
          var pdv = norm(f[P.pdv]); if (!pdv) return;
          var eje = up(f[P.eje]); eje = eje.indexOf('SI') === 0 ? 'SI' : eje.indexOf('NO') === 0 ? 'NO' : eje;
          out.pdv.push([
            pdv, norm(f[P.reg]), norm(f[P.cp]),
            up(P.mes != null ? f[P.mes] : '') || '',
            eje, norm(P.sup != null ? f[P.sup] : '')
          ]);
        });
      }
    } else avisos.push('[actividades] falta la hoja POR PDV');

    if (!out.gen.length && !out.pdv.length) return null;
    return out;
  }

  // ---- 6. PARTICIPACION EN LINEAL (share of shelf) ---------------
  // Marcas propias de Pintuco (para separar Pintuco / competencia). Editar si hace falta.
  var RX_PINTUCO = /(PINTUCO|VINILTEX|KORAZA|PINTULUX|AQUA ?BLOCK|PINTUTRAFICO|PINTUCOAT|ACUALUX|AEROCOLOR|DOM[EÉ]STICO|INTERVINIL|TERINSA|ESTUCOMASTIC|ACRILT|MADETEX|KOLTRON|ICO[ -])/i;
  function construirParticipacion(archivos, wbs, avisos) {
    var nom = elegirArchivo(archivos, [/participaci/, /share.*shelf/, /lineal/]);
    if (!nom) return null;
    var wb = wbs[nom];
    var hoja = buscarHoja(wb, ['report', 'Hoja1', 'Datos']);
    if (!hoja) { avisos.push('[participación] no hallé la hoja de datos'); return null; }
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      marca: ['Marca'], cat: ['Categoria de producto', 'Categoría de producto'],
      reg: ['Regional'], cad: ['Cadena'], pdv: ['PDV'],
      cmCat: ['Total centimetros', 'Total centímetros', 'cent'],
      cmMarca: ['Participacion en CM', 'Participación en CM']
    }, {
      linea: ['Linea de producto', 'Línea de producto'], ciu: ['Ciudad'],
      superc: ['Supercategoria', 'Supercategoría'], enc: ['ID de la encuesta']
    }, avisos, 'participación');
    if (!mc) return null;
    var I = mc.idx, rows = [], encuestas = {}, pdvs = {};
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var marca = norm(f[I.marca]); if (!marca) return;
      var linea = norm(I.linea != null ? f[I.linea] : '');
      var cat = norm(f[I.cat]).replace(/^L_/, '').replace(/_/g, ' ');
      var grp = RX_PINTUCO.test(linea || marca) ? 'Pintuco'
              : /OTRAS MARCAS/i.test(linea || marca) ? 'Otras marcas'
              : 'Competencia';
      var cmM = num(f[I.cmMarca]);
      rows.push([
        norm(f[I.reg]), norm(f[I.cad]), norm(f[I.pdv]), norm(I.ciu != null ? f[I.ciu] : ''),
        cat, marca, grp, cmM
      ]);
      pdvs[norm(f[I.pdv])] = 1;
      if (I.enc != null) encuestas[String(f[I.enc])] = num(f[I.cmCat]);
    });
    if (!rows.length) return null;
    return {
      rows: rows,
      meta: {
        encuestas: Object.keys(encuestas).length || null,
        pdv: Object.keys(pdvs).length,
        archivo: nom
      }
    };
  }

  // ---- 7. VENTAS DETALLADAS (plantilla ventas-detalladas.xlsx) --------
  //  Una fila por  SKU x tienda x mes  (o el detalle que tengas). El ETL agrega.
  //  Columnas: MES/PERIODO, REGIONAL, TIENDA, MARCA, [NEGOCIO], [LINEA],
  //            [SUBNEGOCIO], [SKU], UNIDADES, LITROS, DINERO
  function periodoYM(v) {
    var iso = fechaISO(v);
    if (iso) return iso.slice(0, 7);
    var s = up(v);
    var y = (s.match(/\b(20\d\d)\b/) || [])[1];
    for (var i = 0; i < MESES.length; i++) if (s.indexOf(MESES[i]) > -1)
      return (y || '2026') + '-' + p2(i + 1);
    var m = s.match(/^(\d{4})[-/](\d{1,2})/);
    if (m) return m[1] + '-' + p2(m[2]);
    return null;
  }
  function construirVentas(archivos, wbs, avisos) {
    var nom = elegirArchivo(archivos, [/venta.*detallada/, /detalle.*venta/, /ventas?.*sku/,
      /sell.?out.*sku/, /ventas?-?detalladas/]);
    if (!nom) return null;
    var wb = wbs[nom];
    var hoja = buscarHoja(wb, ['ventas', 'detalle', 'datos', 'Hoja1', 'Sheet1']) || wb.SheetNames[0];
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      per: ['Mes', 'Periodo', 'Período', 'Fecha'],
      reg: ['Regional', 'Region'],
      tie: ['Tienda', 'PDV', 'Almacen', 'Almacén', 'Punto de venta'],
      marca: ['Marca'],
      unid: ['Unidades', 'Cantidad', 'Und'],
      dinero: ['Dinero', 'Venta $', 'Venta pesos', 'Valor', 'Importe'],
      litros: ['Litros', 'Volumen', 'LT']
    }, {
      neg: ['Negocio', 'Unidad de negocio'],
      lin: ['Linea', 'Línea', 'Linea de producto'],
      sub: ['Subnegocio', 'Sub negocio', 'Segmento'],
      sku: ['SKU', 'Codigo', 'Código', 'Referencia', 'Cod producto'],
      anio: ['Año', 'Anio', 'Year']
    }, avisos, 'ventas');
    if (!mc) { avisos.push('[ventas] la hoja no tiene las columnas mínimas (MES, REGIONAL, TIENDA, MARCA, UNIDADES, LITROS, DINERO)'); return null; }
    var I = mc.idx;
    var IN = { per: new Interner(), reg: new Interner(), tienda: new Interner(), marca: new Interner(),
      negocio: new Interner(), linea: new Interner(), subneg: new Interner() };
    var combos = Object.create(null), skus = Object.create(null);
    var filas = 0, sinD = 0, sinL = 0, descart = 0;
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var añoTxt = I.anio != null ? f[I.anio] : '';
      var per = periodoYM((añoTxt ? añoTxt + ' ' : '') + norm(f[I.per]) || f[I.per]);
      var reg = norm(f[I.reg]), tie = norm(f[I.tie]), mar = norm(f[I.marca]);
      if (!per || !reg || !tie || !mar) { descart++; return; }
      var un = num(f[I.unid]), di = num(f[I.dinero]), li = num(f[I.litros]);
      var neg = norm(I.neg != null ? f[I.neg] : '') || '(sin negocio)';
      var lin = norm(I.lin != null ? f[I.lin] : '') || '(sin línea)';
      var sub = norm(I.sub != null ? f[I.sub] : '') || '(sin subnegocio)';
      var pi = IN.per.id(per), ri = IN.reg.id(reg), ti = IN.tienda.id(tie), mi = IN.marca.id(mar),
        gi = IN.negocio.id(neg), li2 = IN.linea.id(lin), si = IN.subneg.id(sub);
      var k = [pi, ri, ti, mi, gi, li2, si].join(',');
      var o = combos[k] || (combos[k] = [0, 0, 0, 0, 0, 0]); // lineas, unid, dinero, litros, ld, ll
      o[0]++; o[1] += un; o[2] += di; o[3] += li;
      if (di > 0) o[4]++; else sinD++;
      if (li > 0) o[5]++; else sinL++;
      if (I.sku != null) { var s = norm(f[I.sku]); if (s) skus[s] = 1; }
      filas++;
    });
    if (!filas) { avisos.push('[ventas] sin filas válidas'); return null; }
    var perOrden = IN.per.arr.map(function (p, i) { return [p, i]; }).sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    // remap período a orden cronológico
    var mapPer = {}; perOrden.forEach(function (e, nuevo) { mapPer[e[1]] = nuevo; });
    var perArr = perOrden.map(function (e) { return e[0]; });
    var rows = Object.keys(combos).map(function (k) {
      var p = k.split(',').map(Number), o = combos[k];
      p[0] = mapPer[p[0]];
      return p.join(',') + ',' + o[0] + ',' + o[1] + ',' + o[2] + ',' + o[3] + ',' + o[4] + ',' + o[5];
    }).join(';');
    if (descart) avisos.push('[ventas] ' + descart + ' filas sin mes/regional/tienda/marca se omitieron');
    return {
      dims: {
        per: perArr, reg: IN.reg.arr, tienda: IN.tienda.arr, marca: IN.marca.arr,
        negocio: IN.negocio.arr, linea: IN.linea.arr, subneg: IN.subneg.arr, asesor: []
      },
      rows: rows,
      meta: {
        filas: filas, combos: Object.keys(combos).length,
        sinDinero: sinD, sinLitros: sinL,
        skus: Object.keys(skus).length || null
      }
    };
  }

  // ================================================================
  //  LIBRO MAESTRO — "DATOS PROMPT PINTUCO ...xlsx" (una sola fuente,
  //  con hojas ya limpias: Tiempos de visita, Efectividad general,
  //  Rotación, Formación, Tomas de almacén, Cumplimientos Sell Out
  //  Cadenas, Actividades general/por PDV, Ventas detalladas cadenas,
  //  Inventario en CEDI, Competencia, Insights de mercado).
  //  Si aparece en fuentes/, se usa en vez de las 7 fuentes sueltas.
  // ================================================================
  function buscarLibroMaestro(wbs) {
    for (var nom in wbs) {
      var wb = wbs[nom];
      if (buscarHoja(wb, ['Tiempos de visita']) && buscarHoja(wb, ['Ventas detalladas cadenas'])) return nom;
    }
    return null;
  }

  function tiemposMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Tiempos de visita']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var IDX = { emp: 0, reg: 2, red: 4, cad: 5, pdv: 6, perf: 9, ciu: 14, fecha: 16, fuera: 18, status: 20, tmin: 43, mes: 51, rol: 52 };
    var filas = aoa.slice(1);
    var brutos = [], fechasSet = Object.create(null), descartadas = 0;
    filas.forEach(function (f) {
      var emp = norm(f[IDX.emp]), pdv = norm(f[IDX.pdv]);
      if (!emp || !pdv) { descartadas++; return; }
      var iso = fechaISO(f[IDX.fecha]);
      var mes = up(f[IDX.mes]) || (iso ? MESES[+iso.slice(5, 7) - 1] : null);
      if (!iso && !mes) { descartadas++; return; }
      if (!iso) iso = '2026-' + p2(MESES.indexOf(mes) + 1) + '-01';
      fechasSet[iso] = 1;
      var seg = Math.round(num(f[IDX.tmin]) * 60);
      var status = norm(f[IDX.status]) || 'Realizada';
      brutos.push({
        iso: iso, mes: mes || MESES[+iso.slice(5, 7) - 1],
        emp: emp, pdv: pdv, red: norm(f[IDX.red]) || '(sin dato)',
        perf: norm(f[IDX.perf]) || '(sin dato)', reg: norm(f[IDX.reg]) || '(sin dato)',
        cad: norm(f[IDX.cad]) || '(sin dato)', rol: up(f[IDX.rol]) || 'PROMOTOR',
        st: status, ciu: norm(f[IDX.ciu]) || '(sin dato)', sc: seg,
        realizada: seg > 0 || /REALIZADA|APROBADA/.test(up(status)),
        fuera: /^S/i.test(norm(f[IDX.fuera]) || '')
      });
    });
    var fechas = Object.keys(fechasSet).sort();
    var fIdx = Object.create(null); fechas.forEach(function (d, i) { fIdx[d] = i; });
    var IN = {
      emp: new Interner(), red: new Interner(), perf: new Interner(), pdv: new Interner(),
      reg: new Interner(), cad: new Interner(), rol: new Interner(), st: new Interner(), ciu: new Interner()
    };
    brutos.sort(function (a, b) { return a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : 0; });
    var recs = brutos.map(function (r) {
      return {
        f: fIdx[r.iso], e: IN.emp.id(r.emp), rd: IN.red.id(r.red), pf: IN.perf.id(r.perf),
        pv: IN.pdv.id(r.pdv), rg: IN.reg.id(r.reg), cd: IN.cad.id(r.cad), rl: IN.rol.id(r.rol),
        st: IN.st.id(r.st), ci: IN.ciu.id(r.ciu), sc: r.sc,
        _mes: r.mes, _red: r.red, _reg: r.reg, _realizada: r.realizada, _fuera: r.fuera
      };
    });
    var out = {
      fechas: fechas, emp: IN.emp.arr, red: IN.red.arr, perf: IN.perf.arr, pdv: IN.pdv.arr,
      reg: IN.reg.arr, cad: IN.cad.arr, rol: IN.rol.arr, st: IN.st.arr, ciu: IN.ciu.arr,
      rows: recs.map(function (r) { return [r.f, r.e, r.rd, r.pf, r.pv, r.rg, r.cd, r.rl, r.st, r.ci, r.sc].join(','); }).join(';'),
      meta: { registros: recs.length, outliers12h: recs.filter(function (r) { return r.sc > 43200; }).length }
    };
    if (descartadas) avisos.push('[tiempos] ' + descartadas + ' filas sin empleado/PDV/fecha se omitieron');
    if (!recs.length) { avisos.push('[tiempos] "Tiempos de visita" sin filas utilizables'); return null; }
    return out;
  }

  function efectividadMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Efectividad general']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      mes: ['MES'], prog: ['Frecuencias programadas'], real: ['Frecuencias realizadas'],
      canal: ['Canal'], reg: ['Regional']
    }, {}, avisos, 'efectividad(maestro)');
    if (!mc) return null;
    var I = mc.idx, rows = [];
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var mes = mesDesde(f[I.mes]); if (!mes) return;
      var reg = norm(f[I.reg]); if (!reg) return;
      rows.push([mes, up(f[I.canal]) || '(sin dato)', reg, num(f[I.prog]), num(f[I.real])]);
    });
    rows.sort(function (a, b) {
      return MESES.indexOf(a[0]) - MESES.indexOf(b[0]) || a[1].localeCompare(b[1])
        || (a[2] === 'NACIONAL' ? -1 : b[2] === 'NACIONAL' ? 1 : a[2].localeCompare(b[2]));
    });
    if (!rows.length) { avisos.push('[efectividad] "Efectividad general" sin filas'); return null; }
    return { rows: rows };
  }

  function tomasMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Tomas de almacén', 'Tomas de almacen']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      mes: ['MES'], reg: ['Regional'], obj: ['Objetivo'], efe: ['Efectivas'], canal: ['Canal']
    }, { anio: ['Año', 'Ano'] }, avisos, 'tomas(maestro)');
    if (!mc) return null;
    var I = mc.idx, rows = [];
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var mes = mesDesde(f[I.mes]); if (!mes) return;
      var reg = norm(f[I.reg]); if (!reg) return;
      rows.push([mes, reg, norm(f[I.canal]) || 'Cadenas', num(f[I.obj]), num(f[I.efe]), (I.anio != null ? num(f[I.anio]) : 0) || 2026]);
    });
    rows.sort(function (a, b) { return MESES.indexOf(a[0]) - MESES.indexOf(b[0]) || a[2].localeCompare(b[2]) || a[1].localeCompare(b[1]); });
    if (!rows.length) { avisos.push('[tomas] "Tomas de almacén" sin filas'); return null; }
    return { rows: rows };
  }

  function selloutMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Cumplimientos Sell Out Cadenas']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      mes: ['MES'], reg: ['Regional'], tie: ['Tienda'],
      cd: ['Cuota Dinero'], cl: ['Cuota Litros'], ad: ['Avance Dinero'], al: ['Avance en Litros', 'Avance Litros']
    }, { anio: ['Año', 'Ano'] }, avisos, 'sellout(maestro)');
    if (!mc) return null;
    var I = mc.idx, rows = [];
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var mes = mesDesde(f[I.mes]); if (!mes) return;
      var reg = norm(f[I.reg]), tie = norm(f[I.tie]); if (!reg || !tie) return;
      var anio = (I.anio != null ? num(f[I.anio]) : 0) || 2026;
      rows.push([mes, anio, reg, tie, num(f[I.cd]), num(f[I.cl]), num(f[I.ad]), num(f[I.al])]);
    });
    rows.sort(function (a, b) {
      return (a[1] - b[1]) || (MESES.indexOf(a[0]) - MESES.indexOf(b[0])) || a[2].localeCompare(b[2]) || a[3].localeCompare(b[3]);
    });
    if (!rows.length) { avisos.push('[sellout] "Cumplimientos Sell Out Cadenas" sin filas'); return null; }
    return { rows: rows };
  }

  function actividadesMaestro(wb, avisos) {
    var hG = buscarHoja(wb, ['Actividades general']);
    var hP = buscarHoja(wb, ['Actividades por PDV']);
    if (!hG && !hP) return null;
    var out = { gen: [], pdv: [] };
    if (hG) {
      var aoaG = hojaAOA(wb, hG);
      var mcG = mapaColumnas(aoaG, {
        gc: ['Nombre general campaña'], cp: ['Nombre campaña'],
        reg: ['Regional'], obj: ['Objetivo'], eje: ['Ejecutado']
      }, { ciu: ['Ciudad'], sup: ['Supervisor'], mes: ['Mes ejecucion', 'Mes ejecución'] }, avisos, 'actividades general(maestro)');
      if (mcG) {
        var G = mcG.idx;
        aoaG.slice(mcG.headerRow + 1).forEach(function (f) {
          var cp = norm(f[G.cp]); if (!cp && !norm(f[G.gc])) return;
          var mes = mesDesde(G.mes != null ? f[G.mes] : '') || null;
          out.gen.push([
            norm(f[G.gc]) || cp, cp || norm(f[G.gc]), norm(f[G.reg]),
            norm(G.ciu != null ? f[G.ciu] : ''), norm(G.sup != null ? f[G.sup] : ''), mes || '',
            num(f[G.obj]), num(f[G.eje])
          ]);
        });
      }
    }
    if (hP) {
      var aoaP = hojaAOA(wb, hP);
      var mcP = mapaColumnas(aoaP, {
        pdv: ['Nombre pdv'], reg: ['Regional'], cp: ['Nombre campaña'], eje: ['Ejecutado']
      }, { mes: ['Mes'], sup: ['Supervisor'] }, avisos, 'actividades pdv(maestro)');
      if (mcP) {
        var P = mcP.idx;
        aoaP.slice(mcP.headerRow + 1).forEach(function (f) {
          var pdv = norm(f[P.pdv]); if (!pdv) return;
          var eje = up(f[P.eje]); eje = eje.indexOf('SI') === 0 ? 'SI' : eje.indexOf('NO') === 0 ? 'NO' : eje;
          out.pdv.push([pdv, norm(f[P.reg]), norm(f[P.cp]), up(P.mes != null ? f[P.mes] : '') || '', eje, norm(P.sup != null ? f[P.sup] : '')]);
        });
      }
    }
    if (!out.gen.length && !out.pdv.length) return null;
    return out;
  }

  function ventasMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Ventas detalladas cadenas']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      reg: ['TIENDA_REGIONAL'], tienda: ['TIENDA_NOMBRE'], marca: ['SKU_MARCA'],
      negocio: ['SKU_NEGOCIO'], linea: ['SKU_LINEA'], subneg: ['SKU_SUBNEGOCIO'],
      anio: ['VENTAS_FECHA (Year)'], unid: ['VENTA_UNIDADES'], dinero: ['VENTA_DINERO'], litros: ['VENTA_LITROS']
    }, { sku: ['SKU_CODIGO'], mesTxt: ['VENTAS_FECHA (Month)'] }, avisos, 'ventas(maestro)');
    if (!mc) return null;
    var I = mc.idx;
    var MESES_EN = { january: 0, february: 1, march: 2, april: 3, may: 4, june: 5, july: 6, august: 7, september: 8, october: 9, november: 10, december: 11 };
    function periodo(f) {
      var anio = num(f[I.anio]);
      var mraw = I.mesTxt != null ? f[I.mesTxt] : null;
      // algunas filas (mayo, en ambos años) traen aquí una fecha cruda en vez
      // del texto "May 2025"; se usa el mes de esa fecha directo.
      if (mraw instanceof Date && !isNaN(mraw)) {
        if (!anio) anio = mraw.getFullYear();
        return anio + '-' + p2(mraw.getMonth() + 1);
      }
      var mtxt = ascii(mraw);
      var mi = -1;
      for (var k in MESES_EN) if (mtxt.indexOf(k) === 0) { mi = MESES_EN[k]; break; }
      if (mi < 0 || !anio) return null;
      return anio + '-' + p2(mi + 1);
    }
    var IN = {
      per: new Interner(), reg: new Interner(), tienda: new Interner(), marca: new Interner(),
      negocio: new Interner(), linea: new Interner(), subneg: new Interner()
    };
    var combos = Object.create(null), skus = Object.create(null);
    var filas = 0, sinD = 0, sinL = 0, descart = 0, corregidos = 0;
    var filasAll = aoa.slice(mc.headerRow + 1);
    for (var idx = 0; idx < filasAll.length; idx++) {
      var f = filasAll[idx];
      var per = periodo(f);
      var reg = norm(f[I.reg]), tie = norm(f[I.tienda]), mar = norm(f[I.marca]);
      if (!per || !reg || !tie || !mar) { descart++; continue; }
      var un = num(f[I.unid]);
      // el export trae algunas celdas de dinero/litros con la coma corrida un
      // numero variable de posiciones (x1e5, x1e9, x1e16...); se reubican por
      // orden de magnitud dentro de un rango plausible para una linea de venta.
      var di = arreglarMagnitud(num(f[I.dinero]), 100, 1e8);
      if (di.corregido) corregidos++;
      var li = arreglarMagnitud(num(f[I.litros]), 0.01, 8000).v;
      di = di.v;
      var neg = norm(I.negocio != null ? f[I.negocio] : '') || '(sin negocio)';
      var lin = norm(I.linea != null ? f[I.linea] : '') || '(sin línea)';
      var sub = norm(I.subneg != null ? f[I.subneg] : '') || '(sin subnegocio)';
      var pi = IN.per.id(per), ri = IN.reg.id(reg), ti = IN.tienda.id(tie), mi2 = IN.marca.id(mar),
        gi = IN.negocio.id(neg), li2 = IN.linea.id(lin), si = IN.subneg.id(sub);
      var k = pi + ',' + ri + ',' + ti + ',' + mi2 + ',' + gi + ',' + li2 + ',' + si;
      var o = combos[k] || (combos[k] = [0, 0, 0, 0, 0, 0]);
      o[0]++; o[1] += un; o[2] += di; o[3] += li;
      if (di > 0) o[4]++; else sinD++;
      if (li > 0) o[5]++; else sinL++;
      if (I.sku != null) { var s = norm(f[I.sku]); if (s) skus[s] = 1; }
      filas++;
    }
    if (!filas) { avisos.push('[ventas] sin filas válidas en "Ventas detalladas cadenas"'); return null; }
    var perOrden = IN.per.arr.map(function (p, i) { return [p, i]; }).sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    var mapPer = {}; perOrden.forEach(function (e, nuevo) { mapPer[e[1]] = nuevo; });
    var perArr = perOrden.map(function (e) { return e[0]; });
    var rowsArr = [];
    for (var k2 in combos) {
      var p = k2.split(',').map(Number), o = combos[k2];
      p[0] = mapPer[p[0]];
      rowsArr.push(p.join(',') + ',' + o[0] + ',' + o[1] + ',' + o[2] + ',' + o[3] + ',' + o[4] + ',' + o[5]);
    }
    if (descart) avisos.push('[ventas] ' + descart + ' filas sin período/regional/tienda/marca se omitieron (de ' + (filas + descart) + ')');
    if (corregidos) avisos.push('[ventas] ' + corregidos + ' montos con formato inválido (>1e12) se corrigieron ÷1e11 — vale la pena auditar el total de dinero');
    return {
      dims: { per: perArr, reg: IN.reg.arr, tienda: IN.tienda.arr, marca: IN.marca.arr, negocio: IN.negocio.arr, linea: IN.linea.arr, subneg: IN.subneg.arr, asesor: [] },
      rows: rowsArr.join(';'),
      meta: { filas: filas, combos: rowsArr.length, sinDinero: sinD, sinLitros: sinL, skus: Object.keys(skus).length || null }
    };
  }

  function inventarioMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Inventario en CEDI']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      cedi: ['nombre_punto'], bodega: ['nombre_bodega'], cod: ['articulo'], art: ['nombre_articulo'],
      pres: ['presentacion'], estado: ['estado_articulo'], cant: ['cantidad'], disp: ['disponible'],
      alistar: ['por_alistar'], valor: ['valor_total']
    }, { doc: ['documento1'], doc2: ['documento2'] }, avisos, 'inventario(maestro)');
    if (!mc) return null;
    var I = mc.idx, rows = [];
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var cedi = norm(f[I.cedi]); if (!cedi) return;
      rows.push([
        cedi, norm(f[I.bodega]), norm(f[I.cod]), norm(f[I.art]), norm(f[I.pres]),
        up(f[I.estado]) || '(sin dato)', norm(I.doc != null ? f[I.doc] : ''),
        num(f[I.cant]), num(f[I.disp]), num(f[I.alistar]), 0,
        num(f[I.valor]), '', norm(I.doc2 != null ? f[I.doc2] : (I.doc != null ? f[I.doc] : ''))
      ]);
    });
    if (!rows.length) { avisos.push('[inventario] "Inventario en CEDI" sin filas'); return null; }
    return { rows: rows };
  }

  function insightsMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Insights de mercado']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      tend: ['Tendencia'], que: ['Que esta pasando', 'Qué esta pasando'], impl: ['Implicacion para Pintuco']
    }, {}, avisos, 'insights(maestro)');
    if (!mc) return null;
    var I = mc.idx, tend = [];
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      var t = norm(f[I.tend]); if (!t) return;
      tend.push([t, norm(f[I.que]), norm(f[I.impl])]);
    });
    if (!tend.length) return null;
    return { tend: tend, items: [] };
  }

  function competenciaMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Competencia']);
    if (!hoja) return null;
    var aoa = hojaAOA(wb, hoja);
    var mc = mapaColumnas(aoa, {
      reg: ['Regional'], cad: ['Cadena'], pdv: ['PDV'], mes: ['fecha'],
      tiene: ['Tiene actividad para reportar'], comp: ['Indica cuál competidor realiza la actividad'],
      tipo: ['Tipo de actividad']
    }, { desc: ['Describe la mecánica de la actividad'] }, avisos, 'competencia(maestro)');
    if (!mc) return null;
    var I = mc.idx;
    var IN = { reg: new Interner(), cad: new Interner(), pdv: new Interner(), comp: new Interner(), cat: new Interner(), conf: new Interner() };
    var det = [], pdvsTodos = Object.create(null), filasTotal = 0, conActividad = 0;
    var catCount = Object.create(null), catWords = Object.create(null);
    aoa.slice(mc.headerRow + 1).forEach(function (f) {
      filasTotal++;
      var pdv = norm(f[I.pdv]); if (pdv) pdvsTodos[pdv] = 1;
      if (up(f[I.tiene]).indexOf('SI') !== 0) return;
      var cat = norm(f[I.tipo]); if (!cat) return;
      conActividad++;
      var mes = mesDesde(f[I.mes]); var mesN = mes ? MESES.indexOf(mes) + 1 : 0;
      var desc = norm(I.desc != null ? f[I.desc] : '');
      catCount[cat] = (catCount[cat] || 0) + 1;
      if (desc && !catWords[cat]) catWords[cat] = desc.slice(0, 80);
      var catIdx = IN.cat.id(cat);
      det.push([
        mesN, 0, IN.cad.id(norm(f[I.cad]) || '(sin dato)'), IN.reg.id(norm(f[I.reg]) || '(sin dato)'), 0,
        IN.pdv.id(pdv || '(sin dato)'), 0, IN.comp.id(norm(f[I.comp]) || '(sin dato)'), 0,
        catIdx, catIdx, IN.conf.id('Alta'), desc, desc
      ]);
    });
    if (!det.length) { avisos.push('[competencia] "Competencia" no tiene filas con actividad reportada'); return null; }
    var dicc = Object.keys(catCount).map(function (c) { return { cat: c, sub: c, n: catCount[c], words: catWords[c] || '' }; });
    return {
      dims: { mes: [], red: [], cad: IN.cad.arr, ciu: [], pdv: IN.pdv.arr, emp: [], comp: IN.comp.arr, tipo: [], cat: IN.cat.arr, sub: IN.cat.arr, conf: IN.conf.arr, reg: IN.reg.arr, det: [], dicc: [], cob: [] },
      det: det, dicc: dicc,
      meta: { filas: filasTotal, conActividad: conActividad, pdvMonitoreados: Object.keys(pdvsTodos).length }
    };
  }

  function rotacionMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Rotación', 'Rotacion']);
    if (!hoja) return null;
    var aoa = hojaAOACompleta(wb, hoja);
    function hallar(regex) {
      for (var i = 0; i < aoa.length; i++) { var v = aoa[i] && aoa[i][0]; if (v && regex.test(up(v))) return i; }
      return -1;
    }
    function filaBloque(i) {
      var r = aoa[i] || [];
      var out = [norm(r[0])];
      for (var m = 1; m <= 12; m++) out.push(r[m] == null ? '' : r[m]);
      out.push(r[13] == null ? '' : r[13]);
      out.push(r[14] == null ? '' : r[14]);
      return out;
    }
    function bloque(inicio, fin) {
      var out = [];
      for (var i = inicio; i < fin; i++) {
        if (!aoa[i] || aoa[i][0] == null || norm(aoa[i][0]) === '') continue;
        out.push(filaBloque(i));
      }
      return out;
    }
    var iGen = hallar(/ROTACION GENERAL/);
    var iCargo = hallar(/ROTACION POR CARGO/);
    var iCanal = hallar(/ROTACION POR CANAL/);
    var iReg = hallar(/ROTACION POR REGIONAL/);
    var iSel = hallar(/INDICADORES DE SELECCION|CICLO DE CONTRATACION/);
    if (iGen < 0) { avisos.push('[rotación] no encontré el bloque "ROTACION GENERAL" en la hoja'); return null; }
    var marcas = [iCargo, iCanal, iReg, iSel, aoa.length].filter(function (x) { return x >= 0; });
    function limite(desde) { for (var k = 0; k < marcas.length; k++) if (marcas[k] > desde) return marcas[k]; return aoa.length; }
    var general = bloque(iGen + 1, limite(iGen));
    var cargo = iCargo >= 0 ? bloque(iCargo + 1, limite(iCargo)) : [];
    var canal = iCanal >= 0 ? bloque(iCanal + 1, limite(iCanal)) : [];
    var regional = iReg >= 0 ? bloque(iReg + 1, limite(iReg)) : [];
    var seleccion = [];
    if (iSel >= 0) {
      for (var i = iSel + 1; i < limite(iSel); i++) {
        if (aoa[i] && aoa[i][0] && norm(aoa[i][0]) && !/INDICADOR/i.test(up(aoa[i][0]))) { seleccion.push(filaBloque(i)); break; }
      }
    }
    if (!general.length) { avisos.push('[rotación] el bloque "ROTACION GENERAL" salió vacío'); return null; }
    return { general: general, cargo: cargo, canal: canal, regional: regional, seleccion: seleccion };
  }

  function formacionMaestro(wb, avisos) {
    var hoja = buscarHoja(wb, ['Formación', 'Formacion']);
    if (!hoja) return null;
    var aoa = hojaAOACompleta(wb, hoja);
    function hallar(regex) {
      for (var i = 0; i < aoa.length; i++) { var v = aoa[i] && aoa[i][0]; if (v && regex.test(up(v))) return i; }
      return -1;
    }
    var iIni = hallar(/FORMACION INICIAL/);
    var iHacku = hallar(/^HACKU/);
    var iHoras = hallar(/HORAS DE FORMACION/);
    var iCursos = -1;
    for (var i = 0; i < aoa.length; i++) {
      var r = aoa[i];
      if (r && r[0] && up(norm(r[0])) === 'MES' && r[1] && /CURSO/i.test(up(r[1]))) { iCursos = i; break; }
    }
    var marcas = [iIni, iHacku, iHoras, iCursos, aoa.length].filter(function (x) { return x >= 0; });
    function limite(desde) { for (var k = 0; k < marcas.length; k++) if (marcas[k] > desde) return marcas[k]; return aoa.length; }

    var inicial = [];
    if (iIni >= 0) {
      var fin = limite(iIni);
      var filaEnc = aoa[iIni + 1] || [];
      var nMeses = 2; for (var c = 2; c < filaEnc.length; c++) if (filaEnc[c]) nMeses = c;
      var ultimaReg = '';
      for (var i2 = iIni + 2; i2 < fin; i2++) {
        var r2 = aoa[i2]; if (!r2) continue;
        var reg = norm(r2[0]), can = norm(r2[1]);
        if (reg) ultimaReg = reg;
        if (!can) continue;
        var fila = [ultimaReg, can];
        for (var m = 2; m <= nMeses; m++) fila.push(r2[m] == null ? '' : r2[m]);
        inicial.push(fila);
      }
    }
    var hacku = [];
    if (iHacku >= 0) {
      for (var i3 = iHacku + 2; i3 < limite(iHacku); i3++) {
        var r3 = aoa[i3]; if (!r3 || !r3[0]) continue;
        hacku.push([norm(r3[0]), r3[1] == null ? '' : r3[1], r3[2] == null ? '' : r3[2]]);
      }
    }
    var horas = [];
    if (iHoras >= 0) {
      for (var i4 = iHoras + 2; i4 < limite(iHoras); i4++) {
        var r4 = aoa[i4]; if (!r4 || !r4[0]) continue;
        horas.push([norm(r4[0]), r4[1] == null ? '' : r4[1]]);
      }
    }
    var cursos = [];
    if (iCursos >= 0) {
      var ultimoMes = '';
      for (var i5 = iCursos + 1; i5 < aoa.length; i5++) {
        var r5 = aoa[i5]; if (!r5 || !r5[1]) continue;
        var mes = norm(r5[0]); if (mes) ultimoMes = mes;
        cursos.push([ultimoMes, norm(r5[1]), num(r5[2]), num(r5[3]), num(r5[4]), num(r5[5]), num(r5[6])]);
      }
    }
    if (!inicial.length && !cursos.length) { avisos.push('[formación] no reconocí los bloques de la hoja "Formación"'); return null; }
    return { inicial: inicial, cursos: cursos, horas: horas, hacku: hacku };
  }

  function construirDesdeMaestro(wbs, avisos, prev) {
    var nomM = buscarLibroMaestro(wbs);
    if (!nomM) return null;
    var wb = wbs[nomM];
    avisos.push('[maestro] usando "' + nomM + '" como fuente consolidada única');
    var prevR = prev.D_RESTO || {};
    return {
      archivo: nomM,
      D_TIEMPOS: tiemposMaestro(wb, avisos),
      D_VENTAS: ventasMaestro(wb, avisos),
      D_COMPET: competenciaMaestro(wb, avisos) || prev.D_COMPET || null,
      D_RESTO: {
        efectividad: efectividadMaestro(wb, avisos) || prevR.efectividad || null,
        tomas: tomasMaestro(wb, avisos) || prevR.tomas || null,
        sellout: selloutMaestro(wb, avisos) || prevR.sellout || null,
        actividades: actividadesMaestro(wb, avisos) || prevR.actividades || null,
        rotacion: rotacionMaestro(wb, avisos),
        formacion: formacionMaestro(wb, avisos),
        inventario: inventarioMaestro(wb, avisos),
        insights: insightsMaestro(wb, avisos) || prevR.insights || null
      }
    };
  }

  // ---- modulos aun sin fuente: marcador "en pausa" --------------
  function marcadorPausa(archivos, patrones, archivoEsperado, columnas) {
    var hay = elegirArchivo(archivos, patrones);
    return {
      pausa: true,
      archivo: hay ? '' : archivoEsperado,
      detalle: hay
        ? ('Detecté <b>' + hay + '</b> en <code>fuentes/</code>, pero el mapeo de este módulo '
           + 'todavía no está configurado. Avísame para terminarlo con ese archivo.')
        : ('Todavía no hay fuente para este módulo. Necesita: ' + columnas + '.')
    };
  }

  // ================================================================ COMPILADOR
  function compilar(archivos, datosPrevios) {
    if (!XLSX) throw new Error('No hay librería XLSX disponible (vendor/xlsx.full.min.js).');
    var avisos = [];
    var prev = datosPrevios || {};
    var prevR = prev.D_RESTO || {};

    // parsear todos los libros una sola vez
    var wbs = {};
    Object.keys(archivos).forEach(function (nom) {
      try {
        var b = archivos[nom];
        var arr = (b instanceof Uint8Array) ? b :
                  (typeof Buffer !== 'undefined' && Buffer.isBuffer(b)) ? new Uint8Array(b) :
                  new Uint8Array(b);
        wbs[nom] = parseLibro(arr);
      } catch (e) { avisos.push('No pude leer ' + nom + ': ' + e.message); }
    });

    var D_TIEMPOS, efect, tomas, sellout, actividades, D_VENTAS, D_COMPET, rotacion, formacion, inventario, insights;
    var maestro = construirDesdeMaestro(wbs, avisos, prev);

    if (maestro) {
      // ---- camino 1: un solo libro maestro con todas las hojas ----
      D_TIEMPOS = maestro.D_TIEMPOS || prev.D_TIEMPOS || null;
      efect = maestro.D_RESTO.efectividad;
      tomas = maestro.D_RESTO.tomas;
      sellout = maestro.D_RESTO.sellout;
      actividades = maestro.D_RESTO.actividades;
      D_VENTAS = maestro.D_VENTAS;
      D_COMPET = maestro.D_COMPET;
      rotacion = maestro.D_RESTO.rotacion;
      formacion = maestro.D_RESTO.formacion;
      inventario = maestro.D_RESTO.inventario;
      insights = maestro.D_RESTO.insights;
    } else {
      // ---- camino 2 (compatibilidad): las 7 fuentes sueltas de antes ----
      D_TIEMPOS = prev.D_TIEMPOS || null;
      efect = prevR.efectividad || null;
      var nomVis = elegirArchivo(archivos, [/informe.*visita/, /visita.*gerencial/, /gerencial.*visita/]);
      if (nomVis && wbs[nomVis]) {
        var v = construirVisitas(wbs[nomVis], avisos);
        if (v) { D_TIEMPOS = v.D_TIEMPOS; efect = construirEfectividad(v.recs, avisos); }
      } else {
        avisos.push('[visitas] no encontré "informe-gerencial-visitas ####.xlsx" — se conserva lo anterior');
      }
      tomas = construirTomas(archivos, wbs, avisos) || prevR.tomas || null;
      sellout = construirSellout(archivos, wbs, avisos) || prevR.sellout || null;
      actividades = construirActividades(archivos, wbs, avisos) || prevR.actividades || null;
      D_VENTAS = construirVentas(archivos, wbs, avisos);
      D_COMPET = prev.D_COMPET || null; // se conserva (homologación manual)
      rotacion = null; formacion = null; inventario = null;
      insights = prevR.insights || null;
    }

    // participación en lineal: fuente aparte (PARTICIPACION*.xlsx), independiente del maestro
    var participacion = construirParticipacion(archivos, wbs, avisos) || prevR.participacion || null;

    var D_RESTO = {};
    Object.keys(prevR).forEach(function (k) { D_RESTO[k] = prevR[k]; }); // arranca del anterior
    if (efect) D_RESTO.efectividad = efect;
    if (tomas) D_RESTO.tomas = tomas;
    if (sellout) D_RESTO.sellout = sellout;
    if (actividades) D_RESTO.actividades = actividades;
    if (participacion) D_RESTO.participacion = participacion;
    if (insights) D_RESTO.insights = insights;

    // ---- módulos que quedan en pausa si ni el maestro ni una fuente suelta los trajeron ----
    D_RESTO.rotacion = rotacion || marcadorPausa(archivos, [/rotaci/, /retiro/, /personal/],
      'rotacion.xlsx', 'mes, retiros, activos y % de rotación por cargo / canal / regional');
    D_RESTO.formacion = formacion || marcadorPausa(archivos, [/formaci/, /capacitaci/, /entrenamiento/],
      'formacion.xlsx', 'formación inicial por regional y mes; cursos (convocados, participantes, satisfacción, transferencia); horas por regional');
    D_RESTO.inventario = inventario || marcadorPausa(archivos, [/inventario/, /\bcedi\b/, /existencia/],
      'inventario-cedi.xlsx', 'CEDI, bodega, código, artículo, presentación, estado, cantidad, disponible, por alistar, valor total');
    if (!D_VENTAS) D_VENTAS = marcadorPausa(archivos, [/venta.*detallada/, /detalle.*venta/, /ventas?-?detalladas/],
      'ventas-detalladas.xlsx', 'MES, REGIONAL, TIENDA, MARCA, UNIDADES, LITROS, DINERO (y si puedes: NEGOCIO, LINEA, SUBNEGOCIO, SKU)');
    if (!D_COMPET) D_COMPET = prev.D_COMPET || null;

    var ahora = new Date();
    var stamp = ahora.getFullYear() + '-' + p2(ahora.getMonth() + 1) + '-' + p2(ahora.getDate()) +
      ' ' + p2(ahora.getHours()) + ':' + p2(ahora.getMinutes());

    var datos = { generado: stamp, D_TIEMPOS: D_TIEMPOS, D_VENTAS: D_VENTAS, D_COMPET: D_COMPET, D_RESTO: D_RESTO };

    var enPausa = [];
    if (D_VENTAS && D_VENTAS.pausa) enPausa.push('ventas');
    ['rotacion', 'formacion', 'inventario'].forEach(function (k) {
      if (D_RESTO[k] && D_RESTO[k].pausa) enPausa.push(k);
    });

    var resumen = {
      maestro: maestro ? maestro.archivo : null,
      tiempos: D_TIEMPOS ? (D_TIEMPOS.meta ? D_TIEMPOS.meta.registros : 0) : 0,
      efectividad: efect ? efect.rows.length : 0,
      tomas: tomas ? tomas.rows.length : 0,
      sellout: sellout ? sellout.rows.length : 0,
      actividades: actividades ? (actividades.gen.length + '/' + actividades.pdv.length) : '0/0',
      participacion: participacion ? participacion.rows.length : 0,
      ventas: (D_VENTAS && !D_VENTAS.pausa && D_VENTAS.meta) ? D_VENTAS.meta.filas : 0,
      rotacion: (D_RESTO.rotacion && !D_RESTO.rotacion.pausa) ? D_RESTO.rotacion.general.length : 0,
      formacion: (D_RESTO.formacion && !D_RESTO.formacion.pausa) ? D_RESTO.formacion.cursos.length : 0,
      inventario: (D_RESTO.inventario && !D_RESTO.inventario.pausa) ? D_RESTO.inventario.rows.length : 0,
      competencia: (D_COMPET && D_COMPET.det) ? D_COMPET.det.length : 0,
      pausa: enPausa,
      archivos: Object.keys(wbs)
    };

    return { datos: datos, avisos: avisos, resumen: resumen };
  }

  root.PINTUCO_ETL = { compilar: compilar, _util: { norm: norm, num: num, fechaISO: fechaISO, segDesdeHora: segDesdeHora } };

  if (typeof module !== 'undefined' && module.exports) module.exports = root.PINTUCO_ETL;

})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
