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

    // 1+2 visitas -> tiempos y efectividad
    var D_TIEMPOS = prev.D_TIEMPOS || null;
    var efect = prevR.efectividad || null;
    var nomVis = elegirArchivo(archivos, [/informe.*visita/, /visita.*gerencial/, /gerencial.*visita/]);
    if (nomVis && wbs[nomVis]) {
      var v = construirVisitas(wbs[nomVis], avisos);
      if (v) {
        D_TIEMPOS = v.D_TIEMPOS;
        efect = construirEfectividad(v.recs, avisos);
      }
    } else {
      avisos.push('[visitas] no encontré "informe-gerencial-visitas ####.xlsx" — se conserva lo anterior');
    }

    // 3 tomas
    var tomas = construirTomas(archivos, wbs, avisos) || prevR.tomas || null;
    // 4 sell out
    var sellout = construirSellout(archivos, wbs, avisos) || prevR.sellout || null;
    // 5 actividades
    var actividades = construirActividades(archivos, wbs, avisos) || prevR.actividades || null;
    // 6 participación en lineal (módulo nuevo)
    var participacion = construirParticipacion(archivos, wbs, avisos) || prevR.participacion || null;

    var D_RESTO = {};
    Object.keys(prevR).forEach(function (k) { D_RESTO[k] = prevR[k]; }); // arranca del anterior
    if (efect) D_RESTO.efectividad = efect;
    if (tomas) D_RESTO.tomas = tomas;
    if (sellout) D_RESTO.sellout = sellout;
    if (actividades) D_RESTO.actividades = actividades;
    if (participacion) D_RESTO.participacion = participacion;

    // ---- módulos en pausa (sin fuente todavía) ----
    D_RESTO.rotacion = marcadorPausa(archivos, [/rotaci/, /retiro/, /personal/],
      'rotacion.xlsx',
      'mes, retiros, activos y % de rotación por cargo / canal / regional');
    D_RESTO.formacion = marcadorPausa(archivos, [/formaci/, /capacitaci/, /entrenamiento/],
      'formacion.xlsx',
      'formación inicial por regional y mes; cursos (convocados, participantes, satisfacción, transferencia); horas por regional');
    D_RESTO.inventario = marcadorPausa(archivos, [/inventario/, /\bcedi\b/, /existencia/],
      'inventario-cedi.xlsx',
      'CEDI, bodega, código, artículo, presentación, estado, cantidad, disponible, por alistar, valor total');

    var ahora = new Date();
    var stamp = ahora.getFullYear() + '-' + p2(ahora.getMonth() + 1) + '-' + p2(ahora.getDate()) +
      ' ' + p2(ahora.getHours()) + ':' + p2(ahora.getMinutes());

    // D_VENTAS: se arma si está el archivo; si no, queda en pausa
    var D_VENTAS = construirVentas(archivos, wbs, avisos)
      || marcadorPausa(archivos, [/venta.*detallada/, /detalle.*venta/, /ventas?-?detalladas/],
        'ventas-detalladas.xlsx',
        'MES, REGIONAL, TIENDA, MARCA, UNIDADES, LITROS, DINERO (y si puedes: NEGOCIO, LINEA, SUBNEGOCIO, SKU)');
    var D_COMPET = prev.D_COMPET || null; // se conserva (homologación manual); Participación en lineal es el módulo vivo

    var datos = {
      generado: stamp,
      D_TIEMPOS: D_TIEMPOS,
      D_VENTAS: D_VENTAS,
      D_COMPET: D_COMPET,
      D_RESTO: D_RESTO
    };

    var enPausa = [];
    if (D_VENTAS && D_VENTAS.pausa) enPausa.push('ventas');
    ['rotacion', 'formacion', 'inventario'].forEach(function (k) {
      if (D_RESTO[k] && D_RESTO[k].pausa) enPausa.push(k);
    });

    var resumen = {
      tiempos: D_TIEMPOS ? (D_TIEMPOS.meta ? D_TIEMPOS.meta.registros : 0) : 0,
      efectividad: efect ? efect.rows.length : 0,
      tomas: tomas ? tomas.rows.length : 0,
      sellout: sellout ? sellout.rows.length : 0,
      actividades: actividades ? (actividades.gen.length + '/' + actividades.pdv.length) : '0/0',
      participacion: participacion ? participacion.rows.length : 0,
      ventas: (D_VENTAS && !D_VENTAS.pausa && D_VENTAS.meta) ? D_VENTAS.meta.filas : 0,
      pausa: enPausa,
      archivos: Object.keys(wbs)
    };

    return { datos: datos, avisos: avisos, resumen: resumen };
  }

  root.PINTUCO_ETL = { compilar: compilar, _util: { norm: norm, num: num, fechaISO: fechaISO, segDesdeHora: segDesdeHora } };

  if (typeof module !== 'undefined' && module.exports) module.exports = root.PINTUCO_ETL;

})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
