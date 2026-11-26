// =====================================================================
//  Remisión de Tareas · "sincronización" SIN NUBE
//  · Las listas (empresas, parcelas, máquinas...) llegan en un archivo que
//    prepara la oficina (Access > "Enviar listas al móvil (archivo)").
//  · Los partes salen en un archivo que se comparte por WhatsApp, correo,
//    Bluetooth... y la oficina importa (Access > "Importar partes desde archivo").
// =====================================================================
(function (g) {
  const FORMATO_LISTAS = 'remision-tareas-listas', FORMATO_PARTES = 'remision-tareas-partes';

  function catDeArchivo(d) {
    const ord = (a, k) => (a || []).slice().sort((x, y) => String(x[k] || '').localeCompare(String(y[k] || ''), 'es'));
    return {
      empresas: ord(d.cat_empresas, 'nombre'),
      parcelas: (d.cat_parcelas || []).slice(),
      faenas: ord(d.cat_faenas, 'nombre'),
      clases: ord(d.cat_clases, 'nombre'),
      unidades: (d.cat_unidades || []).slice().sort((a, b) => a.id - b.id),
      cereales: ord(d.cat_cereales, 'nombre'),
      productos: ord(d.cat_productos, 'nombre'),
      tareas: ord(d.cat_tareas, 'nombre'),
      maquinas: ord(d.cat_maquinas, 'nombre'),
      tipotrabajo: ord(d.cat_tipotrabajo, 'nombre'),
      inventario: ord(d.cat_inventario, 'denominacion'),
      personal: ord(d.personal, 'nombre'),
      entradas: [], parajes: [],
      oficina: d.oficina || '',
      fecha: d.generado || new Date().toISOString()
    };
  }

  // Polilínea codificada (formato de Google, 5 decimales) -> [[lon, lat], ...]
  function decodificar(s) {
    const pts = []; let i = 0, lat = 0, lon = 0;
    while (i < s.length) {
      for (let k = 0; k < 2; k++) {
        let r = 0, sh = 0, b;
        do { b = s.charCodeAt(i++) - 63; r |= (b & 0x1f) << sh; sh += 5; } while (b >= 0x20);
        const d = (r & 1) ? ~(r >> 1) : (r >> 1);
        if (k === 0) lat += d; else lon += d;
      }
      pts.push([lon / 1e5, lat / 1e5]);
    }
    return pts;
  }

  g.Sync = {
    configurado: () => true,
    archivo: true,

    // Lee el archivo de listas que manda la oficina. Devuelve el catálogo.
    async cargarListas(texto) {
      let d;
      try { d = JSON.parse(String(texto).replace(/^\uFEFF/, '')); }
      catch (e) { throw new Error('El archivo no se puede leer. ¿Es el archivo de listas que manda la oficina?'); }
      if (!d || d.formato !== FORMATO_LISTAS) throw new Error('Ese archivo no es de listas de la oficina (Remisión de Tareas).');
      const anterior = await g.BD.leer('catalogos');
      const cat = catDeArchivo(d);
      if (anterior && anterior.entradas) cat.entradas = anterior.entradas;     // entradas guardadas en este móvil
      if (anterior && anterior.parajes && anterior.parajes.length) cat.parajes = anterior.parajes;
      await g.BD.guardar('catalogos', cat);
      await g.BD.guardar('reparacionesNube', { fecha: cat.fecha, lista: d.reparaciones || [] });
      if (d.correo_oficina) await g.BD.guardar('correoOficina', String(d.correo_oficina).trim());
      return cat;
    },

    async descargarCatalogos() { return g.BD.leer('catalogos'); },

    // Entrada: el trabajador elige su nombre de la lista de la oficina
    async entrar(persona) {
      const s = { archivo: true, nombre: persona.nombre, clave: persona.clave || '', email: '' };
      await g.BD.guardar('sesion', s);
      return s;
    },

    // Sin nube: no hay copia de partes en curso ni envío automático
    async guardarBorrador() {}, async borrarBorrador() {}, async recuperarBorradores() { return 0; },
    async enviarPendientes() { return { enviados: 0, errores: 0, sinConexion: false, reintentar: 0 }; },
    // Parajes: vienen con la app (parajes.json, sacado de la nube de la oficina principal)
    async cargarParajes() {
      let d = await g.BD.leer('parajesArchivo');
      if (!d) {
        const r = await fetch('parajes.json', { cache: 'no-cache' });
        if (!r.ok) throw new Error('No se pudieron cargar los parajes');
        d = await r.json();
        await g.BD.guardar('parajesArchivo', d);
      }
      return d.parajes.map(p => ({ id: p.id, ine: p.ine, municipio: p.municipio, nombre: p.nombre, n_parcelas: p.n_parcelas,
        superficie_ha: p.superficie_ha, centro_lat: p.centro_lat, centro_lon: p.centro_lon }));
    },
    async zonaParaje(id) {
      const d = await g.BD.leer('parajesArchivo');
      const p = d && d.parajes.find(x => x.id === id);
      if (!p || !p.z || !p.z.length) return null;
      return { type: 'MultiPolygon', coordinates: p.z.map(poly => poly.map(decodificar)) };
    },
    async guardarEntrada(e) { return e; },

    async reparacionesPendientes() {
      const s = await g.BD.leer('sesion');
      if (!s || s.demo) return (g.CATALOGO_DEMO && g.CATALOGO_DEMO.reparaciones) || [];
      const r = await g.BD.leer('reparacionesNube');
      const fin = ((await g.BD.leer('finalizaciones')) || []).map(f => f.id);
      return ((r && r.lista) || []).filter(x => !fin.includes(x.id));
    },

    // ✅ Reparada: se apunta y viaja a la oficina en el próximo archivo de partes
    async finalizarReparacion(id, idAccess) {
      const s = await g.BD.leer('sesion');
      const lista = (await g.BD.leer('finalizaciones')) || [];
      if (!lista.some(f => f.id === id)) lista.push({ id, id_access: idAccess || null, fecha: new Date().toISOString(),
        trabajador: s ? s.nombre : '', enviada: false });
      await g.BD.guardar('finalizaciones', lista);
      return { estado: 'finalizada' };
    },

    // ¿Qué hay para mandar a la oficina?
    async pendientesDeEnvio() {
      const partes = (await g.BD.partes()).filter(p => p.estado === 'pendiente' || p.estado === 'error');
      const fins = ((await g.BD.leer('finalizaciones')) || []).filter(f => !f.enviada);
      return { partes, fins, total: partes.length + fins.length };
    },

    // Prepara el archivo con todo lo pendiente
    async prepararArchivo() {
      const s = await g.BD.leer('sesion');
      const { partes, fins } = await g.Sync.pendientesDeEnvio();
      const ahora = new Date();
      const d = {
        formato: FORMATO_PARTES, version: 1, generado: ahora.toISOString(),
        trabajador: s ? s.nombre : '', clave: s ? s.clave || '' : '',
        partes: partes.filter(p => p.tipo !== 'horas' && p.tipo !== 'reparacion').map(aEnvio),
        horas: partes.filter(p => p.tipo === 'horas').map(aEnvioHoras),
        reparaciones: partes.filter(p => p.tipo === 'reparacion').map(aEnvioRep),
        finalizaciones: fins.map(f => ({ id: f.id, id_access: f.id_access, fecha: f.fecha, trabajador: f.trabajador }))
      };
      const dos = n => String(n).padStart(2, '0');
      const nombre = 'Partes ' + (d.trabajador || 'movil').replace(/[^\wÁÉÍÓÚÑáéíóúñ ]+/g, '').trim() + ' ' +
        ahora.getFullYear() + '-' + dos(ahora.getMonth() + 1) + '-' + dos(ahora.getDate()) + ' ' + dos(ahora.getHours()) + dos(ahora.getMinutes()) + '.txt';
      const texto = JSON.stringify(d, null, 1);
      return { nombre, texto, ids: partes.map(p => p.id), finIds: fins.map(f => f.id), n: partes.length + fins.length };
    },

    // El trabajador confirma que el archivo ha salido: se marcan como enviados
    async marcarEnviado(env) {
      for (const id of env.ids) {
        const p = await g.BD.parte(id);
        if (p && (p.estado === 'pendiente' || p.estado === 'error')) {
          p.estado = 'enviado'; p.enviado = new Date().toISOString(); delete p.error;
          await g.BD.guardarParte(p);
        }
      }
      const fins = (await g.BD.leer('finalizaciones')) || [];
      fins.forEach(f => { if (env.finIds.includes(f.id)) f.enviada = true; });
      await g.BD.guardar('finalizaciones', fins);
      await g.BD.guardar('ultimoEnvio', { fecha: new Date().toISOString(), enviados: env.n, nombre: env.nombre });
      await g.BD.guardar('ultimoArchivo', { nombre: env.nombre, texto: env.texto, fecha: new Date().toISOString() });
    }
  };

  const hm = iso => { const d = new Date(iso); return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };

  function aEnvioHoras(h) {
    return {
      id: h.id, fecha: h.fecha, hora_entrada: h.hora_entrada || '', hora_salida: h.hora_salida || '',
      descanso_min: h.descanso_min === '' || h.descanso_min == null ? '' : h.descanso_min,
      horas: h.horas === '' || h.horas == null ? '' : h.horas,
      vacaciones: !!h.vacaciones, baja_medica: !!h.baja_medica, observaciones: h.observaciones || '',
      lat: h.gps ? h.gps.lat : '', lon: h.gps ? h.gps.lon : '', creado_movil: h.creado_movil
    };
  }

  function aEnvioRep(r) {
    return {
      id: r.id, fecha: r.fecha, denominacion: r.denominacion || '', matricula: r.matricula || '', clase: r.clase || '',
      reparacion: r.reparacion || '', observaciones: r.observaciones || '', finalizado: !!r.finalizado,
      lat: r.gps ? r.gps.lat : '', lon: r.gps ? r.gps.lon : '', creado_movil: r.creado_movil
    };
  }

  // Lo que viaja a la nube (sin campos internos de la app)
  function aEnvio(p) {
    return {
      id: p.id, tipo: p.tipo || 'maquina', fecha: p.fecha, hora_inicio: p.hora_inicio || '', hora_fin: p.hora_fin || '',
      id_empresa: p.id_empresa, empresa: p.empresa, clase: p.clase || '', faena: p.faena || '',
      maquina: p.maquina || '', kilos: p.kilos === '' || p.kilos == null ? '' : p.kilos,
      problema: p.problema || '',
      productos: (p.productos || []).map((x, i) => ({ orden: i + 1, id_producto: x.id_producto == null ? '' : x.id_producto,
        producto: x.producto, registro: x.registro || '', dosis_ha: x.dosis_ha === '' || x.dosis_ha == null ? '' : x.dosis_ha })),
      observaciones: p.observaciones || '',
      lat: p.gps ? p.gps.lat : '', lon: p.gps ? p.gps.lon : '',
      gps_precision_m: p.gps ? p.gps.precision : '', gps_fecha: p.gps ? p.gps.fecha : '',
      creado_movil: p.creado_movil,
      lineas: (p.lineas || []).map((l, i) => ({
        id: l.id, orden: i + 1, id_parcela: l.id_parcela == null ? '' : l.id_parcela,
        poligono: l.poligono || '', parcela: l.parcela || '', paraje: l.paraje || '',
        tipo_cultivo: l.tipo_cultivo || '', superficie: l.superficie == null ? '' : l.superficie,
        sigpac: l.sigpac || '', cantidad: l.cantidad === '' || l.cantidad == null ? '' : l.cantidad,
        unidad: l.unidad || '', tipo_cereal: l.tipo_cereal || p.cereal || '', cereal: l.cereal || p.riego || '',
        kilos: l.kilos === '' || l.kilos == null ? '' : l.kilos,
        observaciones: (l.observaciones || '') + (l.catastro ? ' · Catastro RC ' + l.catastro.rc + (l.catastro.municipio ? ' · ' + l.catastro.municipio : '') : '') +
          (l.aviso_catastro ? ' · AVISO GPS: ' + l.aviso_catastro : ''),
        lat: l.lat == null ? '' : l.lat, lon: l.lon == null ? '' : l.lon,
        gps_precision_m: l.precision == null ? '' : l.precision,
        inicio: (l.tramos && l.tramos.length) ? l.tramos[0].inicio : '',
        fin: (l.tramos && l.tramos.length) ? (l.tramos[l.tramos.length - 1].fin || '') : '',
        horas: (l.horas_mano != null && l.horas_mano !== '') ? +l.horas_mano
          : (l.tramos && l.tramos.length) ? Math.round(l.tramos.reduce((s, t) => s + (new Date(t.fin || Date.now()) - new Date(t.inicio)) / 36e5, 0) * 100) / 100 : '',
        tramos: (l.tramos || []).map(t => hm(t.inicio) + '-' + (t.fin ? hm(t.fin) : '')).join(', ') +
          ((l.horas_mano != null && l.horas_mano !== '') ? ' [horas corregidas a mano]' : '')
      }))
    };
  }
})(self);
