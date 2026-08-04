(function(){
  const W = window;

  function citas() {
    return (typeof allData !== 'undefined' && allData && Array.isArray(allData.citas)) ? allData.citas : [];
  }

  function eventos() {
    return (typeof allData !== 'undefined' && allData && Array.isArray(allData.eventos)) ? allData.eventos : [];
  }

  function pacientes() {
    return (typeof allData !== 'undefined' && allData && Array.isArray(allData.pacientes)) ? allData.pacientes : [];
  }

  function profesionales() {
    if (typeof allData === 'undefined' || !allData) return [];
    return allData.profesionales || allData.fisioterapeutas || allData.team || [];
  }

  function fDate(value) {
    if (typeof normDate === 'function') return normDate(value || '');
    if (!value) return '';
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return String(value).slice(0,10);
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }

  function peso(value) {
    if (typeof fmtPeso === 'function') return fmtPeso(value || 0);
    return '$' + Math.round(Number(value || 0)).toLocaleString('es-CO');
  }

  function money(value) {
    if (typeof parsePrecio === 'function') return parsePrecio(value);
    if (typeof value === 'number') return value;
    return Number(String(value || '').replace(/[^\d.-]/g,'')) || 0;
  }

  function isServiceRecord(servicio) {
    return typeof esRegistroServ === 'function' ? esRegistroServ(servicio) : false;
  }

  function monthKey(d = new Date()) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  }

  function periodo(d = new Date()) {
    const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
    return `${meses[d.getMonth()]} de ${d.getFullYear()}`;
  }

  function ocupacion(total, date) {
    const y = date.getFullYear();
    const m = date.getMonth();
    const days = new Date(y, m + 1, 0).getDate();
    let capacidad = 0;
    for (let d = 1; d <= days; d++) {
      const dow = new Date(y, m, d).getDay();
      if (dow === 0) continue;
      if (dow === 1) capacidad += 8;
      else if (dow === 6) capacidad += 2;
      else capacidad += 9;
    }
    return capacidad ? Math.round((total / capacidad) * 100) + '%' : 'Sin capacidad calculada';
  }

  function reactivacion(citasAll) {
    const last = {};
    citasAll.forEach(c => {
      if (!c.nombre || String(c.estado || '').toLowerCase().includes('cancel')) return;
      const key = String(c.nombre).trim().toLowerCase();
      const fecha = fDate(c.fecha || '');
      if (!last[key] || fecha > last[key].fecha) {
        last[key] = { nombre:c.nombre, telefono:c.telefono || c.phone || '', fecha, hora:c.hora || '', servicio:c.servicio || '', estado:c.estado || '' };
      }
    });
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 42);
    const min = fDate(cutoff);
    return Object.values(last).filter(p => p.fecha && p.fecha < min).sort((a,b) => a.fecha.localeCompare(b.fecha));
  }

  function candidatosPaquete(citasAll) {
    const map = {};
    citasAll.forEach(c => {
      if (!c.nombre || String(c.estado || '').toLowerCase().includes('cancel')) return;
      const key = String(c.nombre).trim().toLowerCase();
      if (!map[key]) map[key] = { nombre:c.nombre, telefono:c.telefono || c.phone || '', fecha:fDate(c.fecha || ''), hora:c.hora || '', servicio:c.servicio || '', estado:c.estado || '', total:0, paquete:false };
      map[key].total++;
      if (String(c.servicio || '').toLowerCase().includes('paquete')) map[key].paquete = true;
      const fecha = fDate(c.fecha || '');
      if (fecha > map[key].fecha) {
        map[key].fecha = fecha;
        map[key].hora = c.hora || '';
        map[key].servicio = c.servicio || '';
        map[key].estado = c.estado || '';
      }
    });
    return Object.values(map).filter(p => p.total >= 2 && !p.paquete).sort((a,b) => b.total - a.total);
  }

  function gestionData() {
    const now = new Date();
    const mk = monthKey(now);
    const citasAll = citas();
    const citasMesAll = citasAll.filter(c => fDate(c.fecha || '').startsWith(mk) && !isServiceRecord(c.servicio));
    const citasMes = citasMesAll.filter(c => !String(c.estado || '').toLowerCase().includes('cancel'));
    const eventosMes = eventos().filter(e => fDate(e.fecha || '').startsWith(mk));
    const sesionesAtendidas = citasMesAll.filter(c => String(c.estado || '').toLowerCase().includes('atendida')).length;
    const cancelaciones = citasMesAll.filter(c => String(c.estado || '').toLowerCase().includes('cancel')).length;
    const noAsistencias = citasMesAll.filter(c => String(c.estado || '').toLowerCase().includes('no asist')).length;
    const ventasGeneradas = citasMes.reduce((s,c) => s + money(c.precio), 0) + eventosMes.reduce((s,e) => s + money(e.cobro), 0);
    const ingresosCobrados = typeof calcCobradoMes === 'function' ? calcCobradoMes() : ventasGeneradas;
    const pagosPendientesLista = citasMes.filter(c => {
      const estado = String(c.estado || '').toLowerCase();
      return estado.includes('pendiente de pago') || estado.includes('pago por verificar') || estado.includes('rechazado');
    });
    const pendienteCobrar = pagosPendientesLista.reduce((s,c) => s + money(c.precio), 0);
    const egresosMes = (typeof getEgresos === 'function' ? getEgresos() : [])
      .filter(e => fDate(e.fecha || '').startsWith(mk))
      .reduce((s,e) => s + money(e.monto), 0);
    const cfg = typeof getKPIConfig === 'function' ? getKPIConfig() : {};
    const manual = typeof getKPIManual === 'function' ? getKPIManual() : {};
    const metaMensual = cfg.meta_ventas_mes || W.META_VENTAS_MES || 0;
    const cumplimiento = metaMensual > 0 ? Math.round((ingresosCobrados / metaMensual) * 100) : 0;
    const pacienteMes = {};
    citasMes.forEach(c => { if (c.nombre) pacienteMes[String(c.nombre).trim().toLowerCase()] = c.nombre; });
    let personasNuevas = 0;
    let personasRecurrentes = 0;
    Object.keys(pacienteMes).forEach(key => {
      const antes = citasAll.some(c => String(c.nombre || '').trim().toLowerCase() === key && fDate(c.fecha || '') < mk + '-01' && !String(c.estado || '').toLowerCase().includes('cancel'));
      if (antes) personasRecurrentes++; else personasNuevas++;
    });
    const servicios = {};
    const horarios = {};
    citasMes.forEach(c => {
      const serv = c.servicio || 'Sin servicio';
      servicios[serv] = (servicios[serv] || 0) + 1;
      const h = String(c.hora || '').slice(0,2);
      if (h) horarios[h + ':00'] = (horarios[h + ':00'] || 0) + 1;
    });
    const sortMap = obj => Object.entries(obj).sort((a,b) => b[1] - a[1]);
    const serviciosArr = sortMap(servicios);
    const horariosArr = sortMap(horarios);
    const pros = profesionales();
    return {
      periodo: periodo(now),
      metaMensual,
      ingresosCobrados,
      ventasGeneradas,
      pendienteCobrar,
      egresosMes,
      ganancia: ingresosCobrados - egresosMes,
      cumplimiento,
      faltante: Math.max(0, metaMensual - ingresosCobrados),
      citasProgramadas: citasMes.length,
      sesionesAtendidas,
      personasNuevas,
      personasRecurrentes,
      paquetesVendidos: citasMes.filter(c => String(c.servicio || '').toLowerCase().includes('paquete')).length,
      ticketPromedio: citasMes.length ? Math.round(ventasGeneradas / citasMes.length) : 0,
      ocupacion: ocupacion(citasMes.length + eventosMes.length, now),
      cancelaciones,
      noAsistencias,
      leadsRecibidos: typeof getLeadsMes === 'function' ? getLeadsMes() : (manual.leads || 0),
      leadsConvertidos: manual.convertidos || citasMes.length,
      serviciosMasVendidos: serviciosArr.slice(0,5).map(([s,n]) => `${s}: ${n}`).join('\n') || 'Sin datos',
      serviciosMenosVendidos: serviciosArr.slice(-5).map(([s,n]) => `${s}: ${n}`).join('\n') || 'Sin datos',
      horariosMayorOcupacion: horariosArr.slice(0,5).map(([h,n]) => `${h}: ${n} cita(s)`).join('\n') || 'Sin datos',
      horariosMenorOcupacion: horariosArr.slice(-5).map(([h,n]) => `${h}: ${n} cita(s)`).join('\n') || 'Sin datos',
      disponibilidadPros: pros.length ? pros.map(p => `${p.nombre || p.Nombre || 'Profesional'}: ${p.disponibilidad || p.Disponibilidad || 'Sin disponibilidad registrada'}`).join('\n') : 'Sin fisioterapeutas registrados',
      pagosPendientesLista,
      reactivar: reactivacion(citasAll).slice(0,80),
      candidatosPaquete: candidatosPaquete(citasAll).slice(0,80),
      estrategiasEjecutadas: localStorage.getItem('gestion_estrategias_mes') || 'Sin registrar',
      resultadosObtenidos: localStorage.getItem('gestion_resultados_mes') || 'Sin registrar',
      observaciones: localStorage.getItem('gestion_observaciones_mes') || 'Sin registrar'
    };
  }

  function diagnostico(d) {
    const ok = [];
    const att = [];
    if (d.cumplimiento >= 80) ok.push(`Cumplimiento de meta en ${d.cumplimiento}%.`);
    else att.push(`Faltan ${peso(d.faltante)} para llegar a la meta mensual.`);
    if (d.pendienteCobrar > 0) att.push(`Hay ${peso(d.pendienteCobrar)} pendiente por cobrar.`);
    if (d.cancelaciones > 0) att.push(`Se registran ${d.cancelaciones} cancelación(es) este mes.`);
    if (d.reactivar.length) att.push(`Hay ${d.reactivar.length} persona(s) para reactivar.`);
    if (d.candidatosPaquete.length) ok.push(`Hay ${d.candidatosPaquete.length} candidato(s) para ofrecer paquetes.`);
    if (!ok.length) ok.push('Hay información suficiente para tomar decisiones de gestión.');
    if (!att.length) att.push('No se detectan alertas administrativas fuertes con los datos actuales.');
    return {ok, att};
  }

  function acciones(d) {
    return [
      `Contactar ${d.reactivar.length} persona(s) para reactivación.`,
      `Revisar y cerrar ${d.pagosPendientesLista.length} pago(s) pendiente(s).`,
      `Ofrecer paquetes a ${d.candidatosPaquete.length} candidato(s) con sesiones sueltas.`,
      'Revisar horarios de baja ocupación y mover campañas hacia esas franjas.',
      'Revisar el servicio más vendido y crear una oferta complementaria.'
    ];
  }

  function asesorText(d) {
    return [
      'ANÁLISIS MENSUAL DE CUIDÁNDOTE FISIOTERAPIA',
      '',
      `Periodo: ${d.periodo}`,
      `Meta mensual: ${peso(d.metaMensual)}`,
      '',
      'RESUMEN FINANCIERO',
      `* Ingresos cobrados: ${peso(d.ingresosCobrados)}`,
      `* Ventas generadas: ${peso(d.ventasGeneradas)}`,
      `* Pendiente por cobrar: ${peso(d.pendienteCobrar)}`,
      `* Gastos: ${peso(d.egresosMes)}`,
      `* Ganancia estimada: ${peso(d.ganancia)}`,
      `* Cumplimiento de la meta: ${d.cumplimiento}%`,
      '',
      'OPERACIÓN',
      `* Citas programadas: ${d.citasProgramadas}`,
      `* Sesiones atendidas: ${d.sesionesAtendidas}`,
      `* Cancelaciones: ${d.cancelaciones}`,
      `* No asistencias: ${d.noAsistencias}`,
      `* Ocupación total: ${d.ocupacion}`,
      '',
      'CLIENTES Y VENTAS',
      `* Personas nuevas: ${d.personasNuevas}`,
      `* Personas recurrentes: ${d.personasRecurrentes}`,
      `* Leads recibidos: ${d.leadsRecibidos}`,
      `* Leads convertidos: ${d.leadsConvertidos}`,
      `* Paquetes vendidos: ${d.paquetesVendidos}`,
      `* Ticket promedio: ${peso(d.ticketPromedio)}`,
      '',
      'CAPACIDAD DEL EQUIPO',
      `* Disponibilidad por profesional:\n${d.disponibilidadPros}`,
      `* Horarios con baja ocupación:\n${d.horariosMenorOcupacion}`,
      '* Citas que podrían delegarse: revisar citas próximas de servicios presenciales o de descarga muscular.',
      '',
      'OPORTUNIDADES',
      `* Leads sin seguimiento: revisar contador y mensajes pendientes.`,
      `* Personas para reactivar: ${d.reactivar.length}`,
      `* Candidatos para paquetes: ${d.candidatosPaquete.length}`,
      '* Paquetes próximos a terminar: revisar módulo de paquetes.',
      `* Pagos pendientes: ${d.pagosPendientesLista.length}`,
      '',
      'SERVICIOS',
      `* Servicios más vendidos:\n${d.serviciosMasVendidos}`,
      `* Servicios menos vendidos:\n${d.serviciosMenosVendidos}`,
      '* Servicios más rentables: revisar estructura de costos.',
      '* Servicios con menor rentabilidad: revisar estructura de costos.',
      '',
      'ACCIONES DEL MES',
      `* Estrategias ejecutadas: ${d.estrategiasEjecutadas}`,
      `* Resultado: ${d.resultadosObtenidos}`,
      '* Ingreso generado: calcular según campañas registradas.',
      '',
      'OBSERVACIONES',
      d.observaciones,
      '',
      'Actúa como asesor estratégico de Cuidándote Fisioterapia. Analiza esta información y entrégame:',
      '',
      '1. Diagnóstico del mes.',
      '2. Principales problemas.',
      '3. Oportunidades de ingresos.',
      '4. Cinco acciones prioritarias.',
      '5. Personas o segmentos que debemos contactar.',
      '6. Estrategias para llegar a la meta.',
      '7. Actividades que debe realizar administración.',
      '8. Actividades que se pueden delegar a los fisioterapeutas.',
      '9. Riesgos.',
      '10. Próximo paso inmediato.',
      '',
      'No inventes datos. Basa todas las recomendaciones únicamente en la información entregada.'
    ].join('\n');
  }

  W.copyGestionTexto = function(kind) {
    const d = gestionData();
    const diag = diagnostico(d);
    const act = acciones(d);
    const base = [
      `Periodo: ${d.periodo}`,
      '',
      'RESUMEN FINANCIERO',
      `* Ingresos cobrados: ${peso(d.ingresosCobrados)}`,
      `* Ventas generadas: ${peso(d.ventasGeneradas)}`,
      `* Pendiente por cobrar: ${peso(d.pendienteCobrar)}`,
      `* Gastos: ${peso(d.egresosMes)}`,
      `* Ganancia estimada: ${peso(d.ganancia)}`,
      `* Meta mensual: ${peso(d.metaMensual)}`,
      `* Cumplimiento: ${d.cumplimiento}%`,
      `* Dinero faltante: ${peso(d.faltante)}`
    ];
    let text = '';
    if (kind === 'ejecutivo') {
      text = ['RESUMEN EJECUTIVO — CUIDÁNDOTE FISIOTERAPIA', ...base, '', 'PUNTOS CLAVE', ...diag.ok.map(x => `* ${x}`), ...diag.att.map(x => `* ${x}`), '', 'PRÓXIMAS ACCIONES', ...act.map((x,i)=>`${i+1}. ${x}`)].join('\n');
    } else if (kind === 'indicadores') {
      text = ['INDICADORES DE GESTIÓN — CUIDÁNDOTE FISIOTERAPIA', ...base, '', 'OPERACIÓN', `* Citas programadas: ${d.citasProgramadas}`, `* Sesiones atendidas: ${d.sesionesAtendidas}`, `* Ocupación total: ${d.ocupacion}`, `* Cancelaciones: ${d.cancelaciones}`, `* No asistencias: ${d.noAsistencias}`, '', 'CLIENTES Y VENTAS', `* Personas nuevas: ${d.personasNuevas}`, `* Personas recurrentes: ${d.personasRecurrentes}`, `* Leads recibidos: ${d.leadsRecibidos}`, `* Leads convertidos: ${d.leadsConvertidos}`, `* Paquetes vendidos: ${d.paquetesVendidos}`, `* Ticket promedio: ${peso(d.ticketPromedio)}`].join('\n');
    } else if (kind === 'diagnostico') {
      text = ['DIAGNÓSTICO DE GESTIÓN — CUIDÁNDOTE FISIOTERAPIA', `Periodo: ${d.periodo}`, '', 'Lo que está funcionando:', ...diag.ok.map(x => `* ${x}`), '', 'Lo que necesita atención:', ...diag.att.map(x => `* ${x}`)].join('\n');
    } else if (kind === 'estrategias') {
      text = ['ESTRATEGIAS RECOMENDADAS — CUIDÁNDOTE FISIOTERAPIA', `Periodo: ${d.periodo}`, '', ...act.map((x,i)=>`${i+1}. ${x}`), '', 'Estrategias ejecutadas:', d.estrategiasEjecutadas, '', 'Resultados obtenidos:', d.resultadosObtenidos].join('\n');
    } else if (kind === 'plan') {
      text = ['PLAN DE ACCIÓN — CUIDÁNDOTE FISIOTERAPIA', `Periodo: ${d.periodo}`, '', ...act.map((x,i)=>`${i+1}. ${x}`), '', 'Prioridad sugerida:', '1. Cobros pendientes.', '2. Reactivación.', '3. Ofertas de paquetes.', '4. Optimización de horarios.'].join('\n');
    } else if (kind === 'asesor') {
      text = asesorText(d);
    } else {
      text = ['DIRECCIÓN Y CRECIMIENTO', `Periodo: ${d.periodo}`, '', '1. RESUMEN FINANCIERO', `Ingresos cobrados: ${peso(d.ingresosCobrados)}`, `Ventas generadas: ${peso(d.ventasGeneradas)}`, `Pendiente por cobrar: ${peso(d.pendienteCobrar)}`, `Meta mensual: ${peso(d.metaMensual)}`, `Cumplimiento: ${d.cumplimiento}%`, `Dinero faltante: ${peso(d.faltante)}`, '', '2. DIAGNÓSTICO', 'Lo que está funcionando:', ...diag.ok.map(x => `* ${x}`), '', 'Lo que necesita atención:', ...diag.att.map(x => `* ${x}`), '', '3. ACCIONES PRIORITARIAS', ...act.map((x,i)=>`${i+1}. ${x}`), '', '4. SERVICIOS', 'Servicios más vendidos:', d.serviciosMasVendidos, '', 'Servicios menos vendidos:', d.serviciosMenosVendidos, '', '5. CAPACIDAD', 'Horarios con mayor ocupación:', d.horariosMayorOcupacion, '', 'Horarios con menor ocupación:', d.horariosMenorOcupacion].join('\n');
    }
    return copyPlain(text);
  };

  async function copyPlain(text) {
    const clean = String(text || '').replace(/\n{3,}/g, '\n\n').trim();
    try {
      if (navigator.clipboard && W.isSecureContext) {
        await navigator.clipboard.writeText(clean);
        copyOk();
        return true;
      }
    } catch(e) {}
    fallback(clean);
    return false;
  }

  function copyOk() {
    if (typeof toast === 'function') toast('Información copiada correctamente', 'ok');
    document.querySelectorAll('[id="copyGestionStatus"]').forEach(el => {
      el.style.display = 'inline-flex';
      clearTimeout(W._copyGestionStatusTimer);
      W._copyGestionStatusTimer = setTimeout(() => { el.style.display = 'none'; }, 2400);
    });
  }

  function fallback(text, title='Copiar manualmente') {
    let modal = document.getElementById('copyFallbackModal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'copyFallbackModal';
      modal.style.cssText = 'display:none;position:fixed;inset:0;background:rgba(0,0,0,.42);z-index:99999;align-items:center;justify-content:center;padding:18px';
      modal.innerHTML = `<div style="max-width:780px;width:100%;background:var(--s1,#fff);border:1px solid var(--border,#cce);border-radius:18px;padding:22px;box-shadow:0 24px 80px rgba(0,0,0,.22)">
        <div id="copyFallbackTitle" style="font-family:var(--font-h,serif);font-size:1.35rem;margin-bottom:8px">Copiar manualmente</div>
        <p style="font-size:.85rem;color:var(--muted,#667);margin-bottom:10px">Tu navegador no permitió copiar automáticamente. Selecciona el texto y cópialo manualmente.</p>
        <textarea id="copyFallbackText" style="width:100%;min-height:320px;resize:vertical;background:var(--s2,#f7fbfb);border:1px solid var(--border,#cce);border-radius:12px;color:var(--text,#112);padding:12px;font-family:monospace;font-size:.84rem;line-height:1.55"></textarea>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px;flex-wrap:wrap">
          <button class="btn btn-ghost" onclick="document.getElementById('copyFallbackModal').style.display='none'">Cerrar</button>
          <button class="btn btn-teal" onclick="document.getElementById('copyFallbackText').select();document.execCommand('copy');window._copyGestionOk()">Copiar selección</button>
        </div>
      </div>`;
      document.body.appendChild(modal);
    }
    document.getElementById('copyFallbackTitle').textContent = title;
    const ta = document.getElementById('copyFallbackText');
    ta.value = text;
    modal.style.display = 'flex';
    setTimeout(() => { ta.focus(); ta.select(); }, 80);
  }

  W._copyGestionOk = copyOk;
  W._copyPlainText = copyPlain;

  const colLabels = {
    nombre:'Nombre', telefono:'Teléfono', fecha:'Fecha', hora:'Hora', servicio:'Servicio', estado:'Estado', valor:'Valor', total:'Sesiones'
  };

  function listaDatos(tipo) {
    const d = gestionData();
    if (tipo === 'reactivar') return d.reactivar;
    if (tipo === 'paquetes') return d.candidatosPaquete;
    return d.pagosPendientesLista.map(c => ({
      nombre:c.nombre || '',
      telefono:c.telefono || c.phone || '',
      fecha:fDate(c.fecha || ''),
      hora:c.hora || '',
      servicio:c.servicio || '',
      estado:c.estado || '',
      valor:peso(money(c.precio))
    }));
  }

  function activeColumns() {
    return Array.from(document.querySelectorAll('[data-copy-col]:checked')).map(i => i.value);
  }

  function filteredRows() {
    const tipo = document.getElementById('copyListType').value;
    const q = String(document.getElementById('copyListSearch').value || '').toLowerCase();
    const rows = listaDatos(tipo).map((r, idx) => ({...r, _idx:idx}));
    return rows.filter(r => !q || Object.values(r).join(' ').toLowerCase().includes(q));
  }

  function rowText(row, cols) {
    return cols.map(c => `${colLabels[c]}: ${row[c] || 'Sin registrar'}`).join(' | ');
  }

  W.abrirCopiarListaGestion = function() {
    let modal = document.getElementById('copyListModal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'copyListModal';
      modal.style.cssText = 'display:none;position:fixed;inset:0;background:rgba(0,0,0,.42);z-index:99999;align-items:center;justify-content:center;padding:18px';
      modal.innerHTML = `<div style="max-width:980px;width:100%;max-height:92vh;overflow:auto;background:var(--s1,#fff);border:1px solid var(--border,#cce);border-radius:20px;padding:22px;box-shadow:0 24px 80px rgba(0,0,0,.22)">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:12px;flex-wrap:wrap">
          <div>
            <div style="font-family:var(--font-h,serif);font-size:1.4rem">Copiar lista seleccionada</div>
            <p style="font-size:.84rem;color:var(--muted,#667);margin:4px 0 0">Elige registros, columnas y copia texto limpio. Por defecto se excluye teléfono.</p>
          </div>
          <button class="btn btn-ghost" onclick="document.getElementById('copyListModal').style.display='none'">Cerrar</button>
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px;margin-bottom:12px">
          <label style="font-size:.78rem;color:var(--muted,#667)">Lista<select id="copyListType" class="select" onchange="window._copyGestionRefreshListModal()"><option value="pagos">Pagos pendientes</option><option value="reactivar">Personas para reactivar</option><option value="paquetes">Candidatos para paquetes</option></select></label>
          <label style="font-size:.78rem;color:var(--muted,#667)">Buscar<input id="copyListSearch" class="input" placeholder="Nombre, servicio, estado..." oninput="window._copyGestionRefreshListModal()"></label>
        </div>
        <div id="copyListColumns" style="display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 14px"></div>
        <div id="copyListRows" style="display:grid;gap:8px;margin-bottom:12px"></div>
        <textarea id="copyListPreview" style="width:100%;min-height:150px;resize:vertical;background:var(--s2,#f7fbfb);border:1px solid var(--border,#cce);border-radius:12px;color:var(--text,#112);padding:12px;font-family:monospace;font-size:.8rem;line-height:1.45"></textarea>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px;flex-wrap:wrap">
          <button class="btn btn-ghost" onclick="window._copyGestionDownloadCSV()">Descargar CSV</button>
          <button class="btn btn-ghost" onclick="window._copyGestionCopyFiltered()">Copiar toda la lista filtrada</button>
          <button class="btn btn-teal" onclick="window._copyGestionCopySelected()">Copiar registros seleccionados</button>
        </div>
      </div>`;
      document.body.appendChild(modal);
    }
    modal.style.display = 'flex';
    W._copyGestionRefreshListModal();
  };

  W._copyGestionRefreshListModal = function() {
    const colsBox = document.getElementById('copyListColumns');
    const rowsBox = document.getElementById('copyListRows');
    if (!colsBox || !rowsBox) return;
    const tipo = document.getElementById('copyListType').value;
    const cols = tipo === 'paquetes'
      ? ['nombre','total','fecha','hora','servicio','estado','telefono']
      : ['nombre','fecha','hora','servicio','estado','valor','telefono'];
    colsBox.innerHTML = cols.map(c => `<label style="display:inline-flex;align-items:center;gap:6px;font-size:.82rem;color:var(--muted,#667)"><input type="checkbox" data-copy-col value="${c}" ${c === 'telefono' ? '' : 'checked'} onchange="window._copyGestionRefreshPreview()"> ${colLabels[c]}</label>`).join('');
    const rows = filteredRows();
    rowsBox.innerHTML = rows.length ? rows.map((r,i) => `<label style="display:flex;gap:10px;align-items:flex-start;padding:10px 12px;border:1px solid var(--border,#cce);border-radius:12px;background:rgba(255,255,255,.65)"><input type="checkbox" data-copy-row value="${r._idx}" checked onchange="window._copyGestionRefreshPreview()"><span style="font-size:.86rem;line-height:1.45"><strong>${r.nombre || 'Sin nombre'}</strong><br><span style="color:var(--muted,#667)">${r.fecha || ''} ${r.hora || ''} · ${r.servicio || ''} · ${r.estado || ''}</span></span></label>`).join('') : '<div style="padding:12px;border:1px solid var(--border,#cce);border-radius:12px;color:var(--muted,#667)">No hay registros para este filtro.</div>';
    W._copyGestionRefreshPreview();
  };

  W._copyGestionRefreshPreview = function() {
    const rows = filteredRows();
    const cols = activeColumns();
    const selected = new Set(Array.from(document.querySelectorAll('[data-copy-row]:checked')).map(i => Number(i.value)));
    const tipo = document.getElementById('copyListType').selectedOptions[0].textContent;
    const lines = [`${tipo.toUpperCase()} — CUIDÁNDOTE FISIOTERAPIA`, ''].concat(rows.filter(r => selected.has(r._idx)).map((r,i) => `${i+1}. ${rowText(r, cols)}`));
    document.getElementById('copyListPreview').value = lines.join('\n') || 'Sin registros seleccionados';
  };

  W._copyGestionCopySelected = function() {
    W._copyGestionRefreshPreview();
    return copyPlain(document.getElementById('copyListPreview').value);
  };

  W._copyGestionCopyFiltered = function() {
    const rows = filteredRows();
    const cols = activeColumns();
    const tipo = document.getElementById('copyListType').selectedOptions[0].textContent;
    return copyPlain([`${tipo.toUpperCase()} — CUIDÁNDOTE FISIOTERAPIA`, ''].concat(rows.map((r,i) => `${i+1}. ${rowText(r, cols)}`)).join('\n'));
  };

  W._copyGestionDownloadCSV = function() {
    const rows = filteredRows();
    const cols = activeColumns();
    const csv = [cols.map(c => colLabels[c]).join(',')].concat(rows.map(r => cols.map(c => `"${String(r[c] || '').replace(/"/g,'""')}"`).join(','))).join('\n');
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'lista-gestion-cuidandote.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  };

  W.copiarInfoPersonaGestion = function() {
    let modal = document.getElementById('copyPersonModal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'copyPersonModal';
      modal.style.cssText = 'display:none;position:fixed;inset:0;background:rgba(0,0,0,.42);z-index:99999;align-items:center;justify-content:center;padding:18px';
      modal.innerHTML = `<div style="max-width:720px;width:100%;background:var(--s1,#fff);border:1px solid var(--border,#cce);border-radius:18px;padding:22px;box-shadow:0 24px 80px rgba(0,0,0,.22)">
        <div style="font-family:var(--font-h,serif);font-size:1.35rem;margin-bottom:8px">Copiar información de una persona</div>
        <input id="copyPersonName" class="input" placeholder="Escribe nombre o apellido" oninput="window._copyGestionPersonPreview()">
        <textarea id="copyPersonPreview" style="width:100%;min-height:250px;resize:vertical;margin-top:10px;background:var(--s2,#f7fbfb);border:1px solid var(--border,#cce);border-radius:12px;color:var(--text,#112);padding:12px;font-family:monospace;font-size:.82rem;line-height:1.5"></textarea>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px;flex-wrap:wrap"><button class="btn btn-ghost" onclick="document.getElementById('copyPersonModal').style.display='none'">Cerrar</button><button class="btn btn-teal" onclick="window._copyGestionCopyPerson()">Copiar información</button></div>
      </div>`;
      document.body.appendChild(modal);
    }
    modal.style.display = 'flex';
    setTimeout(() => document.getElementById('copyPersonName').focus(), 80);
  };

  W._copyGestionPersonPreview = function() {
    const key = String(document.getElementById('copyPersonName').value || '').trim().toLowerCase();
    const out = document.getElementById('copyPersonPreview');
    if (!key) { out.value = 'Escribe un nombre para ver la información administrativa disponible.'; return; }
    const found = citas().filter(c => String(c.nombre || '').toLowerCase().includes(key) && !isServiceRecord(c.servicio));
    if (!found.length) { out.value = 'No encontré citas para esa persona.'; return; }
    found.sort((a,b) => fDate(b.fecha).localeCompare(fDate(a.fecha)));
    const c0 = found[0];
    out.value = [
      'INFORMACIÓN ADMINISTRATIVA DE PERSONA',
      '',
      `Nombre: ${c0.nombre}`,
      `Teléfono: ${c0.telefono || c0.phone || 'Sin registrar'}`,
      `Correo: ${c0.email || 'Sin registrar'}`,
      `Total de citas registradas: ${found.length}`,
      `Última cita: ${fDate(c0.fecha)} ${c0.hora || ''}`,
      `Último servicio: ${c0.servicio || 'Sin servicio'}`,
      `Estado último registro: ${c0.estado || 'Sin estado'}`,
      '',
      'Historial reciente:',
      ...found.slice(0,8).map((c,i)=>`${i+1}. ${fDate(c.fecha)} ${c.hora || ''} | ${c.servicio || ''} | ${c.estado || ''}`)
    ].join('\n');
  };

  W._copyGestionCopyPerson = function() {
    W._copyGestionPersonPreview();
    return copyPlain(document.getElementById('copyPersonPreview').value);
  };

  W.abrirMensajeWAGestion = function() {
    const d = gestionData();
    const persona = d.reactivar[0] || d.candidatosPaquete[0] || null;
    const nombre = persona ? String(persona.nombre || '').split(' ')[0] : 'Hola';
    const phone = persona ? String(persona.telefono || '').replace(/\D/g,'') : '';
    const msg = `Hola ${nombre}, te saludamos de Cuidándote Fisioterapia. Queríamos saber cómo has seguido y ayudarte a retomar tu proceso si lo necesitas. Tenemos horarios disponibles esta semana. ¿Quieres que te compartamos opciones?`;
    showWAModal(msg, phone);
  };

  function showWAModal(msg, phone='') {
    let modal = document.getElementById('waCopyGestionModal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'waCopyGestionModal';
      modal.style.cssText = 'display:none;position:fixed;inset:0;background:rgba(0,0,0,.42);z-index:99999;align-items:center;justify-content:center;padding:18px';
      modal.innerHTML = `<div style="max-width:660px;width:100%;background:var(--s1,#fff);border:1px solid var(--border,#cce);border-radius:18px;padding:22px;box-shadow:0 24px 80px rgba(0,0,0,.22)">
        <div style="font-family:var(--font-h,serif);font-size:1.35rem;margin-bottom:8px">Mensaje para WhatsApp</div>
        <p style="font-size:.85rem;color:var(--muted,#667);margin-bottom:10px">Revísalo, edítalo y luego cópialo o abre WhatsApp. No se envía automáticamente.</p>
        <input id="waCopyGestionPhone" class="input" placeholder="Teléfono opcional" style="margin-bottom:10px">
        <textarea id="waCopyGestionText" style="width:100%;min-height:170px;resize:vertical;background:var(--s2,#f7fbfb);border:1px solid var(--border,#cce);border-radius:12px;color:var(--text,#112);padding:12px;font-size:.9rem;line-height:1.55"></textarea>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px;flex-wrap:wrap"><button class="btn btn-ghost" onclick="document.getElementById('waCopyGestionModal').style.display='none'">Cerrar</button><button class="btn btn-ghost" onclick="window._copyPlainText(document.getElementById('waCopyGestionText').value)">Copiar mensaje</button><button class="btn btn-teal" onclick="window._openWAGestionPrepared()">Abrir WhatsApp</button></div>
      </div>`;
      document.body.appendChild(modal);
    }
    document.getElementById('waCopyGestionPhone').value = phone || '';
    document.getElementById('waCopyGestionText').value = msg;
    modal.style.display = 'flex';
  }

  W._openWAGestionPrepared = function() {
    const phone = String(document.getElementById('waCopyGestionPhone').value || '').replace(/\D/g,'');
    const text = document.getElementById('waCopyGestionText').value || '';
    const url = phone ? `https://wa.me/57${phone.replace(/^57/,'')}?text=${encodeURIComponent(text)}` : `https://wa.me/?text=${encodeURIComponent(text)}`;
    W.open(url, '_blank');
  };
})();
