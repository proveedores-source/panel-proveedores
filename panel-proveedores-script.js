
(function(){
  "use strict";
  let records = [], notas = [], logEntries = [], miEmail = '';
  let clasificacionesFlujo = [], proveedoresClasif = [], arcaComprobantes = [], mediosPago = [];
  let selectedSource='email', selectedPrio='media', editSelectedPrio='media';
  let lastFileBase64='', lastFileMediaType='';
  let sortKey='fechaRecibida', sortDir=-1;

  const $ = (s,r)=>(r||document).querySelector(s);
  const $$ = (s,r)=>Array.from((r||document).querySelectorAll(s));
  const TIPO_LABEL={factura:'Factura',presupuesto:'Presupuesto',compra:'Compra',consulta:'Consulta'};
  const FUENTE_LABEL={email:'Email',whatsapp:'WhatsApp',presencial:'Presencial'};
  const ESTADO_LABEL={pendiente:'Pendiente',en_proceso:'En proceso',completado:'Completado',eliminado:'Eliminado'};
  const PRIO_ORDER={alta:0,media:1,baja:2};

  function iconFor(src){
    if(src==='email') return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/></svg>';
    if(src==='whatsapp') return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 11.5a8.5 8.5 0 0 1-12.4 7.5L3 20l1.1-5.4A8.5 8.5 0 1 1 21 11.5z"/></svg>';
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="8" r="3.2"/><path d="M4.5 20a7.5 7.5 0 0 1 15 0"/></svg>';
  }
  function todayISO(){ return new Date().toISOString().slice(0,10); }
  function fmtDateShort(iso){ if(!iso) return '—'; const d=new Date(iso+'T00:00:00'); return d.toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit',year:'2-digit'}); }
  function fmtMoney(n){ n=Number(n)||0; return n.toLocaleString('es-AR',{style:'currency',currency:'ARS',maximumFractionDigits:0}); }
  function daysBetween(a,b){ const da=new Date(a+'T00:00:00'), db=new Date(b+'T00:00:00'); return Math.round((db-da)/86400000); }
  function escapeHtml(s){ return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function showToast(msg){ const t=$('#toast'); t.textContent=msg; t.classList.remove('hidden'); clearTimeout(showToast._h); showToast._h=setTimeout(()=>t.classList.add('hidden'),3200); }

  /* ---------------- SERVER BRIDGE ---------------- */
  function runServer(fn, ...args){
    const llamada = new Promise((resolve, reject)=>{
      google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[fn](...args);
    });
    const timeout = new Promise((_, reject)=>{
      setTimeout(()=>reject(new Error('Se agotó el tiempo de espera (25s) para '+fn+'.')), 25000);
    });
    return Promise.race([llamada, timeout]);
  }

  const CLIENT_VERSION = 'v2026-09-16-facturas-paginadas'; // MDG/RVP/Otros se traen en tandas de 50 en vez de todas juntas (evita el fallo de comunicación con hojas grandes)
  let soyAdmin = false;
  function esperar_(ms){ return new Promise(resolve=>setTimeout(resolve, ms)); }
  async function callSeguro(fn, ...args){
    const maxIntentos = 5;
    for(let intento=1; intento<=maxIntentos; intento++){
      try{
        const val = await runServer(fn, ...args);
        console.log('[callSeguro] '+fn+' intento '+intento+':', val===null?'null':(Array.isArray(val)?('Array('+val.length+')'):typeof val));
        if(val===null || val===undefined){
          console.warn(fn+' devolvió null (intento '+intento+' de '+maxIntentos+'), reintentando...');
          await esperar_(600*intento);
          continue;
        }
        return { ok:true, value: val };
      }catch(err){
        console.error('[callSeguro] '+fn+' intento '+intento+' ERROR:', err);
        if(intento===maxIntentos) return { ok:false, error: err };
        await esperar_(600*intento);
      }
    }
    return { ok:false, error: new Error('respuesta vacía (null) tras '+maxIntentos+' intentos') };
  }

  let cargaInicialCompleta = false;
  let arcaCargada = false, arcaCargando = false;
  let logCargado = false, logCargando = false;

  async function cargarArcaPaginado_(fundacionKey){
    let offset = 0, filas = [], terminado = false, errorFinal = null;
    while(!terminado){
      const r = await callSeguro('getArcaComprobantesPagina', fundacionKey, offset, 40);
      if(!r.ok){ errorFinal = r.error; break; }
      filas = filas.concat(r.value.filas || []);
      terminado = r.value.terminado;
      offset += 40;
    }
    return errorFinal ? { ok:false, error: errorFinal } : { ok:true, value: filas };
  }
  async function cargarArca_(mostrarCargando){
    if(arcaCargando) return;
    arcaCargando = true;
    if(mostrarCargando) $('#arca-debug-panel').textContent = 'Cargando comprobantes...';
    const arcaMdgR = await cargarArcaPaginado_('mdg');
    const arcaRvpR = await cargarArcaPaginado_('rvp');
    arcaComprobantes = [];
    if(arcaMdgR.ok) arcaComprobantes = arcaComprobantes.concat(arcaMdgR.value||[]);
    if(arcaRvpR.ok) arcaComprobantes = arcaComprobantes.concat(arcaRvpR.value||[]);
    window.__arcaDebug = {
      mdg: arcaMdgR.ok ? { ok:true, cantidad:(arcaMdgR.value||[]).length } : { ok:false, error:(arcaMdgR.error&&arcaMdgR.error.message)||'error desconocido' },
      rvp: arcaRvpR.ok ? { ok:true, cantidad:(arcaRvpR.value||[]).length } : { ok:false, error:(arcaRvpR.error&&arcaRvpR.error.message)||'error desconocido' },
      totalEnMemoria: arcaComprobantes.length,
      hora: new Date().toLocaleTimeString('es-AR')
    };
    if(!arcaMdgR.ok || !arcaRvpR.ok) showToast('⚠️ No se pudieron cargar los comprobantes ARCA.');
    arcaCargada = true;
    arcaCargando = false;
    renderArca();
  }
  async function cargarArcaSiHaceFalta(){ if(!arcaCargada) await cargarArca_(true); }

  async function cargarLog_(){
    if(logCargando) return;
    logCargando = true;
    const logR = await callSeguro('getLogData');
    if(logR.ok) logEntries = logR.value || [];
    else showToast('⚠️ No se pudo cargar la Actividad.');
    logCargado = true;
    logCargando = false;
    renderLog();
  }
  async function cargarLogSiHaceFalta(){ if(!logCargado) await cargarLog_(); }

  async function refreshAll(){
    cargaInicialCompleta = false;
    const fallidas = [];
    const tiempos = [];
    const t0 = performance.now();
    async function medir_(nombre, fn){
      const ini = performance.now();
      const res = await fn();
      tiempos.push(nombre+': '+Math.round(performance.now()-ini)+'ms');
      return res;
    }

    const metaR = await medir_('getMeta', ()=>callSeguro('getMeta'));
    if(metaR.ok){
      miEmail = metaR.value.currentUserEmail || '';
      soyAdmin = !!metaR.value.isAdmin;
      const version = metaR.value.appVersion || '(sin versión)';
      $('#me-chip').textContent = (miEmail ? ('👤 '+miEmail) : '👤 Usuario no identificado') + '  ·  '+version;
      $('#tab-usuarios').style.display = soyAdmin ? '' : 'none';
    } else fallidas.push('usuario ('+(metaR.error&&metaR.error.message||'error')+')');

    async function cargarRecordsPaginado_(fuente){
      let offset = 0, filas = [], terminado = false, errorFinal = null;
      while(!terminado){
        const r = await callSeguro('getRecordsPagina', fuente, offset, 50);
        if(!r.ok){ errorFinal = r.error; break; }
        filas = filas.concat(r.value.filas || []);
        terminado = r.value.terminado;
        offset += 50;
      }
      return errorFinal ? { ok:false, error: errorFinal } : { ok:true, value: filas };
    }
    const mdgR = await medir_('getRecordsMdg', ()=>cargarRecordsPaginado_('mdg'));
    const rvpR = await medir_('getRecordsRvp', ()=>cargarRecordsPaginado_('rvp'));
    const otrosR = await medir_('getRecordsOtros', ()=>cargarRecordsPaginado_('otros'));
    const partesOk = [mdgR, rvpR, otrosR].filter(r=>r.ok);
    records = partesOk.reduce((acc, r)=>acc.concat(r.value||[]), []);
    // Las "eliminadas" ya no se filtran acá globalmente — cada vista decide
    // si las excluye (Tablero/Dashboard/Transferencias siempre las excluyen)
    // o si las muestra (Listado tiene un selector Activas/Eliminadas/Todas).
    records.forEach(r=>{ r._key = r.fundacion ? (r.fundacion+':'+r._rowNum) : r.id; });
    if(!mdgR.ok) fallidas.push('facturas MDG ('+(mdgR.error&&mdgR.error.message||'error')+')');
    if(!rvpR.ok) fallidas.push('facturas RVP ('+(rvpR.error&&rvpR.error.message||'error')+')');
    if(!otrosR.ok) fallidas.push('otros registros ('+(otrosR.error&&otrosR.error.message||'error')+')');

    const notasR = await medir_('getNotasData', ()=>callSeguro('getNotasData'));
    if(notasR.ok) notas = notasR.value || [];
    else fallidas.push('notas ('+(notasR.error&&notasR.error.message||'error')+')');

    const agendaR = await medir_('getAgendaData', ()=>callSeguro('getAgendaData'));
    if(agendaR.ok) agenda = agendaR.value || [];
    else fallidas.push('agenda ('+(agendaR.error&&agendaR.error.message||'error')+')');

    const listasR = await medir_('getListasChicas', ()=>callSeguro('getListasChicas'));
    if(listasR.ok){
      clasificacionesFlujo = listasR.value.clasificaciones || [];
      mediosPago = listasR.value.mediosPago || [];
      proveedoresClasif = listasR.value.proveedoresClasif || [];
      poblarFiltroClasificacionListado();
    } else fallidas.push('clasificaciones/medios de pago/proveedores ('+(listasR.error&&listasR.error.message||'error')+')');

    // ARCA y Actividad NO se cargan acá — pesan mucho y casi nunca hacen
    // falta en el primer vistazo. Se cargan solos recién cuando el usuario
    // entra a esa pestaña (ver cargarArcaSiHaceFalta / cargarLogSiHaceFalta).

    const totalMs = Math.round(performance.now()-t0);
    window.__tiemposCarga = { total: totalMs, detalle: tiempos };
    console.log('[TIEMPOS DE CARGA] Total: '+totalMs+'ms', tiempos);
    $('#diagnostico-carga-panel').style.display = 'block';
    $('#diagnostico-carga-texto').textContent = 'Total: '+(totalMs/1000).toFixed(1)+'s\n'+tiempos.join('\n');
    $('#me-chip').textContent += '  ·  cliente: '+CLIENT_VERSION+'  ·  facturas: '+records.length+'  ·  carga: '+(totalMs/1000).toFixed(1)+'s';

    renderAll();
    if(soyAdmin) refreshUsuarios();
    cargaInicialCompleta = true;

    if(fallidas.length){
      console.error('Fallaron al cargar:', fallidas);
      showToast('⚠️ No se pudo cargar: '+fallidas.join(', '));
    }
  }

  function renderAll(){ renderCounters(); renderKanban(); renderDashboard(); renderListado(); renderNotas(); renderNotasTabla(); renderAgenda(); renderTransferencias(); renderLog(); renderProvClasif(); renderClasif(); renderArca(); renderMediosPago(); }

  /* ---------------- USUARIOS (solo admins) ---------------- */
  async function refreshUsuarios(){
    try{
      const data = await runServer('getUserManagementData');
      if(!data.isAdmin) return;
      const adminSet = new Set((data.adminEmails||[]));
      const allowed = data.allowedEmails||[];
      const allEmails = Array.from(new Set([...adminSet, ...allowed]));
      $('#usuarios-empty').style.display = allEmails.length ? 'none' : 'block';
      $('#usuarios-body').innerHTML = allEmails.map(email=>{
        const esAdmin = adminSet.has(email);
        return `<tr>
          <td>${escapeHtml(email)}</td>
          <td>${esAdmin ? '⭐ Administrador' : 'Usuario'}</td>
          <td>${esAdmin ? '<span style="font-size:11px;color:var(--ink-soft)">Se edita en Propiedades del script</span>' : `<button class="btn btn-danger" data-quitar-usuario="${escapeHtml(email)}">Quitar</button>`}</td>
        </tr>`;
      }).join('');
      $$('[data-quitar-usuario]').forEach(btn=>btn.addEventListener('click', async ()=>{
        const email = btn.dataset.quitarUsuario;
        if(!confirm('¿Quitar el acceso de '+email+'?')) return;
        try{ await runServer('removeAllowedEmail', email); showToast('Acceso quitado.'); refreshUsuarios(); }
        catch(err){ showToast('Error: '+(err.message||'')); }
      }));
    }catch(err){ console.error(err); }
    renderUsuariosVistos();
  }

  function renderUsuariosVistos(){
    const mapa = {};
    logEntries.forEach(l=>{
      const u = l.usuario || 'Desconocido';
      if(!mapa[u]) mapa[u] = { usuario:u, cantidad:0, ultima:l.ts };
      mapa[u].cantidad++;
      if(new Date(l.ts) > new Date(mapa[u].ultima)) mapa[u].ultima = l.ts;
    });
    const lista = Object.values(mapa).sort((a,b)=> new Date(b.ultima)-new Date(a.ultima));
    $('#usuarios-vistos-empty').style.display = lista.length? 'none':'block';
    $('#usuarios-vistos-body').innerHTML = lista.map(u=>`<tr>
      <td>${escapeHtml(u.usuario)}</td><td class="mono">${fmtLogTs(u.ultima)}</td><td class="mono">${u.cantidad}</td>
    </tr>`).join('');
  }
  $('#btn-agregar-usuario').addEventListener('click', async ()=>{
    const email = $('#nuevo-usuario-email').value.trim();
    if(!email || email.indexOf('@')<0){ showToast('Escribí un mail válido.'); return; }
    try{
      await runServer('addAllowedEmail', email);
      $('#nuevo-usuario-email').value='';
      showToast('Usuario agregado.');
      refreshUsuarios();
    }catch(err){ showToast('Error: '+(err.message||'')); }
  });
  $('#nuevo-usuario-email').addEventListener('keydown',(e)=>{ if(e.key==='Enter'){ e.preventDefault(); $('#btn-agregar-usuario').click(); } });

  /* ---------------- CARGAR ---------------- */
  $('#source-toggle').addEventListener('click',(e)=>{
    const opt=e.target.closest('.source-option'); if(!opt) return;
    selectedSource=opt.dataset.src;
    $$('.source-option').forEach(o=>o.classList.toggle('selected', o===opt));
    const isPres = selectedSource==='presencial';
    $('#paste-block').style.display = isPres?'none':'block';
    $('#manual-only-note').style.display = isPres?'block':'none';
    if(isPres) openReviewForm({});
  });
  $('#file-input').addEventListener('change',(e)=>{ const f=e.target.files[0]; $('#file-name').textContent=f?f.name:''; $('#btn-quitar-archivo').style.display = f?'inline-block':'none'; });
  $('#btn-quitar-archivo').addEventListener('click', ()=>{ $('#file-input').value=''; $('#file-name').textContent=''; $('#btn-quitar-archivo').style.display='none'; });
  $('#btn-manual').addEventListener('click',()=>openReviewForm({}));

  $('#btn-check-folders').addEventListener('click', async ()=>{
    const statusEl = $('#check-folders-status');
    statusEl.innerHTML = '<span class="spinner"></span>Revisando Facturas/Email, WhatsApp y Presencial...';
    $('#btn-check-folders').disabled = true;
    try{
      const res = await runServer('processFacturaFolders');
      const n = res.procesados||0, errs = res.errores||0;
      if(n>0 && errs===0) statusEl.textContent = `Se cargaron ${n} factura(s) nueva(s) automáticamente.`;
      else if(n===0 && errs>0) statusEl.textContent = `⚠️ Había archivos, pero ${errs} no se pudieron procesar (ver pestaña Actividad para el detalle del error). Siguen en la carpeta para reintentar.`;
      else if(n>0 && errs>0) statusEl.textContent = `Se cargaron ${n}, pero ${errs} fallaron (ver Actividad).`;
      else statusEl.textContent = 'No había archivos nuevos.';
      if(n>0) refreshAll();
    }catch(err){
      statusEl.textContent = 'Error: '+(err.message||'no se pudo revisar.');
    }finally{
      $('#btn-check-folders').disabled = false;
      setTimeout(()=>{ statusEl.textContent=''; }, 8000);
    }
  });

  function fileToBase64(file){
    return new Promise((res,rej)=>{
      const r=new FileReader();
      r.onload=()=>res(r.result.split(',')[1]);
      r.onerror=()=>rej(new Error('No se pudo leer el archivo'));
      r.readAsDataURL(file);
    });
  }

  $('#btn-extract').addEventListener('click', async ()=>{
    const text = $('#paste-text').value.trim();
    const file = $('#file-input').files[0];
    if(!text && !file){ showToast('Pegá un texto o adjuntá una imagen primero.'); return; }
    const statusEl = $('#extract-status');
    statusEl.innerHTML = '<span class="spinner"></span>Extrayendo datos...';
    $('#btn-extract').disabled = true;
    try{
      let extracted;
      if(file){
        const base64 = await fileToBase64(file);
        lastFileBase64 = base64; lastFileMediaType = file.type||'image/jpeg';
        extracted = await runServer('extractFromImageSrv', base64, file.type||'image/jpeg', text);
      } else {
        lastFileBase64=''; lastFileMediaType='';
        extracted = await runServer('extractFromTextSrv', text);
      }
      statusEl.textContent = 'Listo. Revisá antes de guardar.';
      openReviewForm(extracted);
    }catch(err){
      console.error(err);
      statusEl.textContent = 'No se pudo procesar automáticamente (' + (err.message||'error') + '). Completá a mano.';
      openReviewForm({});
    }finally{
      $('#btn-extract').disabled = false;
      setTimeout(()=>{ statusEl.textContent=''; }, 5000);
    }
  });

  function uniqueProveedores(){ return Array.from(new Set(records.map(r=>r.proveedor).filter(Boolean))); }

  $('#file-input-comprobante').addEventListener('change',(e)=>{ const f=e.target.files[0]; $('#file-name-comprobante').textContent=f?f.name:''; $('#btn-quitar-archivo-comprobante').style.display = f?'inline-block':'none'; });
  $('#btn-quitar-archivo-comprobante').addEventListener('click', ()=>{ $('#file-input-comprobante').value=''; $('#file-name-comprobante').textContent=''; $('#btn-quitar-archivo-comprobante').style.display='none'; });
  $('#btn-manual-comprobante').addEventListener('click', ()=> openReviewForm({tipo:'compra'}));

  $('#btn-extract-comprobante').addEventListener('click', async ()=>{
    const file = $('#file-input-comprobante').files[0];
    const medio = $('#comprobante-medio').value;
    if(!file){ showToast('Adjuntá primero la foto o PDF del comprobante.'); return; }
    const statusEl = $('#extract-comprobante-status');
    statusEl.innerHTML = '<span class="spinner"></span>Extrayendo datos del comprobante...';
    $('#btn-extract-comprobante').disabled = true;
    try{
      const base64 = await fileToBase64(file);
      const extracted = await runServer('extractComprobantePagoSrv', base64, file.type||'image/jpeg', medio);
      let proveedorFinal = extracted.beneficiario || '', cuitFinal = extracted.cuitDestino || '', clasifFinal = '';
      if(extracted.beneficiario || extracted.cuitDestino){
        try{
          const match = await runServer('buscarProveedorSrv', extracted.beneficiario||'', extracted.cuitDestino||'');
          if(match && match.match && (match.criterio==='cuit' || match.score>=0.85)){
            proveedorFinal = match.match.proveedor; cuitFinal = match.match.cuit || cuitFinal; clasifFinal = match.clasificacionFlujo||'';
          } else if(match && match.match && match.score>=0.5){
            if(confirm('El comprobante dice "'+(extracted.beneficiario||'')+'".\n\n¿Es este proveedor de tu Agenda de Contactos: "'+match.match.proveedor+'"?')){
              proveedorFinal = match.match.proveedor; cuitFinal = match.match.cuit || cuitFinal; clasifFinal = match.clasificacionFlujo||'';
            }
          }
        }catch(errMatch){ console.error('buscarProveedorSrv falló, sigo sin autocompletar proveedor', errMatch); }
      }
      statusEl.textContent = 'Listo. Revisá antes de guardar.';
      openReviewForm({
        tipo: 'compra',
        proveedor: proveedorFinal,
        cuit: cuitFinal,
        importe: extracted.importe || 0,
        descripcion: 'Comprobante de ' + (extracted.medio || medio || 'transferencia'),
        clasificacionFlujo: clasifFinal,
        numeroComprobante: extracted.numeroComprobante || ''
      });
      if(extracted.fecha) $('#f-fecharecibida').value = extracted.fecha;
    }catch(err){
      console.error(err);
      statusEl.textContent = 'No se pudo procesar automáticamente (' + (err.message||'error') + '). Completá a mano.';
      openReviewForm({tipo:'compra'});
    }finally{
      $('#btn-extract-comprobante').disabled = false;
      setTimeout(()=>{ statusEl.textContent=''; }, 6000);
    }
  });

  function openReviewForm(data){
    $('#review-form').style.display='block';
    $('#f-tipo').value = data.tipo || 'compra';
    $('#f-proveedor').value = data.proveedor || '';
    $('#f-cuit').value = data.cuit || '';
    $('#f-numfactura').value = data.numeroFactura || '';
    $('#f-cae').value = data.cae || '';
    $('#f-importe').value = data.importe || '';
    $('#f-fecharecibida').value = todayISO();
    $('#f-descripcion').value = data.descripcion || '';
    $('#f-numcomprobante').value = data.numeroComprobante || '';
    $('#f-fundacion').value = data.fundacion || '';
    $('#f-centrocosto').value = data.centroCosto || '';
    $('#f-clasificacion').value = data.clasificacionFlujo || '';
    $('#f-clasificacion').setAttribute('readonly','readonly');
    ultimoCuitSugerido = '';
    $('#f-tipofactura').value = data.tipoFactura || '';
    $('#f-fechaemision').value = data.fechaEmision || '';
    $('#f-fantasia').value = data.nombreFantasia || '';
    $('#f-vencimiento').value = '';
    selectedPrio = data.prioridad || 'media';
    $$('#prio-select .prio-opt').forEach(o=>o.classList.toggle('selected', o.dataset.p===selectedPrio));
    toggleFacturaFields();
    if(data.tipo==='factura' && !data.fundacion){
      showToast('No se pudo determinar si es MDG o RVP — elegila a mano antes de guardar.');
    }
    if(data.cuit) chequearDatosPorCuit(data.cuit);
    $('#review-form').scrollIntoView({behavior:'smooth', block:'nearest'});
  }
  function toggleFacturaFields(){
    const tipoActual = $('#f-tipo').value;
    const esFactura = tipoActual==='factura';
    $('#factura-fields-tipo').style.display = esFactura?'grid':'none';
    $('#factura-fields-2').style.display = esFactura?'grid':'none';
    $('#f-vencimiento-wrap').style.display = (esFactura && $('#f-tipofactura').value==='Boleta')?'block':'none';
    $('#f-numcomprobante-wrap').style.display = (tipoActual==='compra' || tipoActual==='presupuesto') ? 'block' : 'none';
  }
  $('#f-tipo').addEventListener('change', toggleFacturaFields);
  $('#f-tipofactura').addEventListener('change', toggleFacturaFields);
  $('#prio-select').addEventListener('click',(e)=>{ const o=e.target.closest('.prio-opt'); if(!o) return; selectedPrio=o.dataset.p; $$('#prio-select .prio-opt').forEach(x=>x.classList.toggle('selected',x===o)); });
  $('#btn-cancel-review').addEventListener('click',()=>{ $('#review-form').style.display='none'; $('#paste-text').value=''; $('#file-input').value=''; $('#file-name').textContent=''; });

  $('#review-form').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const tipo = $('#f-tipo').value;
    const proveedor = $('#f-proveedor').value.trim();
    const fundacion = $('#f-fundacion').value;
    const centroCosto = $('#f-centrocosto').value;
    const clasificacionFlujo = $('#f-clasificacion').value;
    const importeVal = parseFloat($('#f-importe').value) || 0;

    const faltantes = [];
    if(!proveedor) faltantes.push('Proveedor');
    if(!fundacion) faltantes.push('Fundación');
    if(!centroCosto) faltantes.push('Centro de Costo');
    if(!clasificacionFlujo) faltantes.push('Clasificación Flujo');
    if(!importeVal) faltantes.push('Importe');
    if(tipo==='factura'){
      if(!$('#f-cuit').value.trim()) faltantes.push('CUIT');
      if(!$('#f-tipofactura').value) faltantes.push('Tipo de comprobante');
      if(!$('#f-numfactura').value.trim()) faltantes.push('N° de factura');
      if($('#f-tipofactura').value!=='Boleta' && !$('#f-cae').value.trim()) faltantes.push('CAE');
      if($('#f-tipofactura').value==='Boleta' && !$('#f-vencimiento').value) faltantes.push('Fecha de vencimiento (Boleta)');
    }
    if(faltantes.length){
      alert('Faltan completar estos campos obligatorios:\n\n• '+faltantes.join('\n• '));
      return;
    }

    let rec;
    if(tipo==='factura'){
      rec = {
        tipo, fundacion, proveedor,
        cuit: $('#f-cuit').value.trim(),
        tipoFactura: $('#f-tipofactura').value,
        numeroFactura: $('#f-numfactura').value.trim(),
        cae: $('#f-cae').value.trim(),
        fechaEmision: $('#f-fechaemision').value || '',
        nombreFantasia: $('#f-fantasia').value.trim(),
        fechaVencimiento: $('#f-vencimiento').value || '',
        importe: importeVal,
        descripcion: $('#f-descripcion').value.trim(),
        prioridad: selectedPrio,
        canal: selectedSource,
        centroCosto, clasificacionFlujo,
        fileBase64: lastFileBase64, fileMediaType: lastFileMediaType
      };
    } else {
      rec = {
        tipo, fuente: selectedSource, proveedor, fundacion,
        cuit: $('#f-cuit').value.trim(), numeroFactura:'', cae:'',
        importe: importeVal,
        descripcion: $('#f-descripcion').value.trim(),
        prioridad: selectedPrio,
        centroCosto, clasificacionFlujo,
        numeroComprobante: (tipo==='compra' || tipo==='presupuesto') ? $('#f-numcomprobante').value.trim() : '',
        fechaRecibida: $('#f-fecharecibida').value || todayISO(),
        fechaResuelta: '', estado: 'pendiente'
      };
    }
    try{
      await runServer('addRecord', rec);
      showToast('Registro guardado.');
      e.target.reset();
      $('#review-form').style.display='none';
      $('#paste-text').value=''; $('#file-input').value=''; $('#file-name').textContent='';
      lastFileBase64=''; lastFileMediaType='';
      const cuitGuardado = rec.cuit || '';
      await refreshAll();
      if(cuitGuardado){
        const match = agenda.find(a=>normCuit_(a.cuit)===normCuit_(cuitGuardado));
        if(match && !match.cbu && !match.email){
          if(confirm('"'+(match.proveedor||proveedor)+'" no tiene CBU ni email cargados en la Agenda de Contactos todavía.\n\n¿Querés completar esos datos ahora?')){
            openAgendaModal(match.id);
          }
        }
      }
    }catch(err){ showToast('No se pudo guardar: '+(err.message||'error')); }
  });

  /* ---------------- TABS ---------------- */
  $$('.tab-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      $$('.tab-btn').forEach(b=>b.classList.toggle('active', b===btn));
      $$('.view').forEach(v=>v.classList.remove('active'));
      $('#view-'+btn.dataset.view).classList.add('active');
      if(btn.dataset.view==='arca') cargarArcaSiHaceFalta();
      if(btn.dataset.view==='actividad') cargarLogSiHaceFalta();
    });
  });

  /* ---------------- KANBAN ---------------- */
  function getFiltered(){
    const tipo=$('#filt-tipo').value, fuente=$('#filt-fuente').value, prio=$('#filt-prioridad').value;
    const prov=$('#filt-proveedor').value.trim().toLowerCase();
    return records.filter(r=> r.estado!=='eliminado' && (!tipo||r.tipo===tipo)&&(!fuente||r.fuente===fuente)&&(!prio||r.prioridad===prio)&&(!prov||(r.proveedor||'').toLowerCase().includes(prov)));
  }
  ['filt-tipo','filt-fuente','filt-prioridad'].forEach(id=>$('#'+id).addEventListener('change', renderKanban));
  $('#filt-proveedor').addEventListener('input', renderKanban);

  function ticketCardHTML(r){
    const d = r.estado==='completado' && r.fechaResuelta ? daysBetween(r.fechaRecibida, r.fechaResuelta) : daysBetween(r.fechaRecibida, todayISO());
    const overdue = r.estado!=='completado' && d>7;
    const srcColor = r.fuente==='whatsapp'?'var(--src-wa)':(r.fuente==='presencial'?'var(--src-pres)':'var(--src-email)');
    const historica = !r.isAppRecord;
    const fundBadge = r.fundacion ? `<span class="ticket-tag" style="background:${r.fundacion==='mdg'?'var(--accent-soft)':'var(--media-soft)'};color:${r.fundacion==='mdg'?'var(--accent-dark)':'#7A5615'};">${r.fundacion.toUpperCase()}</span>` : '';
    return `<div class="ticket-card" data-prio="${r.prioridad}" draggable="true" data-id="${r._key||''}">
      <div class="ticket-stub">
        <span class="src" style="color:${srcColor}">${r.fuente?iconFor(r.fuente):(historica?'📋':'')} ${r.fuente?FUENTE_LABEL[r.fuente]||r.fuente:(historica?'Cargado por Mariana':'')}</span>
        <span class="days ${overdue?'overdue':''}">${d}d</span>
      </div>
      <div class="ticket-tear"></div>
      <div class="ticket-body">
        ${fundBadge}<span class="ticket-tag">${TIPO_LABEL[r.tipo]||r.tipo}</span>
        <div class="prov">${escapeHtml(r.proveedor||'—')}</div>
        <div class="desc">${escapeHtml(r.descripcion||'')}</div>
        <div class="ticket-meta">
          ${r.importe? `<span><b>${fmtMoney(r.importe)}</b></span>`:''}
          ${r.numeroFactura? `<span>N° ${escapeHtml(r.numeroFactura)}</span>`:''}
          <span>${fmtDateShort(r.fechaRecibida)}</span>
          ${r.cargadoPor? `<span>· ${escapeHtml(r.cargadoPor)}</span>`:''}
        </div>
      </div></div>`;
  }
  function renderKanban(){
    const filtered = getFiltered();
    ['pendiente','en_proceso','completado'].forEach(estado=>{
      const list = filtered.filter(r=>r.estado===estado).sort((a,b)=>PRIO_ORDER[a.prioridad]-PRIO_ORDER[b.prioridad]);
      $('#count-'+estado).textContent = list.length;
      const col = $(`.kanban-cards[data-estado="${estado}"]`);
      col.innerHTML = list.length ? list.map(ticketCardHTML).join('') : '<div class="empty-state">Sin registros</div>';
    });
    $$('.ticket-card').forEach(card=>{
      card.addEventListener('click', ()=>openEditModal(card.dataset.id));
      card.addEventListener('dragstart',(e)=>{ e.dataTransfer.setData('text/plain', card.dataset.id); });
    });
  }
  $$('.kanban-cards').forEach(col=>{
    col.addEventListener('dragover',(e)=>{ e.preventDefault(); col.classList.add('drag-over'); });
    col.addEventListener('dragleave',()=>col.classList.remove('drag-over'));
    col.addEventListener('drop', async (e)=>{
      e.preventDefault(); col.classList.remove('drag-over');
      const key = e.dataTransfer.getData('text/plain');
      if(!key) return;
      const nuevoEstado = col.dataset.estado;
      const rec = records.find(r=>r._key===key); if(!rec) return;
      const patch = {estado:nuevoEstado};
      if(nuevoEstado==='completado' && !rec.fechaResuelta) patch.fechaResuelta = todayISO();
      if(nuevoEstado!=='completado') patch.fechaResuelta = '';
      try{
        const kind = rec.tipo==='factura' ? 'factura' : 'otro';
        const idOrRow = kind==='factura' ? rec._rowNum : rec.id;
        await runServer('updateRecordRouter', kind, rec.fundacion||null, idOrRow, patch, `Movió "${rec.proveedor}" de ${ESTADO_LABEL[rec.estado]} a ${ESTADO_LABEL[nuevoEstado]}`);
        showToast('Movido a '+ESTADO_LABEL[nuevoEstado]+'.');
        refreshAll();
      }catch(err){ showToast('No se pudo mover: '+(err.message||'error')); }
    });
  });

  /* ---------------- EDIT MODAL ---------------- */
  let editingId=null, editingKind='otro', editingFundacion=null;
  function openEditModal(key){
    const r = records.find(x=>x._key===key); if(!r) return;
    editingKind = r.tipo==='factura' ? 'factura' : 'otro';
    editingFundacion = r.fundacion || null;
    editingId = editingKind==='factura' ? r._rowNum : r.id;
    $('#e-tipo').value=r.tipo; $('#e-fuente').value=r.fuente||'email'; $('#e-proveedor').value=r.proveedor||'';
    $('#e-cuit').value=formatearCuit_(r.cuit||''); $('#e-importe').value=r.importe||'';
    $('#e-numfactura').value=r.numeroFactura||''; $('#e-cae').value=r.cae||'';
    $('#e-numcomprobante').value=r.numeroComprobante||'';
    $('#e-centrocosto').value=r.centroCosto||''; $('#e-clasificacion').value=r.clasificacionFlujo||'';
    $('#e-clasificacion').setAttribute('readonly','readonly');
    $('#e-fundacion-otros').value = editingKind==='otro' ? (r.fundacion||'') : '';
    $('#e-descripcion').value=r.descripcion||''; $('#e-fecharecibida').value=r.fechaRecibida||'';
    $('#e-fecharesuelta').value=r.fechaResuelta||''; $('#e-estado').value=r.estado;
    editSelectedPrio=r.prioridad;
    $$('#e-prio-select .prio-opt').forEach(o=>o.classList.toggle('selected', o.dataset.p===editSelectedPrio));
    $('#e-factura-fields').style.display = r.tipo==='factura'?'grid':'none';
    $('#e-numcomprobante-wrap').style.display = (r.tipo==='compra' || r.tipo==='presupuesto') ? 'grid' : 'none';
    $('#e-fundacion-otros-wrap').style.display = editingKind==='otro' ? 'grid' : 'none';
    $('#modal-overlay').classList.remove('hidden');
  }
  $('#e-tipo').addEventListener('change',()=>{
    const esFactura = $('#e-tipo').value==='factura';
    const tipoVal = $('#e-tipo').value;
    $('#e-factura-fields').style.display = esFactura?'grid':'none';
    $('#e-fundacion-otros-wrap').style.display = esFactura ? 'none' : 'grid';
    $('#e-numcomprobante-wrap').style.display = (tipoVal==='compra' || tipoVal==='presupuesto') ? 'grid' : 'none';
  });
  $('#e-prio-select').addEventListener('click',(e)=>{ const o=e.target.closest('.prio-opt'); if(!o) return; editSelectedPrio=o.dataset.p; $$('#e-prio-select .prio-opt').forEach(x=>x.classList.toggle('selected',x===o)); });
  function closeModal(){ $('#modal-overlay').classList.add('hidden'); editingId=null; }
  $('#modal-close').addEventListener('click', closeModal);
  $('#btn-cancel-edit').addEventListener('click', closeModal);
  $('#modal-overlay').addEventListener('click',(e)=>{ if(e.target.id==='modal-overlay') closeModal(); });

  $('#edit-form').addEventListener('submit', async (e)=>{
    e.preventDefault();
    if(!editingId) return;
    const tipo = $('#e-tipo').value;
    const proveedorNuevo = $('#e-proveedor').value.trim();
    const patch = editingKind==='factura' ? {
      proveedor:proveedorNuevo, cuit:$('#e-cuit').value.trim(),
      importe:parseFloat($('#e-importe').value)||0,
      numeroFactura:$('#e-numfactura').value.trim(), cae:$('#e-cae').value.trim(),
      centroCosto:$('#e-centrocosto').value, clasificacionFlujo:$('#e-clasificacion').value,
      descripcion:$('#e-descripcion').value.trim(),
      prioridad: editSelectedPrio, estado:$('#e-estado').value
    } : {
      tipo, fuente:$('#e-fuente').value, proveedor:proveedorNuevo, cuit:$('#e-cuit').value.trim(),
      importe:parseFloat($('#e-importe').value)||0,
      numeroFactura: tipo==='factura'?$('#e-numfactura').value.trim():'',
      cae: tipo==='factura'?$('#e-cae').value.trim():'',
      numeroComprobante: (tipo==='compra'||tipo==='presupuesto') ? $('#e-numcomprobante').value.trim() : '',
      centroCosto:$('#e-centrocosto').value, clasificacionFlujo:$('#e-clasificacion').value,
      fundacion: $('#e-fundacion-otros').value,
      descripcion:$('#e-descripcion').value.trim(),
      fechaRecibida:$('#e-fecharecibida').value||todayISO(),
      fechaResuelta:$('#e-fecharesuelta').value||'',
      prioridad: editSelectedPrio, estado:$('#e-estado').value
    };
    try{
      await runServer('updateRecordRouter', editingKind, editingFundacion, editingId, patch, `Editó registro de "${proveedorNuevo}"`);
      showToast('Cambios guardados.'); closeModal(); refreshAll();
    }catch(err){ showToast('No se pudo guardar: '+(err.message||'error')); }
  });
  $('#btn-delete').addEventListener('click', async ()=>{
    if(!editingId && editingId!==0) return;
    const msg = editingKind==='factura'
      ? '¿Eliminar esta factura del panel? No se borra la fila de la planilla real (se mantiene para el registro contable de Mariana), pero deja de mostrarse en Tablero, Listado y Dashboard.'
      : '¿Eliminar este registro? No se puede deshacer.';
    if(confirm(msg)){
      try{ await runServer('deleteRecordRouter', editingKind, editingFundacion, editingId); showToast('Eliminado del panel.'); closeModal(); refreshAll(); }
      catch(err){ showToast('No se pudo completar: '+(err.message||'error')); }
    }
  });

  /* ---------------- DASHBOARD ---------------- */
  function getDashboardFiltered(){
    const fund = $('#dash-filt-fundacion').value, prio = $('#dash-filt-prioridad').value, est = $('#dash-filt-estado').value;
    return records.filter(r=> r.estado!=='eliminado' && (!fund || r.fundacion===fund) && (!prio || r.prioridad===prio) && (!est || r.estado===est) );
  }
  ['dash-filt-fundacion','dash-filt-prioridad','dash-filt-estado'].forEach(id=> $('#'+id).addEventListener('change', renderDashboard));

  function openDetalleModal(titulo, lista){
    $('#detalle-modal-title').textContent = titulo;
    $('#detalle-modal-empty').style.display = lista.length? 'none':'block';
    const mdgRecs = lista.filter(r=>r.fundacion==='mdg');
    const rvpRecs = lista.filter(r=>r.fundacion==='rvp');
    const mdgCount = mdgRecs.length, rvpCount = rvpRecs.length, otrosCount = lista.length - mdgCount - rvpCount;
    const sumaImporte = arr => arr.reduce((s,r)=>s+(Number(r.importe)||0),0);
    const partes = [];
    if(mdgCount) partes.push('MDG: '+mdgCount+' ('+fmtMoney(sumaImporte(mdgRecs))+')');
    if(rvpCount) partes.push('RVP: '+rvpCount+' ('+fmtMoney(sumaImporte(rvpRecs))+')');
    if(otrosCount){
      const otrosRecs = lista.filter(r=>r.fundacion!=='mdg' && r.fundacion!=='rvp');
      partes.push('Otros: '+otrosCount+' ('+fmtMoney(sumaImporte(otrosRecs))+')');
    }
    $('#detalle-modal-resumen').textContent = partes.length ? partes.join(' · ') : '';
    $('#detalle-modal-resumen').style.display = partes.length ? 'block' : 'none';
    const ordenFund = {mdg:0, rvp:1};
    const listaOrdenada = [...lista].sort((a,b)=> (ordenFund[a.fundacion]??2) - (ordenFund[b.fundacion]??2));
    $('#detalle-modal-body').innerHTML = listaOrdenada.map(r=>`<tr ${r._key?'style="cursor:pointer;" data-detalle-edit="'+r._key+'"':''}>
      <td>${r.fundacion?r.fundacion.toUpperCase():''}</td><td>${escapeHtml(r.proveedor||'—')}</td>
      <td>${TIPO_LABEL[r.tipo]||r.tipo}</td><td class="mono">${fmtMoney(r.importe)}</td>
      <td>${ESTADO_LABEL[r.estado]||r.estado}</td><td>${r.prioridad}</td>
    </tr>`).join('');
    $$('#detalle-modal-body [data-detalle-edit]').forEach(tr=>tr.addEventListener('click', ()=>{
      $('#detalle-modal-overlay').classList.add('hidden');
      openEditModal(tr.dataset.detalleEdit);
    }));
    $('#detalle-modal-overlay').classList.remove('hidden');
  }
  $('#detalle-modal-close').addEventListener('click', ()=>$('#detalle-modal-overlay').classList.add('hidden'));
  $('#detalle-modal-overlay').addEventListener('click',(e)=>{ if(e.target.id==='detalle-modal-overlay') $('#detalle-modal-overlay').classList.add('hidden'); });

  function renderDashboard(){
    const recs = getDashboardFiltered();
    const facturas = recs.filter(r=>r.tipo==='factura');
    const compras = recs.filter(r=>r.tipo==='compra');
    const presupuestos = recs.filter(r=>r.tipo==='presupuesto');
    const pendientesTodos = recs.filter(r=>r.estado!=='completado');
    const importePend = pendientesTodos.reduce((s,r)=>s+(Number(r.importe)||0),0);
    const resueltosRecs = recs.filter(r=>r.estado==='completado' && r.fechaRecibida && r.fechaResuelta);
    const diasValidos = resueltosRecs.map(r=>daysBetween(r.fechaRecibida, r.fechaResuelta)).filter(d=>Number.isFinite(d));
    const avgDays = diasValidos.length ? Math.round(diasValidos.reduce((s,d)=>s+d,0)/diasValidos.length) : null;
    const notasPend = notas.filter(n=>!n.done).length;
    const notasVencidas = notas.filter(isOverdue).length;

    const metrics = [
      {lbl:'Facturas', val:facturas.filter(r=>r.estado==='completado').length, of:'de '+facturas.length+' pagadas', click:()=>openDetalleModal('Facturas', facturas)},
      {lbl:'Compras', val:compras.filter(r=>r.estado==='completado').length, of:'de '+compras.length+' realizadas', click:()=>openDetalleModal('Compras', compras)},
      {lbl:'Presupuestos', val:presupuestos.filter(r=>r.estado==='completado').length, of:'de '+presupuestos.length+' respondidos', click:()=>openDetalleModal('Presupuestos', presupuestos)},
      {lbl:'Importe pendiente', val:fmtMoney(importePend), of:'', accent:true, click:()=>openDetalleModal('Pendientes (importe)', pendientesTodos)},
      {lbl:'Días promedio de resolución', val: avgDays===null?'—':avgDays, of: avgDays===null?'':'días', accent:true, click:()=>openDetalleModal('Resueltos (promedio de días)', resueltosRecs)},
      {lbl:'Notas pendientes', val:notasPend, of: notasVencidas>0? notasVencidas+' vencidas':'', warn: notasVencidas>0, click:()=>{
        $$('.tab-btn').forEach(b=>b.classList.toggle('active', b.dataset.view==='notas'));
        $$('.view').forEach(v=>v.classList.remove('active'));
        $('#view-notas').classList.add('active');
      }},
    ];
    $('#metric-grid').innerHTML = metrics.map((m,i)=>`<div class="metric-card ${m.accent?'accent':''} ${m.warn?'warn':''}" ${m.click?'style="cursor:pointer;" data-metric-idx="'+i+'"':''}><div class="top"><span class="val">${m.val}</span><span class="of">${m.of}</span></div><div class="lbl">${m.lbl}</div></div>`).join('') +
      (notasPend>=5? `<div class="alert-banner" style="grid-column:1/-1;">📝 Tenés ${notasPend} notas rápidas sin resolver. Revisalas desde el botón "Notas".</div>`:'');
    $$('[data-metric-idx]').forEach(el=>el.addEventListener('click', ()=>{
      const m = metrics[parseInt(el.dataset.metricIdx,10)];
      if(m && m.click) m.click();
    }));

    const prioColors={alta:'var(--alta)',media:'var(--media)',baja:'var(--baja)'};
    const overdue = recs.filter(r=>r.estado!=='completado' && daysBetween(r.fechaRecibida, todayISO())>7)
      .sort((a,b)=>daysBetween(b.fechaRecibida,todayISO())-daysBetween(a.fechaRecibida,todayISO()));
    $('#alert-list').innerHTML = overdue.length ? overdue.map(r=>`<li class="alert-item">
        <span><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${prioColors[r.prioridad]||'var(--ink-soft)'};margin-right:6px;"></span>${escapeHtml(r.proveedor||'—')} · ${TIPO_LABEL[r.tipo]} · <span style="text-transform:capitalize">${r.prioridad}</span></span>
        <span class="days">${daysBetween(r.fechaRecibida,todayISO())}d pendiente</span>
      </li>`).join('') : '<li class="empty-state">No hay pendientes atrasados. 🎉</li>';

    const prioCounts={alta:0,media:0,baja:0};
    recs.forEach(r=>{ if(r.estado!=='completado') prioCounts[r.prioridad]=(prioCounts[r.prioridad]||0)+1; });
    const maxPrio = Math.max(1, ...Object.values(prioCounts));
    $('#prio-chart').innerHTML = ['alta','media','baja'].map(p=>`<div class="bar-row" style="cursor:pointer;" data-prio-click="${p}"><span class="bar-lbl">${p[0].toUpperCase()+p.slice(1)}</span><div class="bar-track"><div class="bar-fill" style="width:${(prioCounts[p]/maxPrio*100)}%;background:${prioColors[p]}"></div></div><span class="bar-val">${prioCounts[p]}</span></div>`).join('');
    $$('[data-prio-click]').forEach(el=>el.addEventListener('click', ()=>{
      const p = el.dataset.prioClick;
      openDetalleModal('Prioridad: '+p[0].toUpperCase()+p.slice(1), pendientesTodos.filter(r=>r.prioridad===p));
    }));

    const days=[]; for(let i=13;i>=0;i--){ const d=new Date(); d.setDate(d.getDate()-i); days.push(d.toISOString().slice(0,10)); }
    const dayCounts = days.map(d=>recs.filter(r=>r.fechaRecibida===d).length);
    const dayMdg = days.map(d=>recs.filter(r=>r.fechaRecibida===d && r.fundacion==='mdg').length);
    const dayRvp = days.map(d=>recs.filter(r=>r.fechaRecibida===d && r.fundacion==='rvp').length);
    const maxDay = Math.max(1, ...dayCounts);
    $('#daily-chart').innerHTML = `<div style="display:flex;align-items:flex-end;gap:4px;height:92px;">`+
      days.map((d,i)=>{
        const total = dayCounts[i], mdg = dayMdg[i], rvp = dayRvp[i], otros = total - mdg - rvp;
        const alturaTotal = (total/maxDay*60)||2;
        const hMdg = total? (mdg/total*alturaTotal) : 0;
        const hRvp = total? (rvp/total*alturaTotal) : 0;
        const hOtros = total? (otros/total*alturaTotal) : (i===0?0:0);
        return `<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;cursor:pointer;" title="${fmtDateShort(d)}: ${total} (MDG ${mdg} · RVP ${rvp}${otros?' · otros '+otros:''})" data-dia-click="${d}">
          <span style="font-size:10.5px;font-family:var(--font-mono);color:${total?'var(--ink)':'var(--ink-soft)'};margin-bottom:3px;">${total}</span>
          <div style="width:100%;border-radius:3px 3px 0 0;height:${Math.max(alturaTotal,2)}px;min-height:2px;display:flex;flex-direction:column-reverse;overflow:hidden;">
            <div style="width:100%;background:var(--accent);height:${hMdg}px;"></div>
            <div style="width:100%;background:var(--media);height:${hRvp}px;"></div>
            <div style="width:100%;background:var(--ink-soft);opacity:0.5;height:${hOtros}px;"></div>
          </div>
        </div>`;
      }).join('') +
      `</div><div style="display:flex;justify-content:space-between;font-size:10px;color:var(--ink-soft);margin-top:4px;"><span>${fmtDateShort(days[0])}</span><span>${fmtDateShort(days[days.length-1])}</span></div>
       <div style="display:flex;gap:14px;font-size:11px;color:var(--ink-soft);margin-top:8px;">
         <span><span style="display:inline-block;width:9px;height:9px;background:var(--accent);border-radius:2px;margin-right:4px;"></span>MDG</span>
         <span><span style="display:inline-block;width:9px;height:9px;background:var(--media);border-radius:2px;margin-right:4px;"></span>RVP</span>
         <span><span style="display:inline-block;width:9px;height:9px;background:var(--ink-soft);opacity:0.5;border-radius:2px;margin-right:4px;"></span>Otros</span>
       </div>`;
    $$('#daily-chart [data-dia-click]').forEach(el=>el.addEventListener('click', ()=>{
      const dia = el.dataset.diaClick;
      openDetalleModal('Ingresado el '+fmtDateShort(dia), recs.filter(r=>r.fechaRecibida===dia));
    }));
  }

  /* ---------------- LISTADO ---------------- */
  $$('#listado-table th[data-k]').forEach(th=>{
    th.addEventListener('click',()=>{ const k=th.dataset.k; if(sortKey===k) sortDir*=-1; else{ sortKey=k; sortDir=1; } renderListado(); });
  });
  function getListadoFiltered(){
    const fund = $('#listado-filt-fundacion').value, prio = $('#listado-filt-prioridad').value, est = $('#listado-filt-estado').value;
    const clasif = $('#listado-filt-clasificacion').value;
    const activo = $('#listado-filt-activo').value; // activas | eliminadas | todas
    const q = sinTildes_($('#listado-buscar').value);
    return records.filter(r=>
      (activo==='todas' || (activo==='eliminadas' ? r.estado==='eliminado' : r.estado!=='eliminado')) &&
      (!fund || r.fundacion===fund) && (!prio || r.prioridad===prio) && (!est || r.estado===est) &&
      (!clasif || r.clasificacionFlujo===clasif) &&
      (!listadoFechaDesde || (r.fechaRecibida && r.fechaRecibida>=listadoFechaDesde)) &&
      (!listadoFechaHasta || (r.fechaRecibida && r.fechaRecibida<=listadoFechaHasta)) &&
      (!q || sinTildes_(r.proveedor).includes(q) || sinTildes_(r.cuit).includes(q)) );
  }
  ['listado-filt-fundacion','listado-filt-prioridad','listado-filt-estado','listado-filt-activo','listado-filt-clasificacion'].forEach(id=> $('#'+id).addEventListener('change', renderListado));
  function poblarFiltroClasificacionListado(){
    const sel = $('#listado-filt-clasificacion');
    const actual = sel.value;
    sel.innerHTML = '<option value="">Toda clasificación</option>' + clasificacionesFlujo.map(c=>`<option value="${escapeHtml(c.clasificacion)}">${escapeHtml(c.clasificacion)}</option>`).join('');
    sel.value = actual;
  }
  $('#listado-buscar').addEventListener('input', renderListado);

  let listadoFechaDesde = '', listadoFechaHasta = '';
  let notasFechaDesde = '', notasFechaHasta = '';
  let transfFechaDesde = '', transfFechaHasta = '';
  let arcaControlFechaDesde = '', arcaControlFechaHasta = '';
  const FECHA_POPUP_CONFIG = {
    listado: { get:()=>({desde:listadoFechaDesde,hasta:listadoFechaHasta}), set:(d,h)=>{listadoFechaDesde=d;listadoFechaHasta=h;}, chip:'listado-fecha-chip', chipTexto:'listado-fecha-chip-texto', render:()=>renderListado() },
    notas: { get:()=>({desde:notasFechaDesde,hasta:notasFechaHasta}), set:(d,h)=>{notasFechaDesde=d;notasFechaHasta=h;}, chip:'notas-fecha-chip', chipTexto:'notas-fecha-chip-texto', render:()=>renderNotasTabla() },
    transf: { get:()=>({desde:transfFechaDesde,hasta:transfFechaHasta}), set:(d,h)=>{transfFechaDesde=d;transfFechaHasta=h;}, chip:'transf-fecha-chip', chipTexto:'transf-fecha-chip-texto', render:()=>renderTransferencias() },
    arcacontrol: { get:()=>({desde:arcaControlFechaDesde,hasta:arcaControlFechaHasta}), set:(d,h)=>{arcaControlFechaDesde=d;arcaControlFechaHasta=h;}, chip:'arcacontrol-fecha-chip', chipTexto:'arcacontrol-fecha-chip-texto', render:()=>{} }
  };
  let fechaPopupModo = 'listado';
  function actualizarChipFecha(modo){
    const cfg = FECHA_POPUP_CONFIG[modo];
    const {desde, hasta} = cfg.get();
    const chip = $('#'+cfg.chip);
    if(desde || hasta){
      chip.style.display = 'inline-flex';
      $('#'+cfg.chipTexto).textContent = (desde?fmtDateShort(desde):'…')+' → '+(hasta?fmtDateShort(hasta):'…');
    } else {
      chip.style.display = 'none';
    }
  }
  function abrirPopupFecha(modo){
    fechaPopupModo = modo;
    const {desde, hasta} = FECHA_POPUP_CONFIG[modo].get();
    $('#listado-fecha-desde-tmp').value = desde;
    $('#listado-fecha-hasta-tmp').value = hasta;
    $('#listado-fecha-overlay').classList.remove('hidden');
  }
  $('#btn-listado-fecha').addEventListener('click', ()=>abrirPopupFecha('listado'));
  $('#btn-notas-fecha').addEventListener('click', ()=>abrirPopupFecha('notas'));
  $('#btn-transf-fecha').addEventListener('click', ()=>abrirPopupFecha('transf'));
  $('#btn-arcacontrol-fecha').addEventListener('click', ()=>abrirPopupFecha('arcacontrol'));
  $('#listado-fecha-close').addEventListener('click', ()=>$('#listado-fecha-overlay').classList.add('hidden'));
  $('#listado-fecha-cancelar').addEventListener('click', ()=>$('#listado-fecha-overlay').classList.add('hidden'));
  $('#listado-fecha-aceptar').addEventListener('click', ()=>{
    const d = $('#listado-fecha-desde-tmp').value || '', h = $('#listado-fecha-hasta-tmp').value || '';
    FECHA_POPUP_CONFIG[fechaPopupModo].set(d, h);
    actualizarChipFecha(fechaPopupModo);
    $('#listado-fecha-overlay').classList.add('hidden');
    FECHA_POPUP_CONFIG[fechaPopupModo].render();
  });
  ['listado','notas','transf','arcacontrol'].forEach(modo=>{
    $('#'+modo+'-fecha-limpiar').addEventListener('click', ()=>{
      FECHA_POPUP_CONFIG[modo].set('','');
      actualizarChipFecha(modo);
      FECHA_POPUP_CONFIG[modo].render();
    });
  });

  $('#btn-listado-exportar').addEventListener('click', ()=>{
    const filas = getListadoFiltered();
    if(!filas.length){ showToast('No hay filas para exportar con los filtros actuales.'); return; }
    const headers = ['Fuente','Fundación','Pagado','Tipo','Proveedor','CUIT','Tipo Factura','N° Factura','Fecha Emisión','Descripción','Importe','CAE','Medio de Pago','Centro de Costo','Clasificación Flujo','Prioridad','Estado','Cargado Por','Fecha Recibida','Fecha Resuelta'];
    const datos = filas.map(r=>({
      'Fuente': r.fuente?(FUENTE_LABEL[r.fuente]||r.fuente):'',
      'Fundación': r.fundacion?r.fundacion.toUpperCase():'',
      'Pagado': r.pagado?'Sí':'No',
      'Tipo': TIPO_LABEL[r.tipo]||r.tipo,
      'Proveedor': r.proveedor||'',
      'CUIT': r.cuit||'',
      'Tipo Factura': r.tipoFactura||'',
      'N° Factura': r.numeroFactura||'',
      'Fecha Emisión': r.fechaEmision||'',
      'Descripción': r.descripcion||'',
      'Importe': Number(r.importe)||0,
      'CAE': r.cae||'',
      'Medio de Pago': r.medioDePago||'',
      'Centro de Costo': r.centroCosto||'',
      'Clasificación Flujo': r.clasificacionFlujo||'',
      'Prioridad': r.prioridad||'',
      'Estado': ESTADO_LABEL[r.estado]||r.estado||'',
      'Cargado Por': r.cargadoPor||'',
      'Fecha Recibida': r.fechaRecibida||'',
      'Fecha Resuelta': r.fechaResuelta||''
    }));
    const ws = XLSX.utils.json_to_sheet(datos, {header: headers});
    ws['!autofilter'] = { ref: ws['!ref'] };
    ws['!cols'] = headers.map(h => ({ wch: Math.max(12, h.length+2) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Listado');
    XLSX.writeFile(wb, 'Listado_'+todayISO()+'.xlsx');
    showToast('Exportado: '+datos.length+' fila(s).');
  });

  function renderListado(){
    const rows=[...getListadoFiltered()].sort((a,b)=>{
      let va=a[sortKey], vb=b[sortKey];
      if(sortKey==='importe'){ va=Number(va)||0; vb=Number(vb)||0; } else { va=(va||'').toString(); vb=(vb||'').toString(); }
      if(va<vb) return -1*sortDir; if(va>vb) return 1*sortDir; return 0;
    });
    $('#listado-empty').style.display = rows.length?'none':'block';
    $('#listado-body').innerHTML = rows.map(r=>{
      const canPagado = !!r.fundacion && r._rowNum;
      const prioColor = r.prioridad==='alta'?'var(--alta)':(r.prioridad==='baja'?'var(--baja)':'var(--media)');
      return `<tr>
      <td>${r.fuente?FUENTE_LABEL[r.fuente]:''}</td><td>${r.fundacion?r.fundacion.toUpperCase():''}</td>
      <td style="text-align:center">${canPagado ? `<input type="checkbox" data-pagado="${r.tipo==='factura'?'factura':'otro'}:${r.fundacion}:${r._rowNum}" ${r.pagado?'checked':''}>` : '—'}</td>
      <td>${TIPO_LABEL[r.tipo]||r.tipo}</td><td style="width:14ch;max-width:14ch;white-space:normal;word-break:break-word;">${escapeHtml(r.proveedor)}</td>
      <td class="mono"><b>${fmtMoney(r.importe)}</b></td>
      <td class="mono">${escapeHtml(r.cuit)}</td><td class="mono" style="max-width:70px;white-space:normal;word-break:break-word;">${escapeHtml(r.numeroFactura)}</td>
      <td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${prioColor};margin-right:5px;"></span>${r.prioridad}</td>
      <td class="row-actions sticky-col">${r._key? `<button class="btn" data-edit="${r._key}">Editar</button>` : ''}</td>
      <td class="mono">${fmtDateShort(r.fechaRecibida)}</td>
      <td class="mono">${fmtDateShort(r.fechaResuelta)}</td><td>${ESTADO_LABEL[r.estado]||r.estado}</td>
      <td class="mono">${escapeHtml(r.cae)}</td><td>${escapeHtml(r.clasificacionFlujo)}</td><td>${escapeHtml(r.cargadoPor)}</td></tr>`;
    }).join('');
    $$('[data-pagado]').forEach(chk=>chk.addEventListener('change', async ()=>{
      const [kind, fund, rowNum] = chk.dataset.pagado.split(':');
      try{ await runServer('setPagadoRouter', kind, fund, parseInt(rowNum,10), chk.checked); showToast(chk.checked?'Marcado como pagado/completado.':'Desmarcado.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); chk.checked = !chk.checked; }
    }));
    $$('[data-edit]').forEach(b=>b.addEventListener('click',()=>openEditModal(b.dataset.edit)));
  }

  /* ---------------- NOTAS ---------------- */
  function isOverdue(n){ return n.dueAt && !n.done && new Date(n.dueAt).getTime() < Date.now(); }
  function fmtDueDate(iso){ const d=new Date(iso); return d.toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit'})+' '+d.toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'}); }
  function noteItemHTML(n){
    const overdue = isOverdue(n);
    return `<div class="note-item ${n.done?'done':''}" data-id="${n.id}">
      <input type="checkbox" ${n.done?'checked':''} data-toggle="${n.id}">
      <div class="txt">${escapeHtml(n.texto)}
        ${n.dueAt? `<span class="due ${overdue?'overdue':''}">${overdue?'⚠ vencido · ':'⏰ '}${fmtDueDate(n.dueAt)}</span>`:''}
        ${n.autor? `<span class="due">✍ ${escapeHtml(n.autor)}</span>`:''}
      </div>
      <div class="note-actions">
        ${!n.done? `<button class="btn" data-convert="${n.id}">→ Registro</button>`:''}
        <button class="btn btn-ghost" data-delnota="${n.id}">Eliminar</button>
      </div></div>`;
  }
  function renderNotas(){
    const pendientes = notas.filter(n=>!n.done).sort((a,b)=>{ const ao=isOverdue(a), bo=isOverdue(b); if(ao!==bo) return ao?-1:1; return new Date(b.createdAt)-new Date(a.createdAt); });
    const resueltas = notas.filter(n=>n.done).sort((a,b)=>new Date(b.doneAt)-new Date(a.doneAt));
    $('#notas-pendientes-list').innerHTML = pendientes.length? pendientes.map(noteItemHTML).join('') : '<div class="empty-state">Sin notas pendientes.</div>';
    $('#notas-resueltas-list').innerHTML = resueltas.length? resueltas.map(noteItemHTML).join('') : '<div class="empty-state">Todavía no resolviste ninguna.</div>';
    $$('[data-toggle]').forEach(el=>el.addEventListener('change', async ()=>{ try{ await runServer('toggleNota', el.dataset.toggle); refreshAll(); }catch(err){ showToast('Error: '+(err.message||'')); } }));
    $$('[data-delnota]').forEach(el=>el.addEventListener('click', async ()=>{ if(confirm('¿Eliminar esta nota?')){ try{ await runServer('deleteNota', el.dataset.delnota); refreshAll(); }catch(err){ showToast('Error: '+(err.message||'')); } } }));
    $$('[data-convert]').forEach(el=>el.addEventListener('click', ()=>convertNotaToRegistro(el.dataset.convert)));
    const overdueCount = notas.filter(isOverdue).length;
    const pendCount = pendientes.length;
    const badge = $('#notes-badge');
    badge.textContent = pendCount;
    badge.classList.toggle('zero', pendCount===0);
    if(overdueCount>0){ badge.style.background='var(--alta)'; badge.classList.remove('zero'); } else { badge.style.background=''; }
  }
  function convertNotaToRegistro(id){
    const n = notas.find(x=>x.id===id); if(!n) return;
    closeNotesDrawer();
    $$('.tab-btn').forEach(b=>b.classList.toggle('active', b.dataset.view==='cargar'));
    $$('.view').forEach(v=>v.classList.remove('active'));
    $('#view-cargar').classList.add('active');
    openReviewForm({tipo:'consulta', descripcion:n.texto});
    showToast('Completá el proveedor y guardá para crear el registro.');
  }
  function openNotesDrawer(){ $('#notes-drawer').classList.add('open'); $('#notes-overlay-bg').classList.remove('hidden'); }
  function closeNotesDrawer(){ $('#notes-drawer').classList.remove('open'); $('#notes-overlay-bg').classList.add('hidden'); }
  $('#notes-fab').addEventListener('click', openNotesDrawer);
  $('#notes-close').addEventListener('click', closeNotesDrawer);
  $('#notes-overlay-bg').addEventListener('click', closeNotesDrawer);
  $('#nota-toggle-fecha').addEventListener('click', ()=>$('#nota-datetime-wrap').classList.toggle('show'));

  let notaAddPrio = 'media';
  $('#nota-prio-select').addEventListener('click',(e)=>{
    const o = e.target.closest('.prio-opt'); if(!o) return;
    notaAddPrio = o.dataset.p;
    $$('#nota-prio-select .prio-opt').forEach(x=>x.classList.toggle('selected', x===o));
  });
  $('#nota-add-btn').addEventListener('click', async ()=>{
    const texto = $('#nota-text').value; const due = $('#nota-datetime').value;
    if(!texto.trim()){ showToast('Escribí algo primero.'); return; }
    try{
      await runServer('addNota', {texto: texto.trim(), dueAt: due? new Date(due).toISOString() : '', prioridad: notaAddPrio});
      $('#nota-text').value=''; $('#nota-datetime').value=''; $('#nota-datetime-wrap').classList.remove('show');
      notaAddPrio='media'; $$('#nota-prio-select .prio-opt').forEach(x=>x.classList.toggle('selected', x.dataset.p==='media'));
      showToast('Nota guardada.'); refreshAll();
    }catch(err){ showToast('Error: '+(err.message||'')); }
  });
  $('#nota-text').addEventListener('keydown',(e)=>{ if(e.key==='Enter'){ e.preventDefault(); $('#nota-add-btn').click(); } });
  $('#btn-notas-nueva').addEventListener('click', openNotesDrawer);

  /* ---------------- NOTAS: pestaña completa (tabla editable) ---------------- */
  const PRIO_COLORS_ = {alta:'var(--alta)', media:'var(--media)', baja:'var(--baja)'};
  $('#btn-notas-exportar').addEventListener('click', ()=>{
    const filas = getNotasTablaFiltradas();
    if(!filas.length){ showToast('No hay notas para exportar con los filtros actuales.'); return; }
    const headers = ['Estado','Texto','Prioridad','Recordatorio','Autor','Fecha de creación'];
    const datos = filas.map(n=>({
      'Estado': n.done ? 'Resuelta' : 'Pendiente',
      'Texto': n.texto || '',
      'Prioridad': n.prioridad || 'media',
      'Recordatorio': n.dueAt ? fmtDueDate(n.dueAt) : '',
      'Autor': n.autor || '',
      'Fecha de creación': n.createdAt ? fmtDateShort(String(n.createdAt).slice(0,10)) : ''
    }));
    const ws = XLSX.utils.json_to_sheet(datos, {header: headers});
    ws['!autofilter'] = { ref: ws['!ref'] };
    ws['!cols'] = headers.map(h => ({ wch: Math.max(12, h.length+2) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Notas');
    XLSX.writeFile(wb, 'Notas_'+todayISO()+'.xlsx');
    showToast('Exportado: '+datos.length+' nota(s).');
  });

  function getNotasTablaFiltradas(){
    const est = $('#notas-filt-estado').value, prio = $('#notas-filt-prioridad').value;
    return notas.filter(n=>{
      const fechaNota = n.dueAt ? String(n.dueAt).slice(0,10) : '';
      return (!est || (est==='pendiente' ? !n.done : n.done)) &&
        (!prio || (n.prioridad||'media')===prio) &&
        (!notasFechaDesde || (fechaNota && fechaNota>=notasFechaDesde)) &&
        (!notasFechaHasta || (fechaNota && fechaNota<=notasFechaHasta));
    }).sort((a,b)=> new Date(b.createdAt)-new Date(a.createdAt));
  }
  ['notas-filt-estado','notas-filt-prioridad'].forEach(id=> $('#'+id).addEventListener('change', renderNotasTabla));

  function renderNotasTabla(){
    const rows = getNotasTablaFiltradas();
    $('#notas-tabla-empty').style.display = rows.length ? 'none' : 'block';
    $('#notas-tabla-body').innerHTML = rows.map(n=>{
      const overdue = isOverdue(n);
      return `<tr data-notarow="${n.id}">
        <td style="text-align:center"><input type="checkbox" data-notaestado="${n.id}" ${n.done?'checked':''} title="Marcar como resuelta"></td>
        <td><input type="text" class="mono-input" style="font-family:var(--font-body);min-width:220px;" data-notatexto="${n.id}" value="${escapeHtml(n.texto)}"></td>
        <td>
          <select data-notaprio="${n.id}" style="width:auto;">
            <option value="baja" ${(n.prioridad||'media')==='baja'?'selected':''}>Baja</option>
            <option value="media" ${(n.prioridad||'media')==='media'?'selected':''}>Media</option>
            <option value="alta" ${(n.prioridad||'media')==='alta'?'selected':''}>Alta</option>
          </select>
        </td>
        <td class="mono ${overdue?'':''}" style="${overdue?'color:var(--alta);font-weight:700;':''}">${n.dueAt ? fmtDueDate(n.dueAt) + (overdue?' ⚠':'') : '—'}</td>
        <td>${escapeHtml(n.autor||'')}</td>
        <td><button class="btn btn-ghost" data-notadel="${n.id}">Eliminar</button></td>
      </tr>`;
    }).join('');

    $$('[data-notaestado]').forEach(chk=>chk.addEventListener('change', async ()=>{
      try{ await runServer('toggleNota', chk.dataset.notaestado); showToast('Actualizado.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
    $$('[data-notatexto]').forEach(inp=>inp.addEventListener('change', async ()=>{
      try{ await runServer('updateNota', inp.dataset.notatexto, {texto: inp.value.trim()}); showToast('Nota actualizada.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
    $$('[data-notaprio]').forEach(sel=>sel.addEventListener('change', async ()=>{
      try{ await runServer('updateNota', sel.dataset.notaprio, {prioridad: sel.value}); showToast('Prioridad actualizada.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
    $$('[data-notadel]').forEach(btn=>btn.addEventListener('click', async ()=>{
      if(!confirm('¿Eliminar esta nota?')) return;
      try{ await runServer('deleteNota', btn.dataset.notadel); showToast('Nota eliminada.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
  }

  /* ---------------- LOG ---------------- */
  function fmtLogTs(iso){ const d=new Date(iso); return d.toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit',year:'2-digit'})+' '+d.toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'}); }
  function renderLog(){
    const usuarios = Array.from(new Set(logEntries.map(l=>l.usuario))).sort();
    const sel = $('#log-filt-usuario'); const current = sel.value;
    sel.innerHTML = '<option value="">Todos los usuarios</option>' + usuarios.map(u=>`<option value="${escapeHtml(u)}">${escapeHtml(u)}</option>`).join('');
    sel.value = current;
    const filtro = sel.value;
    const rows = filtro ? logEntries.filter(l=>l.usuario===filtro) : logEntries;
    $('#log-empty').style.display = rows.length?'none':'block';
    $('#log-body').innerHTML = rows.map(l=>`<tr><td class="mono">${fmtLogTs(l.ts)}</td><td>${escapeHtml(l.usuario)}</td><td>${escapeHtml(l.accion)}</td></tr>`).join('');
  }
  $('#log-filt-usuario').addEventListener('change', renderLog);

  /* ---------------- COUNTERS ---------------- */
  function renderCounters(){
    const pendRecs = records.filter(r=>r.estado!=='completado');
    const facturasPendRecs = records.filter(r=>r.tipo==='factura' && r.estado!=='completado');
    const overdueRecs = records.filter(r=>r.estado!=='completado' && daysBetween(r.fechaRecibida, todayISO())>7);
    const importePend = pendRecs.reduce((s,r)=>s+(Number(r.importe)||0),0);
    const chips = [
      {n:pendRecs.length, l:'Pendientes', lista:pendRecs, titulo:'Pendientes'},
      {n:facturasPendRecs.length, l:'Facturas x pagar', lista:facturasPendRecs, titulo:'Facturas por pagar'},
      {n:fmtMoney(importePend), l:'Importe pendiente', lista:pendRecs, titulo:'Importe pendiente'},
      {n:overdueRecs.length, l:'+7 días', warn:true, lista:overdueRecs, titulo:'Pendientes hace más de 7 días'}
    ];
    $('#counters-strip').innerHTML = chips.map((c,i)=>`<div class="counter-chip ${c.warn?'warn':''}" style="cursor:pointer;" data-counter-idx="${i}"><span class="num">${c.n}</span><span class="lbl">${c.l}</span></div>`).join('');
    $$('#counters-strip [data-counter-idx]').forEach(el=>el.addEventListener('click', ()=>{
      const c = chips[parseInt(el.dataset.counterIdx,10)];
      openDetalleModal(c.titulo, c.lista);
    }));
  }

  /* ---------------- RESUMEN (manual) ---------------- */
  function buildSummaryHTML(){
    const today = todayISO();
    const pendientes = records.filter(r=>r.estado!=='completado');
    const facturasPend = pendientes.filter(r=>r.tipo==='factura');
    const importePend = pendientes.reduce((s,r)=>s+(Number(r.importe)||0),0);
    const masAntiguos = [...pendientes].sort((a,b)=>daysBetween(b.fechaRecibida,today)-daysBetween(a.fechaRecibida,today)).slice(0,5);
    const cargadosHoy = records.filter(r=>r.fechaRecibida===today);
    const resueltosHoy = records.filter(r=>r.fechaResuelta===today);
    let html = `<div style="margin-bottom:16px;"><h3 style="font-size:12px;text-transform:uppercase;color:var(--ink-soft);">Pendiente acumulado</h3>
      <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:8px;font-family:var(--font-mono);font-size:13px;">
        <div><b style="font-size:18px;display:block;">${pendientes.length}</b>en curso</div>
        <div><b style="font-size:18px;display:block;">${facturasPend.length}</b>facturas x pagar</div>
        <div><b style="font-size:18px;display:block;">${fmtMoney(importePend)}</b>importe pendiente</div>
      </div>
      ${masAntiguos.length? `<ul style="list-style:none;margin:0;padding:0;">${masAntiguos.map(r=>`<li style="font-size:12.5px;display:flex;justify-content:space-between;border-bottom:1px solid var(--line);padding:5px 0;"><span>${escapeHtml(r.proveedor)} · ${TIPO_LABEL[r.tipo]}</span><span class="mono">${daysBetween(r.fechaRecibida,today)}d</span></li>`).join('')}</ul>` : '<div style="font-size:12.5px;color:var(--ink-soft);">No hay nada pendiente. 🎉</div>'}
      </div>
      <div><h3 style="font-size:12px;text-transform:uppercase;color:var(--ink-soft);">Hoy</h3>
      <div style="display:flex;gap:14px;font-family:var(--font-mono);font-size:13px;">
        <div><b style="font-size:18px;display:block;">${cargadosHoy.length}</b>ingresaron hoy</div>
        <div><b style="font-size:18px;display:block;">${resueltosHoy.length}</b>resueltos hoy</div>
      </div></div>`;
    return html;
  }
  $('#btn-cambiar-usuario').addEventListener('click', ()=>{
    const urlActual = window.location.href.split('#')[0];
    window.location.href = 'https://accounts.google.com/AccountChooser?continue=' + encodeURIComponent(urlActual);
  });

  $('#btn-revisar-comprobantes').addEventListener('click', async ()=>{
    const statusEl = $('#comprobantes-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo comprobantes...';
    $('#btn-revisar-comprobantes').disabled = true;
    try{
      const respuesta = await runServer('processTransferenciasFolder');
      statusEl.textContent = '';
      mostrarResumenComprobantes(respuesta.resultados||[], respuesta.sinBanco||[]);
      if((respuesta.resultados||[]).some(r=>r.resultado==='ok')) refreshAll();
    }catch(err){
      statusEl.textContent = 'Error: '+(err.message||'no se pudo revisar.');
    }finally{
      $('#btn-revisar-comprobantes').disabled = false;
    }
  });

  $('#btn-revisar-compras').addEventListener('click', async ()=>{
    const statusEl = $('#compras-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo compras/presupuestos...';
    $('#btn-revisar-compras').disabled = true;
    try{
      const respuesta = await runServer('processComprasPresupuestoFolder');
      const resultados = respuesta.resultados || [];
      const ok = resultados.filter(r=>r.resultado==='ok').length;
      const sinId = resultados.filter(r=>r.resultado==='sinIdentificar').length;
      const err = resultados.filter(r=>r.resultado==='error').length;
      statusEl.textContent = resultados.length ? `Cargados: ${ok} · Sin identificar: ${sinId} · Con error: ${err}.` : 'No había archivos nuevos.';
      if(ok>0) refreshAll();
    }catch(err){
      statusEl.textContent = 'Error: '+(err.message||'no se pudo revisar.');
    }finally{
      $('#btn-revisar-compras').disabled = false;
      setTimeout(()=>{ statusEl.textContent=''; }, 10000);
    }
  });

  function mostrarResumenComprobantes(resultados, sinBanco){
    $('#summary-title').textContent = 'Comprobantes de Transferencia';
    let html = '';
    if(!resultados || !resultados.length){
      html += '<div class="summary-section"><div class="empty-state">No había comprobantes nuevos en la carpeta "Transferencias a renombrar".</div></div>';
    } else {
      const ok = resultados.filter(r=>r.resultado==='ok');
      const sinId = resultados.filter(r=>r.resultado==='sinIdentificar');
      const err = resultados.filter(r=>r.resultado==='error');
      html += `<div class="summary-section"><h3>✅ Renombrados correctamente (${ok.length})</h3>`;
      html += ok.length ? '<ul class="summary-list">' + ok.map(r=>`<li style="flex-direction:column;align-items:flex-start;gap:2px;"><span><b>${escapeHtml(r.archivo)}</b> → ${escapeHtml(r.nuevoNombre)}</span><span style="font-size:11px;color:var(--ink-soft)">${escapeHtml(r.detalle)}</span></li>`).join('') + '</ul>' : '<div class="summary-empty">Ninguno.</div>';
      html += '</div>';
      html += `<div class="summary-section"><h3>⚠️ Sin poder identificar (${sinId.length})</h3>`;
      html += sinId.length ? '<ul class="summary-list">' + sinId.map(r=>`<li style="flex-direction:column;align-items:flex-start;gap:2px;"><span><b>${escapeHtml(r.archivo)}</b></span><span style="font-size:11px;color:var(--alta)">${escapeHtml(r.detalle)}</span></li>`).join('') + '</ul>' : '<div class="summary-empty">Ninguno.</div>';
      html += '</div>';
      if(err.length){
        html += `<div class="summary-section"><h3>❌ Con error (${err.length})</h3><ul class="summary-list">` + err.map(r=>`<li style="flex-direction:column;align-items:flex-start;gap:2px;"><span><b>${escapeHtml(r.archivo)}</b></span><span style="font-size:11px;color:var(--alta)">${escapeHtml(r.detalle)}</span></li>`).join('') + '</ul></div>';
      }
    }
    html += `<div class="summary-section"><h3>🏦 Pagadas/Completadas sin Medio de Pago asignado (${(sinBanco||[]).length})</h3>`;
    html += (sinBanco&&sinBanco.length) ? '<ul class="summary-list">' + sinBanco.map(s=>`<li><span>${s.fundacion} fila ${s.fila} — <b>${escapeHtml(s.proveedor)}</b></span><span class="mono">${fmtMoney(s.importe)}</span></li>`).join('') + '</ul>' : '<div class="summary-empty">Ninguna. 🎉</div>';
    html += '</div>';
    $('#summary-content').innerHTML = html;
    $('#summary-overlay').classList.remove('hidden');
  }
  $('#btn-summary').addEventListener('click', ()=>{ $('#summary-content').innerHTML = buildSummaryHTML(); $('#summary-overlay').classList.remove('hidden'); });
  $('#summary-close').addEventListener('click', ()=>$('#summary-overlay').classList.add('hidden'));
  $('#summary-ok').addEventListener('click', ()=>$('#summary-overlay').classList.add('hidden'));
  $('#summary-overlay').addEventListener('click',(e)=>{ if(e.target.id==='summary-overlay') $('#summary-overlay').classList.add('hidden'); });

  /* ---------------- AGENDA DE CONTACTOS ---------------- */
  let agenda = [];
  let editingAgendaId = null;
  function normTexto_(s){ return String(s||'').trim().toLowerCase(); }
  function normCuit_(s){ return String(s||'').replace(/[^0-9]/g,''); }
  function findAgendaMatch(proveedor, cuit){
    const cuitN = normCuit_(cuit);
    if(cuitN && cuitN.length>5){
      const m = agenda.find(a=>normCuit_(a.cuit)===cuitN);
      if(m) return m;
    }
    const provN = normTexto_(proveedor);
    if(provN){
      const m = agenda.find(a=>normTexto_(a.proveedor)===provN);
      if(m) return m;
    }
    return null;
  }

  function sinTildes_(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim(); }

  /** Arma un combo de texto con sugerencias filtradas en vivo, a partir de
   * una lista de opciones (strings). Reutilizable para cualquier campo. */
  /** Combo de búsqueda para Proveedor/CUIT: busca en la Agenda de
   * Contactos por nombre o CUIT, y al elegir un resultado completa AMBOS
   * campos juntos (y dispara la sugerencia de clasificación). */
  function formatearCuit_(valor){
    const d = String(valor||'').replace(/[^0-9]/g,'').slice(0,11);
    if(d.length<=2) return d;
    if(d.length<=10) return d.slice(0,2)+'-'+d.slice(2);
    return d.slice(0,2)+'-'+d.slice(2,10)+'-'+d.slice(10);
  }
  ['f-cuit','e-cuit','ag-cuit'].forEach(id=>{
    const el = $('#'+id);
    if(el) el.addEventListener('input', ()=>{
      const formateado = formatearCuit_(el.value);
      if(formateado !== el.value) el.value = formateado;
    });
  });

  function wireBusquedaAgenda(inputId, listId){
    const input = $('#'+inputId), lista = $('#'+listId);
    let activo = -1, ocultarTimeout = null;
    function opciones(){ return agenda.filter(a=>a.proveedor).map(a=>({proveedor:a.proveedor, cuit:a.cuit||''})); }
    function filtrar(filtro){
      const q = sinTildes_(filtro), qDigits = normCuit_(filtro);
      const todas = opciones();
      return q ? todas.filter(o=> sinTildes_(o.proveedor).includes(q) || (qDigits && normCuit_(o.cuit).includes(qDigits)) ) : todas;
    }
    function seleccionar(o){
      $('#f-proveedor').value = o.proveedor;
      $('#f-cuit').value = formatearCuit_(o.cuit);
      lista.classList.add('hidden');
      if(o.cuit){ ultimoCuitSugerido=''; chequearDatosPorCuit(o.cuit); }
    }
    function render(filtro){
      if(ocultarTimeout){ clearTimeout(ocultarTimeout); ocultarTimeout=null; }
      const filtradas = filtrar(filtro);
      const top = filtradas.slice(0, 30);
      activo = -1;
      if(!top.length){ lista.innerHTML = '<div class="combo-empty">Sin coincidencias en la Agenda de Contactos.</div>'; }
      else lista.innerHTML = top.map((o,i)=>`<div class="combo-item" data-i="${i}">${escapeHtml(o.proveedor)} <span style="color:var(--ink-soft);font-size:11.5px;">— ${escapeHtml(o.cuit||'sin CUIT')}</span></div>`).join('');
      lista.classList.remove('hidden');
      lista.querySelectorAll('.combo-item').forEach(el=>{
        el.addEventListener('mousedown', (e)=>{ e.preventDefault(); seleccionar(top[parseInt(el.dataset.i,10)]); });
      });
    }
    input.addEventListener('focus', ()=>render(input.value));
    input.addEventListener('input', ()=>render(input.value));
    input.addEventListener('blur', ()=>{ ocultarTimeout = setTimeout(()=>lista.classList.add('hidden'), 150); });
    input.addEventListener('keydown', (e)=>{
      const items = lista.querySelectorAll('.combo-item');
      const filtradas = filtrar(input.value).slice(0,30);
      if(e.key==='ArrowDown'){ e.preventDefault(); activo = Math.min(activo+1, items.length-1); items.forEach((el,i)=>el.classList.toggle('combo-active', i===activo)); }
      else if(e.key==='ArrowUp'){ e.preventDefault(); activo = Math.max(activo-1, 0); items.forEach((el,i)=>el.classList.toggle('combo-active', i===activo)); }
      else if(e.key==='Enter'){ if(activo>-1 && filtradas[activo]){ e.preventDefault(); seleccionar(filtradas[activo]); } }
      else if(e.key==='Escape'){ lista.classList.add('hidden'); }
    });
  }
  wireBusquedaAgenda('f-proveedor', 'f-proveedor-list');
  wireBusquedaAgenda('f-cuit', 'f-cuit-list');

  function wireComboBuscable(inputId, listId, getOpciones){
    const input = $('#'+inputId), lista = $('#'+listId);
    let activo = -1, ocultarTimeout = null;
    function render(filtro){
      if(ocultarTimeout){ clearTimeout(ocultarTimeout); ocultarTimeout=null; }
      const q = sinTildes_(filtro);
      const opciones = getOpciones();
      const esCoincidenciaExacta = q && opciones.some(o=>sinTildes_(o)===q);
      const filtradas = (q && !esCoincidenciaExacta) ? opciones.filter(o=>sinTildes_(o).includes(q)) : opciones;
      const top = filtradas.slice(0, 30);
      activo = -1;
      if(!top.length){ lista.innerHTML = '<div class="combo-empty">Sin coincidencias — se guardará el texto tal cual lo escribas.</div>'; }
      else lista.innerHTML = top.map((o,i)=>`<div class="combo-item" data-i="${i}">${escapeHtml(o)}</div>`).join('');
      lista.classList.remove('hidden');
      lista.querySelectorAll('.combo-item').forEach(el=>{
        el.addEventListener('mousedown', (e)=>{ e.preventDefault(); input.value = top[parseInt(el.dataset.i,10)]; lista.classList.add('hidden'); });
      });
    }
    input.addEventListener('focus', ()=>render(input.value));
    input.addEventListener('input', ()=>render(input.value));
    input.addEventListener('blur', ()=>{ ocultarTimeout = setTimeout(()=>lista.classList.add('hidden'), 150); });
    input.addEventListener('keydown', (e)=>{
      const items = lista.querySelectorAll('.combo-item');
      if(e.key==='ArrowDown'){ e.preventDefault(); activo = Math.min(activo+1, items.length-1); items.forEach((el,i)=>el.classList.toggle('combo-active', i===activo)); if(items[activo]) items[activo].scrollIntoView({block:'nearest'}); }
      else if(e.key==='ArrowUp'){ e.preventDefault(); activo = Math.max(activo-1, 0); items.forEach((el,i)=>el.classList.toggle('combo-active', i===activo)); }
      else if(e.key==='Enter'){ if(activo>-1 && items[activo]){ e.preventDefault(); input.value = items[activo].textContent; lista.classList.add('hidden'); } }
      else if(e.key==='Escape'){ lista.classList.add('hidden'); }
    });
  }
  // Clasificación Flujo (Cargar/Editar/ARCA) usa el modal grande de búsqueda,
  // no el dropdown chico — se conecta más abajo con abrirBuscarClasificacionModal.
  wireComboBuscable('arca-inc-medio-pago', 'arca-inc-medio-pago-list', ()=>mediosPago.map(m=>m.medioDePago));
  function abrirBuscarClasificacionModal(targetInputId){
    window.__clasifModalTarget = targetInputId;
    $('#clasif-buscar-modal-input').value = '';
    renderClasifBuscarModalLista('');
    $('#clasif-buscar-modal-overlay').classList.remove('hidden');
    setTimeout(()=>$('#clasif-buscar-modal-input').focus(), 50);
  }
  function renderClasifBuscarModalLista(filtro){
    const q = sinTildes_(filtro);
    const opciones = clasificacionesFlujo.map(c=>c.clasificacion);
    const filtradas = q ? opciones.filter(o=>sinTildes_(o).includes(q)) : opciones;
    const lista = $('#clasif-buscar-modal-lista');
    if(!filtradas.length){ lista.innerHTML = '<div style="padding:14px;color:var(--ink-soft);">Sin coincidencias.</div>'; return; }
    lista.innerHTML = filtradas.map(o=>`<div class="combo-item" style="padding:10px 14px;border-bottom:1px solid var(--line);" data-valor="${escapeHtml(o)}">${escapeHtml(o)}</div>`).join('');
    $$('#clasif-buscar-modal-lista [data-valor]').forEach(el=>el.addEventListener('click', ()=>{
      const target = $('#'+window.__clasifModalTarget);
      target.value = el.dataset.valor;
      $('#clasif-buscar-modal-overlay').classList.add('hidden');
      if(window.__clasifProveedorPendiente){
        const pendiente = window.__clasifProveedorPendiente;
        window.__clasifProveedorPendiente = null;
        runServer('addProveedorClasificacion', pendiente.cuit, pendiente.proveedor, el.dataset.valor)
          .then(()=> showToast('Clasificación registrada para "'+pendiente.proveedor+'" — se va a sugerir sola la próxima vez.'))
          .catch(()=>{ /* ya existía una clasificación para ese CUIT, no pasa nada */ });
      }
    }));
  }
  $('#clasif-buscar-modal-input').addEventListener('input', ()=>renderClasifBuscarModalLista($('#clasif-buscar-modal-input').value));
  $('#clasif-buscar-modal-close').addEventListener('click', ()=>{ window.__clasifProveedorPendiente = null; $('#clasif-buscar-modal-overlay').classList.add('hidden'); });
  $('#clasif-buscar-modal-overlay').addEventListener('click',(e)=>{ if(e.target.id==='clasif-buscar-modal-overlay'){ window.__clasifProveedorPendiente = null; $('#clasif-buscar-modal-overlay').classList.add('hidden'); } });

  $$('.combo-search-icon').forEach(icon=>icon.addEventListener('click', ()=>{
    if(icon.dataset.target==='arca-inc-clasificacion'){ abrirBuscarClasificacionModal('arca-inc-clasificacion'); return; }
    const el = $('#'+icon.dataset.target);
    el.focus();
    el.dispatchEvent(new Event('input', {bubbles:true}));
  }));
  $$('.combo-edit-icon').forEach(icon=>icon.addEventListener('click', ()=>{
    abrirBuscarClasificacionModal(icon.dataset.target);
  }));
  // También se puede abrir el buscador haciendo clic directo en el campo (es de solo lectura)
  ['f-clasificacion','e-clasificacion'].forEach(id=>{
    const el = $('#'+id);
    if(el) el.addEventListener('mousedown', (e)=>{ e.preventDefault(); abrirBuscarClasificacionModal(id); });
  });
  $('#arca-inc-clasificacion').addEventListener('mousedown', (e)=>{ e.preventDefault(); abrirBuscarClasificacionModal('arca-inc-clasificacion'); });

  let ultimoCuitSugerido = '';
  function chequearDatosPorCuit(cuitCrudo){
    const cuit = normCuit_(cuitCrudo);
    if(cuit.length!==11 || cuit===ultimoCuitSugerido) return;
    const matchAgenda = agenda.find(a=>normCuit_(a.cuit)===cuit);
    const matchClasif = proveedoresClasif.find(p=>normCuit_(p.cuit)===cuit);
    const proveedorSugerido = matchAgenda && matchAgenda.proveedor;
    const clasifSugerida = matchClasif && matchClasif.clasificacion;
    if(!proveedorSugerido && !clasifSugerida) return;
    ultimoCuitSugerido = cuit;
    const partes = [];
    if(proveedorSugerido) partes.push('Proveedor: "'+proveedorSugerido+'"');
    if(clasifSugerida) partes.push('Clasificación de Flujo: "'+clasifSugerida+'"');
    $('#clasif-sugerida-texto').textContent = 'Encontramos esto para el CUIT '+cuit+' — '+partes.join(' y ')+'. ¿Querés cargarlo automáticamente?';
    $('#clasif-sugerida-overlay').dataset.proveedor = proveedorSugerido || '';
    $('#clasif-sugerida-overlay').dataset.valor = clasifSugerida || '';
    $('#clasif-sugerida-overlay').classList.remove('hidden');
  }
  $('#clasif-sugerida-si').addEventListener('click', ()=>{
    const ov = $('#clasif-sugerida-overlay');
    if(ov.dataset.proveedor) $('#f-proveedor').value = ov.dataset.proveedor;
    if(ov.dataset.valor) $('#f-clasificacion').value = ov.dataset.valor;
    ov.classList.add('hidden');
  });
  $('#clasif-sugerida-no').addEventListener('click', ()=>{
    $('#clasif-sugerida-overlay').classList.add('hidden');
  });
  $('#f-cuit').addEventListener('input', ()=> chequearDatosPorCuit($('#f-cuit').value));
  $('#f-cuit').addEventListener('blur', ()=> chequearDatosPorCuit($('#f-cuit').value));

  /** Si el proveedor se eligió/escribió por NOMBRE (sin pasar por el CUIT),
   * busca igual la Clasificación de Flujo asociada — sin popup, solo
   * completa el campo si todavía está vacío (nunca pisa algo ya elegido).
   * Si el proveedor NO tiene ninguna Clasificación registrada todavía
   * (ítem 2), pregunta si querés elegirle una ahora — y si elegís una,
   * queda registrada para que las próximas cargas de ese proveedor la
   * sugieran solas. */
  async function chequearDatosPorProveedor(){
    const nombre = $('#f-proveedor').value.trim();
    if(!nombre || nombre.length<3) return;
    const cuitActual = $('#f-cuit').value.trim();
    try{
      const res = await runServer('buscarProveedorSrv', nombre, cuitActual);
      if(!res) return;
      const cuitFinal = (res.match && res.match.cuit) || cuitActual;
      if(res.match && res.match.cuit && !cuitActual) $('#f-cuit').value = formatearCuit_(res.match.cuit);
      if(res.clasificacionFlujo){
        if(!$('#f-clasificacion').value) $('#f-clasificacion').value = res.clasificacionFlujo;
      } else if(!$('#f-clasificacion').value && cuitFinal){
        if(confirm('"'+nombre+'" todavía no tiene una Clasificación de Flujo asignada.\n\n¿Querés elegirle una ahora? Va a quedar registrada para las próximas cargas.')){
          window.__clasifProveedorPendiente = { cuit: cuitFinal, proveedor: nombre };
          abrirBuscarClasificacionModal('f-clasificacion');
        }
      }
    }catch(e){ console.error('chequearDatosPorProveedor falló (silencioso)', e); }
  }
  $('#f-proveedor').addEventListener('blur', chequearDatosPorProveedor);

  function renderAgenda(){
    const q = sinTildes_($('#agenda-buscar').value);
    const qDigits = normCuit_($('#agenda-buscar').value);
    const rows = agenda.filter(a => !q ||
      sinTildes_(a.proveedor).includes(q) ||
      (qDigits && normCuit_(a.cuit).includes(qDigits)) ||
      sinTildes_(a.cbu).includes(q) ||
      sinTildes_(a.email).includes(q)
    );
    $('#agenda-empty').style.display = rows.length? 'none':'block';
    $('#agenda-body').innerHTML = rows.map(a=>{
      const clasifMatch = proveedoresClasif.find(p=>normCuit_(p.cuit)===normCuit_(a.cuit));
      return `<tr>
      <td class="row-actions sticky-col" style="white-space:nowrap;">
        <button class="btn" data-agenda-edit="${a.id}">Editar</button>
        <button class="btn btn-ghost" data-agenda-del="${a.id}">Eliminar</button>
      </td>
      <td style="max-width:130px;white-space:normal;word-break:break-word;">${escapeHtml(a.proveedor)}</td><td class="mono">${escapeHtml(a.cuit)}</td>
      <td class="mono">${a.cbu? escapeHtml(a.cbu) : '<span style="color:var(--alta)">⚠ falta</span>'}</td>
      <td>${escapeHtml(a.email)}</td>
      <td>${clasifMatch? escapeHtml(clasifMatch.clasificacion) : '<span style="color:var(--ink-soft)">—</span>'}</td>
      <td style="max-width:130px;white-space:normal;word-break:break-word;">${escapeHtml(a.motivo)}</td>
      <td style="max-width:160px;white-space:normal;word-break:break-word;">${escapeHtml(a.descripcion)}</td>
      </tr>`;
    }).join('');
    $$('[data-agenda-edit]').forEach(b=>b.addEventListener('click', ()=>openAgendaModal(b.dataset.agendaEdit)));
    $$('[data-agenda-del]').forEach(b=>b.addEventListener('click', async ()=>{
      if(!confirm('¿Eliminar este contacto de la Agenda?')) return;
      try{ await runServer('deleteContacto', b.dataset.agendaDel); showToast('Contacto eliminado.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
  }
  $('#agenda-buscar').addEventListener('input', ()=>{
    const el = $('#agenda-buscar');
    const sinGuiones = el.value.replace(/-/g,'');
    if(sinGuiones && /^[0-9]+$/.test(sinGuiones)){
      const formateado = formatearCuit_(el.value);
      if(formateado !== el.value) el.value = formateado;
    }
    renderAgenda();
  });

  function openAgendaModal(id){
    editingAgendaId = id || null;
    const a = id ? agenda.find(x=>x.id===id) : null;
    $('#agenda-modal-title').textContent = a ? 'Editar contacto' : 'Nuevo contacto';
    $('#ag-proveedor').value = a? a.proveedor:''; $('#ag-cuit').value = a? a.cuit:'';
    $('#ag-cbu').value = a? a.cbu:''; $('#ag-email').value = a? a.email:'';
    $('#ag-motivo').value = a? a.motivo:'Factura'; $('#ag-descripcion').value = a? a.descripcion:'';
    $('#agenda-modal-overlay').classList.remove('hidden');
  }
  $('#btn-agenda-nuevo').addEventListener('click', ()=>openAgendaModal(null));
  $('#agenda-modal-close').addEventListener('click', ()=>$('#agenda-modal-overlay').classList.add('hidden'));
  $('#agenda-cancel').addEventListener('click', ()=>$('#agenda-modal-overlay').classList.add('hidden'));
  $('#agenda-form').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const data = {
      proveedor: $('#ag-proveedor').value.trim(), cuit: $('#ag-cuit').value.trim(),
      cbu: $('#ag-cbu').value.trim(), email: $('#ag-email').value.trim(),
      motivo: $('#ag-motivo').value.trim(), descripcion: $('#ag-descripcion').value.trim()
    };
    if(!data.proveedor){ showToast('Falta el nombre del proveedor.'); return; }
    try{
      if(editingAgendaId) await runServer('updateContacto', editingAgendaId, data);
      else await runServer('addContacto', data);
      showToast('Guardado.');
      $('#agenda-modal-overlay').classList.add('hidden');
      refreshAll();
    }catch(err){ showToast('Error: '+(err.message||'')); }
  });

  function formatCuitDigits_(digits){
    digits = String(digits||'').replace(/[^0-9]/g,'');
    if(digits.length===11) return digits.slice(0,2)+'-'+digits.slice(2,10)+'-'+digits.slice(10);
    return digits;
  }

  /** Parsea el CSV de contactos exportado por Banco Galicia:
   * Cuenta destino;CUIT/CUIL;Nombre/Razón social;Descripción de Contacto;
   * Tipo de transferencia;Acceso;Mail 1..5;Agregado por
   * A veces el nombre real está en "Descripción de Contacto" y "Nombre/Razón
   * social" repite el número de cuenta — y hay filas duplicadas por contacto. */
  function parseBankContactsCsv(text){
    text = text.replace(/^\uFEFF/, ''); // sacar BOM
    const lines = text.split(/\r?\n/).filter(l=>l.trim().length);
    if(!lines.length) return [];
    const headers = lines[0].split(';').map(h=>h.trim());
    const idx = name => headers.indexOf(name);
    const iCuenta = idx('Cuenta destino'), iCuit = idx('CUIT/CUIL'), iNombre = idx('Nombre/Razón social'),
          iDesc = idx('Descripción de Contacto'), iMail1 = idx('Mail 1'), iMail2 = idx('Mail 2'),
          iMail3 = idx('Mail 3'), iMail4 = idx('Mail 4'), iMail5 = idx('Mail 5');
    const rows = [];
    for(let i=1;i<lines.length;i++){
      const cols = lines[i].split(';');
      const cuenta = (cols[iCuenta]||'').trim();
      const cuitDigits = (cols[iCuit]||'').trim();
      const nombre = (cols[iNombre]||'').trim();
      const desc = (cols[iDesc]||'').trim();
      const soloNumerico = /^[0-9]+$/.test(nombre) || nombre===cuenta;
      const proveedor = (soloNumerico && desc) ? desc : (nombre || desc);
      const email = [iMail1,iMail2,iMail3,iMail4,iMail5].map(ix=>ix>-1?(cols[ix]||'').trim():'').find(m=>m);
      if(!proveedor && !cuenta) continue;
      rows.push({ proveedor: proveedor || cuenta, cbu: cuenta, cuit: formatCuitDigits_(cuitDigits), email: email||'', motivo:'', descripcion:'' });
    }
    // Dedupe: varias filas repiten el mismo contacto (agregado por distintas personas)
    const map = new Map();
    rows.forEach(r=>{
      const key = (r.cuit && r.cuit.length>5 ? r.cuit : normTexto_(r.cbu));
      if(!map.has(key)) map.set(key, r);
      else {
        const ex = map.get(key);
        if(!ex.email && r.email) ex.email = r.email;
        if((!ex.proveedor || /^[0-9]+$/.test(ex.proveedor)) && r.proveedor) ex.proveedor = r.proveedor;
      }
    });
    return Array.from(map.values());
  }

  /* Importar Excel o CSV de Agenda */
  $('#agenda-import-input').addEventListener('change', async (e)=>{
    const file = e.target.files[0]; if(!file) return;
    const statusEl = $('#agenda-import-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo archivo...';
    try{
      let rows;
      if(file.name.toLowerCase().endsWith('.csv')){
        const text = await file.text();
        rows = parseBankContactsCsv(text);
      } else {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, {type:'array'});
        const sheetName = wb.SheetNames.includes('Formulario') ? 'Formulario' : wb.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {defval:''});
        rows = json.map(r=>({
          proveedor: r['Nombre de Proveedor'] || r['Proveedor'] || '',
          cbu: r['CBU/CVU/Alias/Nro cuenta'] || r['CBU/CVU/Alias'] || '',
          cuit: r['CUIT'] || '',
          email: r['Email destinatario (opcional)'] || r['Email'] || '',
          motivo: r['Motivo'] || '',
          descripcion: r['Descripción (opcional)'] || r['Descripcion'] || ''
        })).filter(r=>r.proveedor);
      }
      if(!rows.length){ statusEl.textContent='No se encontraron filas para importar.'; return; }
      statusEl.innerHTML = `<span class="spinner"></span>Importando ${rows.length} contacto(s)...`;
      const res = await runServer('importAgendaBulk', rows);
      statusEl.textContent = `✅ Listo: ${res.creados} nuevos, ${res.actualizados} actualizados.`;
      showToast(`Importado: ${res.creados} nuevos, ${res.actualizados} actualizados.`);
      refreshAll();
    }catch(err){
      console.error(err);
      statusEl.textContent = '❌ Error: '+(err.message||'no se pudo importar.');
      showToast('Error al importar: '+(err.message||''));
    }finally{
      e.target.value='';
      setTimeout(()=>{ statusEl.textContent=''; }, 8000);
    }
  });

  /* Exportar Excel de Agenda (mismo formato/columnas del original) */
  $('#btn-agenda-exportar').addEventListener('click', ()=>{
    const data = agenda.map(a=>({
      'Nombre de Proveedor': a.proveedor, 'CBU/CVU/Alias/Nro cuenta': a.cbu, 'CUIT': a.cuit,
      'Email destinatario (opcional)': a.email, 'Motivo': a.motivo, 'Descripción (opcional)': a.descripcion
    }));
    const ws = XLSX.utils.json_to_sheet(data, {header:['Nombre de Proveedor','CBU/CVU/Alias/Nro cuenta','CUIT','Email destinatario (opcional)','Motivo','Descripción (opcional)']});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Formulario');
    XLSX.writeFile(wb, 'Agenda_Contactos_'+todayISO()+'.xlsx');
  });

  $('#btn-agenda-exportar-macro').addEventListener('click', ()=>{
    const conCuit = agenda.filter(a=> (a.cuit||'').replace(/[^0-9]/g,'').length===11);
    if(!conCuit.length){ showToast('Ningún contacto tiene CUIT válido de 11 dígitos cargado en la Agenda.'); return; }
    const headers = ['Tipo de documento - Beneficiario','Número de documento - Beneficiario','Condición IIBB','Condición Ganancias','Condición de I.V.A.','Razón social - Beneficiario','Calle','Núm. De puerta','Unidad funcional','Código postal','Núm. IIBB','E-mail','Piso','País','Provincia','Ciudad'];
    const data = conCuit.map(a=>({
      'Tipo de documento - Beneficiario':'80',
      'Número de documento - Beneficiario': (a.cuit||'').replace(/[^0-9]/g,''),
      'Condición IIBB':'', 'Condición Ganancias':'', 'Condición de I.V.A.':'',
      'Razón social - Beneficiario': (a.proveedor||'').slice(0,120),
      'Calle':'', 'Núm. De puerta':'', 'Unidad funcional':'', 'Código postal':'',
      'Núm. IIBB':'', 'E-mail': a.email||'', 'Piso':'', 'País':'', 'Provincia':'', 'Ciudad':''
    }));
    const ws = XLSX.utils.json_to_sheet(data, {header: headers});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
    XLSX.writeFile(wb, 'Beneficiarios_Banco_Macro_'+todayISO()+'.xlsx');
    showToast('Generado con '+data.length+' contacto(s). Ojo: Cond. IIBB/Ganancias/IVA y Código Postal quedan en blanco (no los tenemos cargados) — hay que completarlos a mano antes de subir al banco.');
  });

  /* ---------------- TRANSFERENCIAS ---------------- */
  function facturaNumeroCorto(numeroFactura){
    if(!numeroFactura) return '';
    const parts = String(numeroFactura).split('-');
    let last = parts[parts.length-1].trim();
    const sinCeros = last.replace(/^0+/, '');
    return sinCeros || last;
  }
  function getFacturasPendientesDePago(){
    return records.filter(r => r.tipo==='factura' && !r.pagado && r.fundacion && r.estado!=='eliminado');
  }
  function getTransfFiltradas(){
    const fund = $('#transf-filt-fundacion').value;
    const q = $('#transf-buscar').value.trim().toLowerCase();
    return getFacturasPendientesDePago().filter(r=>
      (!fund || r.fundacion===fund) && (!q || (r.proveedor||'').toLowerCase().includes(q)) &&
      (!transfFechaDesde || (r.fechaRecibida && r.fechaRecibida>=transfFechaDesde)) &&
      (!transfFechaHasta || (r.fechaRecibida && r.fechaRecibida<=transfFechaHasta))
    );
  }
  $('#btn-transf-exportar-excel').addEventListener('click', ()=>{
    const filas = getTransfSeleccionadas();
    if(!filas.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    const headers = ['Fundación','Proveedor','CUIT','N° Factura','Importe','Fecha Recibida','CBU/CVU/Alias','Email','En Agenda'];
    const datos = filas.map(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      return {
        'Fundación': r.fundacion.toUpperCase(),
        'Proveedor': r.proveedor || '',
        'CUIT': r.cuit || '',
        'N° Factura': r.numeroFactura || '',
        'Importe': Number(r.importe)||0,
        'Fecha Recibida': r.fechaRecibida || '',
        'CBU/CVU/Alias': match ? (match.cbu||'') : '',
        'Email': match ? (match.email||'') : '',
        'En Agenda': (match && match.cbu) ? 'Sí' : 'Falta CBU'
      };
    });
    const ws = XLSX.utils.json_to_sheet(datos, {header: headers});
    ws['!autofilter'] = { ref: ws['!ref'] };
    ws['!cols'] = headers.map(h => ({ wch: Math.max(12, h.length+2) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pagos');
    XLSX.writeFile(wb, 'Pagos_'+todayISO()+'.xlsx');
    showToast('Exportado: '+datos.length+' fila(s).');
    abrirConfirmacionPagos(filas, '');
  });
  function renderTransferencias(){
    const pendientes = getTransfFiltradas();
    $('#transf-empty').style.display = pendientes.length? 'none':'block';
    $('#transf-body').innerHTML = pendientes.map(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      const tieneCbu = match && match.cbu;
      return `<tr>
        <td style="text-align:center"><input type="checkbox" class="transf-check" data-transfid="${r.fundacion}:${r._rowNum}" ${tieneCbu?'':'disabled'}></td>
        <td>${r.fundacion.toUpperCase()}</td><td>${escapeHtml(r.proveedor)}</td><td class="mono">${escapeHtml(r.cuit)}</td>
        <td class="mono">${escapeHtml(r.numeroFactura)}</td><td class="mono">${fmtMoney(r.importe)}</td>
        <td class="mono">${match? escapeHtml(match.cbu||'—') : '—'}</td>
        <td>${match? escapeHtml(match.email||'—') : '—'}</td>
        <td>${tieneCbu? '<span style="color:var(--baja)">✓ OK</span>' : '<span style="color:var(--alta)">⚠ falta en Agenda</span>'}</td>
        <td>${r._key? `<button class="btn" data-transf-edit="${r._key}">Ver / Editar</button>` : ''}</td>
      </tr>`;
    }).join('');
    $$('.transf-check').forEach(chk=>chk.addEventListener('change', updateTransfResumen));
    $$('[data-transf-edit]').forEach(btn=>btn.addEventListener('click', ()=>openEditModal(btn.dataset.transfEdit)));
    updateTransfResumen();
  }
  ['transf-filt-fundacion','transf-buscar'].forEach(id=>$('#'+id).addEventListener('input', renderTransferencias));
  $('#transf-filt-fundacion').addEventListener('change', renderTransferencias);

  function getTransfSeleccionadas(){
    const pendientes = getFacturasPendientesDePago();
    const seleccionadas = [];
    $$('.transf-check:checked').forEach(chk=>{
      const [fund, rowNum] = chk.dataset.transfid.split(':');
      const rec = pendientes.find(r=>r.fundacion===fund && String(r._rowNum)===rowNum);
      if(rec) seleccionadas.push(rec);
    });
    return seleccionadas;
  }
  function updateTransfResumen(){
    const sel = getTransfSeleccionadas();
    const total = sel.reduce((s,r)=>s+(Number(r.importe)||0),0);
    $('#transf-resumen-cant').textContent = sel.length;
    $('#transf-resumen-total').textContent = fmtMoney(total);
    $('#transf-resumen-warn').style.display = sel.length>50 ? 'inline' : 'none';
  }

  function abrirConfirmacionPagos(lista, banco){
    if(!lista.length) return;
    $('#confirmar-pagos-body').innerHTML = lista.map((r,i)=>`<tr>
      <td style="text-align:center"><input type="checkbox" class="confirmar-pago-check" data-idx="${i}" checked></td>
      <td>${escapeHtml(r.proveedor||'—')}</td>
      <td class="mono">${escapeHtml(r.numeroFactura||'')}</td>
      <td class="mono">${fmtMoney(r.importe)}</td>
    </tr>`).join('');
    $('#confirmar-pagos-overlay').dataset.lista = JSON.stringify(lista.map(r=>({fundacion:r.fundacion, rowNum:r._rowNum, proveedor:r.proveedor})));
    $('#confirmar-pagos-overlay').dataset.banco = banco || '';
    $('#confirmar-pagos-overlay').querySelector('.hint').textContent = banco
      ? `Se tildaron todas por defecto y se van a marcar como pagadas vía "${banco}". Destildá las que no correspondan antes de confirmar.`
      : `Se tildaron todas por defecto y se van a marcar como Pagadas. Destildá las que no correspondan antes de confirmar.`;
    $('#confirmar-pagos-overlay').classList.remove('hidden');
  }
  $('#confirmar-pagos-close').addEventListener('click', ()=>$('#confirmar-pagos-overlay').classList.add('hidden'));
  $('#confirmar-pagos-cancelar').addEventListener('click', ()=>$('#confirmar-pagos-overlay').classList.add('hidden'));
  $('#confirmar-pagos-aceptar').addEventListener('click', async ()=>{
    const lista = JSON.parse($('#confirmar-pagos-overlay').dataset.lista || '[]');
    const banco = $('#confirmar-pagos-overlay').dataset.banco || '';
    const marcar = [];
    $$('.confirmar-pago-check').forEach(chk=>{
      if(chk.checked) marcar.push(lista[parseInt(chk.dataset.idx,10)]);
    });
    $('#confirmar-pagos-aceptar').disabled = true;
    let ok=0, fallidos=0;
    for(const item of marcar){
      try{ await runServer('confirmarPagoConBanco', item.fundacion, item.rowNum, banco); ok++; }
      catch(err){ fallidos++; }
    }
    $('#confirmar-pagos-aceptar').disabled = false;
    $('#confirmar-pagos-overlay').classList.add('hidden');
    showToast('Marcadas como pagadas: '+ok+(fallidos? ' · '+fallidos+' con error':'')+'.');
    refreshAll();
  });

  $('#btn-transf-exportar').addEventListener('click', async ()=>{
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    if(sel.length>50){ showToast('El banco acepta máximo 50 pagos por archivo. Desmarcá algunas.'); return; }
    const filas = sel.map(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      const cbu = match? match.cbu : '';
      const numCorto = facturaNumeroCorto(r.numeroFactura);
      const descripcion = ('FC '+numCorto).slice(0,12);
      const mensajeFallback = ('OP - '+cbu+'- FCFactura - FC '+numCorto+' - $'+Math.round(Number(r.importe)||0));
      const mensaje = (r.nombreArchivoPdf || mensajeFallback).slice(0,200);
      return {
        'CBU/CVU/Alias/Nro cuenta': cbu,
        'Importe': Number(r.importe)||0,
        'Motivo': (match && match.motivo) ? match.motivo : 'Factura',
        'Descripción (opcional)': descripcion,
        'Email destinatario (opcional)': match? (match.email||'') : '',
        'Mensaje del email (opcional)': mensaje
      };
    });
    const ws = XLSX.utils.json_to_sheet(filas, {header:['CBU/CVU/Alias/Nro cuenta','Importe','Motivo','Descripción (opcional)','Email destinatario (opcional)','Mensaje del email (opcional)']});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Formulario');
    const total = sel.reduce((s,r)=>s+(Number(r.importe)||0),0);
    XLSX.writeFile(wb, 'TRANSF_BANCO_GALICIA_'+todayISO()+'.xlsx');
    try{ await runServer('logExportTransferencias', sel.length, Math.round(total)); }catch(e){}
    const marcados = await marcarPdfsComoOpFinalizada(sel);
    showToast('Archivo generado: '+sel.length+' pago(s).'+(marcados.faltantes.length? ' ⚠️ '+marcados.faltantes.length+' PDF(s) no se encontraron para marcar.' : ''));
    abrirConfirmacionPagos(sel, 'Banco Galicia');
  });

  async function marcarPdfsComoOpFinalizada(seleccionadas){
    const faltantes = [];
    for(const r of seleccionadas){
      try{
        const res = await runServer('marcarOpFinalizada', r.fundacion, r._rowNum);
        if(!res || !res.ok) faltantes.push(r.proveedor);
      }catch(err){ faltantes.push(r.proveedor); }
    }
    return { faltantes };
  }

  $('#btn-transf-exportar-macro').addEventListener('click', (e)=>{
    e.stopPropagation();
    $('#menu-macro').classList.toggle('hidden');
  });
  document.addEventListener('click', ()=>{ $('#menu-macro').classList.add('hidden'); $('#menu-caja').classList.add('hidden'); });
  $('#menu-macro').addEventListener('click', (e)=> e.stopPropagation());

  $('#btn-caja').addEventListener('click', (e)=>{
    e.stopPropagation();
    $('#menu-caja').classList.toggle('hidden');
  });
  $('#menu-caja').addEventListener('click', (e)=> e.stopPropagation());
  $$('#menu-caja [data-caja]').forEach(btn=>btn.addEventListener('click', ()=>{
    $('#menu-caja').classList.add('hidden');
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para marcar como pagada.'); return; }
    abrirConfirmacionPagos(sel, btn.dataset.caja);
  }));

  $('#btn-tarjeta-corporativa').addEventListener('click', ()=>{
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para marcar como pagada.'); return; }
    abrirConfirmacionPagos(sel, 'Tarjeta de Crédito Corporativa');
  });

  $('#btn-macro-pago-proveedores').addEventListener('click', async ()=>{
    $('#menu-macro').classList.add('hidden');
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    const sinCbu = [];
    const filas = [];
    sel.forEach(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      const cbuDigits = (match? (match.cbu||'') : '').replace(/[^0-9]/g,'');
      const esNumerico = cbuDigits.length>=15; // CBU (22) o Nro de cuenta (15) — Macro no acepta alias
      if(!esNumerico){ sinCbu.push(r.proveedor); return; }
      const cuitDigits = (r.cuit||'').replace(/[^0-9]/g,'');
      const modalidad = cbuDigits.indexOf('0285')===0 ? '02' : '04'; // 02: cuenta Macro, 04: otros bancos
      const numCorto = facturaNumeroCorto(r.numeroFactura);
      filas.push({
        'Modalidad de Pago': modalidad,
        'Número de documento Beneficiario': cuitDigits,
        'Denominación del Beneficiario': (r.proveedor||'').slice(0,120),
        'Identificación del pago': ('FC '+numCorto).slice(0,30),
        'Importe': Number(r.importe)||0,
        'Cuenta de pago - Nro/CBU o Nº Cheque': cbuDigits,
        'Marca de registración de cheque': '',
        'Fecha de pago - Emisión': new Date().toLocaleDateString('es-AR'),
        'Fecha de pago diferido': ''
      });
    });
    if(!filas.length){ showToast('Ninguna seleccionada tiene CBU/cuenta numérica en la Agenda (Banco Macro no acepta alias).'); return; }
    const ws = XLSX.utils.json_to_sheet(filas, {header:['Modalidad de Pago','Número de documento Beneficiario','Denominación del Beneficiario','Identificación del pago','Importe','Cuenta de pago - Nro/CBU o Nº Cheque','Marca de registración de cheque','Fecha de pago - Emisión','Fecha de pago diferido']});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
    XLSX.writeFile(wb, 'TRANSF_BANCO_MACRO_'+todayISO()+'.xlsx');
    const total = filas.reduce((s,f)=>s+f['Importe'],0);
    try{ await runServer('logExportTransferencias', filas.length, Math.round(total)); }catch(e){}
    const incluidas = sel.filter(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      return match && (match.cbu||'').replace(/[^0-9]/g,'').length>=15;
    });
    const marcados = await marcarPdfsComoOpFinalizada(incluidas);
    if(sinCbu.length) showToast('⚠️ '+filas.length+' incluidas. '+sinCbu.length+' NO se incluyeron por no tener CBU numérico: '+sinCbu.slice(0,5).join(', ')+(sinCbu.length>5?'...':''));
    else showToast('Archivo Banco Macro generado: '+filas.length+' pago(s).'+(marcados.faltantes.length? ' ⚠️ '+marcados.faltantes.length+' PDF(s) no se encontraron para marcar.':''));
    abrirConfirmacionPagos(incluidas, 'Banco Macro');
  });

  $('#btn-macro-transf-masivas').addEventListener('click', async ()=>{
    $('#menu-macro').classList.add('hidden');
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    const sinCbu = [];
    const filas = [];
    sel.forEach(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      const cbuDigits = (match? (match.cbu||'') : '').replace(/[^0-9]/g,'');
      const esNumerico = cbuDigits.length>=15;
      if(!esNumerico){ sinCbu.push(r.proveedor); return; }
      const cuitDigits = (r.cuit||'').replace(/[^0-9]/g,'');
      const modalidad = cbuDigits.indexOf('0285')===0 ? '18' : '12'; // 18: cuentas Macro, 12: otros bancos
      filas.push({
        'Modalidad de Pago': modalidad,
        'Número de documento Beneficiario': cuitDigits,
        'Denominación del Beneficiario': (r.proveedor||'').slice(0,120),
        'Concepto': 'FAC',
        'Importe': Number(r.importe)||0,
        'Cuenta Destino': cbuDigits,
        'Fecha de pago - Emisión': new Date().toLocaleDateString('es-AR')
      });
    });
    if(!filas.length){ showToast('Ninguna seleccionada tiene CBU/cuenta numérica en la Agenda (Banco Macro no acepta alias).'); return; }
    const ws = XLSX.utils.json_to_sheet(filas, {header:['Modalidad de Pago','Número de documento Beneficiario','Denominación del Beneficiario','Concepto','Importe','Cuenta Destino','Fecha de pago - Emisión']});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
    XLSX.writeFile(wb, 'TRANSF_MASIVAS_MACRO_'+todayISO()+'.xlsx');
    const total = filas.reduce((s,f)=>s+f['Importe'],0);
    try{ await runServer('logExportTransferencias', filas.length, Math.round(total)); }catch(e){}
    const incluidas = sel.filter(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      return match && (match.cbu||'').replace(/[^0-9]/g,'').length>=15;
    });
    const marcados = await marcarPdfsComoOpFinalizada(incluidas);
    if(sinCbu.length) showToast('⚠️ '+filas.length+' incluidas. '+sinCbu.length+' NO se incluyeron por no tener CBU numérico: '+sinCbu.slice(0,5).join(', ')+(sinCbu.length>5?'...':''));
    else showToast('Archivo de Transferencias Masivas generado: '+filas.length+' pago(s).'+(marcados.faltantes.length? ' ⚠️ '+marcados.faltantes.length+' PDF(s) no se encontraron para marcar.':''));
    abrirConfirmacionPagos(incluidas, 'Banco Macro');
  });

  $('#btn-transf-exportar-mp').addEventListener('click', async ()=>{
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    if(sel.length>50){ showToast('Máximo 50 pagos por archivo. Desmarcá algunas.'); return; }
    const filas = sel.map(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      const cbu = match? match.cbu : '';
      const numCorto = facturaNumeroCorto(r.numeroFactura);
      const descripcion = ('FC '+numCorto).slice(0,12);
      const mensajeFallback = ('OP - '+cbu+'- FCFactura - FC '+numCorto+' - $'+Math.round(Number(r.importe)||0));
      const mensaje = (r.nombreArchivoPdf || mensajeFallback).slice(0,200);
      return {
        'CBU/CVU/Alias/Nro cuenta': cbu,
        'Importe': Number(r.importe)||0,
        'Motivo': (match && match.motivo) ? match.motivo : 'Factura',
        'Descripción (opcional)': descripcion,
        'Email destinatario (opcional)': match? (match.email||'') : '',
        'Mensaje del email (opcional)': mensaje
      };
    });
    const ws = XLSX.utils.json_to_sheet(filas, {header:['CBU/CVU/Alias/Nro cuenta','Importe','Motivo','Descripción (opcional)','Email destinatario (opcional)','Mensaje del email (opcional)']});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Formulario');
    const total = sel.reduce((s,r)=>s+(Number(r.importe)||0),0);
    XLSX.writeFile(wb, 'TRANSF_MERCADOPAGO_'+todayISO()+'.xlsx');
    try{ await runServer('logExportTransferencias', sel.length, Math.round(total)); }catch(e){}
    const marcados = await marcarPdfsComoOpFinalizada(sel);
    showToast('Archivo generado: '+sel.length+' pago(s).'+(marcados.faltantes.length? ' ⚠️ '+marcados.faltantes.length+' PDF(s) no se encontraron para marcar.' : ''));
    abrirConfirmacionPagos(sel, 'Mercado Pago');
  });


  $('#btn-transf-exportar-provincia').addEventListener('click', async ()=>{
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    // Provincia exige UN convenio (una empresa/CUIT) por archivo, así que
    // se separa automáticamente por fundación si hay una mezcla de MDG y RVP.
    const porFundacion = {};
    sel.forEach(r=>{ (porFundacion[r.fundacion] = porFundacion[r.fundacion]||[]).push(r); });

    let huboExito = false;
    const incluidasTotal = [];
    for(const fundacion of Object.keys(porFundacion)){
      const grupo = porFundacion[fundacion];
      const pagos = grupo.map(r=>{
        const match = findAgendaMatch(r.proveedor, r.cuit);
        return {
          cbu: match? (match.cbu||'') : '',
          importe: Number(r.importe)||0,
          cuit: r.cuit||'',
          nombre: r.proveedor||'',
          referencia: ('FC '+facturaNumeroCorto(r.numeroFactura)).slice(0,15),
          _rec: r
        };
      });
      try{
        const res = await runServer('generarArchivoBancoProvincia', fundacion, pagos.map(p=>({cbu:p.cbu, importe:p.importe, cuit:p.cuit, nombre:p.nombre, referencia:p.referencia})));
        const blob = new Blob([res.contenido], {type:'text/plain;charset=utf-8'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'TRANSF_BANCO_PROVINCIA_'+res.rotulo+'_'+fundacion.toUpperCase()+'_'+todayISO()+'.txt';
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
        try{ await runServer('logExportTransferencias', res.cantidadPagos, Math.round(res.importeTotal)); }catch(e){}
        const cbusValidos = pagos.filter(p=>normalizarCuitCliente(p.cbu).length===22).map(p=>p._rec);
        const marcados = await marcarPdfsComoOpFinalizada(cbusValidos);
        incluidasTotal.push(...cbusValidos);
        huboExito = true;
        showToast('Banco Provincia ('+FUNDACION_INFO_LABEL(fundacion)+'): '+res.cantidadPagos+' pago(s) generado(s).'+(res.excluidos? ' ⚠️ '+res.excluidos+' excluido(s) por no tener CBU numérico completo.':'')+(marcados.faltantes.length? ' ⚠️ '+marcados.faltantes.length+' PDF(s) no encontrados.':''));
      }catch(err){
        showToast('Banco Provincia ('+(FUNDACION_INFO_LABEL?FUNDACION_INFO_LABEL(fundacion):fundacion)+'): '+(err.message||'error al generar el archivo.'));
      }
    }
    if(huboExito) abrirConfirmacionPagos(incluidasTotal, 'Banco Provincia');
  });
  function normalizarCuitCliente(s){ return String(s||'').replace(/[^0-9]/g,''); }
  function FUNDACION_INFO_LABEL(fk){ return fk==='mdg' ? 'MDG' : (fk==='rvp' ? 'RVP' : fk); }

  /* ---------------- CLASIFICACIÓN (proveedores + lista maestra) ---------------- */
  function renderProvClasif(){
    const q = sinTildes_($('#provclasif-buscar').value);
    const rows = proveedoresClasif.filter(p=> !q || sinTildes_(p.proveedor).includes(q) || sinTildes_(p.cuit).includes(q));
    $('#provclasif-empty').style.display = rows.length? 'none':'block';
    $('#provclasif-body').innerHTML = rows.map(p=>`<tr>
      <td class="row-actions"><button class="btn" data-provclasif-edit="${p._rowNum}">Editar</button></td>
      <td class="mono">${escapeHtml(p.cuit)}</td>
      <td>${escapeHtml(p.proveedor)}</td>
      <td>${escapeHtml(p.clasificacion)}</td>
      <td><button class="btn btn-ghost" data-provclasif-del="${p._rowNum}">Eliminar</button></td>
    </tr>`).join('');
    $$('[data-provclasif-edit]').forEach(btn=>btn.addEventListener('click', ()=>abrirEditarProvClasif(parseInt(btn.dataset.provclasifEdit,10))));
    $$('[data-provclasif-del]').forEach(btn=>btn.addEventListener('click', async ()=>{
      if(!confirm('¿Eliminar esta clasificación de proveedor?')) return;
      try{ await runServer('deleteProveedorClasificacion', parseInt(btn.dataset.provclasifDel,10)); showToast('Eliminado.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
  }
  $('#provclasif-buscar').addEventListener('input', renderProvClasif);

  function abrirEditarProvClasif(rowNum){
    const p = rowNum ? proveedoresClasif.find(x=>x._rowNum===rowNum) : null;
    const cuit = prompt('CUIT:', p? p.cuit : '');
    if(cuit===null) return;
    const proveedor = prompt('Nombre del proveedor:', p? p.proveedor : '');
    if(proveedor===null) return;
    const clasificacion = prompt('Clasificación (tal cual figura en la lista maestra, ej: HONORARIOS/Honorarios Legales):', p? p.clasificacion : '');
    if(clasificacion===null) return;
    (async ()=>{
      try{
        if(p) await runServer('updateProveedorClasificacion', rowNum, {cuit, proveedor, clasificacion});
        else await runServer('addProveedorClasificacion', cuit, proveedor, clasificacion);
        showToast('Guardado.'); refreshAll();
      }catch(err){ showToast('Error: '+(err.message||'')); }
    })();
  }
  $('#btn-provclasif-agregar').addEventListener('click', ()=>abrirEditarProvClasif(null));

  $('#btn-provclasif-importar').addEventListener('click', ()=> $('#input-provclasif-importar').click());
  $('#input-provclasif-importar').addEventListener('change', async (e)=>{
    const file = e.target.files[0]; if(!file) return;
    const statusEl = $('#provclasif-import-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo "'+file.name+'"...';
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array'});
      console.log('[Importar Proveedores] hojas encontradas en el archivo:', wb.SheetNames);
      const nombreHoja = wb.SheetNames.includes('Hoja1') ? 'Hoja1' : wb.SheetNames[0];
      const sheet = wb.Sheets[nombreHoja];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
      console.log('[Importar Proveedores] hoja usada:', nombreHoja, '| filas totales leídas:', rows.length, '| primeras 3:', rows.slice(0,3));
      const datos = rows.slice(1) // se saltea la primera fila (encabezado/leyenda del Excel)
        .map(r=>({cuit:String(r[0]||'').trim(), proveedor:String(r[1]||'').trim(), clasificacion:String(r[2]||'').trim()}))
        .filter(r=>r.cuit);
      console.log('[Importar Proveedores] filas con CUIT detectado:', datos.length, '| ejemplo:', datos[0]);
      if(!datos.length){
        statusEl.textContent = '⚠️ Se leyó el archivo ("'+nombreHoja+'", '+rows.length+' filas) pero ninguna tenía CUIT en la columna A. Abrí la consola (F12) para ver el detalle.';
        return;
      }
      statusEl.innerHTML = '<span class="spinner"></span>Importando '+datos.length+' fila(s)...';
      const res = await runServer('importProveedoresClasificacionBulk', datos);
      statusEl.textContent = `✅ Importado: ${res.agregados} nuevo(s), ${res.actualizados} actualizado(s). Total leído del Excel: ${datos.length}.`;
      showToast(`Importado: ${res.agregados} nuevo(s), ${res.actualizados} actualizado(s).`);
      await refreshAll();
      console.log('[Importar Proveedores] proveedoresClasif en memoria tras refrescar:', proveedoresClasif.length);
    }catch(err){
      console.error('[Importar Proveedores] error:', err);
      statusEl.textContent = '❌ Error: '+(err.message||'no se pudo importar.');
    }
    finally{ e.target.value=''; }
  });

  function renderClasif(){
    const q = sinTildes_($('#clasif-buscar').value);
    const rows = clasificacionesFlujo.filter(c=> !q || sinTildes_(c.clasificacion).includes(q));
    $('#clasif-empty').style.display = rows.length? 'none':'block';
    $('#clasif-body').innerHTML = rows.map(c=>`<tr>
      <td class="row-actions"><button class="btn" data-clasif-edit="${c._rowNum}">Editar</button></td>
      <td>${escapeHtml(c.clasificacion)}</td>
      <td><button class="btn btn-ghost" data-clasif-del="${c._rowNum}">Eliminar</button></td>
    </tr>`).join('');
    $$('[data-clasif-edit]').forEach(btn=>btn.addEventListener('click', ()=>abrirEditarClasif(parseInt(btn.dataset.clasifEdit,10))));
    $$('[data-clasif-del]').forEach(btn=>btn.addEventListener('click', async ()=>{
      if(!confirm('¿Eliminar esta clasificación de la lista maestra?')) return;
      try{ await runServer('deleteClasificacionFlujo', parseInt(btn.dataset.clasifDel,10)); showToast('Eliminada.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
  }
  $('#clasif-buscar').addEventListener('input', renderClasif);

  function abrirEditarClasif(rowNum){
    const c = rowNum ? clasificacionesFlujo.find(x=>x._rowNum===rowNum) : null;
    const valor = prompt('Clasificación (formato sugerido: CATEGORIA/Subcategoria):', c? c.clasificacion : '');
    if(valor===null) return;
    (async ()=>{
      try{
        if(c) await runServer('updateClasificacionFlujo', rowNum, valor);
        else await runServer('addClasificacionFlujo', valor);
        showToast('Guardado.'); refreshAll();
      }catch(err){ showToast('Error: '+(err.message||'')); }
    })();
  }
  $('#btn-clasif-agregar').addEventListener('click', ()=>abrirEditarClasif(null));

  $('#btn-clasif-importar').addEventListener('click', ()=> $('#input-clasif-importar').click());
  $('#input-clasif-importar').addEventListener('change', async (e)=>{
    const file = e.target.files[0]; if(!file) return;
    const statusEl = $('#clasif-import-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo "'+file.name+'"...';
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array'});
      console.log('[Importar Clasificaciones] hojas encontradas en el archivo:', wb.SheetNames);
      const nombreHoja = wb.SheetNames.includes('Hoja1') ? 'Hoja1' : wb.SheetNames[0];
      const sheet = wb.Sheets[nombreHoja];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
      console.log('[Importar Clasificaciones] hoja usada:', nombreHoja, '| filas totales leídas:', rows.length, '| primeras 5:', rows.slice(0,5));
      const valores = rows.map(r=>String(r[0]||'').trim()).filter(v=>v.indexOf('/')>-1);
      console.log('[Importar Clasificaciones] valores con formato Categoria/Subcategoria detectados:', valores.length, '| ejemplo:', valores[0]);
      if(!valores.length){
        statusEl.textContent = '⚠️ Se leyó el archivo ("'+nombreHoja+'", '+rows.length+' filas) pero ninguna tenía el formato Categoria/Subcategoria en la columna A. Abrí la consola (F12) para ver el detalle.';
        return;
      }
      statusEl.innerHTML = '<span class="spinner"></span>Importando '+valores.length+' clasificación(es)...';
      const res = await runServer('importClasificacionesBulk', valores);
      statusEl.textContent = `✅ Importado: ${res.agregadas} clasificación(es) nueva(s). Total leído del Excel: ${valores.length}.`;
      showToast(`Importado: ${res.agregadas} clasificación(es) nueva(s).`);
      await refreshAll();
      console.log('[Importar Clasificaciones] clasificacionesFlujo en memoria tras refrescar:', clasificacionesFlujo.length);
    }catch(err){
      console.error('[Importar Clasificaciones] error:', err);
      statusEl.textContent = '❌ Error: '+(err.message||'no se pudo importar.');
    }
    finally{ e.target.value=''; }
  });

  /* ---------------- MEDIOS DE PAGO (lista maestra) ---------------- */
  function renderMediosPago(){
    const q = sinTildes_($('#mediospago-buscar').value);
    const rows = mediosPago.filter(m=> !q || sinTildes_(m.medioDePago).includes(q));
    $('#mediospago-empty').style.display = rows.length? 'none':'block';
    $('#mediospago-body').innerHTML = rows.map(m=>`<tr>
      <td class="row-actions"><button class="btn" data-mediospago-edit="${m._rowNum}">Editar</button></td>
      <td>${escapeHtml(m.medioDePago)}</td>
      <td><button class="btn btn-ghost" data-mediospago-del="${m._rowNum}">Eliminar</button></td>
    </tr>`).join('');
    $$('[data-mediospago-edit]').forEach(btn=>btn.addEventListener('click', ()=>abrirEditarMedioPago(parseInt(btn.dataset.mediospagoEdit,10))));
    $$('[data-mediospago-del]').forEach(btn=>btn.addEventListener('click', async ()=>{
      if(!confirm('¿Eliminar este medio de pago?')) return;
      try{ await runServer('deleteMedioPago', parseInt(btn.dataset.mediospagoDel,10)); showToast('Eliminado.'); refreshAll(); }
      catch(err){ showToast('Error: '+(err.message||'')); }
    }));
  }
  $('#mediospago-buscar').addEventListener('input', renderMediosPago);

  function abrirEditarMedioPago(rowNum){
    const m = rowNum ? mediosPago.find(x=>x._rowNum===rowNum) : null;
    const valor = prompt('Medio de Pago:', m? m.medioDePago : '');
    if(valor===null) return;
    (async ()=>{
      try{
        if(m) await runServer('updateMedioPago', rowNum, valor);
        else await runServer('addMedioPago', valor);
        showToast('Guardado.'); refreshAll();
      }catch(err){ showToast('Error: '+(err.message||'')); }
    })();
  }
  $('#btn-mediospago-agregar').addEventListener('click', ()=>abrirEditarMedioPago(null));

  $('#btn-mediospago-importar').addEventListener('click', ()=> $('#input-mediospago-importar').click());
  $('#input-mediospago-importar').addEventListener('change', async (e)=>{
    const file = e.target.files[0]; if(!file) return;
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
      const valores = rows.map(r=>String(r[0]||'').trim()).filter(v=>v);
      if(!valores.length){ showToast('No se encontraron medios de pago para importar (se espera un valor en la columna A).'); return; }
      const res = await runServer('importMediosPagoBulk', valores);
      showToast(`Importado: ${res.agregados} medio(s) de pago nuevo(s).`);
      refreshAll();
    }catch(err){ showToast('Error al importar: '+(err.message||'')); }
    finally{ e.target.value=''; }
  });

  /* ---------------- VERIFICAR ARCA ---------------- */
  function formatearPtoVtaNumero_(pv, num){
    const pvStr = String(pv||'').padStart(4,'0');
    const numStr = String(num||'').padStart(8,'0');
    return pvStr+'-'+numStr;
  }
  function mapearTipoArca_(tipoTexto){
    const t = String(tipoTexto||'').toUpperCase();
    if(t.includes('FACTURA A')) return 'A';
    if(t.includes('FACTURA B')) return 'B';
    if(t.includes('FACTURA C')) return 'C';
    if(t.includes('NOTA DE CREDITO') || t.includes('NOTA DE CRÉDITO')) return 'NC';
    return '';
  }
  function renderArca(){
    const fund = $('#arca-import-fundacion').value;
    const rows = fund ? arcaComprobantes.filter(a=>a.fundacion===fund) : arcaComprobantes;
    $('#arca-empty').style.display = rows.length? 'none':'block';
    $('#arca-body').innerHTML = rows.map(a=>`<tr>
      <td>${(a.fundacion||'').toUpperCase()}</td>
      <td class="mono">${fmtDateShort(a.fecha)}</td>
      <td>${escapeHtml(a.tipo)}</td>
      <td class="mono">${formatearPtoVtaNumero_(a.puntoVenta,a.numeroDesde)}</td>
      <td class="mono">${escapeHtml(a.cae)}</td>
      <td>${escapeHtml(a.denominacionEmisor)}</td>
      <td class="mono">${escapeHtml(a.cuitEmisor)}</td>
      <td class="mono">${fmtMoney(a.importe)}</td>
    </tr>`).join('');
    if(window.__arcaDebug){
      const d = window.__arcaDebug;
      $('#arca-debug-panel').textContent = 'Diagnóstico ('+d.hora+'): MDG → '+(d.mdg.ok? d.mdg.cantidad+' fila(s)' : 'ERROR: '+d.mdg.error)+' | RVP → '+(d.rvp.ok? d.rvp.cantidad+' fila(s)' : 'ERROR: '+d.rvp.error)+' | Total en memoria del navegador: '+d.totalEnMemoria+' | arcaComprobantes.length ahora: '+arcaComprobantes.length;
    }
  }
  $('#arca-import-fundacion').addEventListener('change', renderArca);

  $('#btn-arca-importar').addEventListener('click', ()=> $('#input-arca-importar').click());
  $('#input-arca-importar').addEventListener('change', async (e)=>{
    const file = e.target.files[0]; if(!file) return;
    if(!cargaInicialCompleta){ showToast('El panel todavía está terminando de cargar — esperá unos segundos y probá de nuevo.'); e.target.value=''; return; }
    const fundacion = $('#arca-import-fundacion').value;
    if(!fundacion){ showToast('Elegí primero para qué fundación es este archivo.'); e.target.value=''; return; }
    const statusEl = $('#arca-import-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo archivo...';
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
      const datos = rows.slice(2).map(r=>{
        const fechaRaw = String(r[0]||'').trim();
        let fechaIso = '';
        const partes = fechaRaw.split('/');
        if(partes.length===3) fechaIso = partes[2]+'-'+String(partes[1]).padStart(2,'0')+'-'+String(partes[0]).padStart(2,'0');
        return {
          fecha: fechaIso,
          tipo: String(r[1]||'').trim(),
          puntoVenta: String(r[2]||'').trim(),
          numeroDesde: String(r[3]||'').trim(),
          numeroHasta: String(r[4]||'').trim(),
          cae: String(r[5]||'').trim(),
          cuitEmisor: String(r[7]||'').trim(),
          denominacionEmisor: String(r[8]||'').trim(),
          cuitReceptor: String(r[10]||'').trim(),
          importeTotal: Number(r[29])||0
        };
      }).filter(d=>d.cae);
      if(!datos.length){ statusEl.textContent = '⚠️ No se encontraron filas válidas en el archivo (¿es el formato correcto de ARCA?).'; return; }
      statusEl.innerHTML = '<span class="spinner"></span>Importando '+datos.length+' comprobante(s)...';
      const resR = await callSeguro('importArcaComprobantes', fundacion, datos);
      if(!resR.ok) throw resR.error;
      const res = resR.value;
      statusEl.textContent = `✅ Importados: ${res.importados} nuevo(s) · Ya existían: ${res.duplicados} · Total en el archivo: ${res.total}.`;
      showToast('Importados '+res.importados+' comprobante(s) nuevo(s).');
      await cargarArca_(true);
    }catch(err){
      statusEl.textContent = '❌ Error: '+(err.message||'no se pudo importar.');
    }
    finally{ e.target.value=''; }
  });

  let arcaSinCoincidir = [];
  function compararArcaLocal_(fund){
    const enRango = (fecha) => {
      if(!fecha) return !arcaControlFechaDesde && !arcaControlFechaHasta ? true : false;
      if(arcaControlFechaDesde && fecha < arcaControlFechaDesde) return false;
      if(arcaControlFechaHasta && fecha > arcaControlFechaHasta) return false;
      return true;
    };
    const candidatasPorFundacion = { mdg: [], rvp: [] };
    records.forEach(r=>{
      if(r.tipo==='factura' && r.fundacion && r.estado!=='eliminado'){
        candidatasPorFundacion[r.fundacion].push(r);
      }
    });
    let arcaFiltradas = fund ? arcaComprobantes.filter(a=>a.fundacion===fund) : arcaComprobantes;
    arcaFiltradas = arcaFiltradas.filter(a=>enRango(a.fecha));

    // Dirección 1: comprobantes ARCA que no tienen una factura correspondiente
    let coincidencias = 0;
    const sinCoincidir = [];
    arcaFiltradas.forEach(a=>{
      const candidatas = candidatasPorFundacion[a.fundacion] || [];
      const cuitArca = normCuit_(a.cuitEmisor), caeArca = String(a.cae||'').trim();
      const hayMatch = candidatas.some(f => normCuit_(f.cuit)===cuitArca && String(f.cae||'').trim()===caeArca && Math.abs((Number(f.importe)||0) - Number(a.importe)) < 1);
      if(hayMatch) coincidencias++;
      else sinCoincidir.push(a);
    });

    // Dirección 2: facturas del sistema que no tienen un comprobante ARCA
    // correspondiente. Solo se revisan las fundaciones para las que ya se
    // importó al menos un comprobante (si nunca se importó nada para una
    // fundación, no tiene sentido marcar TODAS sus facturas como "sin ARCA").
    const fundacionesConArca = new Set(arcaComprobantes.map(a=>a.fundacion));
    const facturasSinArcaLista = [];
    Object.keys(candidatasPorFundacion).forEach(fk=>{
      if(fund && fk!==fund) return;
      if(!fundacionesConArca.has(fk)) return;
      const arcaDeEstaFundacion = arcaComprobantes.filter(a=>a.fundacion===fk);
      candidatasPorFundacion[fk].filter(r=>enRango(r.fechaRecibida)).forEach(r=>{
        const cuitFactura = normCuit_(r.cuit), caeFactura = String(r.cae||'').trim();
        const hayMatch = arcaDeEstaFundacion.some(a => normCuit_(a.cuitEmisor)===cuitFactura && String(a.cae||'').trim()===caeFactura && Math.abs(Number(a.importe) - (Number(r.importe)||0)) < 1);
        if(!hayMatch) facturasSinArcaLista.push(r);
      });
    });

    return { totalArca: arcaFiltradas.length, coincidencias, sinCoincidir, facturasSinArca: facturasSinArcaLista };
  }

  let facturasSinArca = [];
  let vistaResultadoArca = 'enArca';
  function mostrarVistaResultadoArca(vista){
    vistaResultadoArca = vista;
    $('#arca-detalle-en-arca').style.display = vista==='enArca' ? 'block' : 'none';
    $('#arca-detalle-en-listado').style.display = vista==='enListado' ? 'block' : 'none';
    $('#arca-card-en-arca').classList.toggle('metric-card-activa', vista==='enArca');
    $('#arca-card-en-listado').classList.toggle('metric-card-activa', vista==='enListado');
  }
  $('#arca-card-en-arca').addEventListener('click', ()=>mostrarVistaResultadoArca('enArca'));
  $('#arca-card-en-listado').addEventListener('click', ()=>mostrarVistaResultadoArca('enListado'));

  function renderArcaSinCoincidirTabla(){
    const q = sinTildes_($('#arca-sincoincidir-buscar').value);
    const filas = q ? arcaSinCoincidir.filter(a=> sinTildes_(a.cuitEmisor).includes(q) || sinTildes_(a.denominacionEmisor).includes(q)) : arcaSinCoincidir;
    $('#arca-sincoincidir-body').innerHTML = filas.map(a=>{
      const i = arcaSinCoincidir.indexOf(a);
      return `<tr>
        <td style="text-align:center"><input type="checkbox" class="arca-sincoincidir-check" data-idx="${i}" checked></td>
        <td>${(a.fundacion||'').toUpperCase()}</td>
        <td class="mono">${fmtDateShort(a.fecha)}</td>
        <td>${escapeHtml(a.tipo)}</td>
        <td class="mono">${formatearPtoVtaNumero_(a.puntoVenta,a.numeroDesde)}</td>
        <td class="mono">${escapeHtml(a.cae)}</td>
        <td title="${escapeHtml(a.denominacionEmisor)}">${escapeHtml(a.denominacionEmisor)}</td>
        <td class="mono">${escapeHtml(a.cuitEmisor)}</td>
        <td class="mono">${fmtMoney(a.importe)}</td>
      </tr>`;
    }).join('');
    $('#arca-sincoincidir-check-todas').checked = true;
    $$('.arca-sincoincidir-check').forEach(chk=>chk.addEventListener('change', ()=>{
      const todas = $$('.arca-sincoincidir-check');
      $('#arca-sincoincidir-check-todas').checked = todas.every(c=>c.checked);
    }));
  }
  $('#arca-sincoincidir-buscar').addEventListener('input', renderArcaSinCoincidirTabla);

  function renderListadoSinArcaTabla(){
    const q = sinTildes_($('#listado-sinarca-buscar').value);
    const filas = q ? facturasSinArca.filter(r=> sinTildes_(r.cuit).includes(q) || sinTildes_(r.proveedor).includes(q)) : facturasSinArca;
    $('#listado-sinarca-body').innerHTML = filas.map(r=>`<tr>
      <td>${(r.fundacion||'').toUpperCase()}</td>
      <td title="${escapeHtml(r.proveedor)}">${escapeHtml(r.proveedor)}</td>
      <td class="mono">${escapeHtml(r.cuit)}</td>
      <td class="mono">${escapeHtml(r.numeroFactura)}</td>
      <td class="mono">${escapeHtml(r.cae)}</td>
      <td class="mono">${fmtMoney(r.importe)}</td>
      <td class="mono">${fmtDateShort(r.fechaRecibida)}</td>
      <td>${r._key? `<button class="btn" data-listadosinarca-edit="${r._key}">Editar</button>` : ''}</td>
    </tr>`).join('');
    $$('[data-listadosinarca-edit]').forEach(btn=>btn.addEventListener('click', ()=>openEditModal(btn.dataset.listadosinarcaEdit)));
  }
  $('#listado-sinarca-buscar').addEventListener('input', renderListadoSinArcaTabla);

  $('#btn-arca-controlar').addEventListener('click', ()=>{
    if(!cargaInicialCompleta){ showToast('El panel todavía está terminando de cargar — esperá unos segundos y probá de nuevo.'); return; }
    const fund = $('#arca-import-fundacion').value;
    try{
      const resultado = compararArcaLocal_(fund);
      arcaSinCoincidir = resultado.sinCoincidir;
      facturasSinArca = resultado.facturasSinArca;

      $('#arca-card-en-arca-num').textContent = arcaSinCoincidir.length;
      $('#arca-card-en-listado-num').textContent = facturasSinArca.length;
      $('#arca-resumen').innerHTML = `Total comprobantes ARCA revisados: <b>${resultado.totalArca}</b> · Coinciden: <b style="color:var(--baja)">${resultado.coincidencias}</b>. ${(arcaSinCoincidir.length===0 && facturasSinArca.length===0) ? '✅ 0 diferencias encontradas — todo coincide.' : 'Hacé clic en cualquiera de los dos cuadros para ver el detalle.'}`;

      $('#arca-sincoincidir-buscar').value = '';
      renderArcaSinCoincidirTabla();
      $('#listado-sinarca-buscar').value = '';
      renderListadoSinArcaTabla();

      mostrarVistaResultadoArca(arcaSinCoincidir.length>=facturasSinArca.length ? 'enArca' : 'enListado');
      $('#arca-resultado-panel').style.display = 'block';
    }catch(err){
      showToast('Error al comparar: '+(err.message||''));
      console.error(err);
    }
  });

  $('#btn-listado-sinarca-exportar').addEventListener('click', ()=>{
    if(!facturasSinArca.length){ showToast('No hay filas para exportar.'); return; }
    const headers = ['Fundación','Proveedor','CUIT','N° Factura','CAE','Importe','Fecha Recibida'];
    const datos = facturasSinArca.map(r=>({
      'Fundación': (r.fundacion||'').toUpperCase(), 'Proveedor': r.proveedor||'', 'CUIT': r.cuit||'',
      'N° Factura': r.numeroFactura||'', 'CAE': r.cae||'', 'Importe': Number(r.importe)||0, 'Fecha Recibida': r.fechaRecibida||''
    }));
    const ws = XLSX.utils.json_to_sheet(datos, {header:headers});
    ws['!autofilter'] = {ref: ws['!ref']};
    ws['!cols'] = headers.map(h=>({wch:Math.max(12,h.length+2)}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sin ARCA');
    XLSX.writeFile(wb, 'Facturas_SinARCA_'+todayISO()+'.xlsx');
    showToast('Exportado: '+datos.length+' fila(s).');
  });

  $('#arca-sincoincidir-check-todas').addEventListener('change', (e)=>{
    $$('.arca-sincoincidir-check').forEach(chk=>chk.checked = e.target.checked);
  });

  $('#btn-arca-exportar-sincoincidir').addEventListener('click', ()=>{
    if(!arcaSinCoincidir.length){ showToast('No hay filas para exportar.'); return; }
    const headers = ['Fundación','Fecha','Tipo','Punto de Venta','Número','CAE','CUIT Emisor','Denominación Emisor','CUIT Receptor','Importe'];
    const datos = arcaSinCoincidir.map(a=>({
      'Fundación': (a.fundacion||'').toUpperCase(), 'Fecha': a.fecha||'', 'Tipo': a.tipo||'',
      'Punto de Venta': a.puntoVenta||'', 'Número': a.numeroDesde||'', 'CAE': a.cae||'',
      'CUIT Emisor': a.cuitEmisor||'', 'Denominación Emisor': a.denominacionEmisor||'',
      'CUIT Receptor': a.cuitReceptor||'', 'Importe': Number(a.importe)||0
    }));
    const ws = XLSX.utils.json_to_sheet(datos, {header:headers});
    ws['!autofilter'] = {ref: ws['!ref']};
    ws['!cols'] = headers.map(h=>({wch:Math.max(12,h.length+2)}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sin Coincidir');
    XLSX.writeFile(wb, 'ARCA_SinCoincidir_'+todayISO()+'.xlsx');
    showToast('Exportado: '+datos.length+' fila(s).');
  });

  let arcaIncorporarLista = [], arcaIncorporarDatos = [], arcaIncorporarIdx = 0, arcaIncPrio = 'media';
  $('#btn-arca-incorporar-seleccionadas').addEventListener('click', ()=>{
    const seleccionadas = [];
    $$('.arca-sincoincidir-check').forEach(chk=>{ if(chk.checked) seleccionadas.push(arcaSinCoincidir[parseInt(chk.dataset.idx,10)]); });
    if(!seleccionadas.length){ showToast('Tildá al menos un comprobante para incorporar.'); return; }
    arcaIncorporarLista = seleccionadas;
    arcaIncorporarDatos = seleccionadas.map(a=>{
      const match = proveedoresClasif.find(p=>normCuit_(p.cuit)===normCuit_(a.cuitEmisor));
      return { centroCosto:'', clasificacionFlujo: match? (match.clasificacion||'') : '', prioridad:'media', descripcion:'', pagada:true, medioDePago:'' };
    });
    arcaIncorporarIdx = 0;
    abrirArcaIncorporarPaso();
  });
  function abrirArcaIncorporarPaso(){
    const a = arcaIncorporarLista[arcaIncorporarIdx];
    const d = arcaIncorporarDatos[arcaIncorporarIdx];
    $('#arca-incorporar-contador').textContent = (arcaIncorporarIdx+1)+' de '+arcaIncorporarLista.length;
    $('#arca-incorporar-info').innerHTML = `<b>${escapeHtml(a.denominacionEmisor)}</b><br>CUIT: ${escapeHtml(a.cuitEmisor)} · Fundación: ${(a.fundacion||'').toUpperCase()}<br>N° Comprobante: ${formatearPtoVtaNumero_(a.puntoVenta,a.numeroDesde)} (${escapeHtml(a.tipo)}) · Fecha: ${fmtDateShort(a.fecha)}<br>CAE: ${escapeHtml(a.cae)} · Importe: <b>${fmtMoney(a.importe)}</b>`;
    $('#arca-inc-centrocosto').value = d.centroCosto;
    $('#arca-inc-clasificacion').value = d.clasificacionFlujo;
    $('#arca-inc-descripcion').value = d.descripcion;
    $('#arca-inc-pagada').checked = d.pagada;
    $('#arca-inc-medio-pago').value = d.medioDePago;
    arcaIncPrio = d.prioridad;
    $$('#arca-inc-prio-select .prio-opt').forEach(o=>o.classList.toggle('selected', o.dataset.p===arcaIncPrio));
    $('#arca-inc-anterior').disabled = arcaIncorporarIdx===0;
    $('#arca-inc-siguiente').textContent = arcaIncorporarIdx===arcaIncorporarLista.length-1 ? 'Última ✓' : 'Siguiente →';
    $('#arca-inc-estado-actual').textContent = '';
    $('#arca-incorporar-overlay').classList.remove('hidden');
  }
  function guardarPasoArcaActual_(){
    arcaIncorporarDatos[arcaIncorporarIdx] = {
      centroCosto: $('#arca-inc-centrocosto').value,
      clasificacionFlujo: $('#arca-inc-clasificacion').value,
      prioridad: arcaIncPrio,
      descripcion: $('#arca-inc-descripcion').value,
      pagada: $('#arca-inc-pagada').checked,
      medioDePago: $('#arca-inc-medio-pago').value
    };
  }
  $$('#arca-inc-prio-select .prio-opt').forEach(o=>o.addEventListener('click', ()=>{
    arcaIncPrio = o.dataset.p;
    $$('#arca-inc-prio-select .prio-opt').forEach(x=>x.classList.toggle('selected', x===o));
  }));
  // (se conectó arriba con el modal grande, no con wireComboBuscable)
  $('#arca-inc-anterior').addEventListener('click', ()=>{
    if(arcaIncorporarIdx===0) return;
    guardarPasoArcaActual_();
    arcaIncorporarIdx--;
    abrirArcaIncorporarPaso();
  });
  $('#arca-inc-siguiente').addEventListener('click', ()=>{
    guardarPasoArcaActual_();
    if(arcaIncorporarIdx < arcaIncorporarLista.length-1){
      arcaIncorporarIdx++;
      abrirArcaIncorporarPaso();
    }
  });
  $('#arca-incorporar-close').addEventListener('click', ()=>{ guardarPasoArcaActual_(); $('#arca-incorporar-overlay').classList.add('hidden'); });
  $('#arca-inc-cancelar-todo').addEventListener('click', ()=>{
    if(confirm('¿Cancelar? No se va a crear ninguna factura con lo que cargaste en este repaso.')) $('#arca-incorporar-overlay').classList.add('hidden');
  });
  $('#arca-inc-confirmar-final').addEventListener('click', async ()=>{
    guardarPasoArcaActual_();
    const faltantes = [];
    arcaIncorporarDatos.forEach((d,i)=>{
      if(!d.centroCosto || !d.clasificacionFlujo) faltantes.push('#'+(i+1)+' ('+arcaIncorporarLista[i].denominacionEmisor+') — falta Centro de Costo/Clasificación');
      if(d.pagada && !d.medioDePago) faltantes.push('#'+(i+1)+' ('+arcaIncorporarLista[i].denominacionEmisor+') — marcada como pagada pero sin Medio de Pago');
    });
    if(faltantes.length){
      alert('Faltan completar Centro de Costo y/o Clasificación en:\n\n• '+faltantes.join('\n• '));
      return;
    }
    $('#arca-inc-confirmar-final').disabled = true;
    let ok=0, error=0;
    for(let i=0;i<arcaIncorporarLista.length;i++){
      const a = arcaIncorporarLista[i], d = arcaIncorporarDatos[i];
      const rec = {
        tipo:'factura', fundacion:a.fundacion, proveedor:a.denominacionEmisor, cuit:a.cuitEmisor,
        tipoFactura: mapearTipoArca_(a.tipo), numeroFactura: formatearPtoVtaNumero_(a.puntoVenta,a.numeroDesde), cae: a.cae,
        fechaEmision: a.fecha, nombreFantasia:'', fechaVencimiento:'',
        importe: a.importe, descripcion: d.descripcion || 'Incorporada desde ARCA',
        prioridad: d.prioridad, canal:'presencial',
        centroCosto: d.centroCosto, clasificacionFlujo: d.clasificacionFlujo,
        fileBase64:'', fileMediaType:'',
        pagado: !!d.pagada,
        medioDePago: d.pagada ? (d.medioDePago||'') : '',
        estadoInicial: d.pagada ? 'completado' : 'pendiente',
        fechaRecibida: a.fecha || todayISO(),
        fechaResuelta: d.pagada ? todayISO() : ''
      };
      try{ await runServer('addRecord', rec); ok++; }
      catch(err){ error++; console.error('Error incorporando comprobante ARCA', a, err); }
    }
    $('#arca-inc-confirmar-final').disabled = false;
    $('#arca-incorporar-overlay').classList.add('hidden');
    showToast('Incorporadas: '+ok+(error? ' · '+error+' con error':'')+'.');
    await refreshAll();
    $('#btn-arca-controlar').click();
  });

  /* ---------------- VOLVER ARRIBA ---------------- */
  const btnVolverArriba = $('#btn-volver-arriba');
  window.addEventListener('scroll', ()=>{
    btnVolverArriba.style.display = window.scrollY > 300 ? 'block' : 'none';
  });
  btnVolverArriba.addEventListener('click', ()=>{
    window.scrollTo({top:0, behavior:'smooth'});
  });

  /* ---------------- INIT ---------------- */
  $('#today-chip').textContent = new Date().toLocaleDateString('es-AR',{weekday:'long',day:'2-digit',month:'long',year:'numeric'});
  refreshAll();
})();
