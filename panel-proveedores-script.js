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
  function fmtDateShort(iso){ if(!iso) return '—'; const m=/^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso)); return m ? (m[3]+'/'+m[2]+'/'+m[1].slice(2)) : String(iso); }
  function fmtMoney(n){ n=Number(n)||0; return n.toLocaleString('es-AR',{style:'currency',currency:'ARS',maximumFractionDigits:0}); }
  function daysBetween(a,b){ const da=new Date(a+'T00:00:00'), db=new Date(b+'T00:00:00'); return Math.round((db-da)/86400000); }
  function escapeHtml(s){ return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function showToast(msg){ const t=$('#toast'); t.textContent=msg; t.classList.remove('hidden'); clearTimeout(showToast._h); showToast._h=setTimeout(()=>t.classList.add('hidden'),3200); }

  /* ---------- Utilidades de seguridad / idempotencia (Lote 5) ---------- */
  // Identificador único por operación: si la misma operación se reintenta (corte de red, doble clic)
  // el servidor reconoce el identificador y no la repite.
  function nuevoRequestId_(){
    try{
      if(window.crypto && crypto.randomUUID) return crypto.randomUUID();
      const a = new Uint32Array(4); crypto.getRandomValues(a);
      return Array.from(a).map(n=>n.toString(16)).join('-');
    }catch(e){ return 'r'+Date.now().toString(36)+Math.random().toString(36).slice(2,10); }
  }
  // Mismo dato = mismo identificador (reintento); datos distintos = identificador nuevo.
  const requestIdsPorClave_ = {};
  function requestIdParaDatos_(clave, datos){
    const firma = JSON.stringify(datos, (k,v)=> (typeof v==='string' && v.length>200) ? (v.length+':'+v.slice(0,40)+':'+v.slice(-40)) : v);
    const previo = requestIdsPorClave_[clave];
    if(previo && previo.firma===firma) return previo.id;
    const id = nuevoRequestId_();
    requestIdsPorClave_[clave] = { firma, id };
    return id;
  }
  // Valores que van a una celda de Excel: se saca el signo inicial que Excel interpretaría como fórmula.
  function textoXlsx_(s){ return String(s==null?'':s).replace(/^[=+\-@\t\r\n]+/, ''); }
  function fechaHoyAR_(){
    const d = new Date();
    return String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
  }
  // Los datos vienen de una planilla que cualquiera puede editar: lo que se usa como "tipo", "estado", etc.
  // se limita a los valores conocidos antes de armar pantallas con ellos.
  function sanearRegistro_(r){
    if(!r || typeof r!=='object') return r;
    r.tipo = Object.prototype.hasOwnProperty.call(TIPO_LABEL, r.tipo) ? r.tipo : 'consulta';
    r.estado = Object.prototype.hasOwnProperty.call(ESTADO_LABEL, r.estado) ? r.estado : 'pendiente';
    r.prioridad = Object.prototype.hasOwnProperty.call(PRIO_ORDER, r.prioridad) ? r.prioridad : 'media';
    if(r.fuente && !Object.prototype.hasOwnProperty.call(FUENTE_LABEL, r.fuente)) r.fuente = '';
    if(r.canal && !Object.prototype.hasOwnProperty.call(FUENTE_LABEL, r.canal)) r.canal = '';
    r.fundacion = (r.fundacion==='mdg' || r.fundacion==='rvp') ? r.fundacion : (r.fundacion ? '' : r.fundacion);
    if(r._rowNum!=null) r._rowNum = Number(r._rowNum) || 0;
    return r;
  }
  function claveSegura_(s){ return String(s==null?'':s).replace(/[^A-Za-z0-9_:.\-]/g,'_'); }

  /* ---------------- SERVER BRIDGE ---------------- */
  // Tiempo máximo de espera por función (ms). Las que escriben en la planilla pueden esperar el turno del
  // servidor hasta 30 s, por eso tienen más margen que una consulta.
  const LIMITES_MS_ = {
    processTransferenciasFolder: 330000, processComprasPresupuestoFolder: 330000,
    prepararExportacion: 90000, confirmarExportacion: 150000, cancelarExportacion: 60000,
    generarArchivoProvinciaDeExportacion: 60000,
    addRecord: 90000, registrarPagoParcial: 60000, registrarPagoLote: 60000,
    importArcaComprobantes: 150000, importAgendaBulk: 120000
  };
  // Funciones que cambian datos: si se agota la espera, el cambio puede haberse hecho igual.
  const FUNCIONES_QUE_ESCRIBEN_ = /^(add|save|update|delete|remove|import|registrar|confirmar|cancelar|preparar|marcar|generar|autorizar|toggle|set|crear|guardar|editar|eliminar)/i;
  function runServer(fn, ...args){
    const llamada = new Promise((resolve, reject)=>{
      google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[fn](...args);
    });
    const limiteMs = LIMITES_MS_[fn] || 25000;
    let timer;
    const timeout = new Promise((_, reject)=>{
      timer = setTimeout(()=>{
        const aviso = FUNCIONES_QUE_ESCRIBEN_.test(fn) ? ' Es posible que el cambio SÍ se haya guardado: actualizá el panel (F5) y verificá antes de repetirlo.' : '';
        reject(new Error('Se agotó el tiempo de espera ('+(limiteMs/1000)+'s) para '+fn+'.'+aviso));
      }, limiteMs);
    });
    return Promise.race([llamada, timeout]).finally(()=>clearTimeout(timer));
  }

  const CLIENT_VERSION = 'v2026-11-26-lote11-a'; // MDG/RVP/Otros se traen en tandas de 50 en vez de todas juntas (evita el fallo de comunicación con hojas grandes)
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

  /* ---- Centro de Costo y Clasificación de flujo: obligatorios para avanzar ---- */
  function cartelBloqueo_(motivo){
    alert('⛔ NO SE PUEDE AVANZAR\n\n'+motivo+'\n\nCompletá esos datos y volvé a intentar. La operación NO se grabó.');
  }
  function faltaClasif_(centro, clasif){
    const f=[]; if(!String(centro||'').trim()) f.push('Centro de Costo'); if(!String(clasif||'').trim()) f.push('Clasificación de flujo'); return f;
  }
  function esBloqueoServidor_(err){ return /NO SE PUEDE AVANZAR/.test((err&&err.message)||''); }
  function avisarBloqueoServidor_(err){ alert('⛔ '+String(err.message||'').replace(/^NO SE PUEDE AVANZAR:?\s*/,'NO SE PUEDE AVANZAR\n\n')+'\n\nLa operación NO se grabó.'); }
  function sinClasificar_(r){
    return !!r.isAppRecord && r.estado!=='eliminado' && r.estado!=='completado' && !r.pagado && faltaClasif_(r.centroCosto, r.clasificacionFlujo).length>0;
  }
  function sinCuit_(r){ return r.estado!=='eliminado' && !String(r.cuit||'').replace(/[^0-9]/g,''); }
  function renderSinCuit_(){
    const vista = $('#view-listado'); if(!vista) return;
    let el = $('#sin-cuit-banner');
    if(!el){
      el = document.createElement('div'); el.id='sin-cuit-banner';
      el.style.cssText='display:none;margin:0 0 12px;padding:10px 12px;border:1px solid var(--media);background:var(--media-soft);border-radius:8px;font-size:13px;';
      vista.insertBefore(el, vista.firstChild);
      el.addEventListener('click',(e)=>{ if(!e.target.closest('button')) return; const sel=$('#listado-filt-cuit'); sel.value = sel.value==='__sin__' ? '' : '__sin__'; renderSinCuit_(); renderListado(); });
    }
    const n = records.filter(sinCuit_).length;
    const activo = $('#listado-filt-cuit').value==='__sin__';
    if(!n && !activo){ el.style.display='none'; el.innerHTML=''; return; }
    el.style.display='block';
    el.innerHTML = '<b>⚠ '+n+' registro(s) SIN CUIT.</b> Sin CUIT no se pueden identificar ni pagar por transferencia con datos del proveedor. '+
      '<button type="button" class="btn" style="margin-left:8px;">'+(activo?'Ver todos':'Ver solo esos')+'</button>';
  }
  let soloSinClasificar_ = false;
  function renderSinClasificar_(){
    const vista = $('#view-listado'); if(!vista) return;
    let el = $('#sin-clasif-banner');
    if(!el){
      el = document.createElement('div'); el.id='sin-clasif-banner';
      el.style.cssText='display:none;margin:0 0 12px;padding:10px 12px;border:1px solid var(--alta);background:var(--alta-soft);border-radius:8px;font-size:13px;';
      vista.insertBefore(el, vista.firstChild);
      el.addEventListener('click',(e)=>{ if(!e.target.closest('button')) return; soloSinClasificar_ = !soloSinClasificar_; renderSinClasificar_(); renderListado(); });
    }
    const n = records.filter(sinClasificar_).length;
    if(!n && !soloSinClasificar_){ el.style.display='none'; el.innerHTML=''; return; }
    el.style.display='block';
    el.innerHTML = '<b>⚠ '+n+' registro(s) sin Centro de Costo o Clasificación de flujo.</b> No se pueden pagar hasta completarlos (botón Editar). '+
      '<button type="button" class="btn" style="margin-left:8px;">'+(soloSinClasificar_?'Ver todos':'Ver solo esos')+'</button>';
  }

  // Una sola recarga a la vez: si piden otra mientras corre, se repite al terminar (evita lecturas superpuestas).
  let refreshEnCurso_ = null, refreshRepetir_ = false;
  function refreshAll(){
    if(refreshEnCurso_){ refreshRepetir_ = true; return refreshEnCurso_; }
    refreshEnCurso_ = (async ()=>{
      try{ do{ refreshRepetir_ = false; await refreshAll_(); }while(refreshRepetir_); }
      finally{ refreshEnCurso_ = null; }
    })();
    return refreshEnCurso_;
  }
  async function refreshAll_(){
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
      $('#btn-backups').style.display = soyAdmin ? '' : 'none';
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
    records.forEach(sanearRegistro_);
    records.forEach(r=>{ r._key = r.fundacion ? ((r.tipo==='factura' ? 'f:' : 'o:')+r.fundacion+':'+r._rowNum) : claveSegura_(r.id); });  // 'f'/'o': una compra de Otros MDG y una factura de MDG pueden tener la misma fila
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
    cargarExportacionesPendientes_();

    if(fallidas.length){
      console.error('Fallaron al cargar:', fallidas);
      showToast('⚠️ No se pudo cargar: '+fallidas.join(', '));
    }
  }

  function renderAll(){ renderCounters(); renderKanban(); renderDashboard(); renderListado(); renderNotas(); renderNotasTabla(); renderAgenda(); renderTransferencias(); renderLog(); renderProvClasif(); renderClasif(); renderArca(); renderMediosPago(); renderSinClasificar_(); renderSinCuit_(); }

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
      if(extracted.duplicado){
        const d = extracted.duplicado;
        alert('⚠ COMPROBANTE DUPLICADO\n\nEste comprobante (N° '+(extracted.numeroComprobante||'')+') ya está cargado en "'+d.hoja+'" (fila '+d.fila+'):\n• '+(d.proveedor||'—')+' · '+fmtMoney(d.importe)+(d.fecha?' · '+d.fecha:'')+'\n\nNo se vuelve a cargar.');
        statusEl.textContent = 'Duplicado: ya estaba cargado.';
        return;
      }
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
      if(!$('#f-cuit').value.replace(/[^0-9]/g,'')){ cartelBloqueo_('Falta completar: CUIT.\nUna factura no se puede guardar sin CUIT.'); return; }
      if(!$('#f-tipofactura').value) faltantes.push('Tipo de comprobante');
      if(!$('#f-numfactura').value.trim()) faltantes.push('N° de factura');
      if($('#f-tipofactura').value!=='Boleta' && !$('#f-cae').value.trim()) faltantes.push('CAE');
      if($('#f-tipofactura').value==='Boleta' && !$('#f-vencimiento').value) faltantes.push('Fecha de vencimiento (Boleta)');
    }
    const faltClasifAlta = faltaClasif_(centroCosto, clasificacionFlujo);
    if(faltClasifAlta.length){ cartelBloqueo_('Falta completar: '+faltClasifAlta.join(' y ')+'.\nSin esos datos no se puede registrar la '+(tipo==='factura'?'factura':'compra/operación')+', porque después es muy difícil rastrearla.'); return; }
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
      rec.requestId = requestIdParaDatos_('addRecord', rec);
      const resAdd = await runServer('addRecord', rec);
      showToast('Registro guardado.');
      if(resAdd && resAdd.alertaImporte){
        alert('⚠ IMPORTE A AUTORIZAR\n\nLa factura quedó guardada pero RETENIDA: '+resAdd.alertaImporte+'.\n\nNo se va a poder pagar hasta que un administrador (distinto de quien la cargó) la autorice. Ya se les avisó por mail.');
      }
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
    }catch(err){ if(/DUPLICADO/.test(err.message||'')) alert(err.message); else if(esBloqueoServidor_(err)) avisarBloqueoServidor_(err); else showToast('No se pudo guardar: '+(err.message||'error')); }
  });

  /* ---------------- TABS ---------------- */
  $$('.tab-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      $$('.tab-btn').forEach(b=>b.classList.toggle('active', b===btn));
      $$('.view').forEach(v=>v.classList.remove('active'));
      $('#view-'+btn.dataset.view).classList.add('active');
      if(btn.dataset.view==='arca') cargarArcaSiHaceFalta();
      if(btn.dataset.view==='actividad') cargarLogSiHaceFalta();
      if(btn.dataset.view==='transferencias') cargarExportacionesPendientes_();
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
    const parcial = (r.tipo==='factura' && Number(r.cancelado)>0 && Number(r.saldo)>0) ? `<span class="ticket-tag" style="background:var(--media-soft);color:#7A5615;">A cuenta · saldo ${fmtMoney(r.saldo)}</span>` : '';
    const ctrlEst = ctrlImporteEstado_(r);
    const ctrlBadge = ctrlEst==='pendiente' ? '<span class="ticket-tag" style="background:var(--alta-soft);color:var(--alta);">⚠ Importe a autorizar</span>' : (ctrlEst==='autorizado' ? '<span class="ticket-tag" style="background:var(--baja-soft);color:var(--baja);">✓ Importe autorizado</span>' : '');
    const fundBadge = r.fundacion ? `<span class="ticket-tag" style="background:${r.fundacion==='mdg'?'var(--accent-soft)':'var(--media-soft)'};color:${r.fundacion==='mdg'?'var(--accent-dark)':'#7A5615'};">${r.fundacion.toUpperCase()}</span>` : '';
    return `<div class="ticket-card" data-prio="${r.prioridad}" draggable="true" data-id="${r._key||''}">
      <div class="ticket-stub">
        <span class="src" style="color:${srcColor}">${r.fuente?iconFor(r.fuente):(historica?'📋':'')} ${r.fuente?FUENTE_LABEL[r.fuente]||r.fuente:(historica?'Cargado por Mariana':'')}</span>
        <span class="days ${overdue?'overdue':''}">${d}d</span>
      </div>
      <div class="ticket-tear"></div>
      <div class="ticket-body">
        ${fundBadge}${ctrlBadge}${parcial}<span class="ticket-tag">${TIPO_LABEL[r.tipo]||r.tipo}</span>
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
        const idOrRow = (kind==='factura' || (rec.fundacion && rec._rowNum)) ? rec._rowNum : rec.id;
        await runServer('updateRecordRouter', kind, rec.fundacion||null, idOrRow, patch, `Movió "${rec.proveedor}" de ${ESTADO_LABEL[rec.estado]} a ${ESTADO_LABEL[nuevoEstado]}`);
        showToast('Movido a '+ESTADO_LABEL[nuevoEstado]+'.');
        refreshAll();
      }catch(err){ showToast('No se pudo mover: '+(err.message||'error')); }
    });
  });

  /* ---------------- EDIT MODAL ---------------- */
  let editingId=null, editingKind='otro', editingFundacion=null, editOriginalPrio='media';
  // Refleja AL INSTANTE en pantalla el pago que el servidor acaba de registrar (la relectura completa tarda ~25 s).
  function aplicarPagoLocal_(fundacion, rowNum, importe, res){
    const rec = records.find(r=>r.tipo==='factura' && r.fundacion===fundacion && Number(r._rowNum)===Number(rowNum));
    if(!rec || !res) return;
    rec.cancelado = (Number(rec.cancelado)||0) + (Number(importe)||0);
    rec.saldo = Number(res.saldo)||0;
    if(res.completado){ rec.pagado = true; rec.estado = 'completado'; rec.fechaResuelta = todayISO(); }
    else if(rec.estado!=='completado'){ rec.estado = 'en_proceso'; }
  }
  function importePend_(r){ return (r.tipo==='factura' && r.saldo!=null) ? (Number(r.saldo)||0) : (Number(r.importe)||0); }
  function ctrlImporteEstado_(r){ const v=String((r&&r.controlImporte)||''); return v.indexOf('PENDIENTE')===0 ? 'pendiente' : (v.indexOf('AUTORIZADO')===0 ? 'autorizado' : ''); }
  function fmtFechaAR_(iso){ const m=/^(\d{4})-(\d{2})-(\d{2})/.exec(iso||''); return m ? (m[3]+'/'+m[2]+'/'+m[1]) : (iso||''); }
  function cambioPrioVisible_(){ return editOriginalPrio==='alta' && editSelectedPrio!=='alta'; }
  function actualizarCambioPrio_(){
    const visible = cambioPrioVisible_();
    $('#e-cambioprio-wrap').classList.toggle('hidden', !visible);
    if(!visible){ $('#e-cp-preview').textContent=''; return; }
    const quien=$('#e-cp-quien').value.trim(), medio=$('#e-cp-medio').value, fecha=$('#e-cp-fecha').value;
    $('#e-cp-preview').textContent = 'Quedará registrado: «Prioridad bajada de alta a '+editSelectedPrio+' el '+(fmtFechaAR_(fecha)||'[fecha]')+' con autorización de '+(quien||'[nombre y área]')+', confirmada por '+(medio||'[mail / WhatsApp / en persona]')+'.»';
  }
  ['e-cp-quien','e-cp-medio','e-cp-fecha'].forEach(id=>{ const el=document.getElementById(id); el.addEventListener('input',actualizarCambioPrio_); el.addEventListener('change',actualizarCambioPrio_); });
  let editTipoFacturaOrig_ = '', editFechaEmisionOrig_ = '', editingEsDelPanel_ = false;
  function openEditModal(key){
    const r = records.find(x=>x._key===key); if(!r) return;
    editingKind = r.tipo==='factura' ? 'factura' : 'otro';
    editingFundacion = r.fundacion || null;
    editingId = (editingKind==='factura' || (r.fundacion && r._rowNum)) ? r._rowNum : r.id;
    editingEsDelPanel_ = !!r.isAppRecord;
    $('#e-tipo').value=r.tipo; $('#e-fuente').value=r.fuente||'email'; $('#e-proveedor').value=r.proveedor||'';
    $('#e-cuit').value=formatearCuit_(r.cuit||''); $('#e-importe').value=r.importe||'';
    $('#e-numfactura').value=r.numeroFactura||''; $('#e-cae').value=r.cae||'';
    editTipoFacturaOrig_ = r.tipoFactura||''; editFechaEmisionOrig_ = String(r.fechaEmision||'').slice(0,10);
    $('#e-tipofactura').value = ['A','B','C','NC','Boleta'].indexOf(editTipoFacturaOrig_)>=0 ? editTipoFacturaOrig_ : '';
    $('#e-fechaemision').value = editFechaEmisionOrig_;
    $('#e-numcomprobante').value=r.numeroComprobante||'';
    $('#e-centrocosto').value=r.centroCosto||''; $('#e-clasificacion').value=r.clasificacionFlujo||'';
    $('#e-clasificacion').setAttribute('readonly','readonly');
    $('#e-fundacion-otros').value = editingKind==='otro' ? (r.fundacion||'') : '';
    $('#e-descripcion').value=r.descripcion||''; $('#e-fecharecibida').value=r.fechaRecibida||'';
    $('#e-fecharesuelta').value=r.fechaResuelta||''; $('#e-estado').value=r.estado;
    editSelectedPrio=r.prioridad;
    $$('#e-prio-select .prio-opt').forEach(o=>o.classList.toggle('selected', o.dataset.p===editSelectedPrio));
    editOriginalPrio = r.prioridad || 'media';
    cargarPagosModal_(r);
    (function(){
      const est = ctrlImporteEstado_(r), wrap = $('#e-ctrl-wrap');
      wrap.classList.toggle('hidden', !est);
      $('#e-ctrl-autorizar').classList.toggle('hidden', !(est==='pendiente' && soyAdmin && r.tipo==='factura'));
      if(est){
        const pend = est==='pendiente';
        wrap.style.background = pend ? 'var(--alta-soft)' : 'var(--baja-soft)';
        wrap.style.border = '1px solid ' + (pend ? 'var(--alta)' : 'var(--baja)');
        const detalle = escapeHtml(String(r.controlImporte).replace(/^(PENDIENTE|AUTORIZADO)\s*/,'').replace(/^\|\s*/,''));
        $('#e-ctrl-txt').innerHTML = pend
          ? '<b>⚠ Importe a autorizar — factura retenida.</b> No se puede pagar ni pasar a Completado hasta que un administrador (distinto de quien la cargó) la autorice.<br><span style="color:var(--ink-soft)">'+detalle+'</span>'
          : '<b>✓ Importe autorizado.</b><br><span style="color:var(--ink-soft)">'+detalle+'</span>';
      }
    })();
    $('#e-cp-fecha').value = todayISO(); $('#e-cp-medio').value=''; $('#e-cp-quien').value='';
    $('#e-cambioprio-hist').classList.toggle('hidden', !r.cambioPrioridad);
    $('#e-cambioprio-hist-txt').innerHTML = (r.cambioPrioridad||'').split(' | ').filter(Boolean).map(t=>'• '+escapeHtml(t)).join('<br>');
    actualizarCambioPrio_();
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
  $('#e-prio-select').addEventListener('click',(e)=>{ const o=e.target.closest('.prio-opt'); if(!o) return; editSelectedPrio=o.dataset.p; $$('#e-prio-select .prio-opt').forEach(x=>x.classList.toggle('selected',x===o)); actualizarCambioPrio_(); });
  function closeModal(){ $('#modal-overlay').classList.add('hidden'); editingId=null; }
  async function cargarPagosModal_(r){
    const wrap = $('#e-pagos-wrap');
    const esFactura = r.tipo==='factura' && r.fundacion && r._rowNum;
    wrap.classList.toggle('hidden', !esFactura);
    if(!esFactura) return;
    const saldo = Number(r.saldo)||0, cancelado = Number(r.cancelado)||0, completa = !!r.pagado || r.estado==='eliminado';
    $('#e-pagos-resumen').innerHTML = 'Importe '+fmtMoney(r.importe)+' · Registrado '+fmtMoney(cancelado)+' · <b style="color:'+(saldo>0?'var(--alta)':'var(--baja)')+'">Saldo '+fmtMoney(saldo)+'</b>'+(r.pagado?' · ✓ pagada':'');
    $('#e-pagos-form').classList.toggle('hidden', completa);
    $('#e-pg-fecha').value = todayISO(); $('#e-pg-concepto').value='Pago'; $('#e-pg-importe').value = saldo>0 ? saldo : ''; $('#e-pg-medio').value=''; $('#e-pg-nota').value='';
    $('#e-pg-medios').innerHTML = (mediosPago||[]).map(m=>'<option value="'+escapeHtml(m.medioDePago)+'">').join('');
    const lista = $('#e-pagos-lista'); lista.innerHTML = '';
    if(cancelado<=0 && !r.pagado) return;
    lista.innerHTML = '<span class="hint">Cargando movimientos…</span>';
    try{
      const pagos = await runServer('getPagosDeFactura', r.fundacion, r._rowNum);
      lista.innerHTML = pagos.length ? pagos.map(p=>'<div style="display:flex;justify-content:space-between;gap:8px;border-bottom:1px solid var(--line);padding:4px 0;"><span>'+escapeHtml(fmtFechaAR_(p.fecha))+' · <b>'+escapeHtml(p.concepto)+'</b> '+fmtMoney(p.importe)+(p.medioDePago?' · '+escapeHtml(p.medioDePago):'')+(p.nota?' · <i>'+escapeHtml(p.nota)+'</i>':'')+'</span>'+(soyAdmin?'<button type="button" class="btn-ghost" data-del-pago="'+escapeHtml(p.id)+'">quitar</button>':'')+'</div>').join('') : '';
      $$('[data-del-pago]').forEach(b=>b.addEventListener('click', async ()=>{
        if(!confirm('¿Quitar este movimiento? Si con él se había cerrado la factura, vuelve a quedar con saldo pendiente.')) return;
        try{ await runServer('eliminarPagoParcial', b.dataset.delPago); showToast('Movimiento quitado.'); closeModal(); refreshAll(); }
        catch(err){ showToast(err.message||'No se pudo quitar.'); }
      }));
    }catch(err){ lista.innerHTML=''; }
  }
  $('#e-pg-medio').addEventListener('input', ()=>$('#e-pg-usd-hint').classList.toggle('hidden', !/U\$S/i.test($('#e-pg-medio').value)));
  $('#e-pg-registrar').addEventListener('click', async ()=>{
    if(!editingId || editingKind!=='factura') return;
    const datos = { fecha:$('#e-pg-fecha').value, concepto:$('#e-pg-concepto').value, importe:Number($('#e-pg-importe').value)||0, medioDePago:$('#e-pg-medio').value.trim(), nota:$('#e-pg-nota').value.trim() };
    if(!(datos.importe>0)){ showToast('Ingresá un importe mayor a cero.'); return; }
    const btn = $('#e-pg-registrar'); btn.disabled = true;
    try{
      const fundSel = editingFundacion, filaSel = editingId;
      const res = await runServer('registrarPagoParcial', fundSel, filaSel, datos, requestIdParaDatos_('pp:'+fundSel+':'+filaSel, datos));
      showToast(res.completado ? 'Registrado. Saldo $0: la factura quedó pagada por completo.' : 'Registrado. Saldo pendiente: '+fmtMoney(res.saldo));
      aplicarPagoLocal_(fundSel, filaSel, datos.importe, res);
      closeModal(); renderAll(); refreshAll();
    }catch(err){ if(esBloqueoServidor_(err)) avisarBloqueoServidor_(err); else showToast(err.message||'No se pudo registrar.'); }
    btn.disabled = false;
  });
  $('#e-ctrl-autorizar').addEventListener('click', async ()=>{
    if(!editingId || editingKind!=='factura') return;
    const nota = prompt('Vas a autorizar el pago de esta factura con importe fuera de lo habitual.\n\nNota opcional (ej: "confirmado con Dirección por mail"):');
    if(nota===null) return;
    try{
      const fundSel = editingFundacion, filaSel = editingId;
      await runServer('autorizarImporte', fundSel, filaSel, nota);
      showToast('Importe autorizado.');
      const rec = records.find(r=>r.tipo==='factura' && r.fundacion===fundSel && Number(r._rowNum)===Number(filaSel));
      if(rec) rec.controlImporte = String(rec.controlImporte||'').replace(/^PENDIENTE/, 'AUTORIZADO');
      closeModal(); renderAll(); refreshAll();
    }
    catch(err){ showToast(err.message||'No se pudo autorizar.'); }
  });
  $('#modal-close').addEventListener('click', closeModal);
  $('#btn-cancel-edit').addEventListener('click', closeModal);
  $('#modal-overlay').addEventListener('click',(e)=>{ if(e.target.id==='modal-overlay') closeModal(); });

  $('#edit-form').addEventListener('submit', async (e)=>{
    e.preventDefault();
    if(!editingId) return;
    const tipo = $('#e-tipo').value;
    const proveedorNuevo = $('#e-proveedor').value.trim();
    const faltClasifEd = faltaClasif_($('#e-centrocosto').value, $('#e-clasificacion').value);
    if(faltClasifEd.length){ cartelBloqueo_('Falta completar: '+faltClasifEd.join(' y ')+'.\nSin esos datos no se puede guardar este registro.'); return; }
    // Facturas: el CUIT es obligatorio (las compras pueden no tenerlo). Las filas históricas de Mariana no se editan desde acá.
    if(tipo==='factura' && (editingKind!=='factura' || editingEsDelPanel_) && !$('#e-cuit').value.replace(/[^0-9]/g,'')){
      cartelBloqueo_('Falta completar: CUIT.\nUna factura no se puede guardar sin CUIT.'); return;
    }
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
    // Tipo de factura y fecha de emisión: solo se envían si el usuario los cambió (así no se pisa ni se bloquea nada por diferencias de formato).
    if(editingKind==='factura'){
      const tf = $('#e-tipofactura').value, fe = $('#e-fechaemision').value;
      if(tf && tf!==editTipoFacturaOrig_) patch.tipoFactura = tf;
      if(fe && fe!==editFechaEmisionOrig_) patch.fechaEmision = fe;
    }
    if(cambioPrioVisible_()){
      const cp = { fecha:$('#e-cp-fecha').value, autorizadoPor:$('#e-cp-quien').value.trim(), confirmadoPor:$('#e-cp-medio').value };
      if(!cp.fecha || !cp.autorizadoPor || !cp.confirmadoPor){
        showToast('Para bajar una prioridad Alta completá: fecha, quién lo autorizó (nombre y área) y cómo se confirmó.');
        return;
      }
      patch.cambioPrioridad = cp;
    }
    try{
      await runServer('updateRecordRouter', editingKind, editingFundacion, editingId, patch, `Editó registro de "${proveedorNuevo}"`);
      showToast('Cambios guardados.'); closeModal(); refreshAll();
    }catch(err){ if(/DUPLICADO/.test(err.message||'')) alert(err.message); else if(esBloqueoServidor_(err)) avisarBloqueoServidor_(err); else showToast('No se pudo guardar: '+(err.message||'error')); }
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
    const importePend = pendientesTodos.reduce((s,r)=>s+importePend_(r),0);
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
    const centroF = $('#listado-filt-centro').value;
    const cuitF = $('#listado-filt-cuit').value;
    const soloSC = soloSinClasificar_;
    const activo = $('#listado-filt-activo').value; // activas | eliminadas | todas
    const q = sinTildes_($('#listado-buscar').value);
    return records.filter(r=>
      (activo==='todas' || (activo==='eliminadas' ? r.estado==='eliminado' : r.estado!=='eliminado')) &&
      (!fund || r.fundacion===fund) && (!prio || r.prioridad===prio) && (!est || r.estado===est) &&
      (!soloSC || sinClasificar_(r)) &&
      (cuitF!=='__sin__' || sinCuit_(r)) &&
      (!clasif || (clasif==='__sin__' ? !String(r.clasificacionFlujo||'').trim() : r.clasificacionFlujo===clasif)) &&
      (!centroF || (centroF==='__sin__' ? !String(r.centroCosto||'').trim() : sinTildes_(r.centroCosto||'')===sinTildes_(centroF))) &&
      (!listadoFechaDesde || (r.fechaRecibida && r.fechaRecibida>=listadoFechaDesde)) &&
      (!listadoFechaHasta || (r.fechaRecibida && r.fechaRecibida<=listadoFechaHasta)) &&
      (!q || sinTildes_(r.proveedor).includes(q) || sinTildes_(r.cuit).includes(q)) );
  }
  ['listado-filt-fundacion','listado-filt-prioridad','listado-filt-estado','listado-filt-activo','listado-filt-clasificacion','listado-filt-centro','listado-filt-cuit'].forEach(id=> $('#'+id).addEventListener('change', ()=>{ renderSinCuit_(); renderListado(); }));
  function poblarFiltroClasificacionListado(){
    const sel = $('#listado-filt-clasificacion');
    const actual = sel.value;
    sel.innerHTML = '<option value="">Toda clasificación</option><option value="__sin__">⚠ Sin clasificar</option>' + clasificacionesFlujo.map(c=>`<option value="${escapeHtml(c.clasificacion)}">${escapeHtml(c.clasificacion)}</option>`).join('');
    sel.value = actual;
  }
  (function poblarFiltroCentroListado_(){
    const sel = $('#listado-filt-centro');
    const opciones = Array.from($('#e-centrocosto').options).filter(o=>o.value).map(o=>'<option value="'+escapeHtml(o.value)+'">'+escapeHtml(o.textContent)+'</option>').join('');
    sel.innerHTML = '<option value="">Todo centro de costo</option><option value="__sin__">⚠ Sin centro de costo</option>' + opciones;
  })();
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
    const headers = ['Fuente','Fundación','Pagado','Tipo','Proveedor','CUIT','Tipo Factura','N° Factura','Fecha Emisión','Descripción','Importe','CAE','Medio de Pago','Centro de Costo','Clasificación Flujo','Prioridad','Estado','Cargado Por','Fecha Recibida','Fecha Resuelta','Cambio de prioridad autorizado','Control de importe','Pagado/cancelado','Saldo'];
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
      'Fecha Resuelta': r.fechaResuelta||'',
      'Cambio de prioridad autorizado': r.cambioPrioridad||'',
      'Control de importe': r.controlImporte||'',
      'Pagado/cancelado': r.tipo==='factura' ? (Number(r.cancelado)||0) : '',
      'Saldo': r.tipo==='factura' ? (Number(r.saldo)||0) : ''
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
      <td class="mono">${escapeHtml(r.cae)}</td><td>${escapeHtml(r.clasificacionFlujo)}</td><td>${escapeHtml(r.cargadoPor)}</td><td style="white-space:normal;min-width:240px;font-size:11.5px;">${escapeHtml(r.cambioPrioridad||'')}</td><td style="white-space:normal;min-width:200px;font-size:11.5px;">${ctrlImporteEstado_(r)==='pendiente'?'<b style="color:var(--alta)">⚠ A AUTORIZAR</b> ':''}${escapeHtml(String(r.controlImporte||'').replace(/^(PENDIENTE|AUTORIZADO)/, m=>m==='AUTORIZADO'?'✓ AUTORIZADO':'').replace(/^\s*\|\s*/,''))}</td><td class="mono">${(r.tipo==='factura' && Number(r.cancelado)>0) ? fmtMoney(r.saldo) : ''}</td></tr>`;
    }).join('');
    $$('[data-pagado]').forEach(chk=>chk.addEventListener('change', async ()=>{
      const [kind, fund, rowNum] = chk.dataset.pagado.split(':');
      try{ await runServer('setPagadoRouter', kind, fund, parseInt(rowNum,10), chk.checked); showToast(chk.checked?'Marcado como pagado/completado.':'Desmarcado.'); refreshAll(); }
      catch(err){ if(esBloqueoServidor_(err)) avisarBloqueoServidor_(err); else showToast('Error: '+(err.message||'')); chk.checked = !chk.checked; }
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
    const importePend = pendRecs.reduce((s,r)=>s+importePend_(r),0);
    const chips = [
      {n:pendRecs.length, l:'Pendientes', lista:pendRecs, titulo:'Pendientes'},
      {n:facturasPendRecs.length, l:'Facturas x pagar', lista:facturasPendRecs, titulo:'Facturas por pagar'},
      {n:fmtMoney(importePend), l:'Importe pendiente', lista:pendRecs, titulo:'Importe pendiente'},
      {n:overdueRecs.length, l:'+7 días', warn:true, lista:overdueRecs, titulo:'Pendientes hace más de 7 días'}
    ];
    const retenidas = records.filter(r=>r.tipo==='factura' && r.estado!=='eliminado' && ctrlImporteEstado_(r)==='pendiente');
    if(retenidas.length) chips.unshift({n:retenidas.length, l:'Importes a autorizar', warn:true, lista:retenidas, titulo:'Facturas con importe a autorizar'});
    const aCuenta = records.filter(r=>r.tipo==='factura' && r.estado!=='eliminado' && !r.pagado && Number(r.cancelado)>0 && Number(r.saldo)>0);
    if(aCuenta.length) chips.push({n:aCuenta.length, l:'Con pago a cuenta', lista:aCuenta, titulo:'Facturas con pago a cuenta (saldo pendiente)'});
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
    const importePend = pendientes.reduce((s,r)=>s+importePend_(r),0);
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
  const irASelector = (continuar) => {
    const destino = 'https://accounts.google.com/AccountChooser' + (continuar ? '?continue=' + encodeURIComponent(continuar) : '');
    window.open(destino, '_top'); // '_top': el selector de Google no puede abrirse dentro del iframe de Apps Script
  };
  google.script.run
    .withSuccessHandler(url => irASelector(url))
    .withFailureHandler(() => irASelector(''))
    .getWebAppUrl();
});

  $('#btn-revisar-comprobantes').addEventListener('click', async ()=>{
    const statusEl = $('#comprobantes-status');
    statusEl.innerHTML = '<span class="spinner"></span>Leyendo comprobantes… puede tardar unos minutos (cada uno lo lee la IA). No cierres la pestaña.';
    $('#btn-revisar-comprobantes').disabled = true;
    try{
      const respuesta = await runServer('processTransferenciasFolder');
      statusEl.textContent = '';
      mostrarResumenComprobantes(respuesta.resultados||[], respuesta.sinBanco||[]);
      if(respuesta.quedanPendientes) statusEl.textContent = 'Quedan '+respuesta.quedanPendientes+' comprobante(s) sin leer por falta de tiempo: tocá el botón de nuevo.';
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
      if(respuesta.quedanPendientes) statusEl.textContent += ' Quedan '+respuesta.quedanPendientes+' sin leer por falta de tiempo: tocá el botón de nuevo.';
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

  /* ---------------- COPIAS DE SEGURIDAD (solo admins) ---------------- */
  let backupFolderUrl = '';
  function fmtFechaHoraAR_(iso){ if(!iso) return '—'; const d=new Date(iso); return isNaN(d) ? String(iso) : d.toLocaleString('es-AR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'}); }
  async function cargarBackups_(){
    const box = $('#backup-content');
    box.innerHTML = '<div class="hint"><span class="spinner"></span>Leyendo copias…</div>';
    try{
      const info = await runServer('getBackupInfo');
      backupFolderUrl = info.folderUrl || '';
      $('#backup-abrir-carpeta').disabled = !backupFolderUrl;
      if(!info.configurado){
        box.innerHTML = '<div class="alert-banner">Todavía no está activada. Un administrador tiene que ejecutar <b>crearDisparadorBackup</b> una vez desde el editor de Apps Script (crea la carpeta y programa la copia diaria).</div>';
        return;
      }
      const u = info.ultimo;
      let estado;
      if(!u) estado = '<div class="alert-banner">Aún no hay ninguna copia registrada.</div>';
      else if(!u.ok) estado = '<div class="alert-banner" style="background:var(--alta-soft);border-color:var(--alta);color:var(--alta);">⚠ El último intento falló ('+fmtFechaHoraAR_(u.ts)+'): '+escapeHtml(u.error||'error')+'</div>';
      else {
        const horas = (Date.now()-new Date(u.ts).getTime())/3600000;
        estado = horas>36
          ? '<div class="alert-banner">⚠ La última copia exitosa fue el '+fmtFechaHoraAR_(u.ts)+' (hace más de un día y medio). Revisá que el disparador siga activo.</div>'
          : '<div style="font-size:13px;margin-bottom:10px;">✅ Última copia: <b>'+fmtFechaHoraAR_(u.ts)+'</b></div>';
      }
      let lista = '';
      if(info.sinAcceso) lista = '<div class="hint">Tu cuenta no tiene acceso a la carpeta para listar las copias. Pedile a quien la creó que la comparta con vos.</div>';
      else if(!info.backups.length) lista = '<div class="empty-state">No hay copias en la carpeta todavía.</div>';
      else lista = '<div class="table-wrap" style="max-height:240px;"><table><thead><tr><th>Copia</th><th>Creada</th><th></th></tr></thead><tbody>'
        + info.backups.map(b=>'<tr><td style="white-space:normal;">'+escapeHtml(b.nombre.replace('Backup Facturas Procesadas - ',''))+'</td><td class="mono">'+fmtFechaHoraAR_(b.fecha)+'</td><td><button class="btn" data-backup-url="'+escapeHtml(b.url)+'">Abrir</button></td></tr>').join('')
        + '</tbody></table></div><div class="hint" style="margin-top:6px;">Se muestran las últimas '+info.backups.length+' de '+info.total+'. Se conservan '+info.retencionDias+' días.</div>';
      box.innerHTML = estado + lista;
      $$('[data-backup-url]').forEach(b=>b.addEventListener('click',()=>window.open(b.dataset.backupUrl,'_blank','noopener')));
    }catch(err){
      box.innerHTML = '<div class="alert-banner">No se pudo leer: '+escapeHtml(err.message||'error')+'</div>';
    }
  }
  $('#btn-backups').addEventListener('click', ()=>{ $('#backup-overlay').classList.remove('hidden'); cargarBackups_(); });
  $('#backup-close').addEventListener('click', ()=>$('#backup-overlay').classList.add('hidden'));
  $('#backup-overlay').addEventListener('click',(e)=>{ if(e.target.id==='backup-overlay') $('#backup-overlay').classList.add('hidden'); });
  $('#backup-abrir-carpeta').addEventListener('click', ()=>{ if(backupFolderUrl) window.open(backupFolderUrl,'_blank','noopener'); });
  $('#backup-ahora').addEventListener('click', async ()=>{
    const btn = $('#backup-ahora'); btn.disabled = true; const txt = btn.textContent;
    btn.innerHTML = '<span class="spinner"></span>Copiando… (puede tardar un minuto)';
    try{
      // sin el límite de 25 s de runServer: copiar una planilla grande puede demorar más
      const res = await new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).ejecutarBackupAhora());
      showToast('Copia creada: '+(res&&res.nombre||'ok'));
    }catch(err){ showToast('No se pudo hacer la copia: '+(err.message||'error')); }
    btn.disabled = false; btn.textContent = txt;
    cargarBackups_();
  });

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
  // CSV con campos entre comillas (pueden contener ; o saltos de línea).
  function parseCsvTexto_(text, sep){
    const filas = []; let fila = [], campo = '', enComillas = false;
    for(let i=0;i<text.length;i++){
      const ch = text[i];
      if(enComillas){
        if(ch==='"'){ if(text[i+1]==='"'){ campo+='"'; i++; } else enComillas = false; }
        else campo += ch;
      } else if(ch==='"' && campo===''){ enComillas = true; }
      else if(ch===sep){ fila.push(campo); campo=''; }
      else if(ch==='\n'){ fila.push(campo); filas.push(fila); fila=[]; campo=''; }
      else if(ch!=='\r'){ campo += ch; }
    }
    if(campo!=='' || fila.length){ fila.push(campo); filas.push(fila); }
    return filas.filter(f=>f.some(c=>String(c).trim()));
  }
  // Excel guarda los números largos como número (pierde dígitos) o en notación científica: no sirven como CBU.
  function cbuLegible_(v){
    if(typeof v==='number') return Number.isSafeInteger(v) ? { valor:String(v), ilegible:false } : { valor:'', ilegible:true };
    const t = String(v==null?'':v).trim();
    return /^\d+([.,]\d+)?e[+-]?\d+$/i.test(t) ? { valor:'', ilegible:true } : { valor:t, ilegible:false };
  }
  function parseBankContactsCsv(text){
    text = text.replace(/^\uFEFF/, ''); // sacar BOM
    const tabla = parseCsvTexto_(text, ';');
    if(!tabla.length) return [];
    const headers = tabla[0].map(h=>String(h).trim());
    const idx = name => headers.indexOf(name);
    const iCuenta = idx('Cuenta destino'), iCuit = idx('CUIT/CUIL'), iNombre = idx('Nombre/Razón social'),
          iDesc = idx('Descripción de Contacto'), iMail1 = idx('Mail 1'), iMail2 = idx('Mail 2'),
          iMail3 = idx('Mail 3'), iMail4 = idx('Mail 4'), iMail5 = idx('Mail 5');
    const rows = [];
    for(let i=1;i<tabla.length;i++){
      const cols = tabla[i];
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
      const cbuIlegibles = [];
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
      rows.forEach(r=>{ const c = cbuLegible_(r.cbu); if(c.ilegible) cbuIlegibles.push(r.proveedor||'—'); r.cbu = c.valor; });
      if(cbuIlegibles.length){
        if(!confirm('Estos contactos tienen el CBU guardado como número (Excel lo recortó o lo pasó a notación científica) y NO se pueden leer:\n\n• '+cbuIlegibles.slice(0,15).join('\n• ')+(cbuIlegibles.length>15?'\n• ...':'')+'\n\nPara arreglarlo, formateá esa columna como TEXTO antes de pegar los CBU.\n\n¿Importar igual los contactos (con el CBU vacío en esos)?')){ statusEl.textContent='Importación cancelada.'; return; }
      }
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
      // Valores fijos que exige el diseño de registro de Banco Macro (BNF): tipo de documento 10 (CUIT/CUIL),
      // condiciones IIBB/Ganancias/IVA = 99 (no informado), calle "NO INFORMADA", código postal 1001; razón social hasta 40 caracteres.
      'Tipo de documento - Beneficiario':'10',
      'Número de documento - Beneficiario': (a.cuit||'').replace(/[^0-9]/g,''),
      'Condición IIBB':'99', 'Condición Ganancias':'99', 'Condición de I.V.A.':'99',
      'Razón social - Beneficiario': textoXlsx_(a.proveedor||'').slice(0,40),
      'Calle':'NO INFORMADA', 'Núm. De puerta':'', 'Unidad funcional':'', 'Código postal':'1001',
      'Núm. IIBB':'', 'E-mail': a.email||'', 'Piso':'', 'País':'', 'Provincia':'', 'Ciudad':''
    }));
    const ws = XLSX.utils.json_to_sheet(data, {header: headers});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
    XLSX.writeFile(wb, 'Beneficiarios_Banco_Macro_'+todayISO()+'.xlsx');
    showToast('Generado con '+data.length+' contacto(s) con los valores fijos que pide Banco Macro (tipo doc. 10, condiciones 99, calle "NO INFORMADA", CP 1001). El e-mail solo es obligatorio si querés que el banco envíe el aviso de pago al proveedor.');
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
  function renderTransferencias(){
    const pendientes = getTransfFiltradas();
    $('#transf-empty').style.display = pendientes.length? 'none':'block';
    $('#transf-body').innerHTML = pendientes.map(r=>{
      const match = findAgendaMatch(r.proveedor, r.cuit);
      const tieneCbu = match && match.cbu;
      const retenida = ctrlImporteEstado_(r)==='pendiente';
      return `<tr>
        <td style="text-align:center"><input type="checkbox" class="transf-check" data-transfid="${r.fundacion}:${r._rowNum}" ${retenida?'disabled':''}></td>
        <td>${r.fundacion.toUpperCase()}</td><td>${escapeHtml(r.proveedor)}</td><td class="mono">${escapeHtml(r.cuit)}</td>
        <td class="mono">${escapeHtml(r.numeroFactura)}</td><td class="mono">${fmtMoney(r.importe)}${Number(r.cancelado)>0 ? '<br><span style="color:var(--media);font-size:11px;">saldo '+fmtMoney(r.saldo)+'</span>' : ''}</td>
        <td><input type="number" step="0.01" min="0" class="transf-apagar mono-input" style="width:120px;" data-transfid="${r.fundacion}:${r._rowNum}" value="${Number(r.saldo!=null?r.saldo:r.importe)||0}"></td>
        <td class="mono">${match? escapeHtml(match.cbu||'—') : '—'}</td>
        <td>${match? escapeHtml(match.email||'—') : '—'}</td>
        <td>${retenida ? '<span style="color:var(--alta)">🔒 Importe a autorizar</span>' : (tieneCbu? '<span style="color:var(--baja)">✓ OK</span>' : '<span style="color:var(--alta)" title="Sin CBU solo se puede pagar con Caja o Tarjeta, o completar la Agenda">⚠ sin CBU (solo Caja/Tarjeta)</span>')}</td>
        <td>${r._key? `<button class="btn" data-transf-edit="${r._key}">Ver / Editar</button>` : ''}</td>
      </tr>`;
    }).join('');
    $$('.transf-check').forEach(chk=>chk.addEventListener('change', updateTransfResumen));
    $$('.transf-apagar').forEach(inp=>inp.addEventListener('input', updateTransfResumen));
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
      if(!rec) return;
      // Copia con el importe "A pagar" (por defecto el saldo): todos los archivos de banco usan r.importe
      const inp = document.querySelector('.transf-apagar[data-transfid="'+chk.dataset.transfid+'"]');
      const saldo = Number(rec.saldo!=null ? rec.saldo : rec.importe) || 0;
      let monto = inp ? (Number(inp.value)||0) : saldo;
      if(monto > saldo) monto = saldo;
      if(monto > 0) seleccionadas.push(Object.assign({}, rec, { importe: monto, importeOriginal: rec.importe, saldoPrevio: saldo }));
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

  /* ---------------- EXPORTACIONES A BANCOS ---------------- */
  // Flujo: 1) el servidor valida las facturas y deja la exportación PENDIENTE; 2) se descarga el archivo
  // con lo que el servidor aprobó; 3) se confirma qué pagos se subieron al banco. Mientras no se confirma,
  // la exportación queda a la vista de todos en el aviso de Transferencias (también si se cierra el diálogo).
  let exportacionesPendientes = [];
  async function cargarExportacionesPendientes_(){
    const r = await callSeguro('getExportacionesPendientes');
    if(r.ok) exportacionesPendientes = Array.isArray(r.value) ? r.value : [];
    renderExportacionesPendientes_();
  }
  function bannerExportaciones_(){
    let el = $('#exp-pendientes-banner');
    if(el) return el;
    const vista = $('#view-transferencias');
    if(!vista) return null;
    el = document.createElement('div');
    el.id = 'exp-pendientes-banner';
    el.style.cssText = 'display:none;margin:0 0 12px;padding:10px 12px;border:1px solid var(--media);background:var(--media-soft);border-radius:8px;font-size:13px;';
    vista.insertBefore(el, vista.firstChild);
    el.addEventListener('click', (e)=>{
      const b = e.target.closest('button'); if(!b) return;
      if(b.dataset.expConfirmar) confirmarPendienteDesdeAviso_(b.dataset.expConfirmar);
      else if(b.dataset.expCancelar) cancelarPendienteDesdeAviso_(b.dataset.expCancelar);
      else if(b.dataset.expRedescargar) redescargarProvincia_(b.dataset.expRedescargar);
    });
    return el;
  }
  function renderExportacionesPendientes_(){
    const n = exportacionesPendientes.length;
    const tab = document.querySelector('.tab[data-view="transferencias"], [data-view="transferencias"]');
    if(tab){
      let chip = tab.querySelector('.exp-chip');
      if(n && !chip){ chip = document.createElement('span'); chip.className = 'exp-chip'; chip.style.cssText='margin-left:6px;background:var(--media);color:#fff;border-radius:10px;padding:0 6px;font-size:11px;'; tab.appendChild(chip); }
      if(chip){ if(n) chip.textContent = String(n); else chip.remove(); }
    }
    const el = bannerExportaciones_(); if(!el) return;
    if(!n){ el.style.display = 'none'; el.innerHTML = ''; return; }
    el.style.display = 'block';
    el.innerHTML = '<b>⚠ Hay '+n+' exportación(es) a banco sin confirmar.</b> Descargar el archivo no registra ningún pago: los pagos se registran recién al confirmar acá. Hasta entonces esas facturas no se pueden volver a exportar.'+
      exportacionesPendientes.map(ex=>`<div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;">
        <span>${escapeHtml(ex.banco)} · ${escapeHtml(ex.cantidad)} pago(s) · ${fmtMoney(ex.total)} · ${escapeHtml(ex.usuario||'')} · ${escapeHtml(fmtFechaHoraAR_(ex.creada))}</span>
        <button class="btn" data-exp-confirmar="${escapeHtml(ex.id)}">Confirmar pagos</button>
        ${ex.banco==='Banco Provincia' ? `<button class="btn btn-ghost" data-exp-redescargar="${escapeHtml(ex.id)}">Volver a descargar archivo</button>` : ''}
        <button class="btn btn-ghost" data-exp-cancelar="${escapeHtml(ex.id)}">Cancelar</button>
      </div>`).join('');
  }
  function confirmarPendienteDesdeAviso_(id){
    const ex = exportacionesPendientes.find(e=>e.id===id); if(!ex) return;
    abrirConfirmacionExportacion_(ex.id, ex.banco, ex.items);
  }
  async function cancelarPendienteDesdeAviso_(id){
    const ex = exportacionesPendientes.find(e=>e.id===id); if(!ex) return;
    if(!confirm('¿Cancelar la exportación a '+ex.banco+' ('+ex.cantidad+' pago(s), '+fmtMoney(ex.total)+')?\n\nNo se registra ningún pago y las facturas quedan libres.\nSi ya subiste el archivo al banco, tenés que anular esa operación en el banco.')) return;
    try{ await runServer('cancelarExportacion', id); showToast('Exportación cancelada.'); }
    catch(err){ alert('No se pudo cancelar: '+(err&&err.message||'error')); }
    await cargarExportacionesPendientes_();
  }
  function descargarTexto_(contenido, nombre){
    const blob = new Blob([contenido], {type:'text/plain;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nombre;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 2000);
  }
  async function descargarArchivosProvincia_(exportacionId, fundaciones){
    let total = 0, importe = 0;
    for(const fundacion of fundaciones){
      const res = await runServer('generarArchivoProvinciaDeExportacion', exportacionId, fundacion);
      descargarTexto_(res.contenido, 'TRANSF_BANCO_PROVINCIA_'+res.rotulo+'_'+fundacion.toUpperCase()+'_'+todayISO()+'.txt');
      total += Number(res.cantidadPagos)||0; importe += Number(res.importeTotal)||0;
      await esperar_(500); // el navegador puede bloquear descargas simultáneas
    }
    return { cantidad: total, importe };
  }
  async function redescargarProvincia_(id){
    const ex = exportacionesPendientes.find(e=>e.id===id); if(!ex) return;
    const fundaciones = Array.from(new Set(ex.items.map(i=>i.fundacion)));
    try{ const r = await descargarArchivosProvincia_(id, fundaciones); showToast('Archivo(s) Banco Provincia descargado(s): '+r.cantidad+' pago(s).'); }
    catch(err){ alert('No se pudo generar el archivo: '+(err&&err.message||'error')); }
  }

  // Paso 1: el servidor decide qué facturas se pueden exportar. Devuelve null si no hay nada que exportar.
  async function prepararExportacionUI_(bancoKey, sel){
    const items = sel.map(r=>({ fundacion:r.fundacion, rowNum:Number(r._rowNum), importe:Number(r.importe)||0 }));
    let res;
    try{ res = await runServer('prepararExportacion', bancoKey, items, nuevoRequestId_()); }
    catch(err){
      alert('No se pudo preparar la exportación: '+(err&&err.message||'error'));
      await cargarExportacionesPendientes_(); // por si el servidor sí la dejó preparada
      return null;
    }
    if(res.rechazadas && res.rechazadas.length){
      alert('Estas facturas NO se incluyeron en la exportación:\n\n• '+res.rechazadas.map(x=>(x.proveedor||('fila '+x.rowNum))+': '+x.motivo).join('\n• '));
    }
    if(!res.exportacionId || !res.aprobadas || !res.aprobadas.length){ showToast('Ninguna factura pasó los controles: no se generó archivo.'); return null; }
    return res;
  }
  // Si el archivo no se pudo generar o descargar, la exportación se anula (no hay nada que subir al banco).
  async function generarOCancelar_(res, generar){
    try{ await generar(); return true; }
    catch(err){
      try{ await runServer('cancelarExportacion', res.exportacionId); }catch(e){}
      alert('No se pudo generar el archivo ('+(err&&err.message||'error')+'). La exportación se canceló y no se registró ningún pago. Si se alcanzó a descargar algún archivo, NO lo subas al banco.');
      await cargarExportacionesPendientes_();
      return false;
    }
  }
  async function terminarExportacion_(res){
    await cargarExportacionesPendientes_();
    abrirConfirmacionExportacion_(res.exportacionId, res.banco, res.aprobadas);
  }

  /* ---- Diálogo de confirmación (exportación pendiente o pago directo de Caja/Tarjeta) ---- */
  function pintarFilasConfirmacion_(lista){
    $('#confirmar-pagos-body').innerHTML = lista.map((r,i)=>`<tr>
      <td style="text-align:center"><input type="checkbox" class="confirmar-pago-check" data-idx="${i}" checked></td>
      <td>${escapeHtml(r.proveedor||'—')}</td>
      <td class="mono">${escapeHtml(r.numeroFactura||'')}</td>
      <td class="mono">${fmtMoney(r.importe)}${(r.saldoPrevio && r.importe < r.saldoPrevio-0.01) ? ' <span style="color:var(--media);font-size:11px;">(a cuenta)</span>' : ''}</td>
    </tr>`).join('');
  }
  function abrirConfirmacionExportacion_(id, banco, items){
    if(!items || !items.length) return;
    const lista = items.map(r=>({ fundacion:r.fundacion, rowNum:Number(r.rowNum), proveedor:r.proveedor, numeroFactura:r.numeroFactura, importe:Number(r.importe)||0, saldoPrevio:Number(r.saldoPrevio)||0 }));
    const ov = $('#confirmar-pagos-overlay');
    pintarFilasConfirmacion_(lista);
    ov.dataset.lista = JSON.stringify(lista);
    ov.dataset.banco = banco || '';
    ov.dataset.exportacionId = id;
    ov.querySelector('.hint').textContent = 'Exportación '+id+' a '+(banco||'Excel')+'. Dejá tildadas SOLO las facturas cuyo pago se subió al banco; los pagos se registran recién al confirmar. Si cerrás esta ventana, la exportación queda pendiente y se retoma desde el aviso amarillo de Transferencias.';
    ov.classList.remove('hidden');
  }
  function abrirConfirmacionPagos(lista, banco){
    if(!lista.length) return;
    const filas = lista.map(r=>({ fundacion:r.fundacion, rowNum:Number(r._rowNum), proveedor:r.proveedor, numeroFactura:r.numeroFactura, importe:Number(r.importe)||0, saldoPrevio:Number(r.saldoPrevio)||0, requestId:nuevoRequestId_() }));
    const ov = $('#confirmar-pagos-overlay');
    pintarFilasConfirmacion_(filas);
    ov.dataset.lista = JSON.stringify(filas);
    ov.dataset.banco = banco || '';
    ov.dataset.exportacionId = '';
    ov.querySelector('.hint').textContent = `Se tildaron todas por defecto y se van a marcar como pagadas vía "${banco}". Destildá las que no correspondan antes de confirmar.`;
    if(/U\$S/i.test(banco||'')) ov.querySelector('.hint').textContent += ' ⚠ Es una caja en dólares: el importe registrado se descuenta del saldo en PESOS de la factura (cargá el equivalente en pesos en "A pagar").';
    ov.classList.remove('hidden');
  }
  $('#confirmar-pagos-close').addEventListener('click', ()=>$('#confirmar-pagos-overlay').classList.add('hidden'));
  $('#confirmar-pagos-cancelar').addEventListener('click', ()=>$('#confirmar-pagos-overlay').classList.add('hidden'));

  async function confirmarExportacionUI_(id, marcar){
    const ov = $('#confirmar-pagos-overlay');
    if(!marcar.length){
      if(!confirm('No dejaste tildada ninguna factura.\n\nSi seguís, la exportación se CANCELA y no se registra ningún pago. Si ya subiste el archivo al banco, tenés que anular esa operación en el banco.\n\n¿Cancelar la exportación?')) return;
      try{ await runServer('cancelarExportacion', id); ov.classList.add('hidden'); showToast('Exportación cancelada.'); }
      catch(err){ alert('No se pudo cancelar: '+(err&&err.message||'error')); }
      await cargarExportacionesPendientes_();
      return;
    }
    let res;
    try{
      // Sin requestId a propósito: la confirmación es idempotente por estado en el servidor.
      res = await runServer('confirmarExportacion', id, marcar.map(m=>({fundacion:m.fundacion, rowNum:m.rowNum})));
    }catch(err){
      alert('No se pudo confirmar ('+(err&&err.message||'error')+').\n\nLa exportación sigue pendiente: actualizá el panel (F5), fijate si se registraron pagos y volvé a confirmar desde el aviso de Transferencias.');
      await cargarExportacionesPendientes_();
      return;
    }
    ov.classList.add('hidden');
    if(res.yaResuelta){ showToast('Esa exportación ya estaba resuelta ('+res.estado+').'); }
    else{
      showToast('Pagos registrados: '+(res.registrados||0)+(res.descartados? ' · descartados: '+res.descartados : '')+((res.fallidos||[]).length? ' · con error: '+res.fallidos.length : '')+'.');
      const avisos = [];
      if((res.fallidos||[]).length) avisos.push('No se pudieron registrar estos pagos:\n\n• '+res.fallidos.map(f=>(f.proveedor||'—')+' '+(f.numeroFactura||'')+': '+f.motivo).join('\n• ')+(res.sinCambios ? '\n\nNo se registró ninguno: la exportación sigue pendiente.' : '\n\nLos demás quedaron registrados.'));
      if((res.pdfSinMarcar||[]).length) avisos.push('No se encontró el PDF para marcar como "OP finalizada" de: '+res.pdfSinMarcar.join(', '));
      if(avisos.length) alert(avisos.join('\n\n'));
    }
    await cargarExportacionesPendientes_();
    renderAll();
    refreshAll();
  }
  async function confirmarPagosDirectos_(marcar, banco){
    const ov = $('#confirmar-pagos-overlay');
    let ok=0, fallidos=0;
    const motivos = [];
    for(const item of marcar){
      try{
        const res = await runServer('registrarPagoLote', item.fundacion, item.rowNum, item.importe, banco, item.requestId);
        ok++;
        aplicarPagoLocal_(item.fundacion, item.rowNum, item.importe, res);
      }catch(err){ fallidos++; motivos.push((item.proveedor||'—')+': '+(err && err.message ? err.message : 'error')); }
    }
    ov.classList.add('hidden');
    showToast('Pagos registrados: '+ok+(fallidos? ' · '+fallidos+' con error':'')+'.');
    renderAll();
    if(motivos.length) alert('No se pudieron registrar estos pagos:\n\n• '+motivos.join('\n• '));
    refreshAll();
  }
  let confirmandoPagos_ = false;
  $('#confirmar-pagos-aceptar').addEventListener('click', async ()=>{
    if(confirmandoPagos_) return;
    const ov = $('#confirmar-pagos-overlay');
    const lista = JSON.parse(ov.dataset.lista || '[]');
    const banco = ov.dataset.banco || '';
    const expId = ov.dataset.exportacionId || '';
    const marcar = [];
    $$('.confirmar-pago-check').forEach(chk=>{ if(chk.checked) marcar.push(lista[parseInt(chk.dataset.idx,10)]); });
    confirmandoPagos_ = true;
    $('#confirmar-pagos-aceptar').disabled = true;
    try{
      if(expId) await confirmarExportacionUI_(expId, marcar);
      else await confirmarPagosDirectos_(marcar, banco);
    }finally{
      confirmandoPagos_ = false;
      $('#confirmar-pagos-aceptar').disabled = false;
    }
  });

  /* ---- Botones de exportación ---- */
  function hojaDe_(filas, headers, nombreHoja, archivo){
    const ws = XLSX.utils.json_to_sheet(filas, {header:headers});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
    XLSX.writeFile(wb, archivo);
  }
  const HEADERS_GALICIA_ = ['CBU/CVU/Alias/Nro cuenta','Importe','Motivo','Descripción (opcional)','Email destinatario (opcional)','Mensaje del email (opcional)'];
  async function exportarGaliciaOMercadoPago_(bancoKey, archivoPrefijo){
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    if(sel.length>50){ showToast('El banco acepta máximo 50 pagos por archivo. Desmarcá algunas.'); return; }
    const res = await prepararExportacionUI_(bancoKey, sel); if(!res) return;
    const filas = res.aprobadas.map(a=>{
      const numCorto = facturaNumeroCorto(a.numeroFactura);
      const mensajeFallback = 'OP - '+a.cbu+'- FCFactura - FC '+numCorto+' - $'+Math.round(Number(a.importe)||0);
      return {
        'CBU/CVU/Alias/Nro cuenta': a.cbu,
        'Importe': Number(a.importe)||0,
        'Motivo': textoXlsx_(a.motivo || 'Factura'),
        'Descripción (opcional)': ('FC '+numCorto).slice(0,12),
        'Email destinatario (opcional)': a.email || '',
        'Mensaje del email (opcional)': textoXlsx_(a.nombreArchivoPdf || mensajeFallback).slice(0,200)
      };
    });
    const ok = await generarOCancelar_(res, async ()=>{ hojaDe_(filas, HEADERS_GALICIA_, 'Formulario', archivoPrefijo+todayISO()+'.xlsx'); });
    if(!ok) return;
    showToast('Archivo generado: '+filas.length+' pago(s). Falta confirmar cuáles se subieron al banco.');
    await terminarExportacion_(res);
  }
  $('#btn-transf-exportar').addEventListener('click', ()=>exportarGaliciaOMercadoPago_('galicia','TRANSF_BANCO_GALICIA_'));
  $('#btn-transf-exportar-mp').addEventListener('click', ()=>exportarGaliciaOMercadoPago_('mercadopago','TRANSF_MERCADOPAGO_'));

  $('#btn-transf-exportar-excel').addEventListener('click', async ()=>{
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    const res = await prepararExportacionUI_('excel', sel); if(!res) return;
    const headers = ['Fundación','Proveedor','CUIT','N° Factura','Importe','Saldo previo','CBU/CVU/Alias','Email'];
    const filas = res.aprobadas.map(a=>({
      'Fundación': (a.fundacion||'').toUpperCase(), 'Proveedor': textoXlsx_(a.proveedor), 'CUIT': a.cuit||'',
      'N° Factura': a.numeroFactura||'', 'Importe': Number(a.importe)||0, 'Saldo previo': Number(a.saldoPrevio)||0,
      'CBU/CVU/Alias': a.cbu||'', 'Email': a.email||''
    }));
    const ok = await generarOCancelar_(res, async ()=>{
      const ws = XLSX.utils.json_to_sheet(filas, {header:headers});
      ws['!autofilter'] = { ref: ws['!ref'] };
      ws['!cols'] = headers.map(h => ({ wch: Math.max(12, h.length+2) }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Pagos');
      XLSX.writeFile(wb, 'Pagos_'+todayISO()+'.xlsx');
    });
    if(!ok) return;
    showToast('Exportado: '+filas.length+' fila(s). Falta confirmar cuáles se pagaron.');
    await terminarExportacion_(res);
  });

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

  // Cuentas del propio Banco Macro: CBU que empieza con 285 (código de entidad) o número de cuenta de 15 dígitos.
  function esCuentaMacro_(digitos){ return (digitos.length===22 && digitos.indexOf('285')===0) || digitos.length===15; }
  async function exportarMacro_(variante){
    $('#menu-macro').classList.add('hidden');
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    const res = await prepararExportacionUI_('macro', sel); if(!res) return;
    const fecha = fechaHoyAR_();
    const sufijoId = String(res.exportacionId||'').replace(/^E/,'');
    const headersPP = ['Modalidad de Pago','Número de documento Beneficiario','Denominación del Beneficiario','Identificación del pago','Importe','Cuenta de pago - Nro/CBU o Nº Cheque','Marca de registración de cheque','Fecha de pago - Emisión','Fecha de pago diferido'];
    const headersTM = ['Modalidad de Pago','Número de documento Beneficiario','Denominación del Beneficiario','Concepto','Importe','Cuenta Destino','Fecha de pago - Emisión'];
    // Cada fundación tiene su propio convenio (empresa/CUIT y cuenta de débito en Macro): un archivo por fundación, nunca mezcladas.
    // En transferencias masivas, además, el banco no admite mezclar "cuentas Macro" (18) con "otros bancos" (12): archivos separados.
    const fundaciones = Array.from(new Set(res.aprobadas.map(a=>a.fundacion)));
    const archivos = [];
    fundaciones.forEach(fk=>{
      const deFund = res.aprobadas.filter(a=>a.fundacion===fk);
      const etiqueta = FUNDACION_INFO_LABEL(fk);
      if(variante==='pago-proveedores'){
        const filas = deFund.map(a=>{
          const cuenta = String(a.cbu||'').replace(/[^0-9]/g,'');
          return {
            'Modalidad de Pago': esCuentaMacro_(cuenta) ? '02' : '04', // 02: cuenta Macro, 04: otros bancos
            'Número de documento Beneficiario': String(a.cuit||'').replace(/[^0-9]/g,''),
            'Denominación del Beneficiario': textoXlsx_(a.proveedor).slice(0,120),
            // El banco exige que la identificación del pago sea ÚNICA: se agrega el código de esta exportación.
            'Identificación del pago': ('FC '+facturaNumeroCorto(a.numeroFactura)+'-'+sufijoId).slice(0,30),
            'Importe': Number(a.importe)||0,
            'Cuenta de pago - Nro/CBU o Nº Cheque': cuenta,
            'Marca de registración de cheque': '',
            'Fecha de pago - Emisión': fecha,
            'Fecha de pago diferido': ''
          };
        });
        archivos.push({ filas, headers: headersPP, nombre: 'PAGO_PROVEEDORES_MACRO_'+etiqueta+'_'+todayISO()+'.xlsx' });
      } else {
        const filasDe = (macro)=>deFund.filter(a=>esCuentaMacro_(String(a.cbu||'').replace(/[^0-9]/g,''))===macro).map(a=>({
          'Modalidad de Pago': macro ? '18' : '12', // 18: cuentas Macro, 12: otros bancos o billeteras
          'Número de documento Beneficiario': String(a.cuit||'').replace(/[^0-9]/g,''),
          'Denominación del Beneficiario': textoXlsx_(a.proveedor).slice(0,120),
          'Concepto': 'FAC',
          'Importe': Number(a.importe)||0,
          'Cuenta Destino': String(a.cbu||'').replace(/[^0-9]/g,''),
          'Fecha de pago - Emisión': fecha
        }));
        const aMacro = filasDe(true), aOtros = filasDe(false);
        if(aMacro.length) archivos.push({ filas:aMacro, headers: headersTM, nombre: 'TRANSF_MASIVAS_MACRO_'+etiqueta+'_CUENTAS_MACRO_'+todayISO()+'.xlsx' });
        if(aOtros.length) archivos.push({ filas:aOtros, headers: headersTM, nombre: 'TRANSF_MASIVAS_MACRO_'+etiqueta+'_OTROS_BANCOS_'+todayISO()+'.xlsx' });
      }
    });
    const ok = await generarOCancelar_(res, async ()=>{
      for(const a of archivos){ hojaDe_(a.filas, a.headers, 'Plantilla', a.nombre); await esperar_(500); } // el navegador puede bloquear descargas simultáneas
    });
    if(!ok) return;
    showToast('Banco Macro: '+archivos.length+' archivo(s) generado(s) ('+res.aprobadas.length+' pago(s); '+fundaciones.map(FUNDACION_INFO_LABEL).join(' y ')+'). Falta confirmar cuáles se subieron al banco.');
    await terminarExportacion_(res);
  }
  $('#btn-macro-pago-proveedores').addEventListener('click', ()=>exportarMacro_('pago-proveedores'));
  $('#btn-macro-transf-masivas').addEventListener('click', ()=>exportarMacro_('masivas'));

  $('#btn-transf-exportar-provincia').addEventListener('click', async ()=>{
    const sel = getTransfSeleccionadas();
    if(!sel.length){ showToast('Tildá al menos una factura para exportar.'); return; }
    // Provincia exige UN convenio (una empresa/CUIT) por archivo: el servidor arma un archivo por fundación
    // con lo que dejó guardado en la exportación (no con datos del navegador).
    const res = await prepararExportacionUI_('provincia', sel); if(!res) return;
    const fundaciones = Array.from(new Set(res.aprobadas.map(a=>a.fundacion)));
    const ok = await generarOCancelar_(res, async ()=>{ await descargarArchivosProvincia_(res.exportacionId, fundaciones); });
    if(!ok) return;
    showToast('Banco Provincia: archivo(s) generado(s) ('+fundaciones.map(FUNDACION_INFO_LABEL).join(' y ')+'). Falta confirmar cuáles se subieron al banco.');
    await terminarExportacion_(res);
  });
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
    const pvStr = String(pv||'').replace(/\D/g,'').replace(/^0+/,'').padStart(4,'0');
    const numStr = String(num||'').replace(/\D/g,'').replace(/^0+/,'').padStart(8,'0');
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

  /* Lectura del Excel de ARCA ("Mis Comprobantes Recibidos"): se buscan las columnas por el nombre de su
     encabezado (si ARCA agrega o mueve una columna no se corren los datos) y cada fila se valida. */
  const ARCA_COLUMNAS_ = [
    // [campo, encabezado, posición histórica, obligatoria]
    ['fecha',              /^fecha/i,                                   0, true],
    ['tipo',               /^tipo(\s+de)?\s+comprobante$|^tipo$/i,      1, true],
    ['puntoVenta',         /punto\s+de\s+venta/i,                       2, true],
    ['numeroDesde',        /n[uú]mero\s+desde|^n[uú]mero$/i,            3, true],
    ['numeroHasta',        /n[uú]mero\s+hasta/i,                        4, false],
    ['cae',                /c[oó]d\.?\s*autorizaci[oó]n|^cae$/i,        5, true],
    ['cuitEmisor',         /nro\.?\s*doc\.?\s*emisor|cuit\s+emisor/i,   7, true],
    ['denominacionEmisor', /denominaci[oó]n\s+emisor/i,                 8, true],
    ['cuitReceptor',       /nro\.?\s*doc\.?\s*receptor|cuit\s+receptor/i, 10, false],
    ['moneda',             /^moneda$/i,                                -1, false],
    ['importeTotal',       /^imp\.?\s*total|importe\s+total/i,         29, true]
  ];
  function arcaFechaIso_(v){
    if(v instanceof Date && !isNaN(v)) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
    if(typeof v==='number' && isFinite(v) && v>20000 && v<80000){ // número de serie de Excel
      const d = new Date(Math.round((v-25569)*86400000));
      return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+String(d.getUTCDate()).padStart(2,'0');
    }
    const t = String(v==null?'':v).trim();
    let m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
    if(m){ const d=+m[1], mo=+m[2]; const chk = new Date(+m[3], mo-1, d); if(mo<1||mo>12||chk.getFullYear()!==+m[3]||chk.getMonth()!==mo-1||chk.getDate()!==d) return ''; return m[3]+'-'+String(mo).padStart(2,'0')+'-'+String(d).padStart(2,'0'); }
    m = /^(\d{4})-(\d{2})-(\d{2})/.exec(t);
    return m ? (m[1]+'-'+m[2]+'-'+m[3]) : '';
  }
  function arcaMonto_(v){
    if(typeof v==='number') return isFinite(v) ? v : NaN;
    let t = String(v==null?'':v).replace(/[\s$]/g,'');
    if(!t) return NaN;
    if(t.indexOf(',')>-1) t = t.replace(/\./g,'').replace(',','.'); // 1.234,56 → 1234.56
    const n = Number(t);
    return isFinite(n) ? n : NaN;
  }
  function leerArcaDeFilas_(rows){
    const avisos = [], omitidas = [];
    // 1) fila de encabezados: una de las primeras 10 que nombre varias columnas conocidas
    let filaEnc = -1;
    for(let i=0; i<Math.min(rows.length,10) && filaEnc<0; i++){
      const celdas = (rows[i]||[]).map(c=>String(c==null?'':c).trim());
      const claves = [ARCA_COLUMNAS_[0], ARCA_COLUMNAS_[1], ARCA_COLUMNAS_[2], ARCA_COLUMNAS_[5]];
      if(claves.filter(c=>celdas.some(x=>c[1].test(x))).length>=3) filaEnc = i;
    }
    const idx = {};
    if(filaEnc<0){
      avisos.push('No se encontró la fila de encabezados: se usaron las posiciones de columna habituales. Verificá el resultado.');
      filaEnc = 1;
      ARCA_COLUMNAS_.forEach(c=>{ if(c[2]>=0) idx[c[0]] = c[2]; });
    } else {
      const celdas = (rows[filaEnc]||[]).map(c=>String(c==null?'':c).trim());
      ARCA_COLUMNAS_.forEach(c=>{
        const pos = celdas.findIndex(x=>c[1].test(x));
        if(pos>-1) idx[c[0]] = pos;
        else if(c[3] && c[2]>=0){ idx[c[0]] = c[2]; avisos.push('No se encontró la columna "'+c[0]+'": se usó la posición '+(c[2]+1)+'.'); }
      });
    }
    const get = (r, campo)=> (idx[campo]==null ? '' : r[idx[campo]]);
    const datos = [];
    rows.slice(filaEnc+1).forEach((r, k)=>{
      if(!r || !r.some(c=>String(c==null?'':c).trim()!=='')) return;
      const nFila = filaEnc+k+2; // número de fila tal como se ve en Excel
      const cae = String(get(r,'cae')).replace(/\D/g,'');
      if(!cae){ return; } // renglones de totales / vacíos: no son comprobantes
      const fecha = arcaFechaIso_(get(r,'fecha'));
      const cuit = String(get(r,'cuitEmisor')).replace(/\D/g,'');
      const importe = arcaMonto_(get(r,'importeTotal'));
      const moneda = String(get(r,'moneda')||'').trim();
      let motivo = '';
      if(!/^\d{14}$/.test(cae)) motivo = 'CAE inválido';
      else if(!fecha) motivo = 'fecha inválida';
      else if(cuit.length!==11) motivo = 'CUIT del emisor inválido';
      else if(isNaN(importe)) motivo = 'importe inválido';
      else if(moneda && !/^(PES|ARS|\$|PESOS?)$/i.test(moneda)) motivo = 'moneda distinta de pesos ('+moneda+')';
      if(motivo){ omitidas.push({fila:nFila, motivo}); return; }
      datos.push({
        fecha: fecha,
        tipo: String(get(r,'tipo')).trim(),
        puntoVenta: String(get(r,'puntoVenta')).trim(),
        numeroDesde: String(get(r,'numeroDesde')).trim(),
        numeroHasta: String(get(r,'numeroHasta')).trim(),
        cae: cae,
        cuitEmisor: String(get(r,'cuitEmisor')).trim(),
        denominacionEmisor: String(get(r,'denominacionEmisor')).trim(),
        cuitReceptor: String(get(r,'cuitReceptor')).trim(),
        importeTotal: importe
      });
    });
    return { datos, omitidas, avisos };
  }

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
      const lectura = leerArcaDeFilas_(rows);
      const datos = lectura.datos;
      if(lectura.omitidas.length || lectura.avisos.length){
        const det = lectura.avisos.concat(lectura.omitidas.slice(0,15).map(o=>'Fila '+o.fila+': '+o.motivo));
        const hayMas = lectura.omitidas.length>15 ? '\n... y '+(lectura.omitidas.length-15)+' más' : '';
        if(!confirm('Revisá antes de importar'+(lectura.omitidas.length? ' ('+lectura.omitidas.length+' fila(s) se van a omitir)':'')+':\n\n• '+det.join('\n• ')+hayMas+'\n\nSe importarían '+datos.length+' comprobante(s).\n\n¿Continuar?')){ statusEl.textContent='Importación cancelada.'; return; }
      }
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
    // Notas de crédito y tipos no reconocidos no se pueden cargar como factura a pagar.
    const noIncorporables = seleccionadas.filter(a=>{ const t = mapearTipoArca_(a.tipo); return t==='NC' || t===''; });
    if(noIncorporables.length){
      alert('Estos comprobantes NO se pueden incorporar como factura (son notas de crédito u otro tipo que el panel no reconoce). Cargalos a mano:\n\n• '+noIncorporables.map(a=>(a.denominacionEmisor||'—')+' — '+(a.tipo||'sin tipo')+' '+formatearPtoVtaNumero_(a.puntoVenta,a.numeroDesde)).join('\n• '));
    }
    const incorporables = seleccionadas.filter(a=>!noIncorporables.includes(a));
    if(!incorporables.length) return;
    arcaIncorporarLista = incorporables;
    arcaIncorporarDatos = incorporables.map(a=>{
      const match = proveedoresClasif.find(p=>normCuit_(p.cuit)===normCuit_(a.cuitEmisor));
      return { centroCosto:'', clasificacionFlujo: match? (match.clasificacion||'') : '', prioridad:'media', descripcion:'', pagada:false, medioDePago:'' };
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
    $('#arca-inc-pagada').disabled = !soyAdmin;
    $('#arca-inc-pagada').title = soyAdmin ? '' : 'Solo un administrador puede incorporar una factura ya pagada';
    $('#arca-inc-pagada').checked = soyAdmin && !!d.pagada;
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
      pagada: soyAdmin && $('#arca-inc-pagada').checked,
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
    const fallos = [];
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
        pagado: !!(soyAdmin && d.pagada),
        medioDePago: (soyAdmin && d.pagada) ? (d.medioDePago||'') : '',
        estadoInicial: (soyAdmin && d.pagada) ? 'completado' : 'pendiente',
        fechaRecibida: a.fecha || todayISO(),
        fechaResuelta: (soyAdmin && d.pagada) ? todayISO() : ''
      };
      rec.requestId = d.requestId = (d.requestId || nuevoRequestId_());
      try{ await runServer('addRecord', rec); ok++; }
      catch(err){ error++; fallos.push((a.denominacionEmisor||'—')+' '+formatearPtoVtaNumero_(a.puntoVenta,a.numeroDesde)+': '+(err&&err.message||'error')); console.error('Error incorporando comprobante ARCA', a, err); }
    }
    $('#arca-inc-confirmar-final').disabled = false;
    $('#arca-incorporar-overlay').classList.add('hidden');
    showToast('Incorporadas: '+ok+(error? ' · '+error+' con error':'')+'.');
    if(fallos.length) alert('No se pudieron incorporar:\n\n• '+fallos.join('\n• '));
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
