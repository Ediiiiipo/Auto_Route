// ============================================================
// SPX Auto Router v2 — Content Script
// Automação de Cálculos + GeoFixer integrado por CT
// ============================================================
console.log('🚀 SPX Auto Router v2 — Ready');

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const sleepH = async (min = 500, max = 1200) => sleep(Math.floor(Math.random() * (max - min + 1) + min));

function report(action, data = {}) {
  try { chrome.runtime.sendMessage({ action, ...data }); } catch(e) {}
}

// ============================================================
// GEOFIXER — Integrado
// ============================================================
const GFX = {
  state: { outliers: [], allOrders: [] },

  haversine(lat1, lng1, lat2, lng2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLng/2)**2;
    return R * 2 * Math.asin(Math.sqrt(a));
  },

  getTaskId() {
    const m = location.href.match(/taskId=([^&]+)/);
    return m ? m[1] : null;
  },

  // Verifica se está na tela do mapa (step 2 — "Ver pedidos no mapa")
  isOnMapStep() {
    const url = location.href;
    // Cobre: setting?taskId=... (step 2) e displayMap
    return (url.includes('setting?taskId=') || url.includes('displayMap')) && url.includes('taskId=');
  },

  sleep: ms => new Promise(r => setTimeout(r, ms)),

  // Busca pedidos via API
  async fetchOrders(taskId) {
    const url = `/api/spx/lmroute/adminapi/calculation_task/order/list?calculation_task_id=${taskId}&pageno=1&count=200`;
    const res = await fetch(url, { credentials: 'include' });
    const json = await res.json();
    if (json.retcode !== 0) throw new Error(`API list error: ${json.message}`);
    const raw = json.data?.list || json.data?.order_list || json.data || [];
    const arr = Array.isArray(raw) ? raw : Object.values(raw);
    return arr.map(item => ({
      id:      item.order_id || item.shipment_id,
      lat:     parseFloat(item.addr_lat ?? item.lat ?? 0),
      lng:     parseFloat(item.addr_lng ?? item.lng ?? 0),
      regAddr: item.reg_addr || '',
      address: [item.address, item.address_l3, item.address_l2, item.zipcode].filter(Boolean).join(', '),
    })).filter(o => o.id && o.lat !== 0);
  },

  // Detecta outliers usando mediana (resistente a distorções)
  detectOutliersAuto(orders) {
    if (orders.length < 3) return [];

    // Mediana das lats e lngs (mais robusta que média)
    const sortedLat = [...orders].sort((a, b) => a.lat - b.lat);
    const sortedLng = [...orders].sort((a, b) => a.lng - b.lng);
    const mid = Math.floor(orders.length / 2);
    const medLat = sortedLat[mid].lat;
    const medLng = sortedLng[mid].lng;

    // Distância de cada pedido até a mediana
    const dists = orders.map(o => GFX.haversine(o.lat, o.lng, medLat, medLng));

    // MAD: mediana dos desvios (Median Absolute Deviation)
    const sortedDists = [...dists].sort((a, b) => a - b);
    const medDist = sortedDists[mid];
    const mad = [...sortedDists.map(d => Math.abs(d - medDist))].sort((a, b) => a - b)[mid];

    // Threshold: mediana + 3x MAD, com mínimo de 5km e máximo de 15km
    const threshold = Math.min(Math.max(medDist + 3 * mad, 5), 15);

    console.log(`[GeoFixer] Mediana: ${medLat.toFixed(4)},${medLng.toFixed(4)} | medDist: ${medDist.toFixed(2)}km | MAD: ${mad.toFixed(2)}km | threshold: ${threshold.toFixed(2)}km`);

    return orders
      .filter((o, i) => dists[i] > threshold)
      .map(o => ({ ...o, status: 'pending', distToCentroid: dists[orders.indexOf(o)] }));
  },

  // Inicializa geocoder-bridge
  initGeocoder() {
    return new Promise(resolve => {
      if (window.__gfxBridgeReady) return resolve();
      if (document.getElementById('gfx-geocoder-bridge')) {
        const wait = setInterval(() => { if (window.__gfxBridgeReady) { clearInterval(wait); resolve(); } }, 50);
        setTimeout(() => { clearInterval(wait); resolve(); }, 3000);
        return;
      }
      const script = document.createElement('script');
      script.id  = 'gfx-geocoder-bridge';
      script.src = chrome.runtime.getURL('geocoder-bridge.js');
      script.onload = () => {
        const wait = setInterval(() => { if (window.__gfxBridgeReady) { clearInterval(wait); resolve(); } }, 50);
        setTimeout(() => { clearInterval(wait); resolve(); }, 3000);
      };
      document.head.appendChild(script);
    });
  },

  // Geocodifica endereço via google.maps já carregado pelo SPX
  async geocodeAddress(regAddr) {
    await GFX.initGeocoder();
    return new Promise((resolve, reject) => {
      const cleaned    = regAddr.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
      const callbackId = '__gfxGeo_' + Date.now();
      const handler = e => {
        window.removeEventListener(callbackId, handler);
        if (e.detail.error) reject(new Error(e.detail.error));
        else resolve(e.detail);
      };
      window.addEventListener(callbackId, handler);
      window.dispatchEvent(new CustomEvent('__gfxGeocodeRequest', { detail: { callbackId, address: cleaned } }));
      setTimeout(() => { window.removeEventListener(callbackId, handler); reject(new Error('Geocoding timeout')); }, 10000);
    });
  },

  // Corrige um pedido via API
  async modifyOrder(taskId, shipmentId, lat, lng) {
    const res = await fetch('/api/spx/lmroute/adminapi/calculation_task/order/modify', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        calculation_task_id: taskId,
        order_list: [{ shipment_id: shipmentId, verified_lat: lat, verified_lng: lng }]
      })
    });
    const json = await res.json();
    if (json.retcode !== 0) throw new Error(`modify error: ${json.message}`);
    return json;
  },

  // ─── PAINEL INLINE ───────────────────────────────────────────
  showPanel(outliers) {
    document.getElementById('gfx-inline-panel')?.remove();
    const total = outliers.length;

    const itemsHtml = outliers.map(o => `
      <div class="gfx-item" data-gfx-id="${o.id}">
        <div class="gfx-item-dot ${o.status}"></div>
        <div class="gfx-item-info">
          <div class="gfx-item-id">${o.id}</div>
          <div class="gfx-item-dist">${o.distToCentroid?.toFixed(1)}km · ${o.address.substring(0,36)}…</div>
        </div>
        <div class="gfx-item-status ${o.status}">${{pending:'⏳',processing:'⏳',done:'✓',error:'✗'}[o.status]||''}</div>
      </div>`).join('');

    const panel = document.createElement('div');
    panel.id = 'gfx-inline-panel';
    panel.innerHTML = `
      <div class="gfx-header">
        <div class="gfx-header-icon">📍</div>
        <div>
          <div class="gfx-header-title">GeoFixer <span style="font-size:10px;opacity:.5">integrado</span></div>
          <div class="gfx-header-sub">${total} outlier(s) detectado(s) — corrigindo...</div>
        </div>
      </div>
      <div class="gfx-list" style="max-height:200px">
        <div class="gfx-list-header">Corrigindo pontos fora do lugar</div>
        ${itemsHtml}
      </div>
      <div class="gfx-log info" id="gfx-inline-log">Iniciando geocoding...</div>
    `;
    document.body.appendChild(panel);
  },

  updatePanelItem(id, status, logMsg) {
    const item = document.querySelector(`[data-gfx-id="${id}"]`);
    if (item) {
      const dot    = item.querySelector('.gfx-item-dot');
      const st     = item.querySelector('.gfx-item-status');
      if (dot) dot.className = `gfx-item-dot ${status}`;
      if (st)  { st.className = `gfx-item-status ${status}`; st.textContent = {pending:'⏳',processing:'⏳',done:'✓',error:'✗'}[status]||''; }
    }
    if (logMsg) {
      const log = document.getElementById('gfx-inline-log');
      if (log) log.textContent = logMsg;
    }
  },

  hidePanel() {
    document.getElementById('gfx-inline-panel')?.remove();
  },

  // ─── FLUXO PRINCIPAL ─────────────────────────────────────────
  // Chamado após cada CT ser calculada.
  // Retorna true se tudo ok (sem outliers ou tudo corrigido)
  async runForCT(taskId) {
    console.log(`[GeoFixer] Analisando CT: ${taskId}`);

    let orders;
    try {
      orders = await GFX.fetchOrders(taskId);
    } catch(e) {
      console.warn('[GeoFixer] Erro ao buscar pedidos:', e.message);
      return true; // Não bloqueia o fluxo
    }

    const outliers = GFX.detectOutliersAuto(orders);

    if (outliers.length === 0) {
      console.log('[GeoFixer] ✅ Nenhum outlier nesta CT');
      return true;
    }

    console.log(`[GeoFixer] ⚠️ ${outliers.length} outlier(s) — corrigindo...`);
    report('geofix_start', { count: outliers.length });

    // Mostra painel inline
    GFX.showPanel(outliers);
    await GFX.sleep(500);

    let fixed = 0, errors = 0;

    for (const order of outliers) {
      GFX.updatePanelItem(order.id, 'processing', `Geocodificando ${order.id}...`);
      try {
        const coords = await GFX.geocodeAddress(order.regAddr);
        await GFX.modifyOrder(taskId, order.id, coords.lat, coords.lng);
        GFX.updatePanelItem(order.id, 'done', `✓ ${order.id} corrigido → ${coords.lat.toFixed(5)},${coords.lng.toFixed(5)}`);
        console.log(`[GeoFixer] ✓ ${order.id} → ${coords.lat},${coords.lng}`);
        fixed++;
      } catch(e) {
        GFX.updatePanelItem(order.id, 'error', `✗ ${order.id}: ${e.message}`);
        console.error(`[GeoFixer] ✗ ${order.id}:`, e.message);
        errors++;
      }
      await GFX.sleep(400);
    }

    report('geofix_done', { fixed, errors });

    // Mostra resultado brevemente antes de fechar
    const logEl = document.getElementById('gfx-inline-log');
    if (logEl) {
      logEl.className = errors > 0 ? 'gfx-log warn' : 'gfx-log ok';
      logEl.textContent = `✅ ${fixed} corrigidos${errors > 0 ? `, ${errors} erro(s)` : ''} — continuando...`;
    }
    await GFX.sleep(2000);
    GFX.hidePanel();
    return true;
  }
};

// ============================================================
// AUTOMAÇÃO PRINCIPAL
// ============================================================
async function executarAutomacao(config, vehicles) {
  // Garante default true se não enviado
  if (config.geoFixer === undefined) config.geoFixer = true;
  console.log('🚀 Automação v2:', { config, vehicles });
  console.log('📍 GeoFixer ativado?', config.geoFixer);
  try {
    report('progress', { current: 0, total: vehicles.length, message: 'Iniciando cálculos...' });
    await sleep(1000);

    // Clicar em "Cálculo Pendente"
    for (const sp of document.querySelectorAll('span')) {
      if (sp.textContent.includes('Cálculo Pendente')) { sp.click(); await sleep(1500); break; }
    }
    await sleep(1000);

    for (let i = 0; i < vehicles.length; i++) {
      const v = vehicles[i];
      console.log(`>>> CICLO ${i+1}/${vehicles.length} | ${v.type}: ${v.name} (SPR ${v.spr})`);
      report('progress', { current: i, total: vehicles.length, message: `Processando ${v.type}: ${v.name}...` });

      try {
        // Esperar estar na lista
        for (let w = 0; w < 30; w++) {
          if (location.href.includes('#/lmRouteCalculationPool') && !location.href.includes('setting')) break;
          await sleep(200);
        }
        await sleep(500);

        // Clicar Calcular (primeira linha da tabela)
        const btnCalc = document.evaluate(
          "(//table/tbody[2]/tr)[1]//button[contains(., 'Calcular')]",
          document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null
        ).singleNodeValue;
        if (!btnCalc) { console.log('❌ Calcular não encontrado, pulando'); continue; }
        btnCalc.click(); await sleepH(1000, 1500);

        // 1º Próximo → vai para step 2 (Ver pedidos no mapa)
        let btnP = document.evaluate("//button[contains(normalize-space(), 'Próximo')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (btnP) { btnP.click(); await sleepH(800, 1200); }

        // ─── GEOFIXER: roda no step 2 — "Ver pedidos no mapa" ────
        // Neste momento a URL é: setting?taskId=CTBR...&station_id=...
        try {
          // Aguarda URL do mapa estar pronta
          for (let w = 0; w < 50; w++) {
            if (GFX.isOnMapStep()) break;
            await sleep(300);
          }
          await sleep(1200); // Pedidos carregarem no mapa

          const taskId = GFX.getTaskId();
          console.log('[GeoFixer] URL atual:', location.href);
          console.log('[GeoFixer] taskId detectado:', taskId);

          if (!config.geoFixer) {
            console.log('[GeoFixer] Desativado pelo usuário — pulando');
          } else if (taskId) {
            report('progress', { current: i, total: vehicles.length, message: `📍 GeoFixer: verificando CT ${i+1}...` });
            await GFX.runForCT(taskId);
          } else {
            console.warn('[GeoFixer] taskId não encontrado na URL:', location.href);
          }
        } catch(geoErr) {
          console.warn('[GeoFixer] Erro não crítico:', geoErr.message);
        }
        // ─────────────────────────────────────────────────────────

        // 2º Próximo → vai para step 3 (Configuração de Parâmetros)
        await sleep(300);
        const snap = document.evaluate("//button[contains(normalize-space(), 'Próximo')]", document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        if (snap.snapshotLength > 0) { const b = snap.snapshotItem(0); if (b?.offsetParent) { b.click(); await sleepH(800, 1200); } }

      } catch(e) { console.log('Erro fluxo:', e.message); continue; }

      // Preenche dados da CT (passo 3 do wizard SPX)
      await preencherDados({ data: config.date, shift: config.shift, txt1: '11', txt2: '59', veiculo: v.type, spr: v.spr });

      report('progress', { current: i + 1, total: vehicles.length, message: `✅ ${v.type}: ${v.name} concluído` });
    }

    report('complete');
    console.log('✅ Automação v2 concluída!');
  } catch(e) {
    console.error('❌', e);
    report('error', { error: String(e.message || e) });
  }
}

// ============================================================
// PREENCHIMENTO DE DADOS (original — sem alterações)
// ============================================================
async function preencherDados(d) {
  // DATA
  try {
    const cd = document.evaluate("//div[contains(@class,'ssc-form-item')][.//label[contains(.,'Data')]]//div[contains(@class,'ssc-select')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (cd) { cd.click(); await sleep(300); const o = document.evaluate(`//li[normalize-space()='${d.data}']`, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if (o) { o.click(); await sleep(300); } }
  } catch(e){}

  // SHIFT
  try {
    if (d.shift) {
      const cs = document.evaluate("//div[contains(@class,'ssc-form-item')][.//label[contains(.,'Shift') or contains(.,'Turno')]]//div[contains(@class,'ssc-select')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
      if (cs) { cs.click(); await sleep(300); let o = document.evaluate(`//li[normalize-space()='${d.shift}']`, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if (!o) o = document.evaluate(`//li[normalize-space()='${d.shift}-1']`, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if (o) { o.click(); await sleep(300); } }
    }
  } catch(e){}

  // TEMPOS
  try {
    const ct = document.evaluate('//*[@id="fms-container"]/div[2]/div[2]/div/div[2]/div/div[2]/div/form/div[7]/div', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (ct) { const ins = ct.querySelectorAll('input'); if (ins[0]&&d.txt1) { ins[0].click(); await sleep(50); ins[0].select(); ins[0].value=d.txt1; ins[0].dispatchEvent(new Event('input',{bubbles:true})); await sleep(100); } if (ins[1]&&d.txt2) { ins[1].click(); await sleep(50); ins[1].select(); ins[1].value=d.txt2; ins[1].dispatchEvent(new Event('input',{bubbles:true})); await sleep(200); } }
  } catch(e){}

  // CHECKBOX todos
  try { const cbs = document.evaluate("//thead//input[@type='checkbox']", document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null); if (cbs.snapshotLength > 0) { cbs.snapshotItem(0).click(); await sleep(300); } } catch(e){}

  // VEÍCULO
  let row = null, scroll = document.querySelector('.ssc-dialog .ssc-table-body') || document.querySelector('.ssc-table-body');
  for (let s = 0; s < 40; s++) {
    try { const r = document.evaluate(`//tr[contains(.,'${d.veiculo}')]`, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null); for (let j=0; j<r.snapshotLength; j++) { const el = r.snapshotItem(j); if (el?.offsetParent) { row = el; break; } } } catch(e){}
    if (row) break;
    if (scroll) { scroll.scrollTop += 400; await sleep(300); } else break;
  }

  if (row) {
    row.scrollIntoView({ block:'center' }); await sleep(300);
    try { const chk = row.querySelector("input[type='checkbox']"); if (chk) { chk.click(); await sleep(200); } } catch(e){}

    const setCol = async (idx) => {
      try { const td = document.evaluate(`.//td[${idx}]`, row, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if (!td) return; (td.querySelector('.ssc-select-content')||td.querySelector('.ssc-select')||td).click(); await sleep(200); const opt = document.querySelector('body > span > div > div > div > ul > li:nth-child(3)'); if (opt) { for (let i=0;i<10;i++){if(opt.offsetParent)break;await sleep(100);} opt.click(); await sleep(200); } } catch(e){}
    };
    await setCol(5); await setCol(6); await setCol(7);

    // SPR
    try { const inp = document.evaluate('.//td[8]//input', row, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if (inp) { inp.click(); inp.select(); inp.value = String(d.spr); inp.dispatchEvent(new Event('input',{bubbles:true})); await sleep(200); } } catch(e){}
  }

  // CALCULAR
  try {
    const bs = document.evaluate("//button[contains(normalize-space(),'Calcular')]", document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
    for (let b=bs.snapshotLength-1; b>=0; b--) { const btn = bs.snapshotItem(b); if (btn?.offsetParent) { btn.click(); break; } }
    await sleep(1000);
    const pop = document.querySelector('body > div.ssc-dialog > div.ssc-dialog-wrapper > div > div > div > button');
    if (pop) { pop.click(); await sleepH(1000, 2000); }
  } catch(e){}
}

// ============================================================
// LISTENER
// ============================================================
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'startAutomation') {
    executarAutomacao(msg.config, msg.vehicles);
    sendResponse({ ok: true });
  }
  return true;
});
