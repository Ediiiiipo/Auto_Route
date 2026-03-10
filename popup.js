// ============================================================
// SPX Auto Router — Popup (Lógica Principal)
// ============================================================
const $ = id => document.getElementById(id);
let isRunning = false;
let templateData = null; // Parsed template data
let templateBuffer = null; // Raw ArrayBuffer do arquivo carregado
let ctList = []; // Final CT list for execution
let stationNameCache = '—'; // Station name fetched from SPX tab

// ---- Logging ----
function log(msg, type = 'info') {
  const t = new Date().toLocaleTimeString('pt-BR');
  const el = document.createElement('div');
  el.className = `log-entry ${type}`;
  el.innerHTML = `<span class="log-time">[${t}]</span>${msg}`;
  $('logBody').appendChild(el);
  $('logBody').scrollTop = $('logBody').scrollHeight;
  console.log(`[${type}] ${msg}`);
}

// ---- Phase Management ----
function showPhase(n) {
  document.querySelectorAll('.phase').forEach(p => p.classList.remove('active'));
  $(`phase${n}`).classList.add('active');
}

// ---- SPX Tab Detection ----
async function getSPXTab() {
  const tabs = await chrome.tabs.query({ url: 'https://spx.shopee.com.br/*' });
  return tabs[0] || null;
}

// ---- Connection Check ----
async function checkConnection() {
  try {
    const tab = await getSPXTab();
    if (tab) {
      $('statusDot').classList.add('ok');
      $('statusText').textContent = 'Conectado ao SPX';
      // Fetch station name asynchronously (non-blocking)
      runInTab(getStationName).then(name => {
        if (name) {
          stationNameCache = name;
          $('stationLabel').textContent = name;
          document.title = `SPX Auto Router — ${name}`;
        }
      }).catch(() => {});
      return true;
    }
  } catch(e) {}
  $('statusDot').classList.remove('ok');
  $('statusText').textContent = 'Abra o SPX (spx.shopee.com.br) primeiro';
  return false;
}

// ---- Run function in page MAIN world ----
async function runInTab(fn, args = []) {
  const tab = await getSPXTab();
  if (!tab) throw new Error('Aba SPX não encontrada');
  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    world: 'MAIN',
    func: fn,
    args: args
  });
  return results?.[0]?.result;
}

// ============================================================
// TEMPLATE READING
// ============================================================
function parseTemplate(arrayBuffer, filterComerciais = false) {
  // Read workbook in popup context (SheetJS bundled)
  const data = new Uint8Array(arrayBuffer);
  const wb = XLSX.read(data, { type: 'array' });

  // ---- Read Planejamento ----
  const planSheet = wb.Sheets['Planejamento'];
  if (!planSheet) throw new Error('Aba "Planejamento" não encontrada no template');

  // SPR Moto global (M4) and SPR Volumoso (M5)
  const sprMotoGlobal = planSheet['M4']?.v || 95;
  const sprVolGlobal = planSheet['M5']?.v || 25;

  // Read clusters (row 8 onwards until "Dados Gerais")
  const clusters = [];
  for (let r = 8; r <= 100; r++) {
    const cellA = planSheet[`A${r}`];
    if (!cellA || !cellA.v) break;
    const clusterName = String(cellA.v).trim();
    if (clusterName === 'Dados Gerais' || clusterName === '') break;

    const pedidos = planSheet[`B${r}`]?.v || 0;
    const rotas = planSheet[`C${r}`]?.v || 0;
    const sprRaw = planSheet[`D${r}`]?.v;
    const spr = (typeof sprRaw === 'string' && sprRaw.includes('Baixo ADO')) ? 'Baixo ADO' : (sprRaw || 0);
    const sprMin = planSheet[`E${r}`]?.v || 50;
    const sprMax = planSheet[`F${r}`]?.v || 105;
    const ord = planSheet[`G${r}`]?.v || r - 7;
    const dispMoto = planSheet[`H${r}`]?.v || 0;
    const sepMoto = planSheet[`I${r}`]?.v || 0;
    const rotasMoto = planSheet[`J${r}`]?.v || 0;
    const dispVol = planSheet[`K${r}`]?.v || 0;
    const sepVol = planSheet[`L${r}`]?.v || 0;
    const rotasVol = planSheet[`M${r}`]?.v || 0;

    clusters.push({
      name: clusterName, pedidos, rotas, spr, sprMin, sprMax,
      ord: parseInt(ord) || r - 7,
      dispMoto, sepMoto, rotasMoto,
      dispVol, sepVol, rotasVol
    });
  }

  // ---- Read Base ----
  const baseSheet = wb.Sheets['Base'];
  if (!baseSheet) throw new Error('Aba "Base" não encontrada no template');

  const baseData = XLSX.utils.sheet_to_json(baseSheet, { header: 1 });
  const headers = baseData[0] || [];

  // Find column indices
  const colIdx = {};
  const colMap = { 'Shipment Id': 'shipmentId', 'Perfil': 'perfil', 'Clusterização': 'cluster',
    'Validação Arquivos': 'valArq', 'Validação TO': 'valTO', 'Validação Office': 'valOffice', 'LH Trip': 'lhTrip',
    'Tipo de Pedido': 'tipoPedido' };
  headers.forEach((h, i) => {
    const key = colMap[String(h).trim()];
    if (key) colIdx[key] = i;
  });

  // Fallback: try column letters if headers don't match
  if (colIdx.shipmentId === undefined) colIdx.shipmentId = 0;  // A
  if (colIdx.perfil === undefined) colIdx.perfil = 17;          // R
  if (colIdx.cluster === undefined) colIdx.cluster = 21;        // V
  if (colIdx.valArq === undefined) colIdx.valArq = 27;          // AB
  if (colIdx.valTO === undefined) colIdx.valTO = 28;            // AC
  if (colIdx.valOffice === undefined) colIdx.valOffice = 29;    // AD
  if (colIdx.lhTrip === undefined) colIdx.lhTrip = 8;          // I
  if (colIdx.tipoPedido === undefined) colIdx.tipoPedido = 7;   // H

  // Group shipment IDs by cluster+perfil, applying filters
  const shipments = {}; // { clusterName: { MOTO: [...ids], Passeio: [...ids], Volumoso: [...ids] } }
  let totalFiltered = 0;
  const comerciaisRemovidos = []; // Pedidos removidos pelo filtro Office/Commercial
  const lhTrips = {}; // { tripId: count } — LH trips com pedidos roteirizados
  let backlogTotal = 0; // Total de pedidos com LH Trip = BACKLOG

  for (let r = 1; r < baseData.length; r++) {
    const row = baseData[r];
    const sid = row[colIdx.shipmentId];
    if (!sid) continue;

    // Apply exclusion filters
    const valArq = row[colIdx.valArq] ? String(row[colIdx.valArq]).trim() : '';
    const valTO = row[colIdx.valTO] ? String(row[colIdx.valTO]).trim() : '';
    const valOffice = row[colIdx.valOffice] ? String(row[colIdx.valOffice]).trim() : '';
    const tipoPedido = row[colIdx.tipoPedido] ? String(row[colIdx.tipoPedido]).trim().toUpperCase() : '';

    if (valArq === 'Pedidos já reservados') { totalFiltered++; continue; }
    if (valTO === 'TO Removida') { totalFiltered++; continue; }
    if (valOffice === 'Offices Retirados') { totalFiltered++; continue; }

    // Filtro Comerciais (coluna H): Office, Return to Seller, Service Point
    if (filterComerciais) {
      const TIPOS_COMERCIAIS = ['OFFICE', 'RETURN TO SELLER', 'SERVICE POINT'];
      if (TIPOS_COMERCIAIS.some(t => tipoPedido.includes(t))) {
        const cluster = row[colIdx.cluster] ? String(row[colIdx.cluster]).trim() : '';
        comerciaisRemovidos.push({
          id: String(sid),
          tipo: row[colIdx.tipoPedido] ? String(row[colIdx.tipoPedido]).trim() : '',
          cluster,
          perfil: row[colIdx.perfil] ? String(row[colIdx.perfil]).trim() : ''
        });
        totalFiltered++;
        continue;
      }
    }

    const cluster = row[colIdx.cluster] ? String(row[colIdx.cluster]).trim() : '';
    const perfil = row[colIdx.perfil] ? String(row[colIdx.perfil]).trim() : '';
    const lhTrip = row[colIdx.lhTrip] ? String(row[colIdx.lhTrip]).trim() : '';

    if (!cluster || !perfil) continue;

    if (!shipments[cluster]) shipments[cluster] = { MOTO: [], Passeio: [], Volumoso: [], MOTO_BACKLOG: [] };

    // Track LH trips
    if (lhTrip.toUpperCase() === 'BACKLOG') {
      backlogTotal++;
    } else if (lhTrip) {
      lhTrips[lhTrip] = (lhTrips[lhTrip] || 0) + 1;
    }

    // For MOTO, also check BACKLOG
    if (perfil.toUpperCase() === 'MOTO') {
      if (lhTrip.toUpperCase() === 'BACKLOG') {
        shipments[cluster].MOTO_BACKLOG.push(String(sid));
      } else {
        shipments[cluster].MOTO.push(String(sid));
      }
    } else if (perfil === 'Passeio') {
      shipments[cluster].Passeio.push(String(sid));
    } else if (perfil === 'Volumoso') {
      shipments[cluster].Volumoso.push(String(sid));
    }
  }

  return { clusters, shipments, sprMotoGlobal, sprVolGlobal, totalFiltered, comerciaisRemovidos, lhTrips, backlogTotal };
}

// ============================================================
// ============================================================
// CALCULO DE ROTAS E SPR (replica fórmula do Excel)
// =ARREDONDAR.PARA.CIMA(SE(pedidos<sprMin;"-";MÁXIMO(1;MÍNIMO(INT(pedidos/sprMax)+SE(MOD(pedidos,sprMax)>0;1;0);pedidos/sprMin)));0)
// ============================================================
function calcRotasSpr(pedidos, sprMin, sprMax) {
  if (pedidos < sprMin) return { rotas: 0, spr: 'Baixo ADO', isBaixo: true };
  const rotasRaw = Math.min(
    Math.floor(pedidos / sprMax) + (pedidos % sprMax > 0 ? 1 : 0),
    pedidos / sprMin
  );
  const rotas = Math.max(1, Math.ceil(rotasRaw));
  const spr = Math.ceil(pedidos / rotas);
  return { rotas, spr, isBaixo: false };
}

// ============================================================
// BUILD CT LIST
// ============================================================
function buildCTList(data) {
  const cts = [];

  for (const cluster of data.clusters) {
    const sh = data.shipments[cluster.name] || { MOTO: [], Passeio: [], Volumoso: [], MOTO_BACKLOG: [] };

    // IDs reserved for Moto
    const motoReservedIds = new Set();
    // IDs reserved for Volumoso
    const volReservedIds = new Set();

    // CT Moto — if Sep. Moto > 0
    if (cluster.sepMoto > 0 && cluster.rotasMoto > 0) {
      const motoIds = sh.MOTO.slice(0, cluster.sepMoto);
      motoIds.forEach(id => motoReservedIds.add(id));

      // Recalculate SPR Moto based on actual IDs
      const sprMotoMax = data.sprMotoGlobal;
      const sprMotoMin = Math.round(sprMotoMax * 0.6); // approx min for moto
      const motoCalc = calcRotasSpr(motoIds.length, sprMotoMin, sprMotoMax);

      cts.push({
        type: 'MOTO',
        typeLabel: '🏍️ Moto',
        cluster: cluster.name,
        ord: cluster.ord,
        spr: motoIds.length > 0 ? Math.ceil(motoIds.length / cluster.rotasMoto) : data.sprMotoGlobal,
        rotas: cluster.rotasMoto,
        ids: motoIds,
        enabled: true,
        isBaixoADO: false
      });
    }

    // CT Volumoso — if Sep. VOL > 0
    if (cluster.sepVol > 0 && cluster.rotasVol > 0) {
      const volIds = sh.Volumoso.slice(0, cluster.sepVol);
      volIds.forEach(id => volReservedIds.add(id));

      cts.push({
        type: 'FIORINO',
        typeLabel: '🚐 Vol',
        cluster: cluster.name,
        ord: cluster.ord,
        spr: volIds.length > 0 ? Math.ceil(volIds.length / cluster.rotasVol) : data.sprVolGlobal,
        rotas: cluster.rotasVol,
        ids: volIds,
        enabled: true,
        isBaixoADO: false
      });
    }

    // CT Passeio — all remaining IDs (Passeio + leftover MOTO not reserved + leftover Volumoso)
    const passeioIds = [];
    sh.Passeio.forEach(id => passeioIds.push(id));
    sh.MOTO.forEach(id => { if (!motoReservedIds.has(id)) passeioIds.push(id); });
    sh.MOTO_BACKLOG.forEach(id => passeioIds.push(id));
    sh.Volumoso.forEach(id => { if (!volReservedIds.has(id)) passeioIds.push(id); });

    // Recalculate rotas/SPR with actual IDs using Excel formula
    const calc = calcRotasSpr(passeioIds.length, cluster.sprMin, cluster.sprMax);

    if (calc.rotas > 0 || cluster.rotas > 0 || calc.isBaixo) {
      cts.push({
        type: 'PASSEIO',
        typeLabel: '🚗 Passeio',
        cluster: cluster.name,
        ord: cluster.ord,
        spr: calc.isBaixo ? 0 : calc.spr,
        rotas: calc.isBaixo ? 0 : calc.rotas,
        ids: passeioIds,
        enabled: !calc.isBaixo,
        isBaixoADO: calc.isBaixo,
        // Keep original for reference in UI
        sprOriginal: cluster.spr,
        rotasOriginal: cluster.rotas,
        recalculado: (calc.rotas !== cluster.rotas || calc.spr !== cluster.spr)
      });
    }
  }

  // Sort: MOTO first (by ord), then PASSEIO (by ord), then FIORINO (by ord)
  // This is the DISPLAY order in the preview table
  const typeOrder = { 'MOTO': 0, 'PASSEIO': 1, 'FIORINO': 2 };
  cts.sort((a, b) => (typeOrder[a.type] - typeOrder[b.type]) || (a.ord - b.ord));

  return cts;
}

// ============================================================
// RENDER PREVIEW TABLE
// ============================================================
function renderPreview() {
  const tbody = $('ctTableBody');
  tbody.innerHTML = '';

  let totalCTs = 0, totalIDs = 0, motos = 0, passeios = 0, vols = 0;

  ctList.forEach((ct, idx) => {
    const tr = document.createElement('tr');
    if (!ct.enabled) tr.classList.add('disabled');

    const typeClass = ct.type === 'MOTO' ? 'type-moto' : ct.type === 'PASSEIO' ? 'type-passeio' : 'type-vol';

    const recalcTitle = ct.recalculado ? ` title="Recalculado (original: ${ct.rotasOriginal} rotas / SPR ${ct.sprOriginal})"` : '';
    const rotasCell = ct.recalculado
      ? `<span style="color:#f59e0b;" ${recalcTitle}>⚡ ${ct.rotas}</span>`
      : ct.rotas;

    tr.innerHTML = `
      <td><input type="checkbox" data-idx="${idx}" ${ct.enabled ? 'checked' : ''}></td>
      <td class="${typeClass}">${ct.typeLabel}</td>
      <td>${ct.cluster}</td>
      <td>${ct.ids.length}</td>
      <td>${ct.isBaixoADO ? '<span class="baixo-ado">Baixo ADO</span>' :
        `<input type="number" value="${ct.spr}" min="1" data-spr-idx="${idx}" style="width:55px;">`}</td>
      <td>${rotasCell}</td>
    `;
    tbody.appendChild(tr);

    if (ct.enabled) {
      totalCTs++;
      totalIDs += ct.ids.length;
      if (ct.type === 'MOTO') motos++;
      else if (ct.type === 'PASSEIO') passeios++;
      else vols++;
    }
  });

  // Summary badges
  $('ctSummary').innerHTML = `
    <span class="ct-pill total">${totalCTs} CTs</span>
    <span class="ct-pill total">${totalIDs.toLocaleString()} IDs</span>
    ${motos ? `<span class="ct-pill moto">🏍️ ${motos} Moto</span>` : ''}
    ${passeios ? `<span class="ct-pill passeio">🚗 ${passeios} Passeio</span>` : ''}
    ${vols ? `<span class="ct-pill vol">🚐 ${vols} Vol</span>` : ''}
  `;

  // Checkbox listeners
  tbody.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', (e) => {
      const i = parseInt(e.target.dataset.idx);
      ctList[i].enabled = e.target.checked;
      renderPreview();
    });
  });

  // SPR edit listeners
  tbody.querySelectorAll('input[type="number"]').forEach(inp => {
    inp.addEventListener('change', (e) => {
      const i = parseInt(e.target.dataset.sprIdx);
      ctList[i].spr = parseInt(e.target.value) || 0;
    });
  });
}

// ============================================================
// IMPORT: Create CTs + Send Shipment IDs
// ============================================================
function createEmptyCT(stationId) {
  return new Promise(async (resolve) => {
    try {
      const getCookie = (n) => { const m = document.cookie.match(new RegExp('(?:^|;\\s*)' + n + '=([^;]*)')); return m ? m[1] : ''; };
      const res = await fetch('/spx_delivery/admin/delivery/route/smart/routing/push/order/calculate', {
        method: 'POST',
        headers: { 'accept': 'application/json, text/plain, */*', 'content-type': 'application/json;charset=UTF-8',
          'app': 'FMS Portal', 'x-csrftoken': getCookie('csrftoken'), 'device-id': getCookie('spx-admin-device-id') },
        credentials: 'include',
        body: JSON.stringify({ shipment_id_list: [], dest_station_id: stationId, push_type: 1 })
      });
      const data = await res.json();
      if (data.retcode !== 0) return resolve({ success: false, error: data.message || `retcode: ${data.retcode}` });
      resolve({ success: true, ctId: String(data.data?.calculation_task_id || '') });
    } catch(e) { resolve({ success: false, error: e.message }); }
  });
}

function addShipmentsToCT(shipmentIds, stationId, ctId) {
  return new Promise(async (resolve) => {
    try {
      const getCookie = (n) => { const m = document.cookie.match(new RegExp('(?:^|;\\s*)' + n + '=([^;]*)')); return m ? m[1] : ''; };
      const res = await fetch('/spx_delivery/admin/delivery/route/smart/routing/push/order/calculate', {
        method: 'POST',
        headers: { 'accept': 'application/json, text/plain, */*', 'content-type': 'application/json;charset=UTF-8',
          'app': 'FMS Portal', 'x-csrftoken': getCookie('csrftoken'), 'device-id': getCookie('spx-admin-device-id') },
        credentials: 'include',
        body: JSON.stringify({ shipment_id_list: shipmentIds, dest_station_id: stationId, push_type: 1, calculation_task_id: ctId })
      });
      const data = await res.json();
      if (data.retcode !== 0) return resolve({ success: false, error: data.message || `retcode: ${data.retcode}` });
      resolve({ success: true, ctId: String(data.data?.calculation_task_id || ctId) });
    } catch(e) { resolve({ success: false, error: e.message }); }
  });
}

function getStationId() {
  try {
    const raw = localStorage.getItem('station_id') || localStorage.getItem('stationId');
    if (raw) return parseInt(raw);
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      const v = localStorage.getItem(k);
      if (k.toLowerCase().includes('station') && /^\d+$/.test(v)) return parseInt(v);
    }
  } catch(e) {}
  return null;
}

function getStationName() {
  try {
    // Try common localStorage keys
    const candidates = ['station_name', 'stationName', 'hub_name', 'hubName', 'lm_hub_name'];
    for (const k of candidates) {
      const v = localStorage.getItem(k);
      if (v && v.length > 1) return v;
    }
    // Try to find any key with "station" + "name" in it
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k.toLowerCase().includes('station') && k.toLowerCase().includes('name')) {
        const v = localStorage.getItem(k);
        if (v && v.length > 1 && v.length < 80) return v;
      }
    }
    // Fallback: try to read from page DOM (breadcrumb / header)
    const el = document.querySelector('[class*="station-name"], [class*="hub-name"], [data-station], .station-info');
    if (el?.textContent?.trim()) return el.textContent.trim().slice(0, 60);
  } catch(e) {}
  return null;
}

async function executeImport(enabledCTs) {
  log('Iniciando importação...', 'info');

  // Get station ID
  const stationId = await runInTab(getStationId);
  if (!stationId) { log('❌ Station ID não encontrado. Abra o SPX primeiro.', 'error'); return false; }
  log(`Station ID: ${stationId}`, 'info');

  // CREATION ORDER matters! Last created = top of table in SPX.
  // We want Moto on TOP, so Moto must be created LAST.
  // Table order (top→bottom): Moto(ord asc) → Volumoso(ord asc) → Passeio(ord asc)
  // Creation order (first→last): Passeio(ord desc) → Volumoso(ord desc) → Moto(ord desc)
  const passeios = enabledCTs.filter(ct => ct.type === 'PASSEIO').sort((a, b) => b.ord - a.ord);
  const vols = enabledCTs.filter(ct => ct.type === 'FIORINO').sort((a, b) => b.ord - a.ord);
  const motos = enabledCTs.filter(ct => ct.type === 'MOTO').sort((a, b) => b.ord - a.ord);
  const creationOrder = [...passeios, ...vols, ...motos];

  log(`Ordem de criação: ${creationOrder.map(c => `${c.typeLabel} ${c.cluster}`).join(' → ')}`, 'info');

  const results = { created: 0, sent: 0, errors: 0 };
  const pedidosComErro = []; // { id, cluster, typeLabel, error }

  for (let i = 0; i < creationOrder.length; i++) {
    const ct = creationOrder[i];
    const label = `${ct.typeLabel} ${ct.cluster}`;
    updateProgress(i, creationOrder.length, `Criando CT: ${label}`);
    log(`[${i+1}/${creationOrder.length}] Criando CT: ${label} (${ct.ids.length} IDs)`, 'info');

    // Create empty CT
    const createResult = await runInTab(createEmptyCT, [stationId]);
    if (!createResult?.success) {
      log(`❌ Falha ao criar CT: ${label} — ${createResult?.error || 'erro desconhecido'}`, 'error');
      results.errors++;
      continue;
    }
    const ctId = createResult.ctId;
    ct.ctId = ctId;
    results.created++;
    log(`✅ CT criada: ${label} (${ctId})`, 'success');

    // Send IDs in batches of 500
    const BATCH = 500;
    let sentForCT = 0;
    for (let b = 0; b < ct.ids.length; b += BATCH) {
      const batch = ct.ids.slice(b, b + BATCH);
      updateProgress(i, creationOrder.length, `Enviando IDs: ${label} (${b + batch.length}/${ct.ids.length})`);
      const addResult = await runInTab(addShipmentsToCT, [batch, stationId, ctId]);
      if (addResult?.success) {
        sentForCT += batch.length;
      } else {
        // Batch failed — retry each ID individually to isolate the problematic ones
        log(`⚠️ Batch ${b+1}–${b+batch.length} falhou (${addResult?.error || 'erro'}). Revalidando pedidos individualmente...`, 'warn');
        for (const id of batch) {
          const singleResult = await runInTab(addShipmentsToCT, [[id], stationId, ctId]);
          if (singleResult?.success) {
            sentForCT++;
          } else {
            pedidosComErro.push({ id, cluster: ct.cluster, typeLabel: ct.typeLabel, error: singleResult?.error || 'erro desconhecido' });
            log(`❌ Pedido ${id} rejeitado (${ct.cluster}) — ${singleResult?.error || 'erro'}`, 'error');
          }
          await new Promise(r => setTimeout(r, 80));
        }
      }
      await new Promise(r => setTimeout(r, 300));
    }
    results.sent += sentForCT;
    log(`📦 ${sentForCT} IDs enviados para ${label}`, sentForCT > 0 ? 'success' : 'warn');

    await new Promise(r => setTimeout(r, 500));
  }

  log(`Importação: ${results.created} CTs criadas, ${results.sent} IDs enviados, ${results.errors} erros de CT`, results.errors ? 'warn' : 'success');
  if (pedidosComErro.length > 0) {
    log(`⚠️ ${pedidosComErro.length} pedido(s) com erro de importação — veja relatório`, 'warn');
  }
  return { ...results, stationId, pedidosComErro };
}

// ============================================================
// CALC VIA API — Funções auxiliares
// ============================================================
function getDateFormatted(daysOffset = 0) {
  const d = new Date();
  d.setDate(d.getDate() + daysOffset);
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

function getDateTimestamp(daysOffset = 0) {
  const d = new Date();
  d.setDate(d.getDate() + daysOffset);
  // Meia-noite no horário de Brasília (UTC-3) = 03:00 UTC
  return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 3, 0, 0) / 1000);
}

function buildVehiclePayload(vehicleConfig, spr, speedTier) {
  const cfg = JSON.parse(vehicleConfig.config_selections);
  const speed = cfg.speedConfig[speedTier];
  const stopServiceTime = cfg.basicServiceTime[speedTier] * 60;
  const extraServiceTime = cfg.extraServiceTime[speedTier] * 60;
  return {
    config_selections: vehicleConfig.config_selections,
    id: vehicleConfig.id,
    vehicle_name: vehicleConfig.vehicle_name,
    vehicle_id: vehicleConfig.vehicle_id,
    wheel_type: vehicleConfig.wheel_type,
    available_num: vehicleConfig.available_num,
    available_driver_num: vehicleConfig.available_driver_num,
    speed,
    stop_service_time: stopServiceTime,
    extra_service_time: extraServiceTime,
    min_parcels: vehicleConfig.min_parcels,
    max_parcels: vehicleConfig.max_parcels,
    max_distance: vehicleConfig.max_distance,
    max_weight: vehicleConfig.max_weight,
    cur_max_parcels: spr,
    cur_max_distance: vehicleConfig.max_distance,
    cur_max_weight: vehicleConfig.max_weight,
    require_plate_info: vehicleConfig.require_plate_info,
    min_capacity_ratio: vehicleConfig.min_capacity_ratio,
    max_capacity_ratio: vehicleConfig.max_capacity_ratio,
    min_weight_ratio: vehicleConfig.min_weight_ratio,
    max_weight_ratio: vehicleConfig.max_weight_ratio,
    max_distance_ratio: vehicleConfig.max_distance_ratio
  };
}

// Funções executadas no contexto da página (runInTab)
function getCommonConfig(ctId) {
  return new Promise(async (resolve) => {
    try {
      const res = await fetch(`/api/spx/lmroute/adminapi/calculation_task/common_config?calculation_task_id=${ctId}`, {
        headers: { 'accept': 'application/json, text/plain, */*', 'app': 'FMS Portal' },
        credentials: 'include'
      });
      const data = await res.json();
      if (data.retcode !== 0) { resolve({ success: false, error: data.message || `retcode: ${data.retcode}` }); return; }
      resolve({ success: true, vehicles: data.data.vehicles });
    } catch (e) { resolve({ success: false, error: e.message }); }
  });
}

function getEventList(stId, dateTs) {
  return new Promise(async (resolve) => {
    try {
      const getCookie = (n) => { const m = document.cookie.match(new RegExp('(?:^|;\\s*)' + n + '=([^;]*)')); return m ? m[1] : ''; };
      const res = await fetch('/spx_delivery/admin/roster_planning_event/event/brief_event_list', {
        method: 'POST',
        headers: { 'accept': 'application/json, text/plain, */*', 'content-type': 'application/json;charset=UTF-8', 'app': 'FMS Portal', 'x-csrftoken': getCookie('csrftoken'), 'device-id': getCookie('spx-admin-device-id') },
        credentials: 'include',
        body: JSON.stringify({ station_id: stId, event_type: 1, event_date_from: dateTs, event_date_to: dateTs })
      });
      const data = await res.json();
      if (data.retcode !== 0) { resolve({ success: false, error: data.message || `retcode: ${data.retcode}` }); return; }
      const events = data.data?.event_day_list?.[0]?.event_list || [];
      resolve({ success: true, events });
    } catch (e) { resolve({ success: false, error: e.message }); }
  });
}

function triggerCalculation(ctId, rosterEventId, vehiclePayload, maxDeliveryTime) {
  return new Promise(async (resolve) => {
    try {
      const getCookie = (n) => { const m = document.cookie.match(new RegExp('(?:^|;\\s*)' + n + '=([^;]*)')); return m ? m[1] : ''; };
      const res = await fetch('/api/spx/lmroute/adminapi/calculation_task/calculate', {
        method: 'POST',
        headers: { 'accept': 'application/json, text/plain, */*', 'content-type': 'application/json;charset=UTF-8', 'app': 'FMS Portal', 'x-csrftoken': getCookie('csrftoken'), 'device-id': getCookie('spx-admin-device-id') },
        credentials: 'include',
        body: JSON.stringify({
          calculation_task_id: ctId,
          objectives: ['p_shape'],
          vehicles: [vehiclePayload],
          trans_mode: 'oneway',
          max_delivery_time: maxDeliveryTime,
          roster_plan_event_id: rosterEventId
        })
      });
      const data = await res.json();
      if (data.retcode !== 0) { resolve({ success: false, error: data.message || `retcode: ${data.retcode}` }); return; }
      resolve({ success: true });
    } catch (e) { resolve({ success: false, error: e.message }); }
  });
}

// ============================================================
// CALC VIA API — Fluxo principal
// ============================================================
async function executeCalcAPI(enabledCTs, config, stationId) {
  log('Iniciando cálculos via API...', 'info');

  const MAX_DELIVERY_TIME = 43140; // 11h59min em segundos
  const dateOffset = config.date === 'today' ? 0 : 1;
  const dateFormatted = getDateFormatted(dateOffset);
  const dateTs = getDateTimestamp(dateOffset);

  // Ordem: MOTO → FIORINO → PASSEIO
  const motos = enabledCTs.filter(ct => ct.type === 'MOTO').sort((a, b) => a.ord - b.ord);
  const vols = enabledCTs.filter(ct => ct.type === 'FIORINO').sort((a, b) => a.ord - b.ord);
  const passeios = enabledCTs.filter(ct => ct.type === 'PASSEIO').sort((a, b) => a.ord - b.ord);
  const calcOrder = [...motos, ...vols, ...passeios];

  log(`Ordem de cálculo: ${calcOrder.map(c => `${c.typeLabel} ${c.cluster}`).join(' → ')}`, 'info');

  // Buscar evento/turno uma única vez
  log(`Buscando turno ${config.shift} para ${dateFormatted}...`, 'info');
  const eventResult = await runInTab(getEventList, [stationId, dateTs]);
  if (!eventResult?.success) {
    log(`❌ Falha ao buscar eventos: ${eventResult?.error || 'desconhecido'}`, 'error');
    return { successCount: 0, failCount: calcOrder.length };
  }

  const event = eventResult.events.find(e => e.shift_info?.shift_name === config.shift);
  if (!event) {
    log(`❌ Turno "${config.shift}" não encontrado para ${dateFormatted}. Verifique se o turno existe nessa data.`, 'error');
    return { successCount: 0, failCount: calcOrder.length };
  }

  const rosterEventId = event.event_id;
  log(`Turno ${config.shift} → evento ${rosterEventId}`, 'success');

  const tab = await getSPXTab();
  let successCount = 0, failCount = 0;

  for (let i = 0; i < calcOrder.length; i++) {
    const ct = calcOrder[i];
    const label = `${ct.typeLabel} ${ct.cluster}`;
    updateProgress(i, calcOrder.length, `Calculando: ${label}`);
    log(`[${i+1}/${calcOrder.length}] ${label} (CT: ${ct.ctId})`, 'info');

    if (!ct.ctId) {
      log(`⚠️ ${label}: sem CT ID, pulando`, 'warn');
      failCount++;
      continue;
    }

    // GeoFixer antes do cálculo (se habilitado)
    if (config.geoFixer) {
      updateProgress(i, calcOrder.length, `📍 GeoFixer: ${label}`);
      await new Promise((resolve) => {
        chrome.tabs.sendMessage(tab.id, { action: 'runGeoFixerForCT', taskId: ct.ctId }, () => resolve());
      });
    }

    // Buscar configs do veículo na CT
    const configResult = await runInTab(getCommonConfig, [ct.ctId]);
    if (!configResult?.success) {
      log(`❌ ${label}: falha ao buscar config — ${configResult?.error}`, 'error');
      failCount++;
      continue;
    }

    // Encontrar veículo correspondente
    const vehicleMatch = configResult.vehicles.find(cv => {
      const n = cv.vehicle_name.toUpperCase();
      if (ct.type === 'MOTO') return n === 'MOTO';
      if (ct.type === 'PASSEIO') return n === 'PASSEIO';
      if (ct.type === 'FIORINO') return n === 'FIORINO' || n === 'VAN' || n === 'VOLUMOSO';
      return false;
    });

    if (!vehicleMatch) {
      log(`❌ ${label}: tipo "${ct.type}" não encontrado na CT`, 'error');
      failCount++;
      continue;
    }

    // Montar payload e disparar cálculo
    const vehiclePayload = buildVehiclePayload(vehicleMatch, ct.spr, 'fast');
    const calcResult = await runInTab(triggerCalculation, [ct.ctId, rosterEventId, vehiclePayload, MAX_DELIVERY_TIME]);

    if (calcResult?.success) {
      successCount++;
      log(`✅ ${label} → calculado!`, 'success');
    } else {
      failCount++;
      log(`❌ ${label}: ${calcResult?.error}`, 'error');
    }
  }

  updateProgress(calcOrder.length, calcOrder.length, 'Cálculos concluídos');
  log(`Cálculos: ${successCount} sucesso, ${failCount} erros`, successCount > 0 ? 'success' : 'error');
  return { successCount, failCount };
}

// ============================================================
// PROGRESS
// ============================================================
function updateProgress(current, total, message) {
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  $('progressFill').style.width = `${pct}%`;
  $('progressText').textContent = message || `${current}/${total}`;
}

// ============================================================
// EVENT LISTENERS
// ============================================================

// Comerciais toggle
$('comerciaisEnabled').addEventListener('change', () => {
  const on = $('comerciaisEnabled').checked;
  const badge = $('comerciaisBadge');
  badge.textContent = on ? 'ATIVO' : 'DESATIVADO';
  badge.className = on ? 'badge badge-on' : 'badge badge-off';
  // Re-parseia o template com o novo estado do filtro (se já foi carregado)
  if (templateBuffer) applyTemplate(templateBuffer, on);
});

// GeoFixer toggle
$('gfxEnabled').addEventListener('change', () => {
  const on = $('gfxEnabled').checked;
  const badge = $('gfxBadge');
  badge.textContent = on ? 'ATIVO' : 'DESATIVADO';
  badge.className = on ? 'badge badge-on' : 'badge badge-off';
  console.log('📍 Correção de Pins:', on ? 'ATIVO' : 'DESATIVADO');
});

// Shift toggles
$('shiftAM').addEventListener('click', () => { $('shiftAM').classList.add('active'); $('shiftPM1').classList.remove('active'); $('shiftPM2').classList.remove('active'); });
$('shiftPM1').addEventListener('click', () => { $('shiftPM1').classList.add('active'); $('shiftAM').classList.remove('active'); $('shiftPM2').classList.remove('active'); });
$('shiftPM2').addEventListener('click', () => { $('shiftPM2').classList.add('active'); $('shiftAM').classList.remove('active'); $('shiftPM1').classList.remove('active'); });

// Log toggle
$('logToggle').addEventListener('click', () => {
  $('logToggle').classList.toggle('open');   // log-header.open rotates caret
  $('logBody').classList.toggle('open');
});

// Aplica o parse + build CT list + render preview
function applyTemplate(arrayBuffer, filterComerciais) {
  try {
    templateData = parseTemplate(arrayBuffer, filterComerciais);
    if (filterComerciais && templateData.comerciaisRemovidos.length > 0) {
      log(`🏢 Comerciais removidos: ${templateData.comerciaisRemovidos.length} pedidos (Office/RTS/SP)`, 'warn');
    }

    let totalIDs = 0;
    Object.values(templateData.shipments).forEach(s => {
      totalIDs += s.MOTO.length + s.Passeio.length + s.Volumoso.length + s.MOTO_BACKLOG.length;
    });
    $('fileDetails').textContent = `${templateData.clusters.length} clusters | ${totalIDs.toLocaleString()} IDs válidos | SPR Moto: ${templateData.sprMotoGlobal} | SPR Vol: ${templateData.sprVolGlobal}`;

    ctList = buildCTList(templateData);
    renderPreview();
  } catch(err) {
    log(`❌ Erro ao processar template: ${err.message}`, 'error');
  }
}

// File selection
$('btnSelect').addEventListener('click', () => $('fileInput').click());
$('fileInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  $('btnSelect').disabled = true;
  $('btnSelect').innerHTML = '<div class="spinner"></div> Lendo template...';
  log(`Arquivo: ${file.name} (${(file.size/1024).toFixed(0)} KB)`, 'info');

  try {
    templateBuffer = await file.arrayBuffer();
    const filterComerciais = $('comerciaisEnabled').checked;

    applyTemplate(templateBuffer, filterComerciais);
    log(`✅ Template lido: ${templateData.clusters.length} clusters`, 'success');
    log(`📊 Filtrados: ${templateData.totalFiltered} IDs excluídos`, 'info');

    $('fileInfo').style.display = 'block';
    $('fileName').textContent = file.name;
    log(`📋 ${ctList.length} CTs geradas (${ctList.filter(c=>c.enabled).length} ativas)`, 'info');

    showPhase(2);

  } catch(err) {
    log(`❌ Erro ao ler template: ${err.message}`, 'error');
    alert(`Erro ao ler template:\n${err.message}`);
  } finally {
    $('btnSelect').disabled = false;
    $('btnSelect').innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg> Selecionar Template (.xlsm / .xlsx)';
  }
});

// Back button
$('btnBack').addEventListener('click', () => {
  showPhase(1);
});

// Execute button
$('btnExecute').addEventListener('click', async () => {
  if (isRunning) return;

  const connected = await checkConnection();
  if (!connected) { alert('Abra o SPX (spx.shopee.com.br) primeiro!'); return; }

  const enabledCTs = ctList.filter(ct => ct.enabled);
  if (enabledCTs.length === 0) { alert('Nenhuma CT selecionada!'); return; }

  const totalIDs = enabledCTs.reduce((s, ct) => s + ct.ids.length, 0);
  if (!confirm(`Executar ${enabledCTs.length} CTs com ${totalIDs.toLocaleString()} IDs?\n\nOrdem: Moto → Passeio → Volumoso`)) return;

  isRunning = true;
  const shift = $('shiftAM').classList.contains('active') ? 'AM' : $('shiftPM1').classList.contains('active') ? 'PM-1' : 'PM-2';
  const config = { date: $('calcDate').value, shift, geoFixer: $('gfxEnabled').checked };

  showPhase(3);

  // Open logs
  $('logToggle').classList.add('open');
  $('logBody').classList.add('open');

  // Step 1: Import
  $('stepImport').className = 'step active';
  $('stepImport').querySelector('.step-icon').textContent = '⏳';
  const importResult = await executeImport(enabledCTs);

  if (!importResult || importResult.created === 0) {
    log('❌ Nenhuma CT criada. Verifique a conexão.', 'error');
    isRunning = false;
    return;
  }

  $('stepImport').className = 'step done';
  $('stepImport').querySelector('.step-icon').textContent = '✅';

  // Step 2: Calc
  const geoLabel = config.geoFixer ? ' Calculando + GeoFixer por CT...' : ' Calculando roteirizações...';
  $('stepCalc').className = 'step active';
  $('stepCalc').querySelector('.step-icon').textContent = '⏳';
  $('stepCalc').innerHTML = `<span class="step-icon">⏳</span><span>${geoLabel}</span>`;
  const calcResult = await executeCalcAPI(enabledCTs, config, importResult.stationId);

  $('stepCalc').className = 'step done';
  $('stepCalc').innerHTML = `<span class="step-icon">✅</span><span> Cálculos concluídos</span>`;

  // Show results
  $('resCTs').textContent = importResult.created;
  $('resIDs').textContent = importResult.sent.toLocaleString();
  $('resCalc').textContent = calcResult?.successCount ?? importResult.created;

  // Mostrar painel de comerciais se houver removidos
  const removidos = templateData?.comerciaisRemovidos || [];
  if (removidos.length > 0) {
    $('comerciaisResult').style.display = 'block';
    const tipos = removidos.reduce((acc, r) => { acc[r.tipo] = (acc[r.tipo]||0)+1; return acc; }, {});
    const resumo = Object.entries(tipos).map(([t,n]) => `${t}: ${n}`).join(' | ');
    $('comerciaisCount').textContent = `${removidos.length} pedido(s) — ${resumo}`;
  } else {
    $('comerciaisResult').style.display = 'none';
  }

  // Mostrar painel de erros de importação
  const erros = importResult?.pedidosComErro || [];
  if (erros.length > 0) {
    window._pedidosComErro = erros;
    $('importErrorResult').style.display = 'block';
    const porCluster = erros.reduce((acc, r) => { acc[r.cluster] = (acc[r.cluster]||0)+1; return acc; }, {});
    const resumoErros = Object.entries(porCluster).map(([c,n]) => `${c}: ${n}`).join(' | ');
    $('importErrorCount').textContent = `${erros.length} pedido(s) — ${resumoErros}`;
  } else {
    $('importErrorResult').style.display = 'none';
  }

  showPhase(4);
  isRunning = false;

  // Gerar e abrir relatório HTML
  try {
    const html = generateReport(ctList, templateData, importResult, calcResult, config);
    window._lastReport = html;
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    chrome.tabs.create({ url: url });
    $('btnOpenReport').style.display = 'flex';
    log('📊 Relatório de execução gerado e aberto', 'success');
  } catch(e) {
    log(`⚠️ Não foi possível gerar relatório: ${e.message}`, 'warn');
  }
});

// ============================================================
// RELATÓRIO HTML DE EXECUÇÃO
// ============================================================
function generateReport(ctList, templateData, importResult, calcResult, config) {
  const now = new Date().toLocaleString('pt-BR');
  const station = stationNameCache;
  const shiftLabel = config.shift;
  const dateFormatted = config.date === 'tomorrow'
    ? new Date(Date.now() + 86400000).toLocaleDateString('pt-BR')
    : new Date().toLocaleDateString('pt-BR');

  const activeCTs   = ctList.filter(ct => ct.enabled && !ct.isBaixoADO);
  const baixoADO    = ctList.filter(ct => ct.isBaixoADO);
  const totalRoutes = activeCTs.reduce((s, ct) => s + (ct.rotas || 0), 0);
  const totalIDs    = importResult?.sent || 0;
  const backlog     = templateData.backlogTotal || 0;
  const comerciais  = templateData.comerciaisRemovidos || [];
  const erros       = importResult?.pedidosComErro || [];
  const comByType   = comerciais.reduce((a, r) => { a[r.tipo] = (a[r.tipo]||0)+1; return a; }, {});
  const sprGeral    = totalRoutes > 0 ? Math.round(totalIDs / totalRoutes) : 0;

  const typeColor = { MOTO: '#EE4D2D', PASSEIO: '#1E3A8A', FIORINO: '#D97706' };
  const typeLabel = { MOTO: 'Moto', PASSEIO: 'Passeio', FIORINO: 'Volumoso' };
  const typeBg    = { MOTO: '#FFF1EE', PASSEIO: '#EFF6FF', FIORINO: '#FFFBEB' };

  // ── SVG BAR CHART helper ──────────────────────────────
  function svgBar(items, opts = {}) {
    const W = opts.width || 420, H = opts.barH || 140, PAD = 48, GAP = opts.gap || 14;
    const n = items.length;
    if (!n) return '<div style="color:#a1a1aa;font-size:11px;text-align:center;padding:20px;">Sem dados</div>';
    const barW = Math.max(24, Math.min(60, (W - PAD * 2 - GAP * (n - 1)) / n));
    const svgW = PAD * 2 + n * barW + (n - 1) * GAP;
    const maxV = Math.max(...items.map(d => d.value), 1);
    const bars = items.map((d, i) => {
      const x = PAD + i * (barW + GAP);
      const bh = Math.max(4, (d.value / maxV) * H);
      const y = H - bh + 10;
      const lbl = d.label.length > 10 ? d.label.slice(0, 10) + '…' : d.label;
      return `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" fill="${d.color}" rx="3" opacity="0.92"/>
        <text x="${x + barW/2}" y="${y - 5}" text-anchor="middle" font-size="11" font-weight="700" fill="${d.color}">${d.value.toLocaleString('pt-BR')}</text>
        <text x="${x + barW/2}" y="${H + 26}" text-anchor="middle" font-size="10" fill="#71717a">${lbl}</text>`;
    });
    return `<svg viewBox="0 0 ${svgW} ${H + 40}" width="100%" style="overflow:visible;">${bars.join('')}</svg>`;
  }

  // ── Chart data ──────────────────────────────────────
  const clusterChart = activeCTs
    .slice()
    .sort((a, b) => b.ids.length - a.ids.length)
    .slice(0, 12)
    .map(ct => ({ label: ct.cluster.replace(/^\d+\.\s*/, ''), value: ct.ids.length, color: '#EE4D2D' }));

  const typeGroups = activeCTs.reduce((a, ct) => {
    if (!a[ct.type]) a[ct.type] = { routes: 0, ids: 0 };
    a[ct.type].routes += ct.rotas || 0;
    a[ct.type].ids += ct.ids.length;
    return a;
  }, {});
  const routesChart = Object.entries(typeGroups).map(([t, v]) => ({ label: typeLabel[t] || t, value: v.routes, color: typeColor[t] || '#888' }));

  // ── TABLE rows ──────────────────────────────────────
  const clusterRows = activeCTs.map((ct, i) => `
    <tr>
      <td style="color:#a1a1aa;font-size:11px;">${i+1}</td>
      <td style="font-weight:600;">${ct.cluster}</td>
      <td><span style="background:${typeBg[ct.type]||'#f5f5f5'};color:${typeColor[ct.type]||'#333'};padding:1px 8px;border-radius:20px;font-size:10px;font-weight:700;">${typeLabel[ct.type]||ct.type}</span></td>
      <td class="num">${ct.ids.length.toLocaleString('pt-BR')}</td>
      <td class="num">${ct.rotas}</td>
      <td class="num">${ct.spr}</td>
      <td style="font-family:monospace;font-size:10px;color:#a1a1aa;">${ct.ctId||'—'}</td>
    </tr>`).join('');

  const baixoRows = baixoADO.length
    ? baixoADO.map(ct => `<tr><td style="font-weight:600;">${ct.cluster}</td><td class="num">${ct.ids.length}</td></tr>`).join('')
    : `<tr><td colspan="2" style="text-align:center;color:#a1a1aa;padding:10px;font-size:11px;">Nenhum</td></tr>`;

  const comRows = comerciais.length
    ? Object.entries(comByType).map(([t, n]) => `<tr><td style="font-weight:600;">${t}</td><td class="num">${n}</td></tr>`).join('')
    : `<tr><td colspan="2" style="text-align:center;color:#a1a1aa;padding:10px;font-size:11px;">Nenhum</td></tr>`;

  const erroRows = erros.length
    ? erros.slice(0,100).map(r => `<tr><td style="font-family:monospace;font-size:11px;">${r.id}</td><td>${r.cluster}</td><td style="font-size:11px;color:#a1a1aa;">${r.error}</td></tr>`).join('')
    : '';

  const card = (val, lbl, color = '#18181B') =>
    `<div class="kcard"><div class="kval" style="color:${color};">${typeof val==='number'?val.toLocaleString('pt-BR'):val}</div><div class="klbl">${lbl}</div></div>`;

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Relatório — ${station}</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#fff;color:#18181B;font-size:13px;}

/* HEADER */
.hdr{display:flex;align-items:center;justify-content:space-between;padding:12px 24px;border-bottom:3px solid #EE4D2D;background:#fff;}
.hdr-brand{display:flex;align-items:center;gap:10px;}
.hdr-logo{width:36px;height:36px;background:#EE4D2D;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;}
.hdr-name{font-size:16px;font-weight:800;color:#18181B;letter-spacing:-0.3px;}
.hdr-sub{font-size:10px;color:#71717a;}
.hdr-title{font-size:17px;font-weight:800;color:#18181B;text-align:center;letter-spacing:-0.3px;}
.hdr-actions{display:flex;gap:8px;}
.btn-hdr{border:none;border-radius:7px;padding:7px 14px;font-size:11px;font-weight:700;cursor:pointer;}
.btn-print{background:#EE4D2D;color:#fff;}
.btn-share{background:#F0F0F4;color:#52525b;border:1px solid #D4D4DC;}

/* KPIS */
.kpi-row{display:flex;border-bottom:2px solid #e5e5e5;}
.kcard{flex:1;text-align:center;padding:14px 8px;border-right:1px solid #e5e5e5;}
.kcard:last-child{border-right:none;}
.kval{font-size:26px;font-weight:800;line-height:1;letter-spacing:-1px;}
.klbl{font-size:9px;font-weight:600;color:#a1a1aa;text-transform:uppercase;letter-spacing:.6px;margin-top:4px;}

/* CHARTS ROW */
.charts-row{display:grid;grid-template-columns:1fr 1fr;border-bottom:1px solid #e5e5e5;}
.chart-box{padding:16px 20px;}
.chart-box:first-child{border-right:1px solid #e5e5e5;}
.chart-title{font-size:12px;font-weight:700;color:#EE4D2D;text-align:center;margin-bottom:12px;text-transform:uppercase;letter-spacing:.4px;}

/* TABLES */
.tables-section{padding:0;}
.panel{border-bottom:1px solid #e5e5e5;}
.panel-hdr{display:flex;align-items:center;justify-content:space-between;padding:10px 20px 8px;background:#F9F9F9;border-bottom:1px solid #e5e5e5;}
.panel-title{font-size:11px;font-weight:700;color:#18181B;text-transform:uppercase;letter-spacing:.5px;}
.cnt{background:#EE4D2D;color:#fff;font-size:10px;font-weight:700;padding:1px 7px;border-radius:20px;}
.tbl-wrap{overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:11px;}
th{background:#F5F5F8;padding:7px 12px;text-align:left;font-size:9px;font-weight:700;color:#71717a;text-transform:uppercase;letter-spacing:.4px;border-bottom:1px solid #e5e5e5;white-space:nowrap;cursor:pointer;}
th:hover{background:#EBEBEF;}
th.sort-asc::after{content:' ↑';}
th.sort-desc::after{content:' ↓';}
td{padding:7px 12px;border-bottom:1px solid #F0F0F4;vertical-align:middle;}
tr:last-child td{border-bottom:none;}
tr:hover td{background:#fafafa;}
.num{text-align:right;font-weight:600;}

/* BOTTOM PANELS (2-col) */
.bottom-grid{display:grid;grid-template-columns:1fr 1fr;border-top:1px solid #e5e5e5;}
.bottom-grid .panel{border-bottom:none;border-right:1px solid #e5e5e5;}
.bottom-grid .panel:last-child{border-right:none;}

/* FOOTER */
.rpt-footer{display:flex;justify-content:space-between;align-items:center;padding:10px 24px;background:#F9F9F9;border-top:3px solid #EE4D2D;font-size:10px;color:#71717a;flex-wrap:wrap;gap:8px;}
.rpt-footer strong{color:#18181B;}

/* SHARE OVERLAY */
#shareOverlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:99;align-items:center;justify-content:center;}
#shareOverlay.vis{display:flex;}
.share-box{background:#fff;border-radius:14px;padding:26px;width:440px;max-width:92vw;box-shadow:0 20px 50px rgba(0,0,0,.22);}
.share-box h3{font-size:14px;font-weight:700;margin-bottom:4px;}
.share-box p{font-size:11px;color:#71717a;margin-bottom:16px;}
.share-step{display:flex;gap:10px;background:#F5F5F8;border-radius:8px;padding:10px 12px;margin-bottom:8px;align-items:flex-start;}
.step-n{width:20px;height:20px;background:#EE4D2D;color:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;flex-shrink:0;}
.share-step p{font-size:11px;color:#52525b;margin:0;line-height:1.5;}
.share-step p strong{color:#18181B;}
.btn-close-share{width:100%;background:#F0F0F4;border:none;border-radius:8px;padding:10px;font-size:13px;font-weight:600;cursor:pointer;color:#52525b;margin-top:4px;}
.btn-close-share:hover{background:#e4e4e7;}

@media print{
  .hdr-actions,#shareOverlay{display:none!important;}
  body{background:#fff;}
}
</style>
</head>
<body>

<div id="shareOverlay">
  <div class="share-box">
    <h3>📤 Compartilhar Relatório</h3>
    <p>Escolha como salvar e enviar no grupo:</p>
    <div class="share-step"><div class="step-n">1</div><p><strong>Screenshot rápido (Win):</strong> Pressione <strong>Win + Shift + S</strong>, selecione a área — a imagem vai direto para área de transferência.</p></div>
    <div class="share-step"><div class="step-n">2</div><p><strong>Página inteira (Chrome):</strong> F12 → menu ⋮ → <em>Capturar captura de tela de página inteira</em> → salva PNG completo.</p></div>
    <div class="share-step"><div class="step-n">3</div><p><strong>Salvar como PDF:</strong> Clique em <strong>🖨️ Imprimir</strong> e escolha "Salvar como PDF".</p></div>
    <button class="btn-close-share" id="btnCloseShare">Fechar</button>
  </div>
</div>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-brand">
    <div class="hdr-logo">🚀</div>
    <div><div class="hdr-name">SPX Auto Router</div><div class="hdr-sub">Routing Tower</div></div>
  </div>
  <div class="hdr-title">${station} &nbsp;·&nbsp; ${shiftLabel} &nbsp;·&nbsp; ${dateFormatted}</div>
  <div class="hdr-actions">
    <button class="btn-hdr btn-share" id="btnShare">📤 Compartilhar</button>
    <button class="btn-hdr btn-print" id="btnPrint">🖨️ Imprimir / PDF</button>
  </div>
</div>

<!-- KPI CARDS -->
<div class="kpi-row">
  ${card(totalIDs, 'Total Importado', '#EE4D2D')}
  ${card(importResult?.created ?? 0, 'CTs Criadas', '#1E3A8A')}
  ${card(calcResult?.successCount ?? 0, 'Calculados', '#16A34A')}
  ${card(totalRoutes, 'Total Rotas', '#18181B')}
  ${card(sprGeral, 'SPR Geral', '#7C3AED')}
  ${card(backlog, 'Backlog', '#D97706')}
  ${card(baixoADO.length, 'Baixo ADO', '#DC2626')}
  ${card(comerciais.length, 'Comerciais', '#52525b')}
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-box">
    <div class="chart-title">Pedidos por Cluster</div>
    ${svgBar(clusterChart, { width: 500, barH: 130, gap: 8 })}
  </div>
  <div class="chart-box">
    <div class="chart-title">Rotas por Tipo de Veículo</div>
    ${svgBar(routesChart, { width: 300, barH: 130, gap: 30 })}
  </div>
</div>

<!-- CLUSTERS TABLE -->
<div class="panel">
  <div class="panel-hdr">
    <span class="panel-title">Clusters Roteirizados</span>
    <span class="cnt">${activeCTs.length}</span>
  </div>
  <div class="tbl-wrap">
    <table id="tblClusters">
      <thead><tr>
        <th data-sort="0">#</th>
        <th data-sort="1">Cluster</th>
        <th data-sort="2">Tipo</th>
        <th data-sort="3" style="text-align:right;">Pedidos</th>
        <th data-sort="4" style="text-align:right;">Rotas</th>
        <th data-sort="5" style="text-align:right;">SPR</th>
        <th>CT ID</th>
      </tr></thead>
      <tbody>${clusterRows}</tbody>
    </table>
  </div>
</div>

<!-- BOTTOM: Baixo ADO + Comerciais -->
<div class="bottom-grid">
  <div class="panel">
    <div class="panel-hdr"><span class="panel-title">Baixo ADO</span><span class="cnt">${baixoADO.length}</span></div>
    <div class="tbl-wrap"><table><thead><tr><th>Cluster</th><th style="text-align:right;">IDs</th></tr></thead><tbody>${baixoRows}</tbody></table></div>
  </div>
  <div class="panel">
    <div class="panel-hdr"><span class="panel-title">Comerciais Removidos</span><span class="cnt">${comerciais.length}</span></div>
    <div class="tbl-wrap"><table><thead><tr><th>Tipo</th><th style="text-align:right;">Qtd</th></tr></thead><tbody>${comRows}</tbody></table></div>
  </div>
</div>

${erros.length > 0 ? `
<div class="panel">
  <div class="panel-hdr" style="border-color:#DC2626;"><span class="panel-title" style="color:#DC2626;">Erros de Importação</span><span class="cnt">${erros.length}</span></div>
  <div class="tbl-wrap"><table><thead><tr><th>Shipment ID</th><th>Cluster</th><th>Erro</th></tr></thead><tbody>${erroRows}</tbody></table></div>
</div>` : ''}

<!-- FOOTER -->
<div class="rpt-footer">
  <span><strong>Data roteirização:</strong> ${dateFormatted}</span>
  <span><strong>Gerado em:</strong> ${now}</span>
  <span>SPX Auto Router — Routing Tower</span>
</div>

<script>
// Sort tables
function sortTable(id, col) {
  var tbl = document.getElementById(id);
  if (!tbl) return;
  var tbody = tbl.querySelector('tbody');
  var ths = tbl.querySelectorAll('thead th');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var asc = ths[col] && ths[col].classList.contains('sort-asc');
  ths.forEach(function(th){ th.classList.remove('sort-asc','sort-desc'); });
  if (ths[col]) ths[col].classList.add(asc ? 'sort-desc' : 'sort-asc');
  rows.sort(function(a, b) {
    var va = (a.cells[col] && a.cells[col].innerText.trim()) || '';
    var vb = (b.cells[col] && b.cells[col].innerText.trim()) || '';
    var na = parseFloat(va.replace(/\./g,'').replace(',','.'));
    var nb = parseFloat(vb.replace(/\./g,'').replace(',','.'));
    if (!isNaN(na) && !isNaN(nb)) return asc ? nb - na : na - nb;
    return asc ? vb.localeCompare(vb,'pt-BR') : va.localeCompare(vb,'pt-BR');
  });
  rows.forEach(function(r){ tbody.appendChild(r); });
}

// Wire up buttons — runs immediately since script is at end of body
(function() {
  var btnPrint = document.getElementById('btnPrint');
  var btnShare = document.getElementById('btnShare');
  var btnClose = document.getElementById('btnCloseShare');
  var overlay  = document.getElementById('shareOverlay');

  if (btnPrint) btnPrint.addEventListener('click', function() { window.print(); });
  if (btnShare) btnShare.addEventListener('click', function() { overlay.classList.add('vis'); });
  if (btnClose) btnClose.addEventListener('click', function() { overlay.classList.remove('vis'); });
  if (overlay)  overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.classList.remove('vis'); });

  // Table sort on header click
  document.querySelectorAll('thead th[data-sort]').forEach(function(th) {
    th.addEventListener('click', function() {
      var tbl = th.closest('table');
      if (tbl && tbl.id) sortTable(tbl.id, parseInt(th.dataset.sort));
    });
  });
})();
</script>
</body>
</html>`;
}

// ---- Helper: download XLSX blob ----
function downloadXLSX(wb, filename) {
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---- Helper: apply header style ----
function styleSheetHeader(ws, cols) {
  const headerStyle = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'EE4D2D' } }, alignment: { horizontal: 'center' } };
  cols.forEach((_, ci) => {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: ci })];
    if (cell) cell.s = headerStyle;
  });
}

// Exportar lista de comerciais (.xlsx)
$('btnExportComerciais').addEventListener('click', () => {
  const removidos = templateData?.comerciaisRemovidos || [];
  if (removidos.length === 0) return;

  const dataTs = new Date().toLocaleString('pt-BR');
  const rows = removidos.map(r => ({
    'Shipment ID': String(r.id),
    'Tipo':        String(r.tipo),
    'Cluster':     String(r.cluster),
    'Perfil':      String(r.perfil || '')
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 28 }, { wch: 22 }, { wch: 32 }, { wch: 14 }];
  styleSheetHeader(ws, rows[0] ? Object.keys(rows[0]) : []);

  // Info sheet
  const wsInfo = XLSX.utils.aoa_to_sheet([
    ['Relatório de Pedidos Comerciais Removidos'],
    ['Gerado em:', dataTs],
    ['Total:', removidos.length]
  ]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Comerciais Removidos');
  XLSX.utils.book_append_sheet(wb, wsInfo, 'Info');

  const date = new Date().toISOString().slice(0, 10);
  downloadXLSX(wb, `comerciais_removidos_${date}.xlsx`);
  log(`📁 Relatório de comerciais exportado (${removidos.length} pedidos)`, 'success');
});

// Exportar pedidos com erro de importação (.xlsx)
$('btnExportImportErrors').addEventListener('click', () => {
  const erros = window._pedidosComErro || [];
  if (erros.length === 0) return;

  const dataTs = new Date().toLocaleString('pt-BR');
  const rows = erros.map(r => ({
    'Shipment ID': String(r.id),
    'Cluster':     String(r.cluster),
    'Tipo':        String(r.typeLabel).replace(/[^\w\s.]/g, '').trim(),
    'Erro':        String(r.error)
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 28 }, { wch: 32 }, { wch: 12 }, { wch: 50 }];
  styleSheetHeader(ws, rows[0] ? Object.keys(rows[0]) : []);

  const wsInfo = XLSX.utils.aoa_to_sheet([
    ['Relatório de Pedidos com Erro de Importação'],
    ['Gerado em:', dataTs],
    ['Total:', erros.length]
  ]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Erros de Importação');
  XLSX.utils.book_append_sheet(wb, wsInfo, 'Info');

  const date = new Date().toISOString().slice(0, 10);
  downloadXLSX(wb, `pedidos_erro_importacao_${date}.xlsx`);
  log(`📁 Relatório de erros exportado (${erros.length} pedidos)`, 'success');
});

// Reabrir relatório
$('btnOpenReport').addEventListener('click', () => {
  if (!window._lastReport) return;
  const blob = new Blob([window._lastReport], { type: 'text/html;charset=utf-8' });
  window.open(URL.createObjectURL(blob), '_blank');
});

// Restart
$('btnRestart').addEventListener('click', () => {
  templateData = null;
  templateBuffer = null;
  ctList = [];
  window._pedidosComErro = [];
  window._lastReport = null;
  $('fileInput').value = '';
  $('fileInfo').style.display = 'none';
  $('comerciaisResult').style.display = 'none';
  $('importErrorResult').style.display = 'none';
  $('btnOpenReport').style.display = 'none';
  $('logBody').innerHTML = '';
  showPhase(1);
});

// ---- Init ----
checkConnection();

// Toggle rows — clique em qualquer lugar da linha ativa o switch
['gfxToggleRow', 'comerciaisToggleRow'].forEach(rowId => {
  $( rowId).addEventListener('click', (e) => {
    // Evita duplo disparo se clicou direto no input/label/slider
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'LABEL' || e.target.classList.contains('slider')) return;
    const input = $(rowId).querySelector('input[type="checkbox"]');
    if (input) input.click();
  });
});
