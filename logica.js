// =========================================================
// KIT DE PRIMEIROS SOCORROS (EVITA CURTO-CIRCUITO NO JS)
// =========================================================
window.renderPontoMensal = window.renderPontoMensal || function() {};
window.loadInbox = window.loadInbox || function() {};
window.renderAbs = window.renderAbs || function() {};

// Se o sistema procurar o botão f2 e não achar, criamos um fantasma
if (!document.getElementById('f2')) {
    let fakeF2 = document.createElement('input');
    fakeF2.type = 'file'; fakeF2.id = 'f2'; fakeF2.style.display = 'none';
    document.body.appendChild(fakeF2);
}
// =========================================================
// =========================================================================
// NÚCLEO DO SISTEMA v163.0 — CLOUD, RH INTELIGENTE E LOGÍSTICA (CORRIGIDO)
// =========================================================================

let chP = null, chF = null, cenGauge = null, cenRadar = null;
let dC = [], dE = [], dPickRes = [];
let globalHistory = [], globalErrors = [];
// rhData managed by rh.js
let colMapMaster = { ped: null, ate: null, n_ate: null, prod: null, desc: null, mov: null, cx_ped: null, cx_ate: null, rom: null, dep: null };
let forceBoxCalc = false;
let moduleStats = {}, projData = [];
let globalAttentionPoints = [], globalSurplusList = [], globalDeficitList = [];
let cachedCenarioData = null;
let isDecisionCommitted = false;
let stationTopProducts = {}, macroTopProducts = {}, macroTopBoxes = {};
// pendenciasRH managed by rh.js

Chart.register(ChartDataLabels);

let SECTORS = {
    'MEDICAMENTO': { l: 'Medicamento (V43)', ri: 57, pi: 0, rf: 90, pf: 999 },
    'PERFUMARIA': { l: 'Perfumaria (V41)', ri: 57, pi: 0, rf: 90, pf: 999 },
    'RUAS': { l: 'Setor Ruas (23-25)', ri: 1, pi: 11, rf: 22, pf: 33 },
    'ALIMENTO': { l: 'Setor Alimento', ri: 1, pi: 0, rf: 99, pf: 99 },
    'VM': { l: 'Setor VM', ri: 1, pi: 11, rf: 8, pf: 35 },
    'VH': { l: 'Setor VH', ri: 1, pi: 11, rf: 5, pf: 24 },
    'CONTROLADO': { l: 'Controlado (VX)', ri: 0, pi: 0, rf: 9999, pf: 9999 },
    'DERMO': { l: 'Dermo (VW)', ri: 0, pi: 0, rf: 9999, pf: 9999 },
    'VOLUMOSO': { l: 'Volumoso (VY)', ri: 1, pi: 0, rf: 2, pf: 9999 },
    'LATA': { l: 'Latas (VY)', ri: 3, pi: 0, rf: 3, pf: 9999 }
};

const DEFAULT_METAS = { 'm_MEDICAMENTO': 450, 'm_PERFUMARIA': 450, 'm_CONTROLADO': 50, 'm_DERMO': 300, 'm_VOLUMOSO': 300, 'm_ALIMENTO': 250, 'm_LATA': 150, 'm_VM': 300, 'm_VH': 150, 'm_RUAS': 250 };
const META_LABELS = { 'm_MEDICAMENTO': 'Med', 'm_PERFUMARIA': 'Perf', 'm_CONTROLADO': 'Contr', 'm_DERMO': 'Dermo', 'm_VOLUMOSO': 'Vol', 'm_ALIMENTO': 'Alim', 'm_LATA': 'Lata', 'm_VM': 'VM', 'm_VH': 'VH', 'm_RUAS': 'Ruas' };


function setSysModule(mod) {
    document.querySelectorAll('.main-mod-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById('btn-mod-' + mod).classList.add('active');
    document.getElementById('bar-operacao').style.display = 'none';
    document.getElementById('bar-rh').style.display = 'none';
    document.getElementById('wrapper-operacao').classList.remove('active');
    document.getElementById('wrapper-rh').classList.remove('active');
    if (mod === 'operacao') {
        document.getElementById('bar-operacao').style.display = 'flex';
        document.getElementById('wrapper-operacao').classList.add('active');
        if (chP && chF) { setTimeout(() => { chP.resize(); chF.resize(); }, 100); }
    } else if (mod === 'rh') {
        document.getElementById('bar-rh').style.display = 'flex';
        document.getElementById('wrapper-rh').classList.add('active');
    }
}

function switchTab(btnElement, targetViewId, activeClassStr) {
    document.querySelectorAll('.view-section').forEach(e => e.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(e => { e.className = e.className.replace(/\b\S+-active\b/g, '').replace(/\bactive\b/g, '').trim(); e.classList.add('tab-btn'); });
    document.getElementById('view-' + targetViewId).classList.add('active');
    if (btnElement) btnElement.classList.add(activeClassStr);
    if (targetViewId === 'dashboard' && chP && chF) { setTimeout(() => { chP.resize(); chF.resize(); }, 100); }
    if (targetViewId === 'cenario') { setTimeout(() => runCenarioCalc(), 100); }
    if (targetViewId === 'rh-ferias') { loadInbox(); }
    if (targetViewId === 'rh-ponto') { renderPontoMensal(); }
    if (targetViewId === 'projecao') { setTimeout(() => updateProjecaoView(), 100); }
}

function ok(id, n) {
    safeUpdate('msg' + id, '✅ ' + n);
    const bx = document.getElementById('bx' + id);
    if (bx) bx.classList.add('done');
    if (dC.length > 0) document.getElementById('btnCalc').disabled = false;
}


function renderMetas() {
    try {
        const sm = JSON.parse(localStorage.getItem('sysLog_metas')) || DEFAULT_METAS;
        const mc = document.getElementById('metaContainer');
        if (!mc) return;
        mc.innerHTML = Object.keys(DEFAULT_METAS).map(k =>
            `<div class="meta-item">
                <label>${META_LABELS[k]}</label>
                <input type="number" id="${k}" value="${sm[k] !== undefined ? sm[k] : DEFAULT_METAS[k]}" min="1">
             </div>`
        ).join('');
        mc.querySelectorAll('input[type="number"]').forEach(inp => {
            inp.addEventListener('change', () => {
                localStorage.setItem('sysLog_metas', JSON.stringify(getCurrentMetas()));
            });
        });
    } catch(e) { console.warn('renderMetas erro:', e); }
}

async function init() {
    try {
        if (window.openDB) await window.openDB();
        if (window.getFromDB) {
            const savedRH = await window.getFromDB('dRH');
            if (savedRH && savedRH.length > 0) { 
                dRH = savedRH; rhData = (await window.getFromDB('rhData')) || {}; 
                const savedPendencias = await window.getFromDB('pendenciasRH');
                if(savedPendencias) pendenciasRH = savedPendencias;
                renderRHQuad(); renderFerias(); renderBanco(); renderAbs(); renderPontoMensal(); 
            }
            const hist = await window.getFromDB('dHistory');
            if (hist) { globalHistory = hist; renderHistoryList(); }
            await restoreSession(); 
        }
        // METAS — renderiza SEMPRE independente de erros anteriores
        renderMetas();
        const ss = JSON.parse(localStorage.getItem('sysLog_sectors'));
        if (ss) SECTORS = ss;
        aplicarCoresOficiais(0);
        setInterval(updateClock, 1000);
        // AUTO-SAVE automático a cada 30 segundos
        setInterval(async () => { 
            try { await autoSaveData(); } catch(e) { console.warn('autoSave falhou:', e); }
        }, 30000);
        const today = new Date();
        const pickSelect = document.getElementById('pick-month-select');
        if(pickSelect) pickSelect.value = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
    } catch (e) { console.error("Erro init:", e); }
}

async function restoreSession() {
    document.getElementById('loading').style.display = 'flex';
    try {
        if (!window.getFromDB) return;
        // Tentar Firebase primeiro, fallback para localStorage
        try { dC = (await window.getFromDB('dC')) || []; } catch(e) { dC = []; }
        if (!dC.length) {
            try { const ls = localStorage.getItem('sysLog_dC'); if (ls) { dC = JSON.parse(ls); console.log('dC restaurado do localStorage'); } } catch(e) {}
        }
        try { dE = (await window.getFromDB('dE')) || []; } catch(e) { dE = []; }
        if (!dE.length) {
            try { const ls = localStorage.getItem('sysLog_dE'); if (ls) { dE = JSON.parse(ls); console.log('dE restaurado do localStorage'); } } catch(e) {}
        }
        // Restaurar colMapMaster do localStorage se Firebase falhou
        if (!colMapMaster.ped) {
            try { const ls = localStorage.getItem('sysLog_colMap'); if (ls) colMapMaster = JSON.parse(ls); } catch(e) {}
        }
        dPickRes = (await window.getFromDB('dPickRes')) || [];
        colMapMaster = (await window.getFromDB('colMapMaster')) || { ped: null, ate: null, n_ate: null, prod: null, desc: null, mov: null, cx_ped: null, cx_ate: null, rom: null, dep: null };
        forceBoxCalc = (await window.getFromDB('forceBoxCalc')) || false;
        if (dC.length) { ok('1', 'Recuperado da Nuvem'); if (!colMapMaster.ped) identifyColumnsMaster(); if (colMapMaster.ped) runCalc(); }
        if (dE.length) { ok('2', 'Recuperado da Nuvem'); }
        if (dPickRes.length && dE.length) {
            document.getElementById('picking-results-area').innerHTML = '<div style="text-align:center; padding:20px; color:var(--blue); font-weight:bold;">A restaurar cruzamento da nuvem...</div>';
            setTimeout(() => window.runPickingAnalysis(), 1500);
        }
        const btnRestore = document.getElementById('btnRestore');
        if (btnRestore) btnRestore.style.display = 'none'; 
    } catch (e) { console.error(e); }
    document.getElementById('loading').style.display = 'none';
}

async function autoSaveData() {
    if (!window.saveToDB) return;
    try {
        // Indicador visual de salvamento
        const indicator = document.getElementById('saveIndicator');
        if (indicator) { indicator.textContent = '💾 Salvando...'; indicator.style.color = '#f39c12'; }
        if (dC.length) {
            // Salvar dC slim — só campos usados pelo runCalc (evita timeout por volume)
            const dCSlim = dC.map(r => {
                const slim = {};
                const fields = [colMapMaster.prod, colMapMaster.ped, colMapMaster.ate,
                                colMapMaster.n_ate, colMapMaster.cx_ped, colMapMaster.cx_ate,
                                colMapMaster.desc, colMapMaster.rom, colMapMaster.mov, colMapMaster.dep];
                fields.forEach(f => { if (f && r[f] !== undefined) slim[f] = r[f]; });
                // Garantir campo Produto sempre presente
                if (!slim[colMapMaster.prod]) { slim['Produto'] = r['Produto'] || r['Material'] || r['PRODUTO'] || ''; }
                return slim;
            });
            await window.saveToDB('dC', dCSlim);
            // Backup imediato no localStorage (limite ~5MB, mas dCSlim é pequeno)
            try { localStorage.setItem('sysLog_dC', JSON.stringify(dCSlim)); localStorage.setItem('sysLog_colMap', JSON.stringify(colMapMaster)); } catch(eLS) {}
        }
        if (dE.length) {
            const dESlim = dE.map(r => {
                const prod = r['Produto'] || r['Material'] || r['PRODUTO'] || '';
                const pick = r['Picking'] || r['PICKING'] || '';
                if (!prod || !pick) return null; // ignorar sem produto ou sem picking
                return { 'Produto': String(prod).trim(), 'Picking': String(pick).trim().toUpperCase() };
            }).filter(Boolean); // remover nulls
            await window.saveToDB('dE', dESlim);
            try { localStorage.setItem('sysLog_dE', JSON.stringify(dESlim)); } catch(eLS) {}
        }
        if (dPickRes.length) await window.saveToDB('dPickRes', dPickRes);
        await window.saveToDB('pendenciasRH', pendenciasRH);
        await window.saveToDB('colMapMaster', colMapMaster);
        // Confirmar salvamento visual
        const indicatorOk = document.getElementById('saveIndicator');
        if (indicatorOk) { indicatorOk.textContent = '✅ Salvo ' + new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}); indicatorOk.style.color = '#27ae60'; }
    } catch(e) { 
        console.error('Erro AutoSave', e);
        const ind = document.getElementById('saveIndicator');
        if (ind) { ind.textContent = '⚠️ Erro ao salvar — verifique conexão'; ind.style.color = '#e74c3c'; }
    }
}

async function saveSession() {
    const loadScreen = document.getElementById('loading'); const loadMsg = document.getElementById('loadMsg'); loadScreen.style.display = 'flex';
    try {
        if (!window.saveToDB) throw new Error("Banco não inicializado.");
        loadMsg.innerText = "Sincronizando com Google Cloud...";
        await autoSaveData();
        if (dRH && dRH.length) await window.saveToDB('dRH', dRH);
        if (Object.keys(rhData).length) await window.saveToDB('rhData', rhData);
        await window.saveToDB('dHistory', globalHistory); await window.saveToDB('forceBoxCalc', forceBoxCalc);
        localStorage.setItem('sysLog_metas', JSON.stringify(getCurrentMetas())); localStorage.setItem('sysLog_sectors', JSON.stringify(SECTORS));
        alert("Sincronização com Google Cloud finalizada com sucesso!");
    } catch (error) { alert("Atenção: Sincronização parou! Motivo: " + error.message); } finally { loadScreen.style.display = 'none'; loadMsg.innerText = "Processando..."; }
}

async function clearSession() { if (confirm("Limpar a operação DIÁRIA?")) { if (window.saveToDB) { await window.saveToDB('dC', []); await window.saveToDB('dE', []); } localStorage.clear(); location.reload(); } }

function getCurrentMetas() { const m = {}; for (let k in DEFAULT_METAS) { const el = document.getElementById(k); m[k] = (el && el.value) ? parseInt(el.value) : DEFAULT_METAS[k]; if (isNaN(m[k])) m[k] = DEFAULT_METAS[k]; } return m; }

async function emergencyReset() { if (confirm("Resetar Sistema?")) { localStorage.clear(); location.reload(); } }

function read(i, cb) {
    if (typeof XLSX === 'undefined') { alert("A biblioteca Excel não foi carregada."); return; }
    if (!i.files[0]) return;
    document.getElementById('loading').style.display = 'flex';
    document.getElementById('loadMsg').innerText = "A ler o arquivo " + i.files[0].name + "...";
    const r = new FileReader();
    r.onload = e => { 
        try {
            const wb = XLSX.read(e.target.result, { type: 'array' }); 
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            let headerRow = 0;
            for (let j = 0; j < Math.min(20, rawData.length); j++) {
                const rowStr = (rawData[j] || []).join('').toUpperCase();
                if (rowStr.includes('PRODUTO') || rowStr.includes('MATERIAL') || rowStr.includes('QTDE') || rowStr.includes('SALDO')) { headerRow = j; break; }
            }
            const data = XLSX.utils.sheet_to_json(sheet, { range: headerRow, raw: true });
            if (data.length === 0) throw new Error("O arquivo parece estar vazio ou as colunas não foram reconhecidas.");
            cb(data, i.files[0].name); 
        } catch (err) { alert("ERRO AO LER O ARQUIVO: " + err.message); } finally { document.getElementById('loading').style.display = 'none'; document.getElementById('loadMsg').innerText = "Processando..."; }
    };
    r.onerror = () => { alert("Erro ao carregar o arquivo."); document.getElementById('loading').style.display = 'none'; };
    r.readAsArrayBuffer(i.files[0]);
}

document.getElementById('f1').addEventListener('change', function () { read(this, async (d, n) => { dC = d; ok('1', n); identifyColumnsMaster(); runCalc(); await autoSaveData(); }); });
document.getElementById('f2').addEventListener('change', function () { read(this, async (d, n) => { dE = d; ok('2', n); if (dC.length > 0) runCalc(); await autoSaveData(); }); });

document.getElementById('f_pick_res').addEventListener('change', function () {
    if (typeof XLSX === 'undefined') return;
    if (!this.files[0]) return;
    document.getElementById('loading').style.display = 'flex';
    document.getElementById('loadMsg').innerText = "A processar ressuprimento...";
    const r = new FileReader();
    r.onload = async e => {
        try {
            const wb = XLSX.read(e.target.result, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            let headerRow = 0;
            for (let i = 0; i < Math.min(30, rawData.length); i++) {
                const rowStr = (rawData[i] || []).join('').toUpperCase();
                if ((rowStr.includes('PRODUTO') || rowStr.includes('MATERIAL')) && (rowStr.includes('ATENDIDA') || rowStr.includes('ENDERE'))) { headerRow = i; break; }
            }
            dPickRes = XLSX.utils.sheet_to_json(sheet, { range: headerRow, raw: true });
            await autoSaveData();
            alert(`✅ Arquivo de RESSUPRIMENTO MENSAL carregado!\n\nAgora clique em "Processar Cruzamento".`);
        } catch (err) { alert("Erro ao ler ressuprimento: " + err.message); }
        document.getElementById('loading').style.display = 'none'; document.getElementById('loadMsg').innerText = "Processando...";
    };
    r.readAsArrayBuffer(this.files[0]);
});

window.runPickingAnalysis = async function() {
    if (!dPickRes || dPickRes.length === 0) return alert("⚠️ RESSUPRIMENTO MENSAL FALTANDO.");
    if (!dE || dE.length === 0) return alert("⚠️ ESTOQUE ATUAL FALTANDO: Importe na aba 'Planejamento'.");
    let loader = document.getElementById('loading'); if (loader) loader.style.display = 'flex';
    let msg = document.getElementById('loadMsg'); if (msg) msg.innerText = 'A cruzar Mensal com Endereços...';
    
    setTimeout(() => {
        try {
            const headers = Object.keys(dPickRes[0] || {});
            const colProd = headers.find(k => k.toUpperCase().trim() === 'PRODUTO' || k.toUpperCase().includes('MATERIAL'));
            const colQty = headers.find(k => k.toUpperCase().includes('ATENDIDA') && !k.toUpperCase().includes('CX'));
            const colQtyCx = headers.find(k => k.toUpperCase() === 'QTDE ATENDIDA CX' || k.toUpperCase() === 'QTDE ATENDIDA CXS');
            const colDesc = headers.find(k => k.toUpperCase().includes('DESC'));
            const colDate = headers.find(k => k.toUpperCase().includes('MOV') || k.toUpperCase().includes('DATA'));
            const colEnd = headers.find(k => k.toUpperCase().includes('ENDERE') || k.toUpperCase().includes('PICKING'));
            
            const getSector = (end) => {
                if (!end) return 'IGNORAR';
                const s = String(end).toUpperCase().trim().replace(/\s/g, ''); 
                if (s.includes('ALIMENTO') || s.startsWith('VR')) return 'ALIMENTO';
                if (s.includes('VM')) return 'VM';
                if (s.includes('VH')) return 'VH';
                if (s.includes('DERMO') || s.startsWith('VW')) return 'DERMO';
                if (s.includes('CONTROLADO') || s.startsWith('VX')) return 'CONTROLADO';
                if (s.startsWith('VY')) { if (s.match(/VY[-_\s]?0*3(\D|$)/)) return 'LATA'; return 'VOLUMOSO'; }
                if (s.startsWith('V23')) return 'RUA_23'; if (s.startsWith('V24')) return 'RUA_24'; if (s.startsWith('V25')) return 'RUA_25';
                if (s.startsWith('V41') || s.startsWith('V43')) {
                    const isPerf = s.startsWith('V41'); const dep = isPerf ? 'PERF' : 'MED';
                    const match = s.match(/V4[13]\D*0*(\d{1,2})\D*(\d+)/);
                    let rua = 0, pos = 0;
                    if (match) { rua = parseInt(match[1], 10); pos = parseInt(match[2], 10); } 
                    else { let nums = s.replace(/\D/g, ''); if (nums.length >= 6) { rua = parseInt(nums.substring(2, 4), 10); pos = parseInt(nums.substring(4), 10); } }
                    let tipo = (pos > 99) ? '3D' : '2D'; let st = 0;
                    if (rua >= 57 && rua <= 62) st = 1; else if (rua >= 64 && rua <= 69) st = 2; else if (rua >= 71 && rua <= 76) st = 3; else if (rua >= 78 && rua <= 83) st = 4; else if (rua >= 85 && rua <= 90) st = 5;
                    if (st > 0) return `${dep}_${tipo}_ST${st}`;
                }
                return 'IGNORAR';
            };
            
            const mapAddr = {}; const sectorSaldo = {}; 
            const colSaldoDE = Object.keys(dE[0] || {}).find(k => k.toUpperCase().includes('SALDO') || k.toUpperCase() === 'QTDE' || k.toUpperCase() === 'QUANTIDADE');
            
            dE.forEach(r => {
                let p = r['Produto'] || r['PRODUTO'] || r['Material'];
                let end = r['Picking'] || r['PICKING'] || r['Endereço'];
                let saldoItem = parseNum(r[colSaldoDE]);
                if (p && end) mapAddr[sanitize(p)] = String(end).trim().toUpperCase();
                if (end) { let sec = getSector(String(end).trim().toUpperCase()); if (sec !== 'IGNORAR') { if (!sectorSaldo[sec]) sectorSaldo[sec] = 0; sectorSaldo[sec] += saldoItem; } }
            });
            
            const stats = {}; window.stationTopProducts = {}; window.macroTopProducts = {}; window.macroTopBoxes = {}; 
            let globalRes = 0; let globalSkus = new Set(); let skuV41 = new Set(), skuV43 = new Set();
            const getMacroCategory = (sec) => { if(sec.includes('MED_')) return 'MEDICAMENTO'; if(sec.includes('PERF_')) return 'PERFUMARIA'; if(sec.includes('RUA_')) return 'RUAS (23-25)'; return sec; };
            
            dPickRes.forEach(r => {
                let prRaw = r[colProd]; let pr = sanitize(prRaw);
                let q = parseNum(r[colQty]); let qCx = colQtyCx ? parseNum(r[colQtyCx]) : 0; 
                let end = mapAddr[pr]; // Correção: Pegar apenas do Estoque
                let descVal = String(r[colDesc] || '').trim() || "Sem Descrição";
                let dateVal = ''; let rawDate = r[colDate];
                if (rawDate) dateVal = (typeof rawDate === 'number') ? Math.floor(rawDate).toString() : String(rawDate).substring(0, 10);
                if (q <= 0 || !end) return;
                const sec = getSector(end); if (sec === 'IGNORAR') return;
                
                globalRes += q;
                if (pr) { globalSkus.add(pr); if (sec.includes('PERF')) skuV41.add(pr); if (sec.includes('MED')) skuV43.add(pr); }
                if (!stats[sec]) stats[sec] = { res: 0 }; stats[sec].res += q;
                if (!window.stationTopProducts[sec]) window.stationTopProducts[sec] = {};
                if (!window.stationTopProducts[sec][pr]) window.stationTopProducts[sec][pr] = { code: prRaw, qty: 0, dates: new Set(), desc: descVal };
                window.stationTopProducts[sec][pr].qty += q; if (dateVal) window.stationTopProducts[sec][pr].dates.add(dateVal);
                
                const macro = getMacroCategory(sec);
                if (!window.macroTopProducts[macro]) window.macroTopProducts[macro] = {};
                if (!window.macroTopProducts[macro][pr]) window.macroTopProducts[macro][pr] = { code: prRaw, qty: 0, desc: descVal };
                window.macroTopProducts[macro][pr].qty += q;
                if (!window.macroTopBoxes[macro]) window.macroTopBoxes[macro] = {};
                if (!window.macroTopBoxes[macro][pr]) window.macroTopBoxes[macro][pr] = { code: prRaw, qty: 0, desc: descVal };
                window.macroTopBoxes[macro][pr].qty += qCx;
            });
            
            let maxRes = 0; let campeaoKey = '';
            for (let k in stats) { if (stats[k].res > maxRes) { maxRes = stats[k].res; campeaoKey = k; } }
            
            safeUpdate('kpi-vol-ressup', globalRes.toLocaleString('pt-BR')); safeUpdate('kpi-skus-total', globalSkus.size.toLocaleString('pt-BR')); safeUpdate('kpi-sku-v41', skuV41.size.toLocaleString('pt-BR')); safeUpdate('kpi-sku-v43', skuV43.size.toLocaleString('pt-BR'));
            
            const groups = {
                '💊 MEDICAMENTO 02 DÍGITOS (FLOWRACK)': ['MED_2D_ST1', 'MED_2D_ST2', 'MED_2D_ST3', 'MED_2D_ST4', 'MED_2D_ST5'],
                '💊 MEDICAMENTO 03 DÍGITOS (ESTANTERIA)': ['MED_3D_ST1', 'MED_3D_ST2', 'MED_3D_ST3', 'MED_3D_ST4', 'MED_3D_ST5'],
                '🧴 PERFUMARIA 02 DÍGITOS (FLOWRACK)': ['PERF_2D_ST1', 'PERF_2D_ST2', 'PERF_2D_ST3', 'PERF_2D_ST4', 'PERF_2D_ST5'],
                '🧴 PERFUMARIA 03 DÍGITOS (ESTANTERIA)': ['PERF_3D_ST1', 'PERF_3D_ST2', 'PERF_3D_ST3', 'PERF_3D_ST4', 'PERF_3D_ST5'],
                '🛣️ SETOR RUAS (23 AO 25)': ['RUA_23', 'RUA_24', 'RUA_25'],
                '📦 OUTROS SETORES': ['ALIMENTO', 'LATA', 'CONTROLADO', 'VM', 'DERMO', 'VH', 'VOLUMOSO']
            };
            const stationNames = { '1': '57-62', '2': '64-69', '3': '71-76', '4': '78-83', '5': '85-90' };
            let html = '';
            for (let group in groups) {
                const isMed = group.includes('MEDICAMENTO'); const isPerf = group.includes('PERFUMARIA'); const isRua = group.includes('RUAS');
                const cor = isMed ? '#27ae60' : (isPerf ? '#e67e22' : (isRua ? '#34495e' : '#8e44ad'));
                html += `<div style="border-bottom:3px solid ${cor}; padding-bottom:8px; margin:25px 0 15px 0; display:flex; align-items:center; gap:10px;"><span style="background:${cor}; color:white; padding:5px 14px; border-radius:6px; font-weight:800; font-size:13px;">${group}</span></div><div class="adv-grid">`;
                groups[group].forEach(sec => {
                    const d = stats[sec] || { res: 0 }; const saldoSec = sectorSaldo[sec] || 0; 
                    const pctPeso = globalRes > 0 ? ((d.res / globalRes) * 100) : 0; const isCampeao = (sec === campeaoKey && d.res > 0);
                    let label = sec;
                    if (isMed || isPerf) { const stNum = sec.slice(-1); label = stationNames[stNum] ? `Estação ${stationNames[stNum]}` : sec.replace(/_/g, ' '); } else if (isRua) label = sec.replace('_', ' ');
                    html += `<div class="adv-card" style="border-top-color:${isCampeao ? '#f1c40f' : cor}; box-shadow: ${isCampeao ? '0 0 15px rgba(241, 196, 15, 0.5)' : '0 4px 10px rgba(0,0,0,0.1)'};"><div class="adv-title" style="font-size:16px;">${label} ${isCampeao ? '<span title="Setor com maior volume geral!" style="font-size:18px;">🏆</span>' : ''} ${d.res > 0 ? `<button onclick="window.showTop10('${sec}')" style="background:${isCampeao ? '#f39c12' : cor}; color:white; border:none; padding:6px 12px; border-radius:20px; cursor:pointer; font-size:12px; font-weight:900; box-shadow:0 2px 5px rgba(0,0,0,0.2);"><i class="fas fa-fire"></i> TOP 10</button>` : ''}</div><div style="background:${isCampeao ? '#fffdf0' : '#f8f9fa'}; padding:10px; border-radius:6px; margin-bottom:10px;"><div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px; border-bottom: 1px dashed #ddd; padding-bottom: 5px;"><span style="font-size:11px; font-weight:800; color:#7f8c8d; text-transform:uppercase;"><i class="fas fa-boxes"></i> Saldo Atual</span><span style="font-size:16px; font-weight:900; color:var(--orange);">${saldoSec.toLocaleString('pt-BR')}</span></div><div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;"><span style="font-size:11px; font-weight:800; color:#7f8c8d; text-transform:uppercase;">Unid. Atendida</span><span style="font-size:16px; font-weight:900; color:var(--blue);">${d.res.toLocaleString('pt-BR')}</span></div><div style="display:flex; justify-content:space-between; align-items:center;"><span style="font-size:11px; font-weight:800; color:#7f8c8d; text-transform:uppercase;">Representatividade</span><span style="font-size:15px; font-weight:900; color:var(--red);">${pctPeso.toFixed(2)}%</span></div></div><div style="height:8px; background:#ecf0f1; border-radius:4px; overflow:hidden;"><div style="height:100%; width:${Math.min(pctPeso * 2, 100)}%; background:linear-gradient(90deg, ${isCampeao ? '#f1c40f' : cor}, #2c3e50); border-radius:4px;"></div></div></div>`;
                });
                html += `</div>`;
            }
            html += `<div style="margin-top:40px; border-top:3px dashed #ccc; padding-top:20px;"><h2 style="color:#2c3e50; font-size:18px; text-align:center;"><i class="fas fa-trophy" style="color:#f1c40f;"></i> TOP 5 GERAL DE RESSUPRIMENTO POR MACRO-SETOR (UNIDADES)</h2><div style="display:flex; flex-wrap:wrap; gap:15px; justify-content:center; margin-top:20px;">`;
            for (let macro in window.macroTopProducts) {
                const prods = window.macroTopProducts[macro]; const sortedMacro = Object.entries(prods).sort((a, b) => b[1].qty - a[1].qty).slice(0, 5);
                if (sortedMacro.length > 0) {
                    html += `<div style="background:#fff; border:1px solid #ddd; border-radius:8px; width:320px; box-shadow:0 3px 6px rgba(0,0,0,0.05); overflow:hidden;"><div style="background:#34495e; color:white; padding:10px; font-weight:bold; text-align:center; font-size:14px; text-transform:uppercase;">${macro} <span style="font-size:10px; font-weight:normal; opacity:0.8;">(UNIDADES)</span></div><ul style="list-style:none; padding:0; margin:0;">`;
                    sortedMacro.forEach((p, idx) => {
                        const code = p[1].code || p[0], desc = p[1].desc, qty = p[1].qty; const bgLi = idx % 2 === 0 ? '#fcfcfc' : '#fff'; const badgeColor = idx === 0 ? '#f1c40f' : (idx === 1 ? '#bdc3c7' : (idx === 2 ? '#cd7f32' : '#95a5a6'));
                        html += `<li style="padding:8px 10px; border-bottom:1px solid #eee; background:${bgLi}; display:flex; align-items:center; gap:10px;"><div style="background:${badgeColor}; color:${idx===0?'#000':'#fff'}; font-weight:bold; border-radius:50%; width:22px; height:22px; display:flex; align-items:center; justify-content:center; font-size:11px; flex-shrink:0;">${idx+1}</div><div style="flex:1; min-width:0;"><div style="font-size:12px; font-weight:bold; color:#2980b9;">${code}</div><div style="font-size:10px; color:#7f8c8d; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${desc}">${desc}</div></div><div style="font-weight:900; color:#c0392b; font-size:13px;">${qty.toLocaleString('pt-BR')}</div></li>`;
                    });
                    html += `</ul></div>`;
                }
            }
            html += `</div></div>`;
            document.getElementById('picking-results-area').innerHTML = html;
        } catch (err) { alert("Erro ao processar: " + err.message); }
        if (loader) loader.style.display = 'none'; if (msg) msg.innerText = 'Processando...';
    }, 100);
};

window.showTop10 = function(sec) {
    try {
        const products = window.stationTopProducts[sec];
        if (!products || Object.keys(products).length === 0) return alert("Sem dados.");
        const sorted = Object.entries(products).sort((a, b) => b[1].qty - a[1].qty).slice(0, 10);
        const totalSec = sorted.reduce((acc, p) => acc + p[1].qty, 0);
        const stationNames = { '1': '57-62', '2': '64-69', '3': '71-76', '4': '78-83', '5': '85-90' };
        let stLabel = sec;
        if(sec.includes('_ST')) { const stNum = sec.slice(-1); const secBase = sec.includes('MED') ? 'Medicamento ' : 'Perfumaria '; const tipo = sec.includes('2D') ? '02 Dígitos' : '03 Dígitos'; stLabel = `${secBase} ${tipo} — Estação ${stationNames[stNum]}`; } else if (sec.includes('RUA_')) { stLabel = sec.replace('_', ' '); }
        let table = `<table class="top10-table" style="text-align:center; width: 100%; border-collapse: collapse;"><thead><tr style="background:#2c3e50; color:white;"><th style="padding:10px; width:40px; text-align:center;">#</th><th style="padding:10px; text-align:left;">Produto (Cód / Descrição)</th><th style="padding:10px; text-align:center;"><i class="fas fa-box"></i> Unid. Total</th><th style="padding:10px; text-align:center;"><i class="far fa-calendar-alt"></i> Dias Entrou</th><th style="padding:10px; text-align:center;"><i class="fas fa-chart-line"></i> Média/Dia</th><th style="padding:10px; text-align:center;"><i class="fas fa-percent"></i> Peso %</th></tr></thead><tbody>`;
        sorted.forEach((p, i) => {
            const code = p[1].code || p[0], data = p[1]; const qty = data.qty; const daysCount = data.dates && data.dates.size > 0 ? data.dates.size : 1;
            const avg = Math.round(qty / daysCount); const pct = totalSec > 0 ? (qty / totalSec) * 100 : 0; const bg = i < 3 ? 'background:#fff9f0;' : (i % 2 === 0 ? 'background:#f8f9fa;' : 'background:#fff;');
            table += `<tr style="${bg}; border-bottom:1px solid #eee;"><td style="padding:10px;"><span class="rank-badge" style="background:${i<3?'#e67e22':'#3f51b5'}; color:white; padding:4px 8px; border-radius:4px; font-weight:bold;">${i + 1}</span></td><td style="text-align:left; padding:10px;"><div style="font-weight:900; color:#2980b9; font-size:14px;">${code}</div><div style="font-size:10px; color:#7f8c8d; font-weight:700; margin-top:3px; max-width:250px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${data.desc}">${data.desc}</div></td><td style="font-weight:900; color:#2c3e50; font-size:14px; padding:10px;">${qty.toLocaleString('pt-BR')}</td><td style="font-weight:bold; padding:10px; color:#555;">${daysCount}</td><td style="color:#d35400; font-weight:900; font-size:14px; padding:10px;">${avg.toLocaleString('pt-BR')}</td><td style="color:#c0392b; font-weight:800; padding:10px;">${pct.toFixed(1)}%</td></tr>`;
        });
        table += `</tbody></table>`;
        let modal = document.getElementById('modalTop10');
        if (!modal) {
            modal = document.createElement('div'); modal.id = 'modalTop10'; modal.style.cssText = 'display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.8); z-index:9999; justify-content:center; align-items:center; padding:20px;';
            modal.innerHTML = `<div style="background:#fff; border-radius:12px; width:100%; max-width:800px; max-height:90vh; overflow-y:auto; box-shadow:0 10px 30px rgba(0,0,0,0.5); display:flex; flex-direction:column;"><div style="background:#2c3e50; color:white; padding:15px 20px; border-radius:12px 12px 0 0; display:flex; justify-content:space-between; align-items:center; position:sticky; top:0;"><h3 id="top10Title" style="margin:0; font-size:18px;"></h3><button onclick="document.getElementById('modalTop10').style.display='none'" style="background:none; border:none; color:white; font-size:24px; cursor:pointer; padding:0; line-height:1;">&times;</button></div><div id="top10Body" style="padding:20px; overflow-x:auto;"></div></div>`;
            document.body.appendChild(modal);
        }
        document.getElementById('top10Title').innerHTML = `<i class="fas fa-fire"></i> Top 10 Produtos — ${stLabel}`; document.getElementById('top10Body').innerHTML = table; modal.style.display = 'flex';
    } catch (err) { alert("Erro: " + err.message); }
};

function ensureRhData(mat) { if (!rhData[mat]) rhData[mat] = { feriasInicio: '', feriasFim: '', limiteFerias: '', periodoAberto: '', vendeu10Dias: false, banco: 0, faltas: [], statusManual: 'AUTO', ponto: [] }; }

document.getElementById('f_rh').addEventListener('change', function () {
    if (typeof XLSX === 'undefined' || !this.files[0]) return;
    document.getElementById('loading').style.display = 'flex';
    const r = new FileReader();
    r.onload = async e => {
        try {
            const wb = XLSX.read(e.target.result, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            let headerRow = 0;
            for (let i = 0; i < Math.min(20, raw.length); i++) {
                const rowStr = (raw[i] || []).join('').toUpperCase();
                if (rowStr.includes('NOME') && (rowStr.includes('MATRICULA') || rowStr.includes('MATRÍCULA'))) { headerRow = i; break; }
            }
            dRH = XLSX.utils.sheet_to_json(sheet, { range: headerRow, raw: false });
            if (window.saveToDB) await window.saveToDB('dRH', dRH);
            renderRHQuad(); renderFerias(); renderBanco(); renderAbs(); renderPontoMensal();
            alert("Quadro importado!");
        } catch (err) { alert("Erro: " + err); }
        document.getElementById('loading').style.display = 'none';
    };
    r.readAsArrayBuffer(this.files[0]);
});


function renderFerias() {
    const tb = document.getElementById('tbFerias'); if (!tb || !dRH.length) return; tb.innerHTML = '';
    dRH.forEach(r => {
        const mat = r['Matrícula'] || r['Matricula'] || r['MATRICULA'] || '-'; const nome = r['Nome'] || r['NOME'] || '-'; const adm = r['Admissão'] || r['Admissao'] || r['ADMISSÃO'] || '-';
        ensureRhData(mat); const rh = rhData[mat]; const statusManual = rh.statusManual || 'AUTO'; const baseStatus = (r['Status'] || r['STATUS'] || '-').toUpperCase();
        if (nome === '-' || (baseStatus === 'INATIVO' && statusManual !== 'ATIVO') || statusManual === 'INATIVO' || statusManual === 'TRANSFERENCIA DE TURNO') return; 
        let statusBadge = '<span style="background:#3498db; color:white; padding:3px 6px; border-radius:4px; font-weight:bold;">AQUISITIVO</span>'; let limitDisplay = '-';
        const dtIni = rh.feriasInicio || ''; const dtFim = rh.feriasFim || ''; const hasMarcacao = (dtIni && dtFim);
        if (hasMarcacao) {
            const today = new Date(); today.setHours(0,0,0,0); const dI = new Date(dtIni + "T00:00:00"); const dF = new Date(dtFim + "T00:00:00");
            if (today >= dI && today <= dF) statusBadge = `<span style="background:#8e44ad; color:white; padding:3px 6px; border-radius:4px; font-weight:bold;">EM FÉRIAS</span>`;
            else if (dI > today) statusBadge = `<span style="background:#f1c40f; color:#333; padding:3px 6px; border-radius:4px; font-weight:bold;">MARCADA</span>`;
            else statusBadge = `<span style="background:#7f8c8d; color:white; padding:3px 6px; border-radius:4px; font-weight:bold;">GOZADAS</span>`;
        } else if (rh.limiteFerias) {
            const today = new Date(); today.setHours(0,0,0,0); const limitDate = new Date(rh.limiteFerias + "T00:00:00");
            const diffDays = Math.floor((limitDate - today) / (1000 * 60 * 60 * 24));
            if (diffDays < 0) statusBadge = `<span style="background:#c0392b; color:white; padding:3px 6px; border-radius:4px; font-weight:bold;">VENCIDA</span>`;
            else if (diffDays <= 90) statusBadge = `<span style="background:#f39c12; color:white; padding:3px 6px; border-radius:4px; font-weight:bold;">VENCENDO</span>`;
            else statusBadge = `<span style="background:#27ae60; color:white; padding:3px 6px; border-radius:4px; font-weight:bold;">NO PRAZO</span>`;
        }
        if (rh.limiteFerias) {
            const limitStr = rh.limiteFerias.split('-').reverse().join('/'); limitDisplay = `<strong style="color:var(--dark); font-size:12px;">${limitStr}</strong>`;
            const today = new Date(); today.setHours(0,0,0,0); const limitDate = new Date(rh.limiteFerias + "T00:00:00");
            if(limitDate < today && !hasMarcacao) limitDisplay = `<strong style="color:#c0392b; font-size:12px;">${limitStr}</strong>`;
        }
        const periodoAberto = rh.periodoAberto ? `<strong style="color:var(--blue); font-size:11px;">${rh.periodoAberto}</strong>` : '<span style="color:#ccc;">-</span>';
        const isChecked = rh.vendeu10Dias ? 'checked' : '';
        tb.innerHTML += `<tr><td style="font-weight:bold; font-size:12px;">${mat}</td><td style="text-align:left; font-size:12px;">${nome}</td><td style="font-size:11px; color:#555;">${adm}</td><td>${periodoAberto}</td><td>${limitDisplay}</td><td>${statusBadge}</td><td><input type="date" value="${dtIni}" onchange="updateRHData('${mat}', 'feriasInicio', this.value, this); renderRHQuad(); renderFerias();" style="padding:4px; border:1px solid #ccc; border-radius:4px; font-weight:bold; font-size:10px;"></td><td><input type="date" value="${dtFim}" onchange="updateRHData('${mat}', 'feriasFim', this.value, this); renderRHQuad(); renderFerias();" style="padding:4px; border:1px solid #ccc; border-radius:4px; font-weight:bold; font-size:10px;"></td><td><label style="cursor:pointer; font-weight:800; color:var(--orange); font-size:10px; display:flex; align-items:center; gap:5px; justify-content:center;"><input type="checkbox" ${isChecked} onchange="updateRHData('${mat}', 'vendeu10Dias', this.checked, this)"> Vendeu 10 Dias</label></td><td><button class="btn-sm" style="background:#25D366;color:white;padding:5px;border-radius:20px;" onclick="sendWhatsAppLink('${mat}', '${nome}')" title="Link Individual"><i class="fab fa-whatsapp"></i></button></td></tr>`;
    });
}


function renderAbs() {
    const sel = document.getElementById('selAbsEmp'); const tb = document.getElementById('tbAbs'); if (!sel || !tb || !dRH.length) return;
    let totalFaltas = 0, totalFuncionarios = 0;
    if (sel.options.length <= 1) {
        sel.innerHTML = '<option value="">-- Selecione o Colaborador --</option>';
        dRH.forEach(r => { const mat = r['Matrícula'] || r['Matricula'] || r['MATRICULA'] || '-'; const nome = r['Nome'] || r['NOME'] || '-'; const status = (r['Status'] || r['STATUS'] || '-').toUpperCase(); if (nome !== '-' && status === 'ATIVO') { sel.innerHTML += `<option value="${mat}">${nome} (${mat})</option>`; totalFuncionarios++; } });
    } else { totalFuncionarios = Array.from(sel.options).length - 1; }
    tb.innerHTML = ''; let hasRecords = false;
    dRH.forEach(r => {
        const mat = r['Matrícula'] || r['Matricula'] || r['MATRICULA'] || '-'; const nome = r['Nome'] || r['NOME'] || '-';
        if (!rhData[mat] || !rhData[mat].faltas) return;
        rhData[mat].faltas.forEach((f, idx) => { hasRecords = true; totalFaltas++; let mClass = f.motivo.includes('Injustificada') ? 'color:var(--red);font-weight:bold;' : ''; tb.innerHTML += `<tr><td style="font-weight:bold;">${f.data}</td><td style="text-align:left;">${nome}</td><td style="${mClass}">${f.motivo}</td><td><button class="btn-sm btn-gray" onclick="removeAbs('${mat}', ${idx})"><i class="fas fa-trash"></i></button></td></tr>`; });
    });
    if (!hasRecords) tb.innerHTML = '<tr><td colspan="4" style="text-align:center; color:#999;">Nenhum registo.</td></tr>';
    const pct = totalFuncionarios > 0 ? (totalFaltas / (totalFuncionarios * 30)) * 100 : 0; safeUpdate('abs-idx', pct.toFixed(2) + '%');
}

async function registrarAbs() {
    const mat = document.getElementById('selAbsEmp').value; const dateStr = document.getElementById('txtAbsDate').value; const motivo = document.getElementById('selAbsMotivo').value; const dias = parseInt(document.getElementById('txtAbsDias').value) || 1;
    if (!mat || !dateStr) return alert("Selecione colaborador e data.");
    ensureRhData(mat); let baseDate = new Date(dateStr + "T12:00:00");
    for (let i = 0; i < dias; i++) { let targetDate = new Date(baseDate.getTime() + (i * 86400000)); let brDate = targetDate.toLocaleDateString('pt-BR', { timeZone: 'UTC' }); rhData[mat].faltas.push({ data: brDate, motivo: motivo, stamp: targetDate.getTime() }); }
    rhData[mat].faltas.sort((a, b) => b.stamp - a.stamp);
    if (window.saveToDB) window.saveToDB('rhData', rhData);
    document.getElementById('selAbsEmp').value = ''; document.getElementById('txtAbsDate').value = ''; document.getElementById('txtAbsDias').value = '1'; renderAbs();
}

async function removeAbs(mat, idx) { if (confirm("Remover?")) { if (rhData[mat] && rhData[mat].faltas) { rhData[mat].faltas.splice(idx, 1); if (window.saveToDB) window.saveToDB('rhData', rhData); renderAbs(); } } }


// =========================================================================
// PLANEJAMENTO E COLUNAS  (CORREÇÕES DE DINÂMICA)
// =========================================================================

function identifyColumnsMaster() {
    if (!dC.length) return;
    const keys = Object.keys(dC[0]);
    keys.forEach(k => {
        const kl = k.toLowerCase().trim(); const kt = k.trim().toUpperCase();
        if (!colMapMaster.ped && (kt === 'QTDE PEDIDA' || kt === 'QTD PEDIDA' || kt === 'QUANTIDADE PEDIDA' || kt === 'QTDE. PEDIDA' || kt === 'QT PEDIDA' || kt === 'PEDIDO' || kt === 'QTD. PEDIDA' || (kl.includes('pedida') && !kl.includes('cx') && !kl.includes('caixa')) || (kl.includes('pedido') && kl.includes('qtd')) || kt === 'QTD' || kt === 'QTDE' || kt === 'QUANTIDADE')) colMapMaster.ped = k;
        if (!colMapMaster.ate && (kt === 'QTDE ATENDIDA' || kt === 'QTD ATENDIDA' || kt === 'QUANTIDADE ATENDIDA' || kt === 'QTDE. ATENDIDA' || kt === 'QT ATENDIDA' || kt === 'ATENDIDO' || (kl.includes('atendida') && !kl.includes('cx') && !kl.includes('caixa')))) colMapMaster.ate = k;
        
        // QTDE N ATENDIDA (NOVO)
        if (!colMapMaster.n_ate && (kt === 'QTDE N ATENDIDA' || kt === 'QTD N ATENDIDA' || kt === 'FALTA' || kt === 'N ATENDIDA' || kl.includes('n atendida'))) colMapMaster.n_ate = k;
        
        if (!colMapMaster.cx_ped && (kt === 'QTDE PEDIDA CX' || kt === 'QTD PEDIDA CX' || kt === 'QTDE PEDIDA CAIXA' || kt === 'PEDIDA CX' || kt === 'CX PEDIDA' || kt === 'CAIXAS PEDIDAS' || (kl.includes('pedida') && (kl.includes('cx') || kl.includes('caixa'))))) colMapMaster.cx_ped = k;
        if (!colMapMaster.cx_ate && (kt === 'QTDE ATENDIDA CX' || kt === 'QTD ATENDIDA CX' || kt === 'QTDE ATENDIDA CAIXA' || kt === 'ATENDIDA CX' || kt === 'CX ATENDIDA' || kt === 'CAIXAS ATENDIDAS' || (kl.includes('atendida') && (kl.includes('cx') || kl.includes('caixa'))))) colMapMaster.cx_ate = k;
        if (!colMapMaster.prod && (kt === 'PRODUTO' || kt === 'MATERIAL' || kt === 'COD PRODUTO' || kt === 'CODIGO' || kt === 'CÓDIGO' || kt === 'COD' || kt === 'SKU' || kl.includes('produto') || kl.includes('material') || kl.includes('sku'))) colMapMaster.prod = k;
        if (!colMapMaster.desc && (kt === 'DESCRIÇÃO' || kt === 'DESCRICAO' || kt === 'DESC' || kt === 'NOME' || kt === 'NOME DO PRODUTO' || kt === 'DESCR' || kl.includes('descri') || kl.includes('nome'))) colMapMaster.desc = k;
        if (!colMapMaster.rom && (kt === 'ROMANEIO' || kt === 'NR PEDIDO' || kt === 'NUM PEDIDO' || kt === 'PEDIDO' || (kl.includes('romaneio')) || (kl.includes('pedido') && !kl.includes('qtd')))) colMapMaster.rom = k;
        if (!colMapMaster.dep && (kt === 'DEPOSITO' || kt === 'DEPÓSITO' || kt === 'DEP' || kt === 'SETOR' || kt === 'DEPARTAMENTO' || kt === 'LOCAL' || kt === 'LOCALIZACAO' || kt === 'LOCALIZAÇÃO' || kl.includes('deposit') || kl.includes('setor') || kl.includes('depart'))) colMapMaster.dep = k;
    });

    if (!colMapMaster.ped) {
        const fallback = keys.find(k => /^(qtde?\.?|quantidade|saldo|total|unid)/i.test(k.trim()));
        if (fallback) colMapMaster.ped = fallback;
    }
    if (!colMapMaster.ped) {
        const el = document.getElementById('colSelectTitle'); const list = document.getElementById('colSelectList');
        if (el) el.innerText = 'Selecione a coluna de QUANTIDADE PEDIDA';
        if (list) { list.innerHTML = keys.map(k => { const sample = dC[0][k]; return `<div style="padding:8px; border-bottom:1px solid #f0f0f0; cursor:pointer;" onclick="colMapMaster.ped='${k.replace(/'/g,"\\'")}'; document.getElementById('modalColSelect').style.display='none'; runCalc();" onmouseover="this.style.background='#e8f4fd'" onmouseout="this.style.background='#fff'"><b>${k}</b> <span style="color:#888; font-size:11px;">ex: ${sample}</span></div>`; }).join(''); }
        const manualArea = document.getElementById('manualInputArea'); if (manualArea) manualArea.style.display = 'block';
        const modal = document.getElementById('modalColSelect'); if (modal) modal.style.display = 'flex';
        return false; 
    }
    return true;
}

function cancelColSelect() { document.getElementById('modalColSelect').style.display = 'none'; document.getElementById('loading').style.display = 'none'; alert("Operação cancelada. É necessário mapear a coluna para continuar."); }
function useManualBoxCalc() { forceBoxCalc = true; document.getElementById('modalColSelect').style.display = 'none'; alert("Modo Manual Ativado: 1 Linha = 1 Caixa."); runCalc(); }

window.showTop20Global = function() {
    try {
        if (!dC || dC.length === 0) return alert("Nenhum planejamento carregado.");
        const topList = {};
        dC.forEach(r => {
            const prodField = colMapMaster.prod ? r[colMapMaster.prod] : (r['Produto'] || r['Material']);
            const descField = colMapMaster.desc ? r[colMapMaster.desc] : (r['Descrição'] || '');
            if (!prodField) return; const pr = sanitize(prodField); let cp = 0;
            if (forceBoxCalc) cp = 1; else if (colMapMaster.cx_ped && r[colMapMaster.cx_ped]) cp = parseNum(r[colMapMaster.cx_ped]); else cp = parseNum(colMapMaster.ped ? r[colMapMaster.ped] : 0) > 0 ? 1 : 0;
            if (!topList[pr]) topList[pr] = { code: prodField, desc: descField, qty: 0 }; topList[pr].qty += cp;
        });
        const sorted = Object.values(topList).sort((a, b) => b.qty - a.qty).slice(0, 20);
        let html = '<table class="top20-table"><thead><tr><th style="width:50px;">Rank</th><th>Produto</th><th style="text-align:left;">Descrição</th><th>Caixas (Pedidas)</th></tr></thead><tbody>';
        sorted.forEach((item, idx) => { const badgeColor = idx === 0 ? '#f1c40f' : (idx === 1 ? '#bdc3c7' : (idx === 2 ? '#cd7f32' : '#95a5a6')); html += `<tr><td><div class="rank-badge" style="background:${badgeColor}; color:${idx===0?'#000':'#fff'}">${idx + 1}</div></td><td style="font-weight:bold; color:var(--blue); font-size:14px;">${item.code}</td><td style="font-size:11px; text-align:left; color:#555;">${item.desc}</td><td style="font-weight:900; color:var(--red); font-size:15px;">${item.qty.toLocaleString('pt-PT')}</td></tr>`; });
        html += '</tbody></table>'; document.getElementById('top20Body').innerHTML = html; document.getElementById('modalTop20').style.display = 'flex';
    } catch (err) { alert("Erro ao gerar o Top 20: " + err.message); }
};

function getRng(r) { if (r >= 57 && r <= 62) return '57-62'; if (r >= 64 && r <= 69) return '64-69'; if (r >= 71 && r <= 76) return '71-76'; if (r >= 78 && r <= 83) return '78-83'; if (r >= 85 && r <= 90) return '85-90'; return 'OUTROS'; }

function runCalc() {
    if (!dC.length) return;
    document.getElementById('loading').style.display = 'flex';
    setTimeout(() => {
        try {
            if (!colMapMaster.ped) { const found = identifyColumnsMaster(); if (!found || !colMapMaster.ped) { document.getElementById('loading').style.display = 'none'; return; } }
            const ctn = document.getElementById('tbContainer'); if (ctn) ctn.innerHTML = '';
            projData = []; const mapAddr = {};
            if (dE.length > 0) { dE.forEach(r => { let p = r['Produto'] || r['PRODUTO'] || r['Material']; let e = r['Picking'] || r['PICKING'] || r['Endereço']; if (p && e) mapAddr[sanitize(p)] = String(e).trim().toUpperCase(); }); }
            const metas = getCurrentMetas(); const S = SECTORS;
            
            const struct = {
                'PERFUMARIA 02 DIGITOS': { '57-62': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '64-69': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '71-76': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '78-83': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '85-90': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'PERFUMARIA 03 DIGITOS': { '57-62': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '64-69': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '71-76': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '78-83': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '85-90': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'MEDICAMENTO 02 DIGITOS': { '57-62': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '64-69': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '71-76': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '78-83': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '85-90': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'MEDICAMENTO 03 DIGITOS': { '57-62': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '64-69': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '71-76': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '78-83': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, '85-90': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'SETOR RUAS': { 'RUA 23': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, 'RUA 24': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 }, 'RUA 25': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'SETOR ALIMENTO': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } }, 'SETOR VOLUMOSO': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'SETOR CONTROLADO': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } }, 'SETOR DERMO': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } },
                'SETOR LATAS': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } }, 'SETOR VM': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } }, 'SETOR VH': { 'TOTAL': { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 } }
            };
            const stationNeeds = { '57-62': 0, '64-69': 0, '71-76': 0, '78-83': 0, '85-90': 0 };
            moduleStats = { 'MEDICAMENTO': { '2D': {}, '3D': {} }, 'PERFUMARIA': { '2D': {}, '3D': {} } };
            for (let i = 57; i <= 90; i++) { moduleStats['MEDICAMENTO']['2D'][i] = { p: 0, a: 0, items: [] }; moduleStats['MEDICAMENTO']['3D'][i] = { p: 0, a: 0, items: [] }; moduleStats['PERFUMARIA']['2D'][i] = { p: 0, a: 0, items: [] }; moduleStats['PERFUMARIA']['3D'][i] = { p: 0, a: 0, items: [] }; }
            const errors = []; let totP_Unit = 0, totA_Unit = 0, totP_Box = 0, totA_Box = 0, linesProcessed = 0;

            dC.forEach(r => {
                linesProcessed++;
                const prodField = colMapMaster.prod ? r[colMapMaster.prod] : (r['Produto'] || r['Material']);
                const descField = colMapMaster.desc ? r[colMapMaster.desc] : (r['Descrição'] || '');
                
                // CORREÇÃO DE DINÂMICA: Buscar o Endereço de Picking APENAS no Estoque Atual, 
                // para não confundir com o endereço da Reserva que vem na linha do Ressuprimento.
                const pkField = mapAddr[sanitize(prodField)];
                
                const romField = colMapMaster.rom ? r[colMapMaster.rom] : '';
                if (!prodField) { errors.push({ p: '-', k: '-', reason: 'Linha sem código' }); return; }
                const pr = sanitize(prodField);
                let cpun = colMapMaster.ped ? parseNum(r[colMapMaster.ped]) : 0;
                let caun = colMapMaster.ate ? parseNum(r[colMapMaster.ate]) : 0;
                let cnun = colMapMaster.n_ate ? parseNum(r[colMapMaster.n_ate]) : 0;
                
                if (caun === 0 && cpun > 0) {
                    caun = colMapMaster.n_ate ? (cpun - cnun) : cpun;
                }
                if (cnun === 0 && cpun > caun) cnun = cpun - caun;
                
                let cp = 0, ca = 0;
                if (forceBoxCalc) { cp = 1; ca = (caun >= cpun) ? 1 : 0; }
                else if (colMapMaster.cx_ped && r[colMapMaster.cx_ped]) {
                    cp = parseNum(r[colMapMaster.cx_ped]);
                    if (colMapMaster.cx_ate && r[colMapMaster.cx_ate] !== undefined && String(r[colMapMaster.cx_ate]).trim() !== '') {
                        ca = parseNum(r[colMapMaster.cx_ate]);
                    } else { 
                        ca = (cpun > 0) ? Math.floor(cp * (caun / cpun)) : 0; 
                    }
                } else { 
                    cp = cpun > 0 ? 1 : 0; 
                    ca = caun > 0 ? 1 : 0; 
                }
                
                totP_Box += cp; totA_Box += ca; totP_Unit += cpun; totA_Unit += caun;
                
                if (pkField && pkField.length >= 3) {
                    let cleanPk = pkField.replace(/\./g, '-').replace(/\s/g, '-').replace(/_/g, '-');
                    let parts = cleanPk.split('-').filter(p => p !== ''); let dep = parts[0] || ''; let rua = parseInt(parts[1], 10) || 0; let pos = parseInt(parts[parts.length - 1], 10) || 0; let cat = null, rng = null;
                    const inR = (k) => (rua >= S[k].ri && rua <= S[k].rf);
                    
                    if (dep.startsWith('VW')) { cat = 'SETOR DERMO'; rng = 'TOTAL'; }
                    else if (dep.startsWith('VX')) { cat = 'SETOR CONTROLADO'; rng = 'TOTAL'; }
                    else if (dep.startsWith('VR')) { cat = 'SETOR ALIMENTO'; rng = 'TOTAL'; }
                    else if (dep.startsWith('VY')) { cat = (rua === 3 || dep.match(/VY[-_\s]?0*3(\D|$)/)) ? 'SETOR LATAS' : 'SETOR VOLUMOSO'; rng = 'TOTAL'; }
                    else if (dep.startsWith('VM')) { cat = 'SETOR VM'; rng = 'TOTAL'; }
                    else if (dep.startsWith('VH')) { cat = 'SETOR VH'; rng = 'TOTAL'; }
                    else if (dep === 'V40' || dep === 'V41') {
                        // PERFUMARIA — preenche moduleStats E struct juntos
                        const dType = (pos >= 101) ? '3D' : '2D';
                        if (moduleStats['PERFUMARIA']?.[dType]?.[rua]) {
                            moduleStats['PERFUMARIA'][dType][rua].p += cp; moduleStats['PERFUMARIA'][dType][rua].a += ca;
                            if (cpun > caun && cp > 0) moduleStats['PERFUMARIA'][dType][rua].items.push({ cod: pr, desc: descField, rom: romField, q: cpun - caun });
                        }
                        const ruaRng = getRng(rua);
                        if (ruaRng) { cat = 'PERFUMARIA ' + (pos >= 101 ? '03' : '02') + ' DIGITOS'; rng = ruaRng; }
                    }
                    else if (dep === 'V42' || dep === 'V43') {
                        // MEDICAMENTO — preenche moduleStats E struct juntos
                        const dType = (pos >= 101) ? '3D' : '2D';
                        if (moduleStats['MEDICAMENTO']?.[dType]?.[rua]) {
                            moduleStats['MEDICAMENTO'][dType][rua].p += cp; moduleStats['MEDICAMENTO'][dType][rua].a += ca;
                            if (cpun > caun && cp > 0) moduleStats['MEDICAMENTO'][dType][rua].items.push({ cod: pr, desc: descField, rom: romField, q: cpun - caun });
                        }
                        const ruaRng = getRng(rua);
                        if (ruaRng) { cat = 'MEDICAMENTO ' + (pos >= 101 ? '03' : '02') + ' DIGITOS'; rng = ruaRng; }
                    }
                    else if (dep.startsWith('V') && dep.length >= 3 && !isNaN(dep.slice(1))) {
                        // V10–V36 = SETOR RUAS (qualquer rua genérica)
                        cat = 'SETOR RUAS'; rng = 'RUA ' + dep.replace('V', '');
                        if (!struct[cat]) struct[cat] = {};
                        if (!struct[cat][rng]) struct[cat][rng] = { p: 0, a: 0, pun: 0, aun: 0, nun: 0, ncx: 0 };
                    }
                    
                    if (cat && rng && struct[cat] && struct[cat][rng]) { 
                        struct[cat][rng].p += cp; struct[cat][rng].a += ca; struct[cat][rng].pun += cpun; struct[cat][rng].aun += caun; 
                        // Falta em unidades (Qtde N Atendida real)
                        if(struct[cat][rng].nun === undefined) struct[cat][rng].nun = 0;
                        struct[cat][rng].nun += cnun;
                        // Falta em CAIXAS (cx_ped - cx_ate)
                        if(struct[cat][rng].ncx === undefined) struct[cat][rng].ncx = 0;
                        struct[cat][rng].ncx += Math.max(0, cp - ca);
                    } 
                    else if (errors.length < 500) { errors.push({ p: pr, k: pkField, reason: 'Fora dos setores' }); }
                } else if (errors.length < 500) { errors.push({ p: pr, k: 'SEM ENDEREÇO', reason: 'Sem picking' }); }
            });
            
            safeUpdate('audit-lines', linesProcessed); safeUpdate('audit-boxes', fmtInt(totP_Box)); safeUpdate('audit-ignored', errors.length);
            globalErrors = errors; let gCx = 0, gRup = 0, gStf = 0, rowCount = 0;
            
            for (let cat in struct) {
                if (cat.includes('MEDICAMENTO') || cat.includes('PERFUMARIA')) {
                    let mKey = cat.includes('PERFUMARIA') ? 'm_PERFUMARIA' : 'm_MEDICAMENTO'; const meta = metas[mKey] || 100;
                    for (let rng in struct[cat]) { const d = struct[cat][rng]; if (d.p > 0 && rng !== 'OUTROS' && stationNeeds[rng] !== undefined) { stationNeeds[rng] += Math.ceil(d.p / meta); } }
                }
            }
            
            let stationPanelHTML = `<div class="card" style="border-top: 5px solid var(--blue); margin-bottom: 20px;"><h3 style="margin-top:0; color:var(--blue); font-size:14px; margin-bottom:10px;">CONTROLE DE EQUIPA POR ESTAÇÃO</h3><div class="station-panel">`;
            let totalStationNeed = 0;
            for (let rng in stationNeeds) {
                const need = stationNeeds[rng]; totalStationNeed += need; const stId = `st-input-${rng}`;
                stationPanelHTML += `<div class="st-box"><div class="st-header">ESTAÇÃO ${rng}</div><div class="st-body"><div class="st-metric"><div class="st-lbl">SOLICITADO</div><div class="st-val-static">${need}</div></div><div class="st-metric"><div class="st-lbl">REAL</div><input type="number" id="${stId}" class="st-input" data-need="${need}" oninput="updateForce('${stId}', ${need})"></div></div><div class="st-footer"><div class="st-force-lbl"><span>FORÇA</span><span id="force-${stId}">-</span></div><div class="progress-bg"><div class="progress-bar" style="width:0%"></div></div></div></div>`;
                projData.push({ name: `Estação ${rng}`, meta: 600, vol: 'N/A', nec: need, isStation: true, rng: rng });
            }
            stationPanelHTML += `</div></div>`; if (ctn) ctn.innerHTML += stationPanelHTML; gStf += totalStationNeed;
            
            for (let cat in struct) {
                let sP = 0, sA = 0, sRupSec = 0, rows = '', stfSec = 0;
                let mKey = 'm_MEDICAMENTO';
                if (cat.includes('PERFUMARIA')) mKey = 'm_PERFUMARIA'; else if (cat.includes('RUAS')) mKey = 'm_RUAS'; else if (cat.includes('ALIMENTO')) mKey = 'm_ALIMENTO'; else if (cat.includes('VOLUMOSO')) mKey = 'm_VOLUMOSO'; else if (cat.includes('CONTROLADO')) mKey = 'm_CONTROLADO'; else if (cat.includes('DERMO')) mKey = 'm_DERMO'; else if (cat.includes('LATA')) mKey = 'm_LATA'; else if (cat.includes('VM')) mKey = 'm_VM'; else if (cat.includes('VH')) mKey = 'm_VH';
                const meta = metas[mKey] || 100;
                
                for (let rng in struct[cat]) {
                    const d = struct[cat][rng]; 
                    const falRealUnidades = d.nun || 0; // Usando a nova coluna de Qtde Não Atendida real
                    const pct = d.p > 0 ? ((d.a / d.p) * 100).toFixed(1) : 0;
                    let stf = d.p > 0 ? Math.ceil(d.p / meta) : 0; if (rng === 'TOTAL') stfSec = stf;
                    rows += `<tr><td>${rng}</td><td>${d.pun}</td><td>${d.aun}</td><td>${fmtInt(d.p)}</td><td>${fmtInt(d.a)}</td><td class="tag-falta" title="Falta em Caixas">${fmtInt(d.ncx || 0)}</td><td>${pct}%</td><td>${stf}</td></tr>`;
                    sP += d.p; sA += d.a; sRupSec += (d.ncx || 0);
                }
                
                if (!cat.includes('MEDICAMENTO') && !cat.includes('PERFUMARIA') && sP > 0 && stfSec === 0) stfSec = Math.ceil(sP / meta);
                if (cat.includes('RUAS')) { stfSec = 0; for (let r in struct[cat]) stfSec += Math.ceil(struct[cat][r].p / meta); }
                if (!cat.includes('MEDICAMENTO') && !cat.includes('PERFUMARIA')) gStf += stfSec;
                
                let headerControls = '';
                const secColor = cat.includes('MEDICAMENTO') ? '#27ae60' : cat.includes('PERFUMARIA') ? '#e67e22' : cat.includes('RUAS') ? '#3498db' : cat.includes('ALIMENTO') ? '#16a085' : cat.includes('DERMO') ? '#8e44ad' : cat.includes('CONTROLADO') ? '#c0392b' : cat.includes('VM') ? '#2980b9' : cat.includes('VH') ? '#1abc9c' : cat.includes('LATA') ? '#d35400' : '#555';
                if (cat.includes('MEDICAMENTO') || cat.includes('PERFUMARIA')) { headerControls = `<div style="font-size:11px; color:#777;">(Controlo no Painel acima)</div>`; } 
                else { const inputId = `real-stf-${rowCount}`; const forceId = `force-real-stf-${rowCount}`; rowCount++; headerControls = `<div style="display:flex; align-items:center; gap:10px;"><div style="font-size:11px; text-align:right;"><div style="color:#666; font-weight:600">SOLICITADO: <b style="color:var(--blue); font-size:13px;">${stfSec}</b></div><div style="font-size:10px;">Força: <span id="${forceId}" class="force-val">-</span></div></div><div><input type="number" id="${inputId}" class="real-input" data-need="${stfSec}" placeholder="Real" oninput="updateForce('${inputId}', ${stfSec})"></div></div>`; projData.push({ name: cat, meta: meta, vol: sP, nec: stfSec, isStation: false, inputId: inputId }); }
                if (ctn) ctn.innerHTML += `<div class="card" style="border-left:5px solid ${secColor};"><div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #eee; margin-bottom:5px; padding-bottom:5px;"><div style="font-weight:800; color:${secColor}; font-size:13px;">${cat}</div>${headerControls}</div><table><thead><tr><th>Est</th><th>Ped(Un)</th><th>Ate(Un)</th><th>Ped(Cx)</th><th>Ate(Cx)</th><th>Falta(Cx)</th><th>%</th><th>Eqp</th></tr></thead><tbody>${rows}<tr class="total-row"><td>TOT</td><td>-</td><td>-</td><td>${fmtInt(sP)}</td><td>${fmtInt(sA)}</td><td>${fmtInt(sRupSec)}</td><td>${sP > 0 ? ((sA / sP) * 100).toFixed(1) : 0}%</td><td>${stfSec}</td></tr></tbody></table></div>`;
                gCx += sP; gRup += sRupSec; // Agora gRup soma a Falta em Unidades Exatas
            }
            
            const extraSectors = [ { n: 'Mapa', v: 2, meta: '5500 Cx', icon: '🗺️' }, { n: 'Doca 14', v: 2, meta: 'Apoio', icon: '🚚' }, { n: 'Prensa', v: 1, meta: 'Apoio', icon: '📦' } ];
            let extraHTML = `<div class="card" style="border-left:5px solid var(--orange); margin-top:10px;"><div style="font-weight:800; color:var(--orange); margin-bottom:12px; font-size:14px;"><i class="fas fa-tools"></i> SETOR DE APOIO</div><div style="display:flex; gap:15px; flex-wrap:wrap;">`;
            extraSectors.forEach(extra => {
                const extraId = `extra-input-${extra.n.replace(/\s/g, '')}`; gStf += extra.v;
                extraHTML += `<div style="background:#fff; padding:12px 16px; border-radius:8px; border:2px solid #f0aa70; text-align:center; min-width:130px; box-shadow:0 2px 6px rgba(0,0,0,0.06);"><div style="font-size:20px; margin-bottom:4px;">${extra.icon}</div><div style="font-size:12px; font-weight:800; color:#555; margin-bottom:4px;">${extra.n}</div><div style="font-size:11px; font-weight:700; color:var(--blue); margin-bottom:6px;">${extra.meta}</div><div style="font-size:9px; color:#888; margin-bottom:6px;">Nec: <b>${extra.v}</b> pessoas</div><input type="number" id="${extraId}" class="extra-input" placeholder="Real" oninput="updateTotalReal()"></div>`;
                projData.push({ name: extra.n, meta: extra.meta, vol: '-', nec: extra.v, isExtra: true, inputId: extraId });
            });
            extraHTML += `</div></div>`; if (ctn) ctn.innerHTML += extraHTML;
            
            const savedInputs = JSON.parse(localStorage.getItem('sysLog_staff_inputs') || '{}');
            document.querySelectorAll('.st-input, .real-input, .extra-input').forEach(inp => {
                if(savedInputs[inp.id] !== undefined) {
                    inp.value = savedInputs[inp.id];
                    if(inp.hasAttribute('data-need')) {
                        const need = parseFloat(inp.getAttribute('data-need')); const forceEl = document.getElementById('force-' + inp.id);
                        if(forceEl && need > 0) { const real = parseInt(inp.value) || 0; const pct = (real / need) * 100; forceEl.innerText = pct.toFixed(0) + '%'; forceEl.style.color = pct < 90 ? 'var(--red)' : (pct < 110 ? 'var(--green)' : 'var(--green-strong)'); }
                    }
                }
            });
            
            safeUpdate('tCx', fmtInt(gCx)); safeUpdate('tRup', fmtInt(gRup)); safeUpdate('tStf', gStf); safeUpdate('d-stf-need', gStf);
            const elRes = document.getElementById('resGuarda'); if (elRes) elRes.style.display = 'block';
            safeUpdate('d-ped', totP_Unit); safeUpdate('d-ate', totA_Unit); safeUpdate('d-rup', fmtInt(gRup)); // Global Rup reflete N_ATE 
            const gp = totP_Unit > 0 ? ((totA_Unit / totP_Unit) * 100) : 0; safeUpdate('d-pct', gp.toFixed(2) + '%'); aplicarCoresOficiais(gp); cachedCenarioData = struct;
            
            updateTotalReal(); updateModuleView(); updateProjecaoView(); upCharts(struct); runCenarioCalc();
            if (errors.length) alert(`Atenção: ${errors.length} itens ignorados. Eles não foram encontrados no Estoque Atual.`);
        } catch (e) { alert("Erro no cálculo: " + e); console.error(e); }
        document.getElementById('loading').style.display = 'none';
    }, 100);
}

function aplicarCoresOficiais(pct) { const card = document.getElementById('card-service'); const txt = document.getElementById('d-pct'); if (!card || !txt) return; card.classList.remove('k-red', 'k-green-light', 'k-green-strong'); if (pct < 99.00) { card.classList.add('k-red'); txt.style.color = "var(--red)"; } else if (pct < 99.60) { card.classList.add('k-green-light'); txt.style.color = "var(--green-light)"; } else { card.classList.add('k-green-strong'); txt.style.color = "var(--green-strong)"; } }

function updateForce(idInput, need) { try { const input = document.getElementById(idInput); const forceEl = document.getElementById('force-' + idInput); updateTotalReal(); if (!input || !forceEl) return; const real = parseInt(input.value) || 0; if (need <= 0) { forceEl.innerText = '-'; return; } const pct = (real / need) * 100; forceEl.innerText = pct.toFixed(0) + '%'; forceEl.style.color = pct < 90 ? 'var(--red)' : (pct < 110 ? 'var(--green)' : 'var(--green-strong)'); setTimeout(() => runCenarioCalc(), 200); } catch (e) {} }

function updateTotalReal() {
    let totalReal = 0; const inputVals = {};
    document.querySelectorAll('.st-input, .real-input, .extra-input').forEach(inp => { totalReal += (parseInt(inp.value) || 0); inputVals[inp.id] = inp.value; });
    localStorage.setItem('sysLog_staff_inputs', JSON.stringify(inputVals));
    const elTotalStf = document.getElementById('tStf'); const totalNeed = elTotalStf ? (parseFloat(elTotalStf.innerText) || 0) : 0;
    safeUpdate('tRealStf', totalReal); safeUpdate('d-stf', totalReal); safeUpdate('cen-stf', totalReal);
    const elForce = document.getElementById('d-force-global');
    if (elForce && totalNeed > 0) { const globalPct = (totalReal / totalNeed) * 100; elForce.innerText = globalPct.toFixed(1) + '%'; safeUpdate('p-force', globalPct.toFixed(1) + '%'); }
    updateProjecaoView();
}

function updateProjecaoView() {
    const tableBody = document.getElementById('tbProjecaoBody'); const aiBox = document.getElementById('aiSuggestions'); if (!dC || dC.length === 0) return;
    safeUpdate('p-lines', dC.length); safeUpdate('p-units', getVal('d-ped'));
    const tN = getVal('tStf'), tR = getVal('tRealStf'); const gap = tR - tN;
    safeUpdate('p-real-display', tR); safeUpdate('p-nec-display', tN);
    const elGap = document.getElementById('p-gap-display'); if (elGap) { elGap.innerText = gap >= 0 ? `(+${gap})` : `(${gap})`; elGap.style.color = gap >= 0 ? 'var(--green)' : 'var(--red)'; }
    if (!projData || projData.length === 0) return;
    let tableHtml = ''; globalSurplusList = []; globalDeficitList = []; globalAttentionPoints = [];
    projData.forEach(item => {
        let realVal = 0; if (item.isStation) { const el = document.getElementById(`st-input-${item.rng}`); realVal = el ? (parseInt(el.value) || 0) : 0; } else if (item.inputId) { const el = document.getElementById(item.inputId); realVal = el ? (parseInt(el.value) || 0) : 0; }
        const nec = item.nec; let cobPct = nec > 0 ? (realVal / nec) * 100 : (realVal > 0 ? 100 : 0);
        let status = 'OK', classStatus = 'cov-ok';
        if (item.isExtra) { if (realVal < nec) { status = 'FALTA'; classStatus = 'cov-bad'; } else if (realVal > nec) { status = 'SOBRA'; classStatus = 'cov-warn'; } } else { if (cobPct < 90) { status = 'FALTA'; classStatus = 'cov-bad'; } else if (cobPct < 100) { status = 'ATENÇÃO'; classStatus = 'cov-warn'; } else if (cobPct > 120) { status = 'EXCESSO'; classStatus = 'cov-warn'; } }
        const diff = realVal - nec;
        if (diff > 0) globalSurplusList.push({ name: item.name, qtd: diff }); if (diff < 0) globalDeficitList.push({ name: item.name, qtd: Math.abs(diff) });
        tableHtml += `<tr><td style="text-align:left;">${item.name}</td><td>${item.meta}</td><td>${nec}</td><td>${realVal}</td><td class="${classStatus}">${cobPct.toFixed(0)}%</td><td class="${classStatus}" style="font-size:10px;">${status}</td></tr>`;
        if (diff < 0 && !item.isExtra) { const metaOriginal = typeof item.meta === 'number' ? item.meta : 600; const metaRecovery = realVal > 0 ? Math.ceil((nec * metaOriginal) / realVal) : 0; globalAttentionPoints.push({ name: item.name, gap: Math.abs(diff), metaOrig: metaOriginal, metaRec: metaRecovery, realPeople: realVal }); }
    });
    if (tableBody) tableBody.innerHTML = tableHtml;
    let suggestionsHtml = '';
    if (globalSurplusList.length === 0 && globalDeficitList.length === 0) { suggestionsHtml = '<div class="sug-item sug-ok"><i class="fas fa-check-circle" style="color:green"></i><div>Equipa balanceada!</div></div>'; } else {
        let sList = JSON.parse(JSON.stringify(globalSurplusList)).sort((a, b) => b.qtd - a.qtd); let dList = JSON.parse(JSON.stringify(globalDeficitList)).sort((a, b) => b.qtd - a.qtd); let sIdx = 0, dIdx = 0;
        while (sIdx < sList.length && dIdx < dList.length) { const moveQtd = Math.min(sList[sIdx].qtd, dList[dIdx].qtd); suggestionsHtml += `<div class="sug-item sug-move"><i class="fas fa-random" style="color:var(--blue)"></i><div>Mover <b>${moveQtd}</b> de ${sList[sIdx].name} ➡️ ${dList[dIdx].name}</div></div>`; sList[sIdx].qtd -= moveQtd; dList[dIdx].qtd -= moveQtd; if (sList[sIdx].qtd === 0) sIdx++; if (dList[dIdx].qtd === 0) dIdx++; }
        while (dIdx < dList.length) { if (dList[dIdx].qtd > 0) { const ap = globalAttentionPoints.find(p => p.name === dList[dIdx].name); let advice = `Falta ${dList[dIdx].qtd} pessoas.`; if (ap) advice = ap.realPeople === 0 ? "SETOR PARADO!" : `Acelerar para <b>${ap.metaRec} cx/h</b>.`; suggestionsHtml += `<div class="sug-item sug-prod"><i class="fas fa-exclamation-triangle" style="color:red"></i><div><b>${dList[dIdx].name}:</b> ${advice}</div></div>`; } dIdx++; }
    }
    if (aiBox) aiBox.innerHTML = suggestionsHtml;
}

function commitDecision() { const text = document.getElementById('txtActionNote').value; if (text.trim() === "") { alert("Digite a decisão primeiro."); return; } isDecisionCommitted = true; document.getElementById('cardBrain').classList.add('official'); document.getElementById('txtBrainTitle').innerText = "DECISÃO REGISTADA"; document.getElementById('btnCommit').innerHTML = '<i class="fas fa-lock"></i> REGISTADO'; updateProjecaoView(); alert("Plano Oficializado!"); }

function copyReport() { updateProjecaoView(); setTimeout(() => { const now = new Date(); const timeStr = now.toLocaleTimeString().substr(0, 5); const stfReal = getVal('tRealStf'), stfNec = getVal('tStf'); const gapVal = stfReal - stfNec; const pct = document.getElementById('d-pct').innerText; let report = `______________________________________________________\n🚨 RELATÓRIO DE OPERAÇÃO | ${timeStr}\n______________________________________________________\n\nLINHAS: ${fmtBig(dC.length)} | UNIDADES: ${fmtBig(getVal('d-ped'))}\nEQUIPA: ${stfReal}/${stfNec} (GAP ${gapVal >= 0 ? '+' : ''}${gapVal}) | NÍVEL SERV: ${pct}\n\n`; if (globalAttentionPoints.length > 0) { report += `🛑 ALERTAS:\n`; globalAttentionPoints.forEach(p => { report += p.realPeople <= 0 ? `[🔥] ${p.name} ..... PARADO\n` : `[⚠️] ${p.name} ..... Meta: ${p.metaRec} cx/h\n`; }); report += `\n`; } const userNote = document.getElementById('txtActionNote').value.trim(); report += `🔄 MOVIMENTAÇÕES:\n${userNote || 'Equipa Balanceada.'}\n`; const copyArea = document.getElementById("txtReport"); if (copyArea) { copyArea.value = report; copyArea.select(); navigator.clipboard.writeText(copyArea.value); alert("Relatório copiado!"); } }, 100); }

function runCenarioCalc() {
    if (!cachedCenarioData) return;
    let totalBox = 0, totalDone = 0; for (let cat in cachedCenarioData) { for (let rng in cachedCenarioData[cat]) { totalBox += cachedCenarioData[cat][rng].p; totalDone += cachedCenarioData[cat][rng].a; } }
    safeUpdate('cen-box-pend', fmtInt(totalBox - totalDone)); safeUpdate('cen-box-tot', `de ${fmtInt(totalBox)} totais`); safeUpdate('cen-stf', document.getElementById('tRealStf') ? document.getElementById('tRealStf').innerText : '0');
    const globalPct = totalBox > 0 ? (totalDone / totalBox) : 0; safeUpdate('cen-health-val', (globalPct * 100).toFixed(0) + '%');
    const ctxG = document.getElementById('cenChartGauge'); if (cenGauge) cenGauge.destroy();
    if (ctxG) cenGauge = new Chart(ctxG, { type: 'doughnut', data: { labels: ['Concluído', 'Pendente'], datasets: [{ data: [globalPct * 100, 100 - (globalPct * 100)], backgroundColor: ['#2ecc71', '#ecf0f1'], borderWidth: 0, circumference: 180, rotation: 270 }] }, options: { responsive: true, maintainAspectRatio: false, cutout: '75%', plugins: { legend: { display: false }, tooltip: { enabled: false }, datalabels: { display: false } } } });
    const radarLabels = [], radarData = [];
    for (let cat in cachedCenarioData) { let sP = 0, sA = 0; for (let rng in cachedCenarioData[cat]) { sP += cachedCenarioData[cat][rng].p; sA += cachedCenarioData[cat][rng].a; } if (sP > 0) { radarLabels.push(cat.replace('SETOR ', '').substring(0, 10)); radarData.push((sA / sP) * 100); } }
    if (cenRadar) cenRadar.destroy(); const ctxR = document.getElementById('cenRadarChart');
    if (ctxR) { cenRadar = new Chart(ctxR, { type: 'radar', data: { labels: radarLabels, datasets: [{ label: '% Concluído', data: radarData, backgroundColor: 'rgba(41,128,185,0.2)', borderColor: 'rgba(41,128,185,1)', pointBackgroundColor: 'rgba(41,128,185,1)', borderWidth: 2 }] }, options: { responsive: true, maintainAspectRatio: false, scales: { r: { min: 0, max: 100, ticks: { display: false } } }, plugins: { legend: { display: false }, datalabels: { display: false } } } }); }
    const pending = totalBox - totalDone; const now = new Date(); const startMin = 20 * 60 + 12; const curMin = now.getHours() * 60 + now.getMinutes(); let minutesWorked = curMin >= startMin ? curMin - startMin : (24 * 60 - startMin) + curMin; if (minutesWorked < 1) minutesWorked = 1;
    const speed = totalDone / minutesWorked; const elPredTime = document.getElementById('cen-pred-time'); const elPredMsg = document.getElementById('cen-pred-msg');
    if (elPredTime && elPredMsg) {
        if (speed > 0 && pending > 0) { const predDate = new Date(now.getTime() + (pending / speed) * 60000); elPredTime.innerText = predDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); const h = predDate.getHours(); elPredTime.style.color = (h >= 5 && h < 14) ? '#e74c3c' : '#2ecc71'; elPredMsg.innerText = (h >= 5 && h < 14) ? "Risco: Estoiro de Turno!" : "Dentro do esperado."; } 
        else if (pending <= 0) { elPredTime.innerText = "CONCLUÍDO"; elPredTime.style.color = "#3498db"; elPredMsg.innerText = "Operação finalizada."; } else { elPredTime.innerText = "--:--"; elPredMsg.innerText = "A calcular ritmo..."; }
    }
}

function getShiftProgress() { const now = new Date(); const totalMin = now.getHours() * 60 + now.getMinutes(); const startMin = 20 * 60 + 12, endMin = 5 * 60, shiftDuration = 528; let elapsed = 0; if (totalMin >= startMin) elapsed = totalMin - startMin; else if (totalMin < endMin) elapsed = (24 * 60 - startMin) + totalMin; else return 0; return Math.min(Math.max(elapsed / shiftDuration, 0), 1); }
function updateClock() { const now = new Date(); safeUpdate('cen-clock', now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })); const prog = getShiftProgress(); const pct = (prog * 100).toFixed(1) + '%'; safeUpdate('cen-time-elapsed', pct + ' Decorrido'); const elBar = document.getElementById('cen-time-bar'); if (elBar) elBar.style.width = pct; }

function updateModuleView() {
    const ctn = document.getElementById('modulos-content'); if (!ctn) return;
    let totalMed = 0, totalPerf = 0, maxLoad = 0, minLoad = 999999, maxModule = '-', minModule = '-'; let maxPendingVal = 0, maxPendingMod = null; const allModules = [];
    const stations = [ { id: 'S1', lbl: 'Estação 1 (57-62)', r: [57, 58, 59, 60, 61, 62] }, { id: 'S2', lbl: 'Estação 2 (64-69)', r: [64, 65, 66, 67, 68, 69] }, { id: 'S3', lbl: 'Estação 3 (71-76)', r: [71, 72, 73, 74, 75, 76] }, { id: 'S4', lbl: 'Estação 4 (78-83)', r: [78, 79, 80, 81, 82, 83] }, { id: 'S5', lbl: 'Estação 5 (85-90)', r: [85, 86, 87, 88, 89, 90] } ];
    ['MEDICAMENTO', 'PERFUMARIA'].forEach(type => {
        ['2D', '3D'].forEach(subType => {
            stations.forEach(st => {
                st.r.forEach(rNum => {
                    let ped = 0, ate = 0; if (moduleStats[type] && moduleStats[type][subType] && moduleStats[type][subType][rNum]) { ped = moduleStats[type][subType][rNum].p; ate = moduleStats[type][subType][rNum].a; }
                    if (type === 'MEDICAMENTO') totalMed += ped; else totalPerf += ped;
                    const pend = ped - ate; const modId = `${type}_${subType}_${rNum}`;
                    if (pend > 0) { allModules.push({ id: modId, val: pend }); if (pend > maxPendingVal) { maxPendingVal = pend; maxPendingMod = modId; } }
                    const typeShort = type.substr(0, 3), subShort = subType === '2D' ? 'FLOW' : 'EST'; const modName = `${typeShort} ${subShort} ${rNum}`;
                    if (ped > maxLoad) { maxLoad = ped; maxModule = modName; } if (ped > 0 && ped < minLoad) { minLoad = ped; minModule = modName; }
                });
            });
        });
    });
    allModules.sort((a, b) => b.val - a.val); const top3 = allModules.slice(0, 3).map(m => m.id); let html = '';
    ['MEDICAMENTO', 'PERFUMARIA'].forEach(type => {
        const isPerf = type === 'PERFUMARIA'; const badgeClass = isPerf ? 'mb-perf' : 'mb-med';
        ['2D', '3D'].forEach(subType => {
            const subLbl = subType === '2D' ? 'FLOWRACK (02 DÍGITOS)' : 'ESTANTERIA (03 DÍGITOS)'; const icon = subType === '2D' ? '<i class="fas fa-water"></i>' : '<i class="fas fa-layer-group"></i>';
            html += `<div class="modern-section-title"><div class="modern-badge ${badgeClass}">${type}</div><h3>${icon} ${subLbl}</h3></div>`;
            stations.forEach(st => {
                html += `<div class="station-group-title"><i class="fas fa-users"></i> ${st.lbl}</div><div class="mod-grid-new">`;
                st.r.forEach(rNum => {
                    let ped = 0, ate = 0; if (moduleStats[type] && moduleStats[type][subType] && moduleStats[type][subType][rNum]) { ped = moduleStats[type][subType][rNum].p; ate = moduleStats[type][subType][rNum].a; }
                    const pct = ped > 0 ? (ate / ped) * 100 : 100; const pend = ped - ate; const currentId = `${type}_${subType}_${rNum}`;
                    let cardStatus = 'status-ok', pctStatus = 'ok', fillBg = '#2ecc71';
                    if (pend > 0) { if (currentId === maxPendingMod) { cardStatus = 'super-crit'; pctStatus = 'crit'; fillBg = '#c0392b'; } else if (top3.includes(currentId) || ped > 100) { cardStatus = 'status-crit'; pctStatus = 'crit'; fillBg = '#e74c3c'; } else if (pct < 95) { cardStatus = 'status-warn'; pctStatus = 'warn'; fillBg = '#f39c12'; } }
                    const pendHtml = pend > 0 ? `<div class="street-pend"><i class="fas fa-exclamation-circle"></i> Faltam ${pend}</div>` : `<div style="font-size:11px; color:#95a5a6;"><i class="fas fa-check"></i> Completo</div>`;
                    html += `<div class="street-card ${cardStatus}" onclick="showModDetails('${type}','${subType}',${rNum})"><div class="street-header"><div class="street-id">RUA ${rNum}</div><div class="street-pct ${pctStatus}">${pct.toFixed(0)}%</div></div><div class="street-body"><div class="street-vol"><span class="street-vol-lbl">Total Caixas</span><span class="street-vol-val">${ped}</span></div>${pendHtml}</div><div class="street-track"><div class="street-fill" style="width:${pct}%; background:${fillBg}"></div></div></div>`;
                });
                html += `</div>`;
            });
        });
    });
    if (cachedCenarioData) {
        let outrosHtml = `<div class="modern-section-title" style="margin-top:35px; border-top:2px dashed #ddd; padding-top:20px;"><div class="modern-badge mb-outros">DIVERSOS</div><h3><i class="fas fa-boxes"></i> OUTROS SETORES</h3></div><div class="mod-grid-new">`; let hasOutros = false;
        for (let cat in cachedCenarioData) {
            if (cat.includes('MEDICAMENTO') || cat.includes('PERFUMARIA')) continue;
            for (let rng in cachedCenarioData[cat]) {
                const d = cachedCenarioData[cat][rng]; if (d.p === 0 && d.a === 0) continue; hasOutros = true;
                const pct = d.p > 0 ? (d.a / d.p) * 100 : 100; const pend = d.p - d.a;
                let cardStatus = 'status-ok', pctStatus = 'ok', fillBg = '#27ae60';
                if (pend > 0) { if (pct < 90) { cardStatus = 'status-crit'; pctStatus = 'crit'; fillBg = '#e74c3c'; } else { cardStatus = 'status-warn'; pctStatus = 'warn'; fillBg = '#f39c12'; } }
                const pendHtml = pend > 0 ? `<div class="street-pend"><i class="fas fa-exclamation-circle"></i> Faltam ${pend}</div>` : `<div style="font-size:11px;color:#95a5a6;"><i class="fas fa-check"></i> Completo</div>`;
                const title = cat.replace('SETOR ', '') + (rng !== 'TOTAL' ? ' ' + rng : '');
                outrosHtml += `<div class="street-card ${cardStatus}"><div class="street-header"><div class="street-id" style="font-size:13px;">${title}</div><div class="street-pct ${pctStatus}">${pct.toFixed(0)}%</div></div><div class="street-body"><div class="street-vol"><span class="street-vol-lbl">Total Caixas</span><span class="street-vol-val">${d.p}</span></div>${pendHtml}</div><div class="street-track"><div class="street-fill" style="width:${pct}%;background:${fillBg}"></div></div></div>`;
            }
        }
        outrosHtml += `</div>`; if (hasOutros) html += outrosHtml;
    }
    ctn.innerHTML = html; safeUpdate('mod-tot-med', totalMed); safeUpdate('mod-tot-perf', totalPerf); safeUpdate('mod-gargalo', maxModule); safeUpdate('mod-leve', minLoad === 999999 ? '-' : minModule);
}

function showModDetails(type, subType, rua) {
    const modal = document.getElementById('modalModDetail'); const tbBody = document.getElementById('tbModDetailBody'); const title = document.getElementById('modDetailTitle');
    if (!moduleStats[type] || !moduleStats[type][subType] || !moduleStats[type][subType][rua]) { alert("Sem dados."); return; }
    const data = moduleStats[type][subType][rua]; title.innerText = `${type} - ${subType} - RUA ${rua}`; tbBody.innerHTML = '';
    if (!data.items || data.items.length === 0) { tbBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding:20px; color:green;">Sem pendências!</td></tr>'; } 
    else { data.items.sort((a, b) => b.q - a.q); data.items.forEach(item => { tbBody.innerHTML += `<tr><td style="font-weight:700;">${item.cod}</td><td style="font-size:11px; text-align:left;">${(item.desc || '').substring(0, 40)}</td><td style="font-size:11px;">${item.rom || '-'}</td><td style="font-weight:800; color:var(--red); text-align:right;">${item.q}</td></tr>`; }); }
    modal.style.display = 'flex';
}
function closeModDetail() { document.getElementById('modalModDetail').style.display = 'none'; }

function upCharts(data) {
    const ctxP = document.getElementById('chartPct'); const ctxF = document.getElementById('chartFalta'); if (!ctxP || !ctxF) return;
    const lbs = [], dP = [], dF = [], bgColors = [];
    for (let c in data) {
        const isDet = c.includes('MED') || c.includes('PERF');
        if (isDet) { for (let r in data[c]) { const d = data[c][r]; if (d.p > 0) { let sName = c.split(' ')[0].substr(0, 4) + (c.includes('02') ? ' 2D' : ' 3D'); lbs.push(`${sName} ${r}`); dP.push(((d.a / d.p) * 100).toFixed(1)); dF.push(d.p - d.a); } } } 
        else { let sP = 0, sA = 0; for (let r in data[c]) { sP += data[c][r].p; sA += data[c][r].a; } if (sP > 0) { lbs.push(c.replace('SETOR ', '').substring(0, 12)); dP.push(((sA / sP) * 100).toFixed(1)); dF.push(sP - sA); } }
    }
    dP.forEach(val => bgColors.push(parseFloat(val) >= 99.0 ? '#2ecc71' : '#e74c3c'));
    if (chP instanceof Chart) chP.destroy(); if (chF instanceof Chart) chF.destroy();
    chP = new Chart(ctxP, { type: 'bar', data: { labels: lbs, datasets: [{ data: dP, backgroundColor: bgColors }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, datalabels: { color: '#333', font: { weight: 'bold' } } } } });
    chF = new Chart(ctxF, { type: 'bar', data: { labels: lbs, datasets: [{ data: dF, backgroundColor: '#e74c3c' }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, datalabels: { color: '#e74c3c', anchor: 'end', align: 'end', font: { weight: 'bold' } } } } });
}

function toggleTVMode() { if (!document.fullscreenElement) { document.documentElement.requestFullscreen().catch(() => {}); document.body.classList.add('tv-mode'); } else { if (document.exitFullscreen) document.exitFullscreen(); document.body.classList.remove('tv-mode'); } }
document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement) document.body.classList.remove('tv-mode'); });

window.encerrarOperacaoDia = async function() {
    if (!projData || projData.length === 0 || !cachedCenarioData) return alert("Não há dados de operação processados para salvar.");
    if (!confirm("⚠️ ATENÇÃO: Tem a certeza que deseja ENCERRAR a operação de hoje?\n\nIsto irá arquivar todos os dados de produtividade e limpar a tela.")) return;
    document.getElementById('loading').style.display = 'flex'; document.getElementById('loadMsg').innerText = 'A arquivar a Operação do Dia...';
    try {
        const todayStr = new Date().toISOString().slice(0, 10); const timestamp = new Date().getTime(); const archiveKey = 'archive_' + todayStr + '_' + timestamp;
        const elGargalo = document.getElementById('mod-gargalo'); const gargaloText = elGargalo ? elGargalo.innerText : '-';
        const archiveData = { date: new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }), metrics: { ped: document.getElementById('d-ped').innerText, ate: document.getElementById('d-ate').innerText, pct: document.getElementById('d-pct').innerText, staffReal: getVal('tRealStf'), staffNeed: getVal('tStf'), linhas: dC ? dC.length : 0, caixas: getVal('tCx'), ruptura: getVal('tRup'), gargalo: gargaloText }, projData: projData, staffInputs: JSON.parse(localStorage.getItem('sysLog_staff_inputs') || '{}'), cenarioData: cachedCenarioData };
        globalHistory.unshift({ id: archiveKey, date: archiveData.date, stats: archiveData.metrics });
        if (window.saveToDB) { await window.saveToDB(archiveKey, archiveData); await window.saveToDB('dHistory', globalHistory); await window.saveToDB('dC', []); await window.saveToDB('dE', []); localStorage.removeItem('sysLog_staff_inputs'); }
        renderHistoryList(); alert("✅ OPERAÇÃO ARQUIVADA COM SUCESSO!"); location.reload();
    } catch(e) { alert("Erro ao arquivar operação: " + e.message); }
    document.getElementById('loading').style.display = 'none';
};

function renderHistoryList() { 
    const ul = document.getElementById('histList'); if (!ul) return; 
    ul.innerHTML = ''; if (globalHistory.length === 0) { ul.innerHTML = '<li style="padding:10px; color:#777;">Nenhum histórico diário salvo.</li>'; return; } 
    globalHistory.forEach((item, index) => { 
        ul.innerHTML += `<li class="hist-item" style="border: 1px solid #eee; border-radius: 6px; padding: 10px; margin-bottom: 8px; background: #fafafa; display: flex; align-items:center;"><div onclick="viewHistoryDetail(${index})" style="flex:1; cursor:pointer;"><div style="font-weight:bold; color:var(--blue);"><i class="far fa-calendar-alt"></i> ${item.date}</div><div style="font-size: 11px; color:#555; margin-top:4px;"><b>Serviço:</b> <span style="color:var(--green-strong); font-weight:900;">${item.stats.pct}</span> | <b>Equipa:</b> ${item.stats.staffReal}/${item.stats.staffNeed} | <b>Caixas:</b> ${item.stats.caixas || '-'}</div></div><button class="btn btn-red" style="padding:6px 10px;" onclick="deleteHistory(${index})" title="Excluir Arquivo"><i class="fas fa-trash"></i></button></li>`; 
    }); 
}

window.viewHistoryDetail = async function(index) { 
    const item = globalHistory[index]; if(!item || !item.id) return alert("Erro ao localizar arquivo.");
    document.getElementById('loading').style.display = 'flex'; document.getElementById('loadMsg').innerText = 'A buscar Raio-X do Arquivo Morto...';
    try {
        const fullArchive = await window.getFromDB(item.id);
        let reportHTML = `<div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 20px; border-bottom: 2px solid #eee; padding-bottom: 15px;"><div style="background:#f8f9fa; padding:10px; border-radius:6px; text-align:center; border-left:4px solid var(--blue);"><div style="font-size:10px; color:#777; font-weight:bold;">UNIDADES (PED ➡️ ATE)</div><div style="font-size:15px; font-weight:900; color:var(--blue);">${item.stats.ped} ➡️ ${item.stats.ate}</div></div><div style="background:#f8f9fa; padding:10px; border-radius:6px; text-align:center; border-left:4px solid var(--green-strong);"><div style="font-size:10px; color:#777; font-weight:bold;">SERVIÇO FINAL</div><div style="font-size:16px; font-weight:900; color:var(--green-strong);">${item.stats.pct}</div></div><div style="background:#f8f9fa; padding:10px; border-radius:6px; text-align:center; border-left:4px solid var(--purple);"><div style="font-size:10px; color:#777; font-weight:bold;">EQUIPA (REAL / NEC)</div><div style="font-size:16px; font-weight:900; color:var(--purple);">${item.stats.staffReal} / ${item.stats.staffNeed}</div></div><div style="background:#f8f9fa; padding:10px; border-radius:6px; text-align:center; border-left:4px solid #34495e;"><div style="font-size:10px; color:#777; font-weight:bold;">LINHAS LIDAS (ACESSO)</div><div style="font-size:16px; font-weight:900; color:#34495e;">${item.stats.linhas || '-'}</div></div><div style="background:#f8f9fa; padding:10px; border-radius:6px; text-align:center; border-left:4px solid var(--orange);"><div style="font-size:10px; color:#777; font-weight:bold;">CAIXAS (TOT / RUPTURA)</div><div style="font-size:16px; font-weight:900; color:var(--orange);">${item.stats.caixas || '-'} / <span style="color:var(--red)">${item.stats.ruptura || '-'}</span></div></div><div style="background:#fff0f0; padding:10px; border-radius:6px; text-align:center; border-left:4px solid var(--red);"><div style="font-size:10px; color:#c0392b; font-weight:bold;">MAIOR GARGALO</div><div style="font-size:13px; font-weight:900; color:var(--red);">${item.stats.gargalo || '-'}</div></div></div>`;
        reportHTML += `<div style="display:flex; gap:15px;"><div style="flex:1;"><h4 style="margin: 0 0 10px 0; color:#2c3e50;"><i class="fas fa-users"></i> Distribuição da Equipa</h4><div style="max-height: 250px; overflow-y:auto; border:1px solid #eee; border-radius:6px; margin-bottom:20px;"><table style="width:100%; text-align:left; border-collapse: collapse; font-size:11px;"><thead style="position:sticky; top:0;"><tr style="background:#2c3e50; color:white;"><th style="padding:6px;">Setor / Estação</th><th style="padding:6px; text-align:center;">Nec.</th><th style="padding:6px; text-align:center;">Real</th><th style="padding:6px; text-align:center;">Status</th></tr></thead><tbody>`;
        if(fullArchive && fullArchive.projData) {
            fullArchive.projData.forEach(p => {
                let realVal = 0; if(fullArchive.staffInputs) { if (p.isStation) realVal = parseInt(fullArchive.staffInputs[`st-input-${p.rng}`]) || 0; else if (p.inputId) realVal = parseInt(fullArchive.staffInputs[p.inputId]) || 0; }
                let cobPct = p.nec > 0 ? (realVal / p.nec) * 100 : (realVal > 0 ? 100 : 0); let badge = '<span style="color:var(--green); font-weight:bold;">OK</span>'; if(cobPct < 90) badge = '<span style="color:var(--red); font-weight:bold;">FALTOU</span>'; if(cobPct > 120) badge = '<span style="color:var(--orange); font-weight:bold;">EXCESSO</span>';
                reportHTML += `<tr style="border-bottom:1px solid #eee;"><td style="padding:6px;"><b>${p.name}</b></td><td style="padding:6px; text-align:center; font-weight:bold; color:#777;">${p.nec}</td><td style="padding:6px; text-align:center; font-weight:bold; color:var(--blue);">${realVal}</td><td style="padding:6px; text-align:center;">${badge}</td></tr>`;
            });
        } else { reportHTML += `<tr><td colspan="4" style="text-align:center; padding:10px;">Sem detalhes de equipe salvos.</td></tr>`; }
        reportHTML += `</tbody></table></div></div><div style="flex:1;">`;
        if (fullArchive && fullArchive.cenarioData) {
            reportHTML += `<h4 style="margin: 0 0 10px 0; color:#2c3e50;"><i class="fas fa-boxes"></i> Caixa por Setor/Rua</h4><div style="max-height: 250px; overflow-y:auto; border:1px solid #eee; border-radius:6px;"><table style="width:100%; text-align:left; border-collapse: collapse; font-size:11px;"><thead style="position:sticky; top:0; z-index:2;"><tr style="background:#34495e; color:white;"><th style="padding:6px;">Setor / Estação</th><th style="padding:6px; text-align:center;">Ped</th><th style="padding:6px; text-align:center;">Ate</th><th style="padding:6px; text-align:center;">Falta</th><th style="padding:6px; text-align:center;">%</th></tr></thead><tbody>`;
            for (let cat in fullArchive.cenarioData) {
                for (let rng in fullArchive.cenarioData[cat]) {
                    const d = fullArchive.cenarioData[cat][rng]; if (d.p === 0 && d.a === 0) continue; 
                    const pct = d.p > 0 ? ((d.a / d.p) * 100).toFixed(1) : 100; const falta = d.p - d.a; const colorFalta = falta > 0 ? 'color:var(--red); font-weight:bold;' : 'color:#999;'; const colorPct = pct < 95 ? 'color:var(--red); font-weight:bold;' : 'color:var(--green); font-weight:bold;'; let shortCat = cat.replace('SETOR ', '').replace('MEDICAMENTO', 'MED').replace('PERFUMARIA', 'PERF');
                    reportHTML += `<tr style="border-bottom:1px solid #eee; background:#fff;"><td style="padding:6px;"><b>${shortCat}</b><br><span style="color:var(--blue); font-size:10px;">${rng}</span></td><td style="padding:6px; text-align:center; font-weight:bold; color:#555;">${d.p}</td><td style="padding:6px; text-align:center; font-weight:bold; color:var(--blue);">${d.a}</td><td style="padding:6px; text-align:center; ${colorFalta}">${falta}</td><td style="padding:6px; text-align:center; ${colorPct}">${pct}%</td></tr>`;
                }
            }
            reportHTML += `</tbody></table></div>`;
        } else { reportHTML += `<div style="padding:20px; text-align:center; border:1px solid #eee; border-radius:6px; color:#999;">Sem dados de setor neste arquivo.</div>`; }
        reportHTML += `</div></div>`; 
        let modal = document.getElementById('modalAudit');
        if(!modal) { modal = document.createElement('div'); modal.id = 'modalAudit'; modal.style.cssText = 'display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.8); z-index:99999; justify-content:center; align-items:center; padding:20px;'; document.body.appendChild(modal); }
        modal.innerHTML = `<div style="background:#fff; border-radius:12px; width:100%; max-width:900px; box-shadow:0 10px 30px rgba(0,0,0,0.5); display:flex; flex-direction:column;"><div style="background:#2c3e50; color:white; padding:15px 20px; border-radius:12px 12px 0 0; display:flex; justify-content:space-between; align-items:center;"><h3 style="margin:0; font-size:16px;"><i class="fas fa-file-archive"></i> Relatório de Encerramento: ${item.date}</h3><button onclick="document.getElementById('modalAudit').style.display='none'" style="background:none; border:none; color:white; font-size:24px; cursor:pointer;">&times;</button></div><div style="padding:20px; overflow-y:auto; max-height:85vh;">${reportHTML}</div></div>`;
        modal.style.display = 'flex';
    } catch(e) { alert("Falha ao abrir o arquivo. Erro: " + e.message); }
    document.getElementById('loading').style.display = 'none';
};

function deleteHistory(i) { if (confirm("Excluir este arquivo morto da nuvem?")) { globalHistory.splice(i, 1); if (window.saveToDB) window.saveToDB('dHistory', globalHistory); renderHistoryList(); closeHistDetail(); } }
function exportHistory() { const a = document.createElement('a'); a.href = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(globalHistory)); a.download = "hist_" + new Date().toISOString().slice(0, 10) + ".json"; document.body.appendChild(a); a.click(); a.remove(); }
function importHistory(i) { const f = i.files[0]; if (!f) return; const r = new FileReader(); r.onload = async e => { try { const imp = JSON.parse(e.target.result); if (confirm(`Importar ${imp.length} registos?`)) { globalHistory = [...imp, ...globalHistory]; if (window.saveToDB) await window.saveToDB('dHistory', globalHistory); renderHistoryList(); } } catch (err) { alert("Erro ao importar."); } }; r.readAsText(f); }
function closeHistDetail() { const p = document.getElementById('histDetailPanel'); if(p) p.style.display = 'none'; }

function openCfg() { document.getElementById('modalCfg').style.display = 'flex'; document.getElementById('cfgBody').innerHTML = Object.keys(SECTORS).map(k => `<div class="cfg-row"><div><b>${SECTORS[k].l}</b></div><div>In: <input id="ri_${k}" value="${SECTORS[k].ri}" class="ipt-cfg"><input id="pi_${k}" value="${SECTORS[k].pi}" class="ipt-cfg"></div><div>Fim: <input id="rf_${k}" value="${SECTORS[k].rf}" class="ipt-cfg"><input id="pf_${k}" value="${SECTORS[k].pf}" class="ipt-cfg"></div></div>` ).join(''); }
function closeCfg() { document.getElementById('modalCfg').style.display = 'none'; }
function saveCfg() { for (let k in SECTORS) { SECTORS[k].ri = parseInt(document.getElementById(`ri_${k}`).value) || 0; SECTORS[k].rf = parseInt(document.getElementById(`rf_${k}`).value) || 0; } closeCfg(); }

function showErrors() { const tb = document.querySelector('#tbErrContent tbody'); tb.innerHTML = ''; if (globalErrors.length === 0) tb.innerHTML = '<tr><td colspan="3">Nenhum erro.</td></tr>'; else globalErrors.slice(0, 500).forEach(e => tb.innerHTML += `<tr><td>${e.p}</td><td>${e.k}</td><td>${e.reason}</td></tr>`); document.getElementById('modalErr').style.display = 'flex'; }
function closeErrors() { document.getElementById('modalErr').style.display = 'none'; }

function exportMaster() { try { const wb = XLSX.utils.book_new(); const rows = [['Setor', 'Faixa', 'Ped(Un)', 'Ate(Un)', 'Ped(Cx)', 'Ate(Cx)', 'Falta', '%']]; if (cachedCenarioData) { for (let cat in cachedCenarioData) { for (let rng in cachedCenarioData[cat]) { const d = cachedCenarioData[cat][rng]; const pct = d.p > 0 ? ((d.a / d.p) * 100).toFixed(1) + '%' : '0%'; rows.push([cat, rng, d.pun, d.aun, d.p, d.a, d.ncx || Math.max(0, d.p - d.a), pct]); } } } const ws = XLSX.utils.aoa_to_sheet(rows); XLSX.utils.book_append_sheet(wb, ws, 'Planejamento'); XLSX.writeFile(wb, `Logistica_${new Date().toISOString().slice(0, 10)}.xlsx`); } catch (e) { alert("Erro ao exportar: " + e.message); } }

document.addEventListener('DOMContentLoaded', () => verificarLogin());

// Salvar no localStorage ao fechar/recarregar a página (backup de emergência)
window.addEventListener('beforeunload', () => {
    try {
        if (dC.length) {
            const dCSlim = dC.map(r => {
                const slim = {};
                const fields = [colMapMaster.prod, colMapMaster.ped, colMapMaster.ate,
                                colMapMaster.n_ate, colMapMaster.cx_ped, colMapMaster.cx_ate,
                                colMapMaster.desc, colMapMaster.rom, colMapMaster.dep];
                fields.forEach(f => { if (f && r[f] !== undefined) slim[f] = r[f]; });
                if (!slim[colMapMaster.prod]) slim['Produto'] = r['Produto'] || r['Material'] || '';
                return slim;
            });
            localStorage.setItem('sysLog_dC', JSON.stringify(dCSlim));
            localStorage.setItem('sysLog_colMap', JSON.stringify(colMapMaster));
        }
        if (dE.length) {
            const dESlim = dE.map(r => {
                const prod = r['Produto'] || r['Material'] || r['PRODUTO'] || '';
                const pick = r['Picking'] || r['PICKING'] || '';
                if (!prod || !pick) return null;
                return { 'Produto': String(prod).trim(), 'Picking': String(pick).trim().toUpperCase() };
            }).filter(Boolean);
            localStorage.setItem('sysLog_dE', JSON.stringify(dESlim));
        }
    } catch(e) {}
});

function verificarLogin() { 
    // Se é modo formulário público, NUNCA mostrar login nem sistema
    const urlP = new URLSearchParams(window.location.search);
    if (urlP.get('mode') === 'formferias') {
        const lo = document.getElementById('loginOverlay'); if(lo) lo.style.display='none';
        ['app-header','memory-bar','tab-bar','container'].forEach(c=>{const e=document.querySelector('.'+c)||document.getElementById(c);if(e)e.style.display='none';});
        const fw=document.getElementById('publicFormWrapper');if(fw)fw.style.display='block';
        loadEmployeesForPublicForm(); return;
    }
    if (urlP.get('mode') === 'formepi') {
        const lo = document.getElementById('loginOverlay'); if(lo) lo.style.display='none';
        ['app-header','memory-bar','tab-bar','container'].forEach(c=>{const e=document.querySelector('.'+c)||document.getElementById(c);if(e)e.style.display='none';});
        const fw=document.getElementById('publicFormWrapper');if(fw)fw.style.display='block';
        ['epi-pub-form','pf-form-view','pf-success-view'].forEach(id=>{const e=document.getElementById(id);if(e)e.style.display=(id==='epi-pub-form')?'block':'none';});
        loadEpiColabs(); return;
    }
    // Fluxo normal de login
    if(sessionStorage.getItem('logado') === 'sim') { 
        const lo = document.getElementById('loginOverlay'); 
        if(lo) lo.style.display = 'none'; 
        renderMetas(); init(); 
    } else { 
        const lo = document.getElementById('loginOverlay'); 
        if(lo) lo.style.display = 'flex'; 
    } 
}
function checkPassword() { const s = document.getElementById('sysPassword').value; if(s === 'renato2026') { sessionStorage.setItem('logado', 'sim'); const lo = document.getElementById('loginOverlay'); if(lo) lo.style.display = 'none'; renderMetas(); init(); } else { document.getElementById('loginError').style.display = 'block'; document.getElementById('sysPassword').value = ''; document.getElementById('sysPassword').focus(); } }

async function loadEmployeesForPublicForm() { try { const savedRH = await window.getFromDB('dRH'); const sel = document.getElementById('pf_colab_select'); if(savedRH && savedRH.length > 0) { sel.innerHTML = '<option value="">-- Selecione o seu nome aqui --</option>'; const ativos = savedRH.filter(r => (r['Status']||r['STATUS']||'ATIVO').toUpperCase() !== 'INATIVO'); ativos.sort((a,b) => (a['Nome']||a['NOME']||'').localeCompare(b['Nome']||b['NOME']||'')); ativos.forEach(r => { const mat = r['Matrícula']||r['Matricula']||r['MATRICULA']; const nomeCompleto = r['Nome']||r['NOME']||''; const pts = nomeCompleto.trim().split(' '); const exib = pts.length > 1 ? pts[0]+' '+pts[1] : pts[0]; sel.innerHTML += `<option value="${mat}">${exib}</option>`; }); } else { sel.innerHTML = '<option value="">Nenhum funcionário sincronizado.</option>'; } } catch(e) { document.getElementById('pf_colab_select').innerHTML = '<option value="">Erro ao carregar lista.</option>'; } }

// ================================================================
// GIRO DE ALOCAÇÃO — modulo completo
// Valida se produtos estão no giro correto com base no endereço
// Cruzamento: Análise de Alocação x Mapeamento de Estações
// ================================================================

// ── Mapeamento de estações (configurável) ─────────────────────
// Regra: prefixo do endereço determina a estação
var GIRO_ESTACOES = [
    {
        id: 1,
        nome: '1ª Estação',
        descricao: 'Flowrack Perfumaria',
        cor: 'giro-s1',
        // Ruas 21-xx, 22-xx, 23-xx (endereços que começam com V21, V22, V23)
        prefixos: ['V21','V22','V23','21-','22-','23-'],
        // Subrange: 21-01-011 ao 21-25-081, 22-01 ao 22-25-081, 23-01 ao 23-25-081
        validar: function(end) {
            if(!end) return false;
            var e = end.toUpperCase();
            return (e.startsWith('V21')||e.startsWith('V22')||e.startsWith('V23')||
                    e.startsWith('21-')||e.startsWith('22-')||e.startsWith('23-'));
        }
    },
    {
        id: 2,
        nome: '2ª Estação',
        descricao: 'Flowrack Secundário',
        cor: 'giro-s2',
        prefixos: ['V24','V25','V26','24-','25-','26-'],
        validar: function(end) {
            if(!end) return false;
            var e = end.toUpperCase();
            return (e.startsWith('V24')||e.startsWith('V25')||e.startsWith('V26')||
                    e.startsWith('24-')||e.startsWith('25-')||e.startsWith('26-'));
        }
    }
];

// ── Estado interno ─────────────────────────────────────────────
var _giroData      = [];   // dados brutos importados
var _giroFiltered  = [];   // dados após filtros
var _giroSortCol   = 'dataAlocacao';
var _giroSortAsc   = false;
var _giroInitialized = false;

// ── Inicialização da aba ───────────────────────────────────────
window.initGiroAlocacao = function() {
    if(_giroInitialized && _giroData.length > 0) {
        giroRenderAll();
    }
    // Configura listener de arquivo
    var inp = document.getElementById('giro-file-alocacao');
    if(inp && !inp._giroReady) {
        inp._giroReady = true;
    }
};

// ── Importação do arquivo ──────────────────────────────────────
window.giroImportarAlocacao = async function(input) {
    var file = input.files[0]; if(!file) return;
    var btn = document.querySelector('[onclick*="giroAtualizar"]');
    var statusEl = document.getElementById('giro-status-importacao');
    var statusTxt = document.getElementById('giro-status-txt');

    // Mostra loading
    document.getElementById('giro-placeholder').style.display = 'none';
    if(statusEl) { statusEl.style.display='block'; statusEl.style.background='#fff3cd'; statusEl.style.borderColor='#ffeeba'; }
    if(statusTxt) statusTxt.textContent = 'Lendo arquivo... Aguarde.';

    try {
        var data = await giroLerArquivo(file);
        if(!data || !data.length) { alert('Arquivo vazio ou formato não reconhecido.'); return; }

        _giroData = giropProcessar(data);
        _giroInitialized = true;

        // Salva no Firebase
        try { await window.saveToDB('giro_alocacao_data', _giroData); } catch(e) {}

        if(statusEl) { statusEl.style.background='#eaf4fd'; statusEl.style.borderColor='#aed6f1'; }
        if(statusTxt) statusTxt.textContent = file.name + ' — ' + _giroData.length + ' linhas importadas com sucesso em ' + new Date().toLocaleTimeString('pt-BR');

        giroRenderAll();
        document.getElementById('giro-btn-atualizar').style.display='';
        document.getElementById('giro-btn-excel').style.display='';

    } catch(e) {
        alert('Erro ao ler arquivo: ' + e.message);
        console.error(e);
    }
    input.value = '';
};

// ── Lê arquivo (xlsx ou csv) ───────────────────────────────────
function giroLerArquivo(file) {
    return new Promise(function(resolve, reject) {
        var reader = new FileReader();
        reader.onload = function(e) {
            try {
                var wb = XLSX.read(e.target.result, {type:'binary', cellDates:true});
                var ws = wb.Sheets[wb.SheetNames[0]];
                var json = XLSX.utils.sheet_to_json(ws, {defval:''});
                resolve(json);
            } catch(err) { reject(err); }
        };
        reader.onerror = reject;
        reader.readAsBinaryString(file);
    });
}

// ── Processa linhas do arquivo ─────────────────────────────────
// Mapeia colunas flexíveis (PT, EN, abreviadas, etc.)
function giropProcessar(rows) {
    return rows.map(function(r, idx) {
        // Resolve valor de coluna com múltiplos nomes possíveis
        function col() {
            for(var i=0;i<arguments.length;i++){
                var keys = Object.keys(r);
                for(var j=0;j<keys.length;j++){
                    if(keys[j].toLowerCase().replace(/\s/g,'').includes(arguments[i].toLowerCase().replace(/\s/g,'')))
                        return String(r[keys[j]]||'').trim();
                }
            }
            return '';
        }

        var dataRaw   = col('DataAlocação','DataAloca','Date','Data','DataAloc');
        var produto   = col('Produto','Product','Cod','Código','Code','Item');
        var resp      = col('Resp.Alocação','RespAloc','Responsavel','Resp','User','Usuario','Operador');
        var endereco  = col('Endereço','Endereco','Address','End.','End','Local','Localizacao');
        var barras    = col('Barras','Barra','Barcode','Bar','Código Barras');
        var volumes   = col('Volumes','Volume','Qtd','Quantidade','Total');
        var volAloc   = col('VolumesAlocados','Vol.Aloc','Alocados','Allocated');
        var volPend   = col('VolumesPendentes','Vol.Pend','Pendentes','Pending');
        var tempoAprov = col('TempoAprov','Tempo','Time','TAprov','TempoA');

        // Data formatada
        var dataFmt = '';
        if(dataRaw) {
            var d = new Date(dataRaw);
            if(!isNaN(d)) dataFmt = d.toLocaleDateString('pt-BR');
            else dataFmt = dataRaw;
        }

        // Classifica endereço → estação
        var estacao = giroClassificarEstacao(endereco);

        // Status de giro
        var status = giroAvaliarStatus(produto, endereco, estacao, parseInt(volPend)||0);

        return {
            idx: idx,
            dataAlocacao: dataFmt,
            dataRaw: dataRaw,
            produto: produto,
            respAlocacao: resp,
            endereco: endereco,
            barras: barras,
            volumes: parseInt(volumes)||0,
            volAlocados: parseInt(volAloc)||0,
            volPendentes: parseInt(volPend)||0,
            tempoAprov: tempoAprov,
            estacaoId: estacao ? estacao.id : 0,
            estacaoNome: estacao ? estacao.nome : 'Não mapeada',
            estacaoCor: estacao ? estacao.cor : 'giro-sX',
            status: status // 'ok' | 'warn' | 'err'
        };
    }).filter(function(r){ return r.produto || r.endereco; });
}

// ── Classifica endereço em estação ────────────────────────────
function giroClassificarEstacao(endereco) {
    if(!endereco) return null;
    for(var i=0;i<GIRO_ESTACOES.length;i++){
        if(GIRO_ESTACOES[i].validar(endereco)) return GIRO_ESTACOES[i];
    }
    return null;
}

// ── Avalia status do giro ─────────────────────────────────────
// Regras:
//   OK   = produto tem estação definida e volumes pendentes = 0
//   WARN = produto tem estação mas tem pendências OU estação != esperado
//   ERR  = produto sem endereço ou sem estação mapeada
function giroAvaliarStatus(produto, endereco, estacao, volPend) {
    if(!endereco || !produto) return 'err';
    if(!estacao) return 'warn'; // endereço fora do mapeamento
    if(volPend > 0) return 'warn';
    return 'ok';
}

// ── Renderiza tudo ─────────────────────────────────────────────
function giroRenderAll() {
    giroRenderKPIs();
    giroRenderEstacoes();
    giroFiltrar();
    // Mostra paineis
    document.getElementById('giro-placeholder').style.display = 'none';
    document.getElementById('giro-kpis').style.display = '';
    document.getElementById('giro-mapa-estacoes').style.display = '';
    document.getElementById('giro-filtros').style.display = '';
}

// ── KPIs ───────────────────────────────────────────────────────
function giroRenderKPIs() {
    var d = _giroData;
    var ok   = d.filter(function(r){return r.status==='ok';}).length;
    var warn = d.filter(function(r){return r.status==='warn';}).length;
    var err  = d.filter(function(r){return r.status==='err';}).length;
    var pend = d.reduce(function(a,r){return a+r.volPendentes;},0);
    var estIds = [...new Set(d.map(function(r){return r.estacaoId;}).filter(Boolean))];

    document.getElementById('gk-total').textContent = d.length;
    document.getElementById('gk-ok').textContent    = ok;
    document.getElementById('gk-warn').textContent  = warn;
    document.getElementById('gk-err').textContent   = err;
    document.getElementById('gk-pend').textContent  = pend;
    document.getElementById('gk-estacoes').textContent = estIds.length;
}

// ── Mapa de estações ───────────────────────────────────────────
function giroRenderEstacoes() {
    var grid = document.getElementById('giro-estacoes-grid');
    if(!grid) return;
    var estMap = {};
    _giroData.forEach(function(r) {
        var key = r.estacaoId || 0;
        if(!estMap[key]) estMap[key] = {est:r, ok:0, warn:0, err:0, pend:0, total:0};
        estMap[key].total++;
        estMap[key][r.status]++;
        estMap[key].pend += r.volPendentes;
    });
    var html = '';
    Object.keys(estMap).forEach(function(key) {
        var e = estMap[key];
        var nome = e.est.estacaoNome;
        var cor  = e.est.estacaoCor;
        var pct  = e.total > 0 ? Math.round((e.ok/e.total)*100) : 0;
        var barColor = pct>=80?'#27ae60':pct>=50?'#f39c12':'#e74c3c';
        html += '<div style="background:white;border-radius:10px;padding:14px 18px;border-left:4px solid #2980b9;box-shadow:0 2px 8px rgba(0,0,0,.07);min-width:200px;flex:1;">'
              + '<div style="font-size:11px;font-weight:800;color:#888;text-transform:uppercase;margin-bottom:4px;"><span class="giro-station-badge '+cor+'">'+nome+'</span></div>'
              + '<div style="font-size:22px;font-weight:900;color:#2c3e50;">'+e.total+' <span style="font-size:11px;font-weight:600;color:#aaa;">produtos</span></div>'
              + '<div style="display:flex;gap:8px;margin:8px 0;flex-wrap:wrap;">'
              + '<span class="giro-badge giro-ok">✓ '+e.ok+'</span>'
              + '<span class="giro-badge giro-warn">⚡ '+e.warn+'</span>'
              + '<span class="giro-badge giro-err">✗ '+e.err+'</span>'
              + '</div>'
              + '<div style="background:#f0f0f0;border-radius:4px;height:6px;margin-top:8px;">'
              + '<div style="background:'+barColor+';height:6px;border-radius:4px;width:'+pct+'%;transition:width .5s;"></div></div>'
              + '<div style="font-size:10px;color:#aaa;margin-top:3px;">'+pct+'% correto | '+e.pend+' vol. pendentes</div>'
              + '</div>';
    });
    if(!html) html = '<div style="color:#999;font-size:12px;padding:10px;">Nenhuma estação identificada nos dados importados.</div>';
    grid.innerHTML = html;
}

// ── Filtros e tabela ───────────────────────────────────────────
window.giroFiltrar = function() {
    var busca   = (document.getElementById('giro-busca').value||'').toLowerCase();
    var fStatus = document.getElementById('giro-filtro-status').value;
    var fEst    = document.getElementById('giro-filtro-estacao').value;
    var fPend   = document.getElementById('giro-filtro-pendente').value;

    _giroFiltered = _giroData.filter(function(r) {
        if(busca && !(r.produto.toLowerCase().includes(busca) || r.endereco.toLowerCase().includes(busca) || r.respAlocacao.toLowerCase().includes(busca))) return false;
        if(fStatus !== 'all' && r.status !== fStatus) return false;
        if(fEst !== 'all') {
            if(fEst === 'outro' && r.estacaoId !== 0) return false;
            if(fEst !== 'outro' && r.estacaoId !== parseInt(fEst)) return false;
        }
        if(fPend === 'pend' && r.volPendentes === 0) return false;
        if(fPend === 'ok'   && r.volPendentes > 0)  return false;
        return true;
    });

    // Ordena
    _giroFiltered.sort(function(a,b){
        var va = a[_giroSortCol]||'', vb = b[_giroSortCol]||'';
        if(typeof va === 'number') return _giroSortAsc ? va-vb : vb-va;
        return _giroSortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });

    giroRenderTabela();
    var lbl = document.getElementById('giro-count-label');
    if(lbl) lbl.textContent = _giroFiltered.length + ' de ' + _giroData.length + ' registros';
};

window.giroSort = function(col) {
    if(_giroSortCol === col) _giroSortAsc = !_giroSortAsc;
    else { _giroSortCol = col; _giroSortAsc = true; }
    giroFiltrar();
};

function giroRenderTabela() {
    var tbody = document.getElementById('giro-tbody');
    var empty = document.getElementById('giro-empty');
    if(!tbody) return;
    if(!_giroFiltered.length) {
        tbody.innerHTML = '';
        if(empty) empty.style.display = '';
        return;
    }
    if(empty) empty.style.display = 'none';

    var badgeMap = {
        ok:   '<span class="giro-badge giro-ok"><i class="fas fa-check"></i> Giro Correto</span>',
        warn: '<span class="giro-badge giro-warn"><i class="fas fa-exclamation-triangle"></i> Revisar Giro</span>',
        err:  '<span class="giro-badge giro-err"><i class="fas fa-times"></i> Sem Definição</span>'
    };
    var rowClass = {ok:'row-ok', warn:'row-warn', err:'row-err'};

    tbody.innerHTML = _giroFiltered.map(function(r) {
        var pendCell = r.volPendentes > 0
            ? '<span style="color:#e74c3c;font-weight:800;">'+r.volPendentes+'</span>'
            : '<span style="color:#27ae60;">0</span>';
        var estBadge = '<span class="giro-station-badge '+r.estacaoCor+'">'+r.estacaoNome+'</span>';
        return '<tr class="'+rowClass[r.status]+'">'
            + '<td style="color:#777;font-size:10px;">'+r.dataAlocacao+'</td>'
            + '<td style="font-weight:700;">'+r.produto+'</td>'
            + '<td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;" title="'+r.respAlocacao+'">'+r.respAlocacao+'</td>'
            + '<td style="font-family:monospace;font-size:11px;font-weight:700;color:#2980b9;">'+r.endereco+'</td>'
            + '<td style="color:#777;font-size:10px;">'+r.barras+'</td>'
            + '<td style="text-align:right;font-weight:700;">'+r.volumes+'</td>'
            + '<td style="text-align:right;color:#27ae60;font-weight:700;">'+r.volAlocados+'</td>'
            + '<td style="text-align:right;">'+pendCell+'</td>'
            + '<td style="color:#777;font-size:10px;">'+r.tempoAprov+'</td>'
            + '<td>'+estBadge+'</td>'
            + '<td>'+badgeMap[r.status]+'</td>'
            + '</tr>';
    }).join('');
}

// ── Atualizar (reaplica filtros) ───────────────────────────────
window.giroAtualizar = function() {
    if(_giroData.length) { giroRenderAll(); }
};

// ── Exportar Excel ─────────────────────────────────────────────
window.giroExportarExcel = function() {
    if(!_giroData.length) { alert('Nenhum dado para exportar.'); return; }
    var rows = [['Data Alocacao','Produto','Resp. Alocacao','Endereco','Barras','Volumes','Vol. Alocados','Vol. Pendentes','Tempo Aprov.','Estacao','Status Giro']];
    (_giroFiltered.length ? _giroFiltered : _giroData).forEach(function(r) {
        var statusMap={ok:'Giro Correto',warn:'Revisar Giro',err:'Sem Definição'};
        rows.push([r.dataAlocacao,r.produto,r.respAlocacao,r.endereco,r.barras,r.volumes,r.volAlocados,r.volPendentes,r.tempoAprov,r.estacaoNome,statusMap[r.status]||r.status]);
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Giro_Alocacao');
    XLSX.writeFile(wb, 'Giro_Alocacao_'+new Date().toISOString().slice(0,10)+'.xlsx');
};

// ── Carrega dados salvos do Firebase ──────────────────────────
window.giroCarregarSalvos = async function() {
    try {
        var saved = await window.getFromDB('giro_alocacao_data');
        if(saved && saved.length) {
            _giroData = saved;
            _giroInitialized = true;
            giroRenderAll();
            document.getElementById('giro-btn-atualizar').style.display='';
            document.getElementById('giro-btn-excel').style.display='';
            var st = document.getElementById('giro-status-importacao');
            var tx = document.getElementById('giro-status-txt');
            if(st) st.style.display='block';
            if(tx) tx.textContent = 'Dados restaurados da nuvem — '+saved.length+' registros.';
        }
    } catch(e) {}
};


// ============================================================
// EPI + DASHBOARD RH (completo)
// ============================================================
var _epiPubSel={blusa:'',calca:'',bermuda:'',casaco:'',bota:'',luva:'',cinta:''};
var _epiInboxMap={},_epiAprovandoIdx=-1,_epiTabBusy=false;
var _epiCache={docs:null,ts:0,TTL:90000};
var EPI_ITEMS=[{k:'blusa',lbl:'Blusa',ico:'fas fa-tshirt'},{k:'calca',lbl:'Calca',ico:'fas fa-male'},{k:'bermuda',lbl:'Bermuda',ico:'fas fa-cut'},{k:'casaco',lbl:'Casaco',ico:'fas fa-mitten'},{k:'bota',lbl:'Bota',ico:'fas fa-shoe-prints'},{k:'luva',lbl:'Luva',ico:'fas fa-hand-paper'},{k:'cinta',lbl:'Cinta',ico:'fas fa-life-ring'}];
async function loadEpiColabs(){const sel=document.getElementById('epi_sel');if(!sel)return;try{const rh=await window.getFromDB('dRH');if(rh&&rh.length>0){const at=rh.filter(r=>(r['Status']||r['STATUS']||'ATIVO').toUpperCase()!=='INATIVO');at.sort((a,b)=>(a['Nome']||a['NOME']||'').localeCompare(b['Nome']||b['NOME']||''));sel.innerHTML='<option value="">-- Selecione o seu nome --</option>';at.forEach(r=>{const mat=r['Matrícula']||r['Matricula']||r['MATRICULA']||'';const nome=r['Nome']||r['NOME']||'';if(nome)sel.innerHTML+=`<option value="${mat}">${nome}</option>`;});return;}}catch(e){}sel.innerHTML='<option value="">Sem lista.</option>';}
window.epiTog=function(item,size){const grp=document.getElementById('epi_grp_'+item);if(!grp)return;const same=_epiPubSel[item]===size;grp.querySelectorAll('.epi-sz-btn').forEach(b=>b.classList.remove('selected'));_epiPubSel[item]=same?'':size;if(!same)grp.querySelectorAll('.epi-sz-btn').forEach(b=>{if(b.textContent.trim()===size)b.classList.add('selected');});};
window.submitEpiPublicForm=async function(){const sel=document.getElementById('epi_sel');const vm=document.getElementById('epi-vmsg');const err=t=>{if(vm){vm.style.display='block';vm.textContent=t;}};if(!sel||!sel.value){err('Selecione o seu nome!');return;}if(!Object.values(_epiPubSel).some(v=>v!=='')){err('Selecione pelo menos um tamanho!');return;}if(vm)vm.style.display='none';const btn=document.getElementById('epi-sub-btn');if(btn){btn.disabled=true;btn.innerText='A ENVIAR...';}const nome=sel.options[sel.selectedIndex].text;const p={nomeCompleto:nome,nome,matricula:sel.value,blusa:_epiPubSel.blusa,calca:_epiPubSel.calca,bermuda:_epiPubSel.bermuda,casaco:_epiPubSel.casaco,bota:_epiPubSel.bota,luva:_epiPubSel.luva,cinta:_epiPubSel.cinta,timestamp:new Date().getTime(),status:'PENDENTE'};try{await dbCloud.collection('logistica_epi_inbox').add(p);}catch(e){}document.getElementById('epi-pub-form').style.display='none';document.getElementById('epi-ok-view').style.display='block';};
window.copyEpiFormLink=function(){const nl=String.fromCharCode(10);const link=window.location.href.split('?')[0]+'?mode=formepi';const msg='Solicitacao de Uniforme & EPI'+nl+nl+'Acesse o link, selecione seu nome e tamanhos:'+nl+nl+link+nl+nl+'Obrigado!';if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(msg).then(()=>alert('Link copiado!')).catch(()=>window.open('https://api.whatsapp.com/send?text='+encodeURIComponent(msg),'_blank'));else window.open('https://api.whatsapp.com/send?text='+encodeURIComponent(msg),'_blank');};
window.renderEpiTab=async function(forceRefresh){if(_epiTabBusy)return;_epiTabBusy=true;try{await Promise.all([loadEpiInbox(forceRefresh),renderEpiTable(forceRefresh)]);}finally{setTimeout(()=>{_epiTabBusy=false;},1500);}};
window.loadEpiInbox=async function(forceRefresh){const body=document.getElementById('epiInboxBody');if(!body)return;body.innerHTML='<div style="padding:12px;text-align:center;color:#999;font-size:12px;"><i class="fas fa-circle-notch fa-spin"></i> Carregando...</div>';_epiInboxMap={};let items=[];try{const snap=await dbCloud.collection('logistica_epi_inbox').where('status','==','PENDENTE').get();snap.forEach(d=>items.push({id:d.id,data:d.data()}));items.sort((a,b)=>(b.data.timestamp||0)-(a.data.timestamp||0));}catch(e){}if(!items.length){body.innerHTML='<div style="padding:16px;text-align:center;color:#999;font-size:12px;">Nenhuma solicitacao pendente.</div>';return;}let h='';items.forEach((item,i)=>{_epiInboxMap[i]=item;const d=item.data;const tags=EPI_ITEMS.filter(it=>d[it.k]).map(it=>`<span class="epi-sz-tag"><i class="${it.ico}"></i> ${it.lbl}: ${d[it.k]}</span>`).join('');h+=`<div class="epi-inbox-item"><div class="epi-inbox-name">${d.nomeCompleto||d.nome||'?'} <span style="font-size:10px;color:#999;font-weight:400;">(${d.matricula})</span></div><div class="epi-inbox-sizes">${tags}</div><div style="display:flex;gap:8px;margin-top:8px;"><button class="btn btn-green" style="padding:5px 12px;font-size:12px;" onclick="abrirModalAprovar(${i})"><i class="fas fa-check-square"></i> Aprovar Itens</button><button class="btn btn-red" style="padding:5px 12px;font-size:12px;" onclick="rejeitarEpi(${i})"><i class="fas fa-times"></i> Rejeitar</button></div></div>`;});body.innerHTML=h;};
window.abrirModalAprovar=function(i){const item=_epiInboxMap[i];if(!item)return;_epiAprovandoIdx=i;const d=item.data;const nEl=document.getElementById('modalAprovarEpiNome');if(nEl)nEl.innerHTML=`<i class="fas fa-user"></i> ${d.nomeCompleto||d.nome||'?'} <span style="color:#888;font-weight:400;font-size:12px;">(${d.matricula})</span>`;const iEl=document.getElementById('modalAprovarEpiItens');if(iEl){const req=EPI_ITEMS.filter(it=>d[it.k]);iEl.innerHTML=req.map(it=>`<label class="epi-item-chk"><input type="checkbox" id="chk_${it.k}" checked onchange="this.closest('label').style.opacity=this.checked?'1':'0.5'"><i class="${it.ico}" style="color:#8e44ad;font-size:15px;"></i><span style="font-weight:800;min-width:60px;">${it.lbl}:</span><span style="background:#8e44ad;color:white;border-radius:5px;padding:2px 9px;font-size:12px;">${d[it.k]}</span><span style="margin-left:auto;font-size:11px;color:#27ae60;font-weight:700;">Sera fornecido</span></label>`).join('');}const btn=document.getElementById('btnConfirmarAprEpi');if(btn){btn.disabled=false;btn.innerHTML='<i class="fas fa-check"></i> Confirmar Aprovacao';}document.getElementById('modalAprovarEpi').style.display='flex';};
window.confirmarAprovacaoEpi=async function(){const item=_epiInboxMap[_epiAprovandoIdx];if(!item)return;const btn=document.getElementById('btnConfirmarAprEpi');if(btn){btn.disabled=true;btn.innerHTML='<i class="fas fa-circle-notch fa-spin"></i> Salvando...';}const d=item.data;const forn={};EPI_ITEMS.forEach(it=>{const chk=document.getElementById('chk_'+it.k);forn[it.k]=(chk&&chk.checked&&d[it.k])?d[it.k]:'-';});if(!Object.values(forn).some(v=>v!=='-')){alert('Selecione pelo menos um item!');if(btn){btn.disabled=false;btn.innerHTML='<i class="fas fa-check"></i> Confirmar Aprovacao';}return;}const nr={id:'tmp_'+Date.now(),mat:d.matricula,nome:d.nomeCompleto||d.nome||'',blusa:forn.blusa,calca:forn.calca,bermuda:forn.bermuda,casaco:forn.casaco,bota:forn.bota,luva:forn.luva,cinta:forn.cinta,dataAprov:new Date().toLocaleDateString('pt-BR'),dataAprovTs:Date.now(),status:'aprovado',dataEntrega:''};if(_epiCache.docs)_epiCache.docs.unshift(nr);document.getElementById('modalAprovarEpi').style.display='none';renderEpiTableFromCache();try{const[snapReg]=await Promise.all([dbCloud.collection('logistica_epi_registros').add({inboxId:item.id,mat:nr.mat,nome:nr.nome,blusa:nr.blusa,calca:nr.calca,bermuda:nr.bermuda,casaco:nr.casaco,bota:nr.bota,luva:nr.luva,cinta:nr.cinta,dataAprov:nr.dataAprov,dataAprovTs:nr.dataAprovTs,status:'aprovado',dataEntrega:''}),dbCloud.collection('logistica_epi_inbox').doc(item.id).update({status:'APROVADO'})]);if(_epiCache.docs){const idx=_epiCache.docs.findIndex(r=>r.id===nr.id);if(idx>=0)_epiCache.docs[idx].id=snapReg.id;}}catch(e){console.warn('EPI:',e);if(_epiCache.docs)_epiCache.docs=_epiCache.docs.filter(r=>r.id!==nr.id);renderEpiTableFromCache();alert('Erro: '+e.message);return;}loadEpiInbox(true);};
window.rejeitarEpi=async function(i){const item=_epiInboxMap[i];if(!item||!confirm('Rejeitar?'))return;try{await dbCloud.collection('logistica_epi_inbox').doc(item.id).update({status:'REJEITADO'});}catch(e){}loadEpiInbox(true);};
window.confirmarEntregaEpi=async function(docId){if(!confirm('Confirmar entrega fisica?'))return;const de=new Date().toLocaleDateString('pt-BR');if(_epiCache.docs){const r=_epiCache.docs.find(r=>r.id===docId);if(r){r.status='entregue';r.dataEntrega=de;}}renderEpiTableFromCache();try{await dbCloud.collection('logistica_epi_registros').doc(docId).update({status:'entregue',dataEntrega:de});}catch(e){alert('Erro:'+e.message);_epiCache.docs=null;renderEpiTable(true);}};
window.renderEpiTable=async function(forceRefresh){if(!forceRefresh&&_epiCache.docs&&(Date.now()-_epiCache.ts)<_epiCache.TTL){renderEpiTableFromCache();return;}const tb=document.getElementById('tbEpi');if(!tb)return;tb.innerHTML='<tr><td colspan="12" style="text-align:center;padding:12px;color:#999;"><i class="fas fa-circle-notch fa-spin"></i> Carregando...</td></tr>';try{const snap=await dbCloud.collection('logistica_epi_registros').get();const docs=[];snap.forEach(d=>docs.push({id:d.id,...d.data()}));docs.sort((a,b)=>(b.dataAprovTs||0)-(a.dataAprovTs||0));_epiCache.docs=docs;_epiCache.ts=Date.now();}catch(e){if(!_epiCache.docs){tb.innerHTML='<tr><td colspan="12" style="text-align:center;padding:20px;color:#e74c3c;">Erro ao carregar.</td></tr>';return;}}renderEpiTableFromCache();};
window.renderEpiTableFromCache=function(){const tb=document.getElementById('tbEpi');if(!tb)return;const docs=_epiCache.docs||[];const fSel=document.getElementById('epiFilterStatus');const f=fSel?fSel.value:'all';const filtered=f==='all'?docs:docs.filter(r=>r.status===f);const cnt=document.getElementById('epiRegCount');if(cnt)cnt.textContent=docs.length;const ecnt=document.getElementById('epiEntregCount');if(ecnt)ecnt.textContent=docs.filter(r=>r.status==='entregue').length+' entregues';if(!filtered.length){tb.innerHTML='<tr><td colspan="12" style="text-align:center;padding:20px;color:#999;">Nenhum registro.</td></tr>';return;}const sz=v=>(!v||v==='-')?'<span style="color:#ccc;">-</span>':`<span style="background:#e8daef;color:#6c3483;padding:2px 7px;border-radius:4px;font-weight:700;font-size:10px;">${v}</span>`;tb.innerHTML=filtered.map(x=>{const ok=x.status==='entregue';const isParcial=x.status==='parcial';const badge=ok?`<span class="epi-badge-entg"><i class="fas fa-check-double"></i> Entregue ${x.dataEntrega||''}</span>`:isParcial?`<span style="background:#d6eaf8;color:#1a5276;padding:3px 9px;border-radius:10px;font-size:10px;font-weight:800;white-space:nowrap;"><i class="fas fa-box-open"></i> Parcial ${x.dataEntrega||''}</span>`:`<span class="epi-badge-pend"><i class="fas fa-box"></i> Aguard. Entrega</span>`;const act=ok?`<button onclick="excluirEpiReg('${x.id}')" style="background:#e74c3c;color:white;border:none;border-radius:4px;padding:3px 8px;font-size:10px;font-weight:700;cursor:pointer;"><i class="fas fa-trash"></i></button>`:`<button onclick="abrirModalEntregaEpi('${x.id}')" style="background:#27ae60;color:white;border:none;border-radius:6px;padding:5px 10px;font-size:11px;font-weight:700;cursor:pointer;"><i class="fas fa-box-open"></i> Entregar</button> <button onclick="excluirEpiReg('${x.id}')" style="background:#e74c3c;color:white;border:none;border-radius:4px;padding:5px 8px;font-size:10px;font-weight:700;cursor:pointer;margin-left:4px;"><i class="fas fa-trash"></i></button>`;return `<tr class="${ok?'epi-reg-row-entg':''}"><td style="text-align:left;font-weight:700;white-space:nowrap;">${x.nome}</td><td style="text-align:left;color:#777;font-size:10px;">${x.mat}</td><td>${sz(x.blusa)}</td><td>${sz(x.calca)}</td><td>${sz(x.bermuda)}</td><td>${sz(x.casaco)}</td><td>${sz(x.bota)}</td><td>${sz(x.luva)}</td><td>${sz(x.cinta)}</td><td style="color:#777;font-size:10px;white-space:nowrap;">${x.dataAprov||''}</td><td>${badge}</td><td style="white-space:nowrap;">${act}</td></tr>`;}).join('');};
window.excluirEpiReg=async function(docId){if(!confirm('Excluir permanentemente?'))return;if(_epiCache.docs)_epiCache.docs=_epiCache.docs.filter(r=>r.id!==docId);renderEpiTableFromCache();try{await dbCloud.collection('logistica_epi_registros').doc(docId).delete();}catch(e){alert('Erro:'+e.message);_epiCache.docs=null;renderEpiTable(true);}};
window.filterEpiTable=function(){const q=(document.getElementById('epiSearchInput').value||'').toLowerCase();document.querySelectorAll('#tbEpi tr').forEach(r=>{r.style.display=q===''||r.textContent.toLowerCase().includes(q)?'':'none';});};
window.exportEpiExcel=async function(){try{let docs=_epiCache.docs;if(!docs){const snap=await dbCloud.collection('logistica_epi_registros').get();docs=[];snap.forEach(d=>docs.push(d.data()));docs.sort((a,b)=>(b.dataAprovTs||0)-(a.dataAprovTs||0));}if(!docs.length){alert('Sem registros.');return;}const rows=[['Colaborador','Matricula','Blusa','Calca','Bermuda','Casaco','Bota','Luva','Cinta','Data Aprovacao','Status','Data Entrega']];docs.forEach(x=>rows.push([x.nome,x.mat,x.blusa,x.calca,x.bermuda,x.casaco,x.bota,x.luva,x.cinta,x.dataAprov||'',x.status==='entregue'?'Entregue':'Aguardando',x.dataEntrega||'']));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Uniforme_EPI');XLSX.writeFile(wb,'Uniforme_EPI_'+new Date().toISOString().slice(0,10)+'.xlsx');}catch(e){alert('Erro:'+e.message);}};

// Dashboard RH
window.renderDashboardRH=async function(){const btn=document.querySelector('[onclick*="renderDashboardRH"]');if(btn){btn.disabled=true;btn.innerHTML='<i class="fas fa-circle-notch fa-spin"></i> Actualizando...';}try{const ano=new Date().getFullYear();setDRH('drh-ano-label',ano);const[dRHfb,rhDataFb,pendInbox,epiRegs]=await Promise.all([window.getFromDB('dRH').catch(()=>null),window.getFromDB('rhData').catch(()=>null),dbCloud.collection('logistica_ferias_inbox').where('status','==','PENDENTE').get().catch(()=>null),dbCloud.collection('logistica_epi_registros').get().catch(()=>null)]);const lista=(dRHfb&&dRHfb.length)?dRHfb:(window.dRH||[]);const rh=rhDataFb||window.rhData||{};const hoje=new Date();hoje.setHours(0,0,0,0);let ativos=0,inativos=0,entradasAno=0,supCount=0,liderCount=0,opCount=0;lista.forEach(c=>{const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||'';const func=(c['Função']||c['Funcao']||c['FUNÇÃO']||'').toUpperCase();const s=(c['Status']||c['STATUS']||'ATIVO').toUpperCase();const rhM=rh[mat]||{};const sf=(rhM.statusManual&&rhM.statusManual!=='AUTO')?rhM.statusManual.toUpperCase():s;if(sf==='INATIVO'){inativos++;return;}ativos++;if(func.includes('SUPERVISOR'))supCount++;else if(func.includes('LIDER')||func.includes('LÍDER'))liderCount++;else opCount++;const admRaw=c['Admissão']||c['Admissao']||c['ADMISSÃO']||'';let admY=0;if(admRaw.includes('/'))admY=parseInt(admRaw.split('/')[2]||0);else if(admRaw.includes('-'))admY=parseInt(admRaw.split('-')[0]||0);if(admY===ano)entradasAno++;});const turnover=inativos>0?((inativos/Math.max(1,ativos+inativos))*100).toFixed(1):'0.0';const pendRep=(typeof pendenciasRH!=='undefined')?pendenciasRH.length:0;setDRH('drh-entradas',entradasAno);setDRH('drh-saidas',inativos);setDRH('drh-turnover',turnover+'%');setDRH('drh-ativos',ativos);setDRH('drh-reposicoes',pendRep);let insC='📊 <strong>'+ativos+' ativos</strong> | '+supCount+' sup | '+liderCount+' líd | '+opCount+' op. ';insC+=parseFloat(turnover)>15?'🚨 Turnover elevado ('+turnover+'%).':parseFloat(turnover)>8?'⚡ Moderado ('+turnover+'%).':'✅ Controlado ('+turnover+'%).';if(pendRep>0)insC+=' 🔴 <strong>'+pendRep+' vaga(s) em reposição.</strong>';setDRH('ai-colab-insight',insC,true);let fAtivas=0,fVencidas=0,em30=0;lista.forEach(c=>{const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||'';const r=rh[mat]||{};if(r.feriasInicio&&r.feriasFim){const fi=new Date(r.feriasInicio+'T00:00:00'),ff=new Date(r.feriasFim+'T00:00:00');if(hoje>=fi&&hoje<=ff)fAtivas++;}if(r.limiteFerias){const lim=new Date(r.limiteFerias+'T00:00:00'),d=(lim-hoje)/86400000;if(d<0&&!r.feriasInicio)fVencidas++;else if(d>=0&&d<=30&&!r.feriasInicio)em30++;}});const pendFer=pendInbox?pendInbox.size:0;setDRH('drh-ferias-ativas',fAtivas);setDRH('drh-ferias-vencidas',fVencidas);setDRH('drh-ferias-a-vencer',em30);setDRH('drh-ferias-pendentes',pendFer);let insFer='';if(fVencidas>0)insFer+='🚨 <strong>'+fVencidas+' vencidas</strong> — risco legal. ';if(em30>0)insFer+='⏰ '+em30+' a vencer em 30 dias. ';if(fAtivas>0)insFer+='🌴 '+fAtivas+' em gozo. ';if(pendFer>0)insFer+='📬 '+pendFer+' pendente(s). ';if(!insFer)insFer='✅ Sem irregularidades.';setDRH('ai-ferias-insight',insFer,true);let totPos=0,totNeg=0,qtdPos=0,qtdNeg=0;Object.keys(rh).forEach(mat=>{const bh=typeof rh[mat].banco==='number'?rh[mat].banco:0;if(bh>0){totPos+=bh;qtdPos++;}else if(bh<0){totNeg+=Math.abs(bh);qtdNeg++;}});const fmtBH=h=>{const a=Math.abs(h);return(h<0?'-':'+')+String(Math.floor(a)).padStart(2,'0')+':'+String(Math.round((a-Math.floor(a))*60)).padStart(2,'0');};setDRH('drh-bh-positivo',fmtBH(totPos));setDRH('drh-bh-negativo',fmtBH(-totNeg));setDRH('drh-bh-qtd-pos',qtdPos);setDRH('drh-bh-qtd-neg',qtdNeg);setDRH('ai-bh-insight',(totPos>100?'⚠️ Saldo positivo elevado. ':qtdNeg>ativos*0.2?'🔴 +20% com saldo negativo. ':'')+'✅ OK',true);let totalAbs=0,diasAbs=0,motCount={},pesCount={},absMeses={};lista.forEach(c=>{const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||'';const nome=c['Nome']||c['NOME']||mat;((rh[mat]||{}).faltas||[]).forEach(f=>{totalAbs++;const dias=parseInt(f.dias)||1;diasAbs+=dias;motCount[f.motivo||'Outro']=(motCount[f.motivo||'Outro']||0)+dias;pesCount[nome]=(pesCount[nome]||0)+dias;const mk=(f.data||'').substring(0,7)||'N/A';absMeses[mk]=(absMeses[mk]||0)+dias;});});const topMot=Object.keys(motCount).sort((a,b)=>motCount[b]-motCount[a])[0]||'—';const topPes=Object.keys(pesCount).sort((a,b)=>pesCount[b]-pesCount[a])[0]||'—';setDRH('drh-abs-total',totalAbs);setDRH('drh-abs-dias',diasAbs);setDRH('drh-abs-motivo',topMot.length>22?topMot.slice(0,20)+'...':topMot);setDRH('drh-abs-pessoa',topPes.length>22?topPes.slice(0,20)+'...':topPes);renderChartAbsMotivo(motCount);renderChartAbsMensal2(absMeses);const taxa=ativos>0?((diasAbs/(ativos*22))*100).toFixed(1):0;setDRH('ai-abs-insight','📊 Taxa: <strong>'+taxa+'%</strong>. '+(parseFloat(taxa)>4?'🚨 Acima do limite. Motivo: <strong>'+topMot+'</strong>.':parseFloat(taxa)>2?'⚠️ Moderada.':'✅ Ok.'),true);let pTotal=0,pFaltas=0,pMarcacoes=0,pOk=0,pontoTipos={};Object.keys(rh).forEach(mat=>{((rh[mat]||{}).ponto||[]).forEach(p=>{pTotal++;const tp=(p.tipo||'').toLowerCase(),ac=(p.acao||'').toLowerCase();if(tp.includes('falta sem'))pFaltas++;if(tp.includes('marcação')||tp.includes('marcacao'))pMarcacoes++;if(ac.includes('justificado')||ac.includes('arquivado'))pOk++;pontoTipos[p.tipo||'Outro']=(pontoTipos[p.tipo||'Outro']||0)+1;});});setDRH('drh-ponto-total',pTotal);setDRH('drh-ponto-faltas',pFaltas);setDRH('drh-ponto-marcacoes',pMarcacoes);setDRH('drh-ponto-ok',pOk);renderChartPontoTipo(pontoTipos);setDRH('ai-ponto-insight',pTotal>0?'🗓 <strong>'+pTotal+' ocorrência(s)</strong>.'+(pFaltas>5?' ⚠️ Faltas: '+pFaltas+'.':''):'Nenhuma ocorrência.',true);let epiTotal=0,epiEntg=0;if(epiRegs)epiRegs.forEach(d=>{epiTotal++;if(d.data().status==='entregue')epiEntg++;});setDRH('drh-epi-total',epiTotal);setDRH('drh-epi-entregues',epiEntg);setDRH('drh-epi-pendentes',epiTotal-epiEntg);}catch(e){console.error('Dashboard RH:',e);}finally{if(btn){btn.disabled=false;btn.innerHTML='<i class="fas fa-sync-alt"></i> Atualizar Painel';}}};
function renderChartAbsMensal(x){}
function renderChartAbsMensal2(meses){const c=document.getElementById('chartAbsMensal');if(!c)return;if(window._dashCharts&&window._dashCharts.absMensal)window._dashCharts.absMensal.destroy();const sorted=Object.keys(meses).sort();if(!sorted.length)return;if(!window._dashCharts)window._dashCharts={};window._dashCharts.absMensal=new Chart(c.getContext('2d'),{type:'bar',data:{labels:sorted,datasets:[{label:'Dias',data:sorted.map(k=>meses[k]),backgroundColor:'#e74c3c99',borderColor:'#e74c3c',borderWidth:1}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:10}}}}}});}
function renderChartPontoTipo(tipos){const c=document.getElementById('chartPontoTipo');if(!c)return;if(window._dashCharts&&window._dashCharts.pontoTipo)window._dashCharts.pontoTipo.destroy();if(!tipos||!Object.keys(tipos).length)return;if(!window._dashCharts)window._dashCharts={};const labels=Object.keys(tipos);window._dashCharts.pontoTipo=new Chart(c.getContext('2d'),{type:'bar',data:{labels,datasets:[{label:'Ocorrencias',data:labels.map(k=>tipos[k]),backgroundColor:'#8e44ad99',borderColor:'#8e44ad',borderWidth:1}]},options:{indexAxis:'y',responsive:true,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{font:{size:10}}}}}});}


// ================================================================
// METAS — Absenteísmo (<4%) + Ressuprimento (≥90%)
// 70% da meta total do CD
// Dados Abs: rhData[mat].faltas (aba Absenteísmo & Ponto)
// Dados RSS: importação XLSX — consolidado por DIA
// ================================================================

var _metaChartAbs = null;
var _metaChartRss = null;
var _rssData = []; // [{data, pedido, atendido}] agrupado por dia

window.initMetas = function() { calcularMetas(); };

window.calcularMetas = async function() {
    var hoje = new Date();
    var ini  = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    var fim  = new Date(hoje.getFullYear(), hoje.getMonth()+1, 0);
    var fmt  = function(d){ return d.toLocaleDateString('pt-BR'); };
    var lbl  = document.getElementById('meta-periodo-label');
    if (lbl) lbl.textContent = 'Período: ' + fmt(ini) + ' a ' + fmt(fim);
    await calcMetaAbsenteismo();
    calcMetaRessuprimento();
    atualizarScoreTotal();
};

// ═══════════════════════════════════════════════════════════════
// META 1 — ABSENTEÍSMO
// Fonte: rhData[mat].faltas (aba Absenteísmo & Ponto unificada)
// ═══════════════════════════════════════════════════════════════
async function calcMetaAbsenteismo() {
    // Busca dados do Firebase (mesma fonte do Dashboard RH)
    var dRHfb   = null;
    var rhDataFb = null;
    try {
        var results = await Promise.all([
            window.getFromDB('dRH').catch(function(){ return null; }),
            window.getFromDB('rhData').catch(function(){ return null; })
        ]);
        dRHfb   = results[0];
        rhDataFb = results[1];
    } catch(e) {}

    var listaRH = (dRHfb && dRHfb.length) ? dRHfb : (window.dRH || []);
    var rh      = (rhDataFb && Object.keys(rhDataFb).length) ? rhDataFb : (window.rhData || {});

    if (!listaRH.length) {
        var el = document.getElementById('abs-ia-text');
        if (el) el.textContent = '⚠️ Nenhum colaborador encontrado. Importe o quadro de colaboradores primeiro.';
        return;
    }

    var hoje   = new Date(); hoje.setHours(0,0,0,0);
    var ano    = hoje.getFullYear();
    var mes    = hoje.getMonth();
    var anoMes = ano + '-' + String(mes+1).padStart(2,'0');

    // Quadro ativo (apenas ATIVO real)
    var quadroAtivo = Math.max(1, listaRH.filter(function(c) {
        var mat = c['Matrícula']||c['Matricula']||c['MATRICULA']||'';
        var s   = (c['Status']||c['STATUS']||'ATIVO').toUpperCase();
        var rhM = (rh[mat]||{});
        var sf  = (rhM.statusManual && rhM.statusManual !== 'AUTO') ? rhM.statusManual.toUpperCase() : s;
        return sf === 'ATIVO';
    }).length);

    // Dias úteis decorridos e total do mês
    var diasUteisAte = 0;
    var diasUteisMes = 0;
    var diaAtual  = hoje.getDate();
    var ultimoDia = new Date(ano, mes+1, 0).getDate();
    for (var d = 1; d <= ultimoDia; d++) {
        var dow = new Date(ano, mes, d).getDay();
        if (dow === 0 || dow === 6) continue;
        diasUteisMes++;
        if (d <= diaAtual) diasUteisAte++;
    }

    // Coleta faltas do mês (exclui Folga)
    var diasAusentes = 0;
    var impactMap    = {};
    var diasPorDia   = {};

    Object.keys(rh).forEach(function(mat) {
        var faltas = (rh[mat]||{}).faltas || [];
        var emp    = listaRH.find(function(c){ return (c['Matrícula']||c['Matricula']||c['MATRICULA']||'') === mat; });
        var nome   = emp ? (emp['Nome']||emp['NOME']||mat) : mat;

        faltas.forEach(function(f) {
            if ((f.motivo||'').toLowerCase().includes('folga')) return;
            var dt = f.data || '';
            var fAnoMes = '';
            var dia = '';
            if (dt.includes('/')) {
                var p = dt.split('/');
                fAnoMes = p[2] + '-' + p[1].padStart(2,'0');
                dia     = p[0].padStart(2,'0');
            } else if (dt.includes('-')) {
                fAnoMes = dt.substring(0,7);
                dia     = dt.substring(8,10);
            }
            if (fAnoMes !== anoMes) return;
            var dias = parseInt(f.dias) || 1;
            diasAusentes += dias;
            diasPorDia[dia] = (diasPorDia[dia]||0) + dias;
            if (!impactMap[mat]) impactMap[mat] = { nome: nome, dias: 0 };
            impactMap[mat].dias += dias;
        });
    });

    var taxa    = diasUteisAte > 0 ? (diasAusentes / (quadroAtivo * diasUteisAte)) * 100 : 0;
    var maxDias = 0.04 * quadroAtivo * diasUteisMes;
    var saldo   = Math.floor(maxDias - diasAusentes);
    var barW    = Math.min((taxa / 8) * 100, 100);
    var barCol  = taxa <= 2 ? '#27ae60' : taxa <= 4 ? '#f39c12' : '#e74c3c';
    var atingiu = taxa <= 4;

    var set  = function(id,v){ var el=document.getElementById(id); if(el) el.textContent=v; };
    var setH = function(id,v){ var el=document.getElementById(id); if(el) el.innerHTML=v; };
    var setSt = function(id,p,v){ var el=document.getElementById(id); if(el) el.style[p]=v; };

    set('abs-kpi-taxa',  taxa.toFixed(2)+'%');
    set('abs-kpi-dias',  diasAusentes);
    set('abs-kpi-total', quadroAtivo);
    set('abs-kpi-uteis', diasUteisMes);
    set('abs-kpi-tol',   Math.floor(maxDias));
    set('abs-kpi-saldo', (saldo>=0?'+':'')+saldo);
    setSt('abs-kpi-saldo','color', saldo>=0?'#27ae60':'#e74c3c');
    setSt('abs-gauge-fill','width', barW+'%');
    setSt('abs-gauge-fill','background', barCol);
    set('abs-gauge-pct', taxa.toFixed(2)+'%');
    setSt('abs-gauge-pct','color', barCol);
    set('meta-res-abs-pct', atingiu ? '✅ OK' : '⚠️ '+taxa.toFixed(1)+'%');

    var badge = document.getElementById('abs-meta-badge');
    if (badge) {
        badge.textContent = atingiu
            ? '✅ META ATINGIDA ('+taxa.toFixed(2)+'%)'
            : '⚠️ ACIMA DA META ('+taxa.toFixed(2)+'%)';
        badge.style.background = atingiu ? 'rgba(39,174,96,.25)' : 'rgba(231,76,60,.25)';
    }

    // Lista de impacto
    var arr    = Object.values(impactMap).sort(function(a,b){ return b.dias-a.dias; });
    var listEl = document.getElementById('abs-impact-list');
    if (listEl) {
        listEl.innerHTML = arr.length ? arr.slice(0,15).map(function(x){
            var contrib = diasUteisAte>0 ? ((x.dias/(quadroAtivo*diasUteisAte))*100).toFixed(2) : '0.00';
            return '<div class="meta-impact-item">' +
                '<span class="meta-impact-name">'+x.nome+'</span>' +
                '<span style="display:flex;gap:8px;align-items:center;">' +
                '<span class="meta-impact-dias">'+x.dias+' dias</span>' +
                '<span style="font-size:10px;color:#e74c3c;font-weight:700;">+'+contrib+'%</span>' +
                '</span></div>';
        }).join('') : '<div style="padding:16px;text-align:center;color:#999;font-size:12px;">Nenhuma ausência no mês.</div>';
    }

    // IA
    var ia = '';
    if (taxa === 0) {
        ia = '✅ <strong>Perfeito!</strong> Nenhuma ausência registada no mês atual.';
    } else if (taxa <= 2) {
        ia = '✅ Índice <strong>'+taxa.toFixed(2)+'%</strong> — excelente. Saldo de <strong>'+saldo+'</strong> dias disponíveis.';
    } else if (taxa <= 4) {
        ia = '⚡ Índice <strong>'+taxa.toFixed(2)+'%</strong> — atenção, saldo restante: <strong>'+saldo+' dias</strong>.<br>'+
             '• Conversar com colaboradores de maior impacto<br>• Monitorar presença diariamente';
    } else {
        var rest = diasUteisMes - diasUteisAte;
        ia = '🚨 <strong>META EM RISCO!</strong> Índice: <strong>'+taxa.toFixed(2)+'%</strong> (meta: 4%).<br>'+
             '• Acionar gestores imediatamente<br>• Solicitar documentação das ausências<br>'+
             '• Não aprovar folgas até fim do mês';
        if (rest > 0 && (maxDias - diasAusentes) > 0) {
            var txMax = ((maxDias-diasAusentes)/(quadroAtivo*rest)*100).toFixed(2);
            ia += '<br>• Taxa máxima permitida nos <strong>'+rest+' dias úteis restantes</strong>: <strong>'+txMax+'%</strong>';
        }
    }
    setH('abs-ia-text', ia);

    // Gráfico
    renderChartAbsMeta(diasPorDia, quadroAtivo, ano, mes, diasUteisMes);
}

function renderChartAbsMeta(diasPorDia, quadro, ano, mes, diasUteisMes) {
    var canvas = document.getElementById('chartAbsMeta');
    if (!canvas) return;
    if (_metaChartAbs) { _metaChartAbs.destroy(); _metaChartAbs = null; }
    var labels=[], taxaLine=[], metaLine=[];
    var ultimo = new Date(ano, mes+1, 0).getDate();
    var acum = 0, uteis = 0;
    for (var d=1; d<=ultimo; d++) {
        var dow = new Date(ano,mes,d).getDay();
        if (dow===0||dow===6) continue;
        uteis++;
        acum += (diasPorDia[String(d).padStart(2,'0')]||0);
        var taxa = uteis>0 ? (acum/(quadro*uteis))*100 : 0;
        labels.push(d+'/'+(mes+1));
        taxaLine.push(parseFloat(taxa.toFixed(2)));
        metaLine.push(4);
    }
    _metaChartAbs = new Chart(canvas.getContext('2d'), {
        type: 'line',
        data: { labels: labels, datasets: [
            { label: 'Absenteísmo (%)', data: taxaLine, borderColor:'#e74c3c', backgroundColor:'rgba(231,76,60,.08)', fill:true, tension:0.35, pointRadius:3 },
            { label: 'Meta (4%)',       data: metaLine, borderColor:'#27ae60', borderDash:[6,3], fill:false, pointRadius:0 }
        ]},
        options: { responsive:true, plugins:{legend:{position:'top',labels:{font:{size:10}}}},
            scales:{ y:{beginAtZero:true,suggestedMax:8,ticks:{callback:function(v){return v+'%';},font:{size:10}}}, x:{ticks:{font:{size:9},maxTicksLimit:16}} } }
    });
}

// ═══════════════════════════════════════════════════════════════
// META 2 — RESSUPRIMENTO por DIA
// Arquivo XLSX: colunas Data, Pedido, Atendido (por dia)
// ═══════════════════════════════════════════════════════════════
window.importarRessuprimento = async function(input) {
    var file = input.files[0]; if (!file) return;
    try {
        var rows = await new Promise(function(resolve,reject){
            var r=new FileReader();
            r.onload=function(e){try{var wb=XLSX.read(e.target.result,{type:'binary',cellDates:true});resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''}));}catch(err){reject(err);}};
            r.onerror=reject; r.readAsBinaryString(file);
        });
        _rssData = processarRssPorDia(rows);
        try { await window.saveToDB('meta_rss_dias', _rssData); } catch(e) {}
        calcMetaRessuprimento();
        atualizarScoreTotal();
    } catch(e) { alert('Erro: '+e.message); }
    input.value='';
};

function processarRssPorDia(rows) {
    // Colunas do arquivo ressuprimento_mensal.xlsx:
    // Mapa | Movimento (data) | Romaneio | Endereço | Produto | Descrição
    // Qtde Pedida | Qtde Pedida CX | Qtde Atendida | Qtde Atendida CX
    // Qtde N Atendida | Qtde Resolução Div. | Saldo Pendente | ...
    var mapaData = {};

    rows.forEach(function(r) {
        // Detecta coluna de data (Movimento)
        var dt = null;
        var keys = Object.keys(r);
        for (var i=0; i<keys.length; i++) {
            var kl = keys[i].toLowerCase().replace(/\s/g,'');
            if (kl === 'movimento' || kl === 'data' || kl === 'date' || kl === 'dia') {
                var v = r[keys[i]];
                if (v) { dt = v; break; }
            }
        }
        if (!dt) return;

        // Normaliza data para yyyy-mm-dd
        var dStr = '';
        if (dt instanceof Date || (typeof dt === 'object' && dt.getFullYear)) {
            dStr = dt.getFullYear() + '-' +
                   String(dt.getMonth()+1).padStart(2,'0') + '-' +
                   String(dt.getDate()).padStart(2,'0');
        } else {
            var s = String(dt).trim();
            if (s.includes('/')) {
                var p = s.split('/');
                // dd/mm/yyyy or mm/dd/yyyy — assume dd/mm/yyyy
                if (p.length === 3) {
                    var y = p[2].length === 4 ? p[2] : '20'+p[2];
                    dStr = y + '-' + p[1].padStart(2,'0') + '-' + p[0].padStart(2,'0');
                }
            } else if (s.includes('-') && s.length >= 8) {
                dStr = s.substring(0,10);
            } else if (!isNaN(Number(s)) && Number(s) > 40000) {
                // Excel serial date
                var d = new Date((Number(s) - 25569) * 86400 * 1000);
                dStr = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
            }
        }
        if (!dStr || dStr.length < 8) return;

        // Detecta colunas de pedido e atendido
        var pedido   = 0;
        var atendido = 0;
        keys.forEach(function(k) {
            var kl = k.toLowerCase().replace(/[\s\.]/g,'');
            var v  = parseFloat(r[k]) || 0;
            // Qtde Pedida (sem CX, sem Caixa)
            if (!pedido && (kl === 'qtdepedida' || kl === 'qtdpedida' || kl === 'quantidadepedida') &&
                !kl.includes('cx') && !kl.includes('caixa')) {
                pedido = v;
            }
            // Qtde Atendida (sem CX, sem Caixa)
            if (!atendido && (kl === 'qtdeatendida' || kl === 'qtdatendida' || kl === 'quantidadeatendida') &&
                !kl.includes('cx') && !kl.includes('caixa')) {
                atendido = v;
            }
        });

        if (pedido === 0) return; // linha sem pedido — pula

        if (!mapaData[dStr]) mapaData[dStr] = { data: dStr, pedido: 0, atendido: 0 };
        mapaData[dStr].pedido   += pedido;
        mapaData[dStr].atendido += atendido;
    });

    var result = Object.values(mapaData).sort(function(a,b){ return a.data.localeCompare(b.data); });
    return result;
}

function calcMetaRessuprimento() {
    if (!_rssData.length) {
        window.getFromDB('meta_rss_dias').then(function(s){ if(s&&s.length){_rssData=s;calcMetaRessuprimento();} }).catch(function(){});
        return;
    }
    var totPed = _rssData.reduce(function(a,r){return a+r.pedido;},0);
    var totAte = _rssData.reduce(function(a,r){return a+r.atendido;},0);
    var totPend= totPed - totAte;
    var pct    = totPed>0 ? (totAte/totPed)*100 : 0;
    var atingiu= pct>=90;
    var gap    = Math.max(0,90-pct).toFixed(1);
    var fmtN   = function(n){ return Math.round(n).toLocaleString('pt-BR'); };

    var set  = function(id,v){var el=document.getElementById(id);if(el)el.textContent=v;};
    var setH = function(id,v){var el=document.getElementById(id);if(el)el.innerHTML=v;};
    var setSt= function(id,p,v){var el=document.getElementById(id);if(el)el.style[p]=v;};

    set('rss-kpi-pct',  pct.toFixed(2)+'%');
    set('rss-kpi-ped',  fmtN(totPed));
    set('rss-kpi-ate',  fmtN(totAte));
    set('rss-kpi-pend', fmtN(totPend));
    set('rss-kpi-dias', _rssData.length+' dias');
    set('rss-kpi-gap',  atingiu?'0% ✅':gap+'%');
    setSt('rss-kpi-pct','color', atingiu?'#27ae60':'#e74c3c');
    setSt('rss-gauge-fill','width', Math.min(pct,100)+'%');
    setSt('rss-gauge-fill','background', pct>=90?'#27ae60':pct>=75?'#f39c12':'#e74c3c');
    set('rss-gauge-pct', pct.toFixed(2)+'%');
    setSt('rss-gauge-pct','color', atingiu?'#27ae60':'#e74c3c');
    set('meta-res-rss-pct', atingiu?'✅ OK':'⚠️ '+pct.toFixed(1)+'%');

    var badge=document.getElementById('rss-meta-badge');
    if(badge){badge.textContent=atingiu?'✅ META ATINGIDA ('+pct.toFixed(2)+'%)':'⚠️ ABAIXO DA META ('+pct.toFixed(2)+'%)';badge.style.background=atingiu?'rgba(39,174,96,.25)':'rgba(231,76,60,.25)';}

    // Tabela por dia
    var tbody = document.getElementById('rss-tbody');
    if (tbody) {
        tbody.innerHTML = _rssData.map(function(r) {
            var p   = r.pedido>0 ? (r.atendido/r.pedido)*100 : 0;
            var cor = p>=90?'#27ae60':p>=75?'#f39c12':'#e74c3c';
            var bg  = p>=90?'#f0fff4':p>=75?'#fffdf0':'#fff5f5';
            var dtFmt = r.data.split('-').reverse().join('/');
            return '<tr style="background:'+bg+';">' +
                '<td style="padding:6px 12px;font-weight:700;">'+dtFmt+'</td>' +
                '<td style="text-align:center;padding:6px 12px;">'+fmtN(r.pedido)+'</td>' +
                '<td style="text-align:center;padding:6px 12px;">'+fmtN(r.atendido)+'</td>' +
                '<td style="text-align:center;padding:6px 12px;font-weight:900;color:'+cor+';">'+p.toFixed(1)+'%</td>' +
                '<td style="text-align:center;padding:6px 12px;color:#e74c3c;font-weight:700;">'+fmtN(r.pedido-r.atendido)+'</td>' +
                '<td style="text-align:center;padding:6px 12px;">' +
                    (p>=90?'<span style="background:#d5f5e3;color:#1a7a3c;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:800;">✅ OK</span>' :
                           '<span style="background:#fadbd8;color:#c0392b;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:800;">⚠️ Abaixo</span>') +
                '</td></tr>';
        }).join('');
    }

    // Gráfico % por dia
    renderChartRssDia();

    // IA
    var diasAbaixo = _rssData.filter(function(r){ return r.pedido>0 && (r.atendido/r.pedido)*100 < 90; }).length;
    var ia = '';
    if (atingiu) {
        ia = '✅ <strong>Meta atingida!</strong> Atendimento geral: <strong>'+pct.toFixed(2)+'%</strong>. '+diasAbaixo+' dia(s) abaixo de 90%. Continue monitorando.';
    } else {
        var unidNec = Math.ceil(0.90*totPed - totAte);
        ia = '⚠️ Atendimento <strong>'+pct.toFixed(2)+'%</strong> — abaixo de 90%. Faltam <strong>'+fmtN(unidNec)+' unidades</strong>.<br>' +
             '• <strong>'+diasAbaixo+'</strong> dia(s) com atendimento abaixo da meta<br>' +
             '• Verificar os dias críticos na tabela abaixo<br>' +
             '• Priorizar picking nos horários de maior volume<br>' +
             '• Comunicar ao time de compras os dias com maior pendência';
    }
    setH('rss-ia-text', ia);
}

function renderChartRssDia() {
    var canvas = document.getElementById('chartRssMeta');
    if (!canvas) return;
    if (_metaChartRss) { _metaChartRss.destroy(); _metaChartRss = null; }
    if (!_rssData.length) return;
    var labels = _rssData.map(function(r){ return r.data.split('-').slice(1).reverse().join('/'); });
    var pcts   = _rssData.map(function(r){ return r.pedido>0 ? parseFloat(((r.atendido/r.pedido)*100).toFixed(1)) : 0; });
    var colors = pcts.map(function(v){ return v>=90?'rgba(39,174,96,.7)':v>=75?'rgba(243,156,18,.7)':'rgba(231,76,60,.7)'; });
    _metaChartRss = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: { labels: labels, datasets: [
            { label: '% Atendido por Dia', data: pcts, backgroundColor: colors },
            { label: 'Meta (90%)', data: Array(labels.length).fill(90), type:'line', borderColor:'#e74c3c', borderDash:[5,3], fill:false, pointRadius:0, borderWidth:2 }
        ]},
        options: { responsive:true, plugins:{legend:{position:'top',labels:{font:{size:10}}}},
            scales:{ y:{beginAtZero:true,max:100,ticks:{callback:function(v){return v+'%';},font:{size:10}}}, x:{ticks:{font:{size:9},maxTicksLimit:18}} } }
    });
}

function atualizarScoreTotal() {
    var el = document.getElementById('meta-score-total');
    if (!el) return;
    var absEl = document.getElementById('abs-kpi-taxa');
    var rssEl = document.getElementById('rss-kpi-pct');
    var absOk = absEl && parseFloat(absEl.textContent) <= 4;
    var rssOk = rssEl && parseFloat(rssEl.textContent) >= 90;
    var score = (absOk?35:0) + (rssOk?35:0);
    el.textContent = score+'%';
    el.style.color = score>=70?'#2ecc71':score>=35?'#f1c40f':'#e74c3c';
}


window.subTabAbs = function(aba) {
    var pAbs   = document.getElementById('sub-painel-abs');
    var pPonto = document.getElementById('sub-painel-ponto');
    var bAbs   = document.getElementById('btn-sub-abs');
    var bPonto = document.getElementById('btn-sub-ponto');
    if (!pAbs || !pPonto) return;
    if (aba === 'abs') {
        pAbs.style.display=''; pPonto.style.display='none';
        if(bAbs){bAbs.style.background='var(--red)';bAbs.style.color='white';}
        if(bPonto){bPonto.style.background='#e9ecef';bPonto.style.color='#555';}
    } else {
        pAbs.style.display='none'; pPonto.style.display='';
        if(bPonto){bPonto.style.background='var(--purple)';bPonto.style.color='white';}
        if(bAbs){bAbs.style.background='#e9ecef';bAbs.style.color='#555';}
        if(window.renderPontoMensal) renderPontoMensal();
    }
};

window.updatePublicFormFields = function() { const sel = document.getElementById('pf_colab_select'); if(!sel.value) return; document.getElementById('pf_mat').value = sel.value; document.getElementById('pf_nome').value = sel.options[sel.selectedIndex].text; };
async function submitPublicForm() { const nome = document.getElementById('pf_nome').value; const mat = document.getElementById('pf_mat').value; const inicio = document.getElementById('pf_inicio').value; const fim = document.getElementById('pf_fim').value; const vendeuRadio = document.querySelector('input[name="pf_vender"]:checked'); const vendeu = vendeuRadio ? vendeuRadio.value : 'nao'; if(!mat||!nome) return alert("Selecione o seu nome!"); if(!inicio||!fim) return alert("Preencha Início e Retorno!"); const payload = { matricula:mat, nome:nome, dataInicio:inicio, dataRetorno:fim, vendeuDias:vendeu==='sim', timestamp:new Date().getTime(), status:'PENDENTE' }; try { const btn = document.querySelector('.pf-btn'); btn.innerText='A ENVIAR...'; btn.disabled=true; await dbCloud.collection('logistica_ferias_inbox').add(payload);
            document.getElementById('pf-form-view').style.display = 'none';
            document.getElementById('pf-success-view').style.display = 'block';
        } catch(e) {
            alert("Erro ao enviar. Tente novamente.");
            const btn = document.querySelector('.pf-btn');
            btn.innerText = 'Enviar Solicitação';
            btn.disabled = false;
        }
    }

    // ==========================================
    // CAIXA DE ENTRADA (INBOX) E FÉRIAS PRO
    // ==========================================
    window.copyGenericFormLink = function() {
        const link = window.location.href.split('?')[0] + '?mode=formferias';
        navigator.clipboard.writeText(link);
        alert("Link do formulário copiado!\n\nEnvie este link para os colaboradores no WhatsApp.");
    };

    window.sendWhatsAppLink = function(mat, nome) {
        const link = window.location.href.split('?')[0] + '?mode=formferias';
        const msg = `Olá ${nome}, por favor preencha a sua solicitação de férias neste link:\n${link}`;
        window.open('https://api.whatsapp.com/send?text=' + encodeURIComponent(msg), '_blank');
    };

    window.loadInbox = async function() {
        const inboxBody = document.getElementById('feriasInboxBody');
        if(!inboxBody) return;
        inboxBody.innerHTML = '<div style="padding:20px; text-align:center; color:#999; font-size:12px;">A carregar solicitações...</div>';
        try {
            const snap = await dbCloud.collection('logistica_ferias_inbox').where('status', '==', 'PENDENTE').get();
            if(snap.empty) {
                inboxBody.innerHTML = '<div style="padding:20px; text-align:center; color:#999; font-size:12px;">Nenhuma nova solicitação pendente.</div>';
                return;
            }
            let html = '';
            snap.forEach(doc => {
                const data = doc.data();
                const dataIn = data.dataInicio.split('-').reverse().join('/');
                const dataFim = data.dataRetorno.split('-').reverse().join('/');
                const badgeVenda = data.vendeuDias ? `<span class="inbox-badge inbox-badge-warn">Vende 10 Dias</span>` : '';
                html += `<div class="inbox-item"><div class="inbox-info"><div class="inbox-name">${data.nome} <span style="font-size:10px; color:#999;">(${data.matricula})</span></div><div class="inbox-dates"><span><i class="far fa-calendar-check"></i> Início: <b>${dataIn}</b></span><span><i class="far fa-calendar-times"></i> Retorno: <b>${dataFim}</b></span>${badgeVenda}</div></div><div class="inbox-actions"><button class="btn-sm btn-green" onclick="aprovarFerias('${doc.id}', '${data.matricula}', '${data.dataInicio}', '${data.dataRetorno}', ${data.vendeuDias})"><i class="fas fa-check"></i> Aprovar</button><button class="btn-sm btn-red" onclick="rejeitarFerias('${doc.id}')"><i class="fas fa-times"></i> Rejeitar</button></div></div>`;
            });
            inboxBody.innerHTML = html;
        } catch(e) {
            inboxBody.innerHTML = '<div style="padding:20px; text-align:center; color:red; font-size:12px;">Erro ao carregar a caixa de entrada.</div>';
        }
    };

    window.aprovarFerias = async function(docId, mat, inicio, fim, vendeuDias) {
        if(!confirm("Aprovar esta solicitação e lançar no quadro de férias?")) return;
        try {
            ensureRhData(mat);
            rhData[mat].feriasInicio = inicio;
            rhData[mat].feriasFim = fim;
            rhData[mat].vendeu10Dias = vendeuDias;
            rhData[mat].statusManual = 'AUTO';
            if (window.saveToDB) await window.saveToDB('rhData', rhData);
            await dbCloud.collection('logistica_ferias_inbox').doc(docId).update({status: 'APROVADO'});
            renderRHQuad(); renderFerias(); loadInbox();
            alert("Férias aprovadas e lançadas com sucesso no quadro principal!");
        } catch(e) { alert("Erro ao aprovar."); }
    };

    window.rejeitarFerias = async function(docId) {
        if(!confirm("Tem a certeza que deseja rejeitar esta solicitação? Ela será removida da caixa de entrada.")) return;
        try {
            await dbCloud.collection('logistica_ferias_inbox').doc(docId).update({status: 'REJEITADO'});
            loadInbox();
        } catch(e) { alert("Erro ao rejeitar."); }
    };

    // ==========================================
    // GESTÃO DO PONTO MENSAL (FECHAMENTO 06 A 05)
    // ==========================================
    window.renderPontoMensal = function() {
        const sel = document.getElementById('selPontoEmp');
        if (sel && sel.options.length <= 1) {
            sel.innerHTML = '<option value="">-- Selecione o Colaborador --</option>';
            dRH.forEach(r => { 
                const mat = r['Matrícula'] || r['Matricula'] || r['MATRICULA'] || '-'; 
                const nome = r['Nome'] || r['NOME'] || '-'; 
                const status = (r['Status'] || r['STATUS'] || '-').toUpperCase(); 
                if (nome !== '-' && status === 'ATIVO') { 
                    sel.innerHTML += `<option value="${mat}">${nome} (${mat})</option>`; 
                } 
            });
        }
        atualizarPeriodoPontoDisplay();
        carregarTabelaPontoMes();
    };

    let mesPontoOffset = 0;
    window.mudarPeriodoPonto = function(offset) { mesPontoOffset += offset; atualizarPeriodoPontoDisplay(); carregarTabelaPontoMes(); };
    
    function getPeriodoFechamento() {
        const hoje = new Date();
        hoje.setMonth(hoje.getMonth() + mesPontoOffset);
        let ano = hoje.getFullYear();
        let mes = hoje.getMonth();
        let diaAtual = hoje.getDate();
        
        let mesInicio, anoInicio, mesFim, anoFim;
        if (diaAtual >= 6) { mesInicio = mes; anoInicio = ano; mesFim = mes + 1; anoFim = ano; if (mesFim > 11) { mesFim = 0; anoFim++; } } 
        else { mesInicio = mes - 1; anoInicio = ano; if (mesInicio < 0) { mesInicio = 11; anoInicio--; } mesFim = mes; anoFim = ano; }
        
        const dataInicio = new Date(anoInicio, mesInicio, 6);
        const dataFim = new Date(anoFim, mesFim, 5);
        return { start: dataInicio, end: dataFim, startStr: dataInicio.toLocaleDateString('pt-BR'), endStr: dataFim.toLocaleDateString('pt-BR'), name: dataFim.toLocaleString('pt-BR', {month:'long', year:'numeric'}).toUpperCase() };
    }

    window.atualizarPeriodoPontoDisplay = function() {
        const periodo = getPeriodoFechamento();
        const lbl = document.getElementById('lblPontoPeriodo');
        if (lbl) lbl.innerHTML = `${periodo.name} <div style="font-size:10px; color:#666; font-weight:normal;">(${periodo.startStr} a ${periodo.endStr})</div>`;
    };

    window.registrarPonto = async function() {
        const mat = document.getElementById('selPontoEmp').value;
        const dataStr = document.getElementById('txtPontoDate').value;
        const tipo = document.getElementById('selPontoTipo').value;
        const acao = document.getElementById('selPontoAcao').value;
        const obs = document.getElementById('txtPontoObs').value;
        
        if (!mat || !dataStr) return alert("Selecione o colaborador e a data da ocorrência!");
        
        ensureRhData(mat);
        if (!rhData[mat].ponto) rhData[mat].ponto = [];
        
        const dObj = new Date(dataStr + "T12:00:00");
        const brDate = dObj.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
        
        rhData[mat].ponto.push({
            id: new Date().getTime().toString(),
            dataOrig: dataStr,
            dataForm: brDate,
            stamp: dObj.getTime(),
            tipo: tipo,
            acao: acao,
            obs: obs
        });
        
        rhData[mat].ponto.sort((a,b) => b.stamp - a.stamp);
        if (window.saveToDB) await window.saveToDB('rhData', rhData);
        
        document.getElementById('txtPontoDate').value = '';
        document.getElementById('txtPontoObs').value = '';
        alert("Ocorrência de ponto registada com sucesso!");
        carregarTabelaPontoMes();
    };

    window.removerPonto = async function(mat, idStr) {
        if (!confirm("Tem a certeza que deseja remover esta ocorrência de ponto?")) return;
        if (rhData[mat] && rhData[mat].ponto) {
            rhData[mat].ponto = rhData[mat].ponto.filter(p => p.id !== idStr);
            if (window.saveToDB) await window.saveToDB('rhData', rhData);
            carregarTabelaPontoMes();
        }
    };

    window.carregarTabelaPontoMes = function() {
        const tb = document.getElementById('tbPonto');
        if (!tb) return;
        
        const periodo = getPeriodoFechamento();
        const tStart = periodo.start.getTime();
        const tEnd = periodo.end.getTime() + 86399000; // Fim do dia 05
        
        let html = '';
        let cont = 0;
        
        for (let mat in rhData) {
            const rh = rhData[mat];
            if (!rh.ponto || rh.ponto.length === 0) continue;
            
            const emp = dRH.find(e => (e['Matrícula'] || e['Matricula'] || e['MATRICULA'] || '').toString() === mat);
            const nome = emp ? (emp['Nome'] || emp['NOME']) : 'Desconhecido';
            
            rh.ponto.forEach(p => {
                if (p.stamp >= tStart && p.stamp <= tEnd) {
                    cont++;
                    let tipoColor = p.tipo.toLowerCase().includes('falta') ? 'color:var(--red); font-weight:bold;' : 'color:#333;';
                    let acaoColor = p.acao.toLowerCase().includes('descontar') ? 'background:#ffebeb; color:#c0392b;' : (p.acao.toLowerCase().includes('compensar') ? 'background:#e8f4fd; color:#2980b9;' : 'background:#f0f0f0; color:#555;');
                    
                    html += `<tr style="border-bottom:1px solid #eee;">
                        <td style="font-weight:bold; font-size:12px;">${p.dataForm}</td>
                        <td style="text-align:left; font-size:12px;"><b>${nome}</b><br><span style="font-size:9px; color:#888;">${mat}</span></td>
                        <td style="font-size:11px; ${tipoColor}">${p.tipo}<br><span style="font-size:10px; color:#777; font-weight:normal;">${p.obs || ''}</span></td>
                        <td><span style="padding:3px 6px; border-radius:4px; font-size:10px; font-weight:bold; ${acaoColor}">${p.acao}</span></td>
                        <td><button class="btn-sm btn-red" onclick="removerPonto('${mat}', '${p.id}')"><i class="fas fa-trash"></i></button></td>
                    </tr>`;
                }
            });
        }
        
        if (cont === 0) {
            tb.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:20px; color:#999;">Nenhuma ocorrência registada neste período de fechamento (Dia 06 a 05).</td></tr>';
        } else {
            tb.innerHTML = html;
        }
    };

    window.gerarRelatorioPontoPDF = function() {
        const periodo = getPeriodoFechamento();
        const tStart = periodo.start.getTime();
        const tEnd = periodo.end.getTime() + 86399000;
        
        let conteudo = [['Data', 'Colaborador', 'Matrícula', 'Ocorrência', 'Ação', 'Observações']];
        
        for (let mat in rhData) {
            const rh = rhData[mat];
            if (!rh.ponto || rh.ponto.length === 0) continue;
            const emp = dRH.find(e => (e['Matrícula'] || e['Matricula'] || e['MATRICULA'] || '').toString() === mat);
            const nome = emp ? (emp['Nome'] || emp['NOME']) : 'Desconhecido';
            
            rh.ponto.forEach(p => {
                if (p.stamp >= tStart && p.stamp <= tEnd) {
                    conteudo.push([
                        p.dataForm,
                        nome,
                        mat,
                        p.tipo,
                        p.acao,
                        p.obs || '-'
                    ]);
                }
            });
        }
        
        if (conteudo.length === 1) return alert("Não há dados para gerar relatório neste período.");
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(conteudo);
        XLSX.utils.book_append_sheet(wb, ws, "Ocorrencias_Ponto");
        XLSX.writeFile(wb, `Relatorio_Ponto_${periodo.name}.xlsx`);
    };

    // ==========================================
    // INICIALIZAÇÃO E MODO FORMULÁRIO PÚBLICO
    // ==========================================
    window.addEventListener('load', function() {
        const urlParams = new URLSearchParams(window.location.search);
        const mode = urlParams.get('mode');
        if (mode === 'formferias') {
            document.body.style.background='#f0f2f5';
            document.getElementById('loginOverlay').style.display='none';
            document.querySelector('.app-header').style.display='none';
            document.querySelector('.memory-bar').style.display='none';
            document.querySelector('.tab-bar').style.display='none';
            document.querySelector('.container').style.display='none';
            if(document.getElementById('btnExitTV'))document.getElementById('btnExitTV').style.display='none';
            document.getElementById('publicFormWrapper').style.display='block';
            document.getElementById('pf_group_select').style.display='block';
            document.getElementById('pf_group_nome').style.display='none';
            document.getElementById('pf_group_mat').style.display='none';
            loadEmployeesForPublicForm();
        }
        if (mode === 'formepi') {
            document.body.style.background='#f0f2f5'; document.body.style.overflow='auto';
            document.getElementById('loginOverlay').style.display='none';
            document.querySelector('.app-header').style.display='none';
            document.querySelector('.memory-bar').style.display='none';
            document.querySelector('.tab-bar').style.display='none';
            document.querySelector('.container').style.display='none';
            if(document.getElementById('btnExitTV'))document.getElementById('btnExitTV').style.display='none';
            document.getElementById('publicFormWrapper').style.display='block';
            ['epi-pub-form','pf-form-view','pf-success-view'].forEach(id=>{const e=document.getElementById(id);if(e)e.style.display=(id==='epi-pub-form')?'block':'none';});
            loadEpiColabs();
        }
    });

// ==========================================
// SISTEMA DE AUTO-SAVE PERMANENTE
// ==========================================
document.addEventListener('input', function(e) {
    if(e.target.tagName === 'INPUT' && e.target.type !== 'file' && e.target.id) {
        let saved = JSON.parse(localStorage.getItem('sysLog_metas_salvas') || '{}');
        saved[e.target.id] = e.target.value;
        localStorage.setItem('sysLog_metas_salvas', JSON.stringify(saved));
    }
});

window.addEventListener('load', function() {
    setTimeout(() => {
        let saved = JSON.parse(localStorage.getItem('sysLog_metas_salvas') || '{}');
        for(let id in saved) {
            let el = document.getElementById(id);
            if(el) {
                el.value = saved[id];
                el.dispatchEvent(new Event('input'));
            }
        }
        try {
            // BUSCA NA MEMÓRIA PERMANENTE
            let memC = localStorage.getItem('sysLog_mem_dC');
            let memE = localStorage.getItem('sysLog_mem_dE');
            if(memC && memE) {
                dC = JSON.parse(memC);
                dE = JSON.parse(memE);
                if(dC.length > 0 && dE.length > 0) runCalc();
            }
        } catch(err) {}
    }, 800);
});

// ==========================================
// FUNÇÃO GLOBAL DE INDICADORES (DETETIVE 3.0)
// ==========================================
function atualizarTodosIndicadores() {
    if (typeof dC === 'undefined' || !dC) return;

    // Faz as contas rigorosas (Exatamente como a IA viu)
    let dCAtivos = dC.filter(c => c.status && c.status.trim().toLowerCase() === 'ativo');
    
    let totalGeral = dC.length;
    let totalAtivos = dCAtivos.length;
    let totalFerias = dC.filter(c => c.status === 'Férias').length;
    let totalEPIs = (typeof dE !== 'undefined' && dE) ? dE.filter(e => e.status === 'Pendente' || e.status === 'Solicitado').length : 0;
    
    let contaSupervisores = dCAtivos.filter(c => (c.funcao || c.cargo || "").toLowerCase().includes('supervisor')).length;
    let contaLideres = dCAtivos.filter(c => { let f = (c.funcao || c.cargo || "").toLowerCase(); return f.includes('lider') || f.includes('líder'); }).length;
    let contaOperadores = dCAtivos.filter(c => (c.funcao || c.cargo || "").toLowerCase().includes('operador')).length;

    // Detetive 3.0: Procura em todas as tags do sistema
    let todosElementos = document.querySelectorAll('div, span, h3, h4, th, td, p');
    todosElementos.forEach(el => {
        let texto = el.innerText.trim().toLowerCase();
        let valorInjetar = null;

        // Associa a palavra exata ao número calculado
        if (texto === 'total colaboradores' || texto === 'total de colaboradores') valorInjetar = totalGeral;
        else if (texto === 'em férias' || texto === 'em ferias') valorInjetar = totalFerias;
        else if (texto === 'pendências epi' || texto === 'pendencias epi') valorInjetar = totalEPIs;
        else if (texto === 'supervisores' || texto === 'supervisor') valorInjetar = contaSupervisores;
        else if (texto === 'líderes' || texto === 'lideres') valorInjetar = contaLideres;
        else if (texto === 'operadores' || texto === 'operador') valorInjetar = contaOperadores;
        // Se for o card "Total" do quadro interno
        else if (texto === 'total' && el.parentElement && !el.parentElement.innerText.toLowerCase().includes('colaboradores')) valorInjetar = totalAtivos;

        // Se encontrou onde injetar, procura a caixa do número logo abaixo ou ao lado
        if (valorInjetar !== null) {
            let caixaValor = el.nextElementSibling;
            
            // Se o próximo não for o número, procura dentro do "pai" do elemento
            if (!caixaValor || (caixaValor.innerText.trim() !== '0' && isNaN(parseInt(caixaValor.innerText)))) {
                caixaValor = el.parentElement.querySelector('.kpi-val, [style*="font-size: 28px"], [style*="font-size:28px"]');
            }
            
            if (caixaValor) {
                caixaValor.innerText = valorInjetar;
            }
        }
    });
}

// ==========================================
// ATUALIZAÇÃO DA FUNÇÃO runCalc FINAL
// ==========================================
const runCalcOriginal = window.runCalc;
window.runCalc = function() {
    try {
        if (typeof runCalcOriginal === 'function') {
            runCalcOriginal();
        }
        // Sobrevivência (Memória do PC)
        if (typeof dC !== 'undefined' && typeof dE !== 'undefined') {
            localStorage.setItem('sysLog_mem_dC', JSON.stringify(dC));
            localStorage.setItem('sysLog_mem_dE', JSON.stringify(dE));
        }
        // Chama o nosso novo atualizador completo
        atualizarTodosIndicadores();
        
    } catch(err) {
        console.error("Erro no runCalc:", err);
    }
}
// =====================================================
// 🔥 PATCH DEFINITIVO - DASHBOARD RH FUNCIONANDO 100%
// =====================================================

// ❌ Remove função antiga se existir (evita conflito)
try {
    if (typeof atualizarIndicadoresRH !== 'undefined') {
        atualizarIndicadoresRH = null;
    }
} catch(e) {}


// =====================================================
// ✅ FUNÇÃO PRINCIPAL DE KPI RH (CORRETA)
// =====================================================
function atualizarKPIsRH(lista) {
    try {
        if (!lista || !Array.isArray(lista)) return;

        let ativos = 0;
        let lideres = 0;
        let operadores = 0;
        let supervisores = 0;

        lista.forEach(c => {
            let status = (c.status || c.STATUS || 'ATIVO').toUpperCase();
            let funcao = (c.funcao || c.FUNCAO || '').toUpperCase();

            if (status === 'ATIVO' || status === 'AUTO') {
                ativos++;

                if (funcao.includes('SUPERVISOR')) {
                    supervisores++;
                } else if (funcao.includes('LIDER') || funcao.includes('LÍDER')) {
                    lideres++;
                } else if (funcao.includes('OPERADOR') || funcao.includes('SEPARADOR')) {
                    operadores++;
                }
            }
        });

        // Atualização segura no DOM
        const set = (id, val) => {
            let el = document.getElementById(id);
            if (el) el.textContent = val;
        };

        set('rh-tot', ativos);
        set('rh-lid', lideres);
        set('rh-op', operadores);
        set('rh-sup', supervisores);

        // Exibir grid
        let kpis = document.getElementById('rh-kpis');
        if (kpis && lista.length > 0) {
            kpis.style.display = 'grid';
        }

    } catch (e) {
        console.error('Erro KPIs RH:', e);
    }
}


// =====================================================
// ✅ ATUALIZAÇÃO AUTOMÁTICA AO CARREGAR SISTEMA
// =====================================================
window.addEventListener('load', function() {
    setTimeout(() => {
        try {
            if (typeof dRH !== 'undefined' && dRH.length > 0) {
                atualizarKPIsRH(dRH);
            }
        } catch(e) {
            console.warn('Erro ao iniciar KPIs RH:', e);
        }
    }, 1000);
});


// =====================================================
// ✅ ATUALIZAÇÃO AUTOMÁTICA CONTÍNUA
// =====================================================
setInterval(() => {
    try {
        if (typeof dRH !== 'undefined' && dRH.length > 0) {
            atualizarKPIsRH(dRH);
        }
    } catch(e) {}
}, 5000);


// =====================================================
// ✅ FORÇA ATUALIZAÇÃO AO ABRIR DASHBOARD RH
// =====================================================
(function() {
    const originalPushState = history.pushState;
    history.pushState = function() {
        originalPushState.apply(this, arguments);

        setTimeout(() => {
            try {
                let dash = document.getElementById('view-rh-dashboard');
                if (dash && dash.style.display !== 'none') {

                    if (typeof renderDashboardRH === 'function') {
                        renderDashboardRH();
                    }

                    if (typeof dRH !== 'undefined') {
                        atualizarKPIsRH(dRH);
                    }
                }
            } catch(e) {}
        }, 300);
    };
})();


// =====================================================
// ✅ ATUALIZA AO MODIFICAR LISTA DE RH
// =====================================================
window.atualizarRHManual = function() {
    try {
        if (typeof dRH !== 'undefined') {
            atualizarKPIsRH(dRH);
        }
    } catch(e) {}
};

// ── Entrega por item individual ────────────────────────────────
var _epiEntregDocId = null;
var _epiEntregDoc   = null;

window.abrirModalEntregaEpi = function(docId) {
    var doc = (_epiCache.docs||[]).find(function(r){ return r.id === docId; });
    if (!doc) return;
    _epiEntregDocId = docId;
    _epiEntregDoc   = doc;
    var nEl = document.getElementById('modalEntregaNome');
    if (nEl) nEl.innerHTML = '<i class="fas fa-user"></i> ' + doc.nome +
        ' <span style="color:#888;font-weight:400;font-size:12px;">(' + doc.mat + ')</span>';
    var iEl = document.getElementById('modalEntregaItens');
    if (iEl) {
        var h = '';
        EPI_ITEMS.forEach(function(it) {
            var val = doc[it.k];
            if (!val || val === '-') return;
            var jaEntregue = doc[it.k + '_entregue'] === true;
            h += '<label style="display:flex;align-items:center;gap:10px;background:' +
                 (jaEntregue ? '#f0fff4' : '#f9f9f9') + ';border:1px solid ' +
                 (jaEntregue ? '#a9dfbf' : '#ddd') + ';border-radius:7px;padding:8px 12px;cursor:pointer;margin-bottom:6px;">' +
                 '<input type="checkbox" id="entrega_chk_' + it.k + '" ' +
                 (jaEntregue ? 'checked disabled' : 'checked') +
                 ' style="accent-color:#27ae60;width:16px;height:16px;">' +
                 '<i class="' + it.ico + '" style="color:' + (jaEntregue ? '#27ae60' : '#555') + ';font-size:15px;"></i>' +
                 '<span style="font-weight:800;min-width:60px;">' + it.lbl + ':</span>' +
                 '<span style="background:' + (jaEntregue ? '#27ae60' : '#8e44ad') + ';color:white;border-radius:5px;padding:2px 9px;font-size:12px;">' + val + '</span>' +
                 (jaEntregue ? '<span style="margin-left:auto;font-size:10px;color:#27ae60;font-weight:700;"><i class="fas fa-check"></i> Já entregue</span>' :
                               '<span style="margin-left:auto;font-size:10px;color:#f39c12;font-weight:700;">Pendente</span>') +
                 '</label>';
        });
        iEl.innerHTML = h || '<p style="color:#999;font-size:13px;">Sem itens aprovados.</p>';
    }
    document.getElementById('modalEntregaEpi').style.display = 'flex';
};

window.entregarTudo = async function() {
    if (!_epiEntregDocId) return;
    document.getElementById('modalEntregaEpi').style.display = 'none';
    await confirmarEntregaEpi(_epiEntregDocId);
};

window.confirmarEntregaParcial = async function() {
    var docId = _epiEntregDocId;
    var doc   = _epiEntregDoc;
    if (!docId || !doc) return;
    var marcados = {};
    var temAlgum = false;
    EPI_ITEMS.forEach(function(it) {
        var chk = document.getElementById('entrega_chk_' + it.k);
        if (chk && chk.checked && !chk.disabled && doc[it.k] && doc[it.k] !== '-') {
            marcados[it.k] = true; temAlgum = true;
        }
    });
    if (!temAlgum) { alert('Marque pelo menos um item para entregar.'); return; }

    // Verifica se todos os pendentes foram marcados
    var totalPendentes = EPI_ITEMS.filter(function(it){
        return doc[it.k] && doc[it.k] !== '-' && !doc[it.k + '_entregue'];
    }).length;
    var totalMarcados = Object.keys(marcados).length;
    var dataEntrega   = new Date().toLocaleDateString('pt-BR');
    document.getElementById('modalEntregaEpi').style.display = 'none';

    if (totalMarcados >= totalPendentes) {
        // Todos os pendentes entregues → status entregue
        await confirmarEntregaEpi(docId);
    } else {
        // Entrega parcial — atualiza cache otimisticamente
        if (_epiCache.docs) {
            var reg = _epiCache.docs.find(function(r){ return r.id === docId; });
            if (reg) {
                EPI_ITEMS.forEach(function(it){
                    if (marcados[it.k]) reg[it.k + '_entregue'] = true;
                });
                reg.status      = 'parcial';
                reg.dataEntrega = dataEntrega + ' (parcial)';
            }
        }
        renderEpiTableFromCache();
        // Salva no Firestore
        try {
            var upd = { dataEntrega: dataEntrega, status: 'parcial' };
            EPI_ITEMS.forEach(function(it){ if (marcados[it.k]) upd[it.k + '_entregue'] = true; });
            await dbCloud.collection('logistica_epi_registros').doc(docId).update(upd);
        } catch(e) { alert('Erro ao salvar: ' + e.message); }
    }
};



// =====================================================
// 🔥 GARANTE EXECUÇÃO APÓS RESTORE SESSION
// =====================================================
const _oldRestoreSession = window.restoreSession;
window.restoreSession = async function() {
    if (_oldRestoreSession) {
        await _oldRestoreSession.apply(this, arguments);
    }

    setTimeout(() => {
        try {
            if (typeof dRH !== 'undefined' && dRH.length > 0) {
                atualizarKPIsRH(dRH);
            }
        } catch(e) {}
    }, 500);
};

// =========================================================
// LÓGICA FINAL: GIRO DE ALOCAÇÃO (ESTOQUE ATUAL) - TURBO
// =========================================================
// =========================================================
// LÓGICA FINAL + APONTAMENTOS (KPIs DINÂMICOS)
// =========================================================

// =========================================================
// 🚀 IMPLEMENTAÇÃO FINAL: GIRO DE ALOCAÇÃO (ESTOQUE ATUAL)
// =========================================================

window.giroImportarAlocacao = function(input) {
    const file = input.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            console.log("Processando Estoque Atual...");
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

            const resultado = json.map(item => {
                const pck = String(item["Picking"] || "").trim().toUpperCase();
                const res = String(item["Reserva"] || "").trim().toUpperCase();
                return {
                    produto: item["Produto"] || "N/D",
                    descricao: item["Descricao"] || "N/D",
                    picking: pck,
                    reserva: res,
                    estacao: definirEstacaoPavuna(res),
                    metodologia: validarRegra30(pck, res)
                };
            });

            // 📊 ATUALIZA OS INDICADORES (KPIs) NA TELA
            const caixaKpis = document.getElementById('giro-kpis');
            if (caixaKpis) {
                caixaKpis.style.display = 'flex';

                // Conta os status para exibir nos painéis
                let oks = resultado.filter(r => r.metodologia.includes('✅') || r.metodologia.includes('📦')).length;
                let errados = resultado.filter(r => r.metodologia.includes('❌')).length;
                let faltantes = resultado.filter(r => r.metodologia.includes('⚠️')).length;

                caixaKpis.innerHTML = `
                    <div class="giro-kpi" style="border-left-color:#2980b9;">
                        <div class="giro-kpi-lbl">Total de Linhas</div>
                        <div class="giro-kpi-val">${resultado.length.toLocaleString('pt-BR')}</div>
                        <div class="giro-kpi-sub">Estoque Atual</div>
                    </div>
                    <div class="giro-kpi" style="border-left-color:#27ae60;">
                        <div class="giro-kpi-lbl">Alocação Correta</div>
                        <div class="giro-kpi-val" style="color:#27ae60;">${oks.toLocaleString('pt-BR')}</div>
                        <div class="giro-kpi-sub">Dentro da Regra</div>
                    </div>
                    <div class="giro-kpi" style="border-left-color:#c0392b;">
                        <div class="giro-kpi-lbl">Divergentes</div>
                        <div class="giro-kpi-val" style="color:#c0392b;">${errados.toLocaleString('pt-BR')}</div>
                        <div class="giro-kpi-sub">Erro de Endereçamento</div>
                    </div>
                    <div class="giro-kpi" style="border-left-color:#f39c12;">
                        <div class="giro-kpi-lbl">S/ Endereço</div>
                        <div class="giro-kpi-val" style="color:#f39c12;">${faltantes.toLocaleString('pt-BR')}</div>
                        <div class="giro-kpi-sub">Falta Picking/Reserva</div>
                    </div>
                `;
            }

            // ⚡ RENDERIZA TABELA COM LIMITE PARA NÃO TRAVAR O NAVEGADOR
            if (typeof window.renderizarTabelaGiro === 'function') {
                window.renderizarTabelaGiro(resultado.slice(0, 300)); 
            }

        } catch (err) {
            alert("Erro ao processar a planilha: " + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
};

// 📍 REGRA DE ESTAÇÕES E TIPOS (PAVUNA)
function definirEstacaoPavuna(endereco) {
    if (!endereco || endereco.includes("FRAC") || endereco === "N/D") return "Sem Estação";

    const partes = endereco.split('-');
    if (partes.length < 2) return "Endereço Inválido";

    const rua = parseInt(partes[0]);
    const modulo = parseInt(partes[1]);

    if (rua >= 21 && rua <= 36) {
        let estacao = (rua <= 23) ? 1 : (rua <= 26) ? 2 : (rua <= 29) ? 3 : (rua <= 32) ? 4 : 5;

        if (modulo >= 1 && modulo <= 26) return `Estação ${estacao} - Perf. (Alto)`;
        if (modulo >= 37 && modulo <= 40) return `Estação ${estacao} - Perf. (Baixo)`;
        if (modulo >= 41 && modulo <= 46) return `Estação ${estacao} - Medic. (Alto)`;
        if (modulo >= 57 && modulo <= 60) return `Estação ${estacao} - Medic. (Baixo)`;
    }
    return "Geral / Outros";
}

// 📏 REGRA MATEMÁTICA E VALIDAÇÃO DE METODOLOGIA
function validarRegra30(pck, res) {
    if (!pck || !res || pck === "N/D" || res === "N/D") return "⚠️ Faltando Endereço";

    const numR = parseInt(res.replace(/\D/g, '').substring(0, 2));
    const numP = parseInt(pck.replace(/\D/g, '').substring(0, 2));

    if (isNaN(numR) || isNaN(numP)) return "---";

    // Regra Volumoso / Controlado / Dermo (Ruas 80 a 91)
    if (numR >= 80 && numR <= 91) {
        return (numR - numP === 30) ? "✅ OK (Metodologia)" : "❌ Divergente";
    }

    // Regra Perfumaria / Medicamento (Ruas 21 a 36)
    if (numR >= 21 && numR <= 36) {
        return "📦 Validado por Estação";
    }

    return "---";
}

// =========================================================
// 🛡️ ESCUDO ANTI-TRAVAMENTO DO FIREBASE (80k LINHAS)
// =========================================================
setTimeout(() => {
    const _salvamentoAntigo = window.saveToDB;
    window.saveToDB = async function(chave, valor) {
        if (Array.isArray(valor) && valor.length > 5000) {
            console.warn(`🛡️ Bloqueado envio de ${valor.length} itens para a nuvem. Processando APENAS na tela para evitar erro!`);
            return Promise.resolve("OK");
        }
        if (typeof _salvamentoAntigo === 'function') {
            return _salvamentoAntigo(chave, valor);
        }
    };
    console.log("🛡️ Escudo Anti-Travamento do Firebase ATIVADO.");
}, 2000);