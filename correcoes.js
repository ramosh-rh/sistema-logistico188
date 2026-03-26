// ====================================================
// ARQUIVO DE CORREÇÕES (V8.0 - BASEADO NO SEU DIAGNÓSTICO!)
// ====================================================

// 1. A FUNÇÃO CORRETA QUE ESCREVE NA TELA (Fallback)
if (typeof window.setDRH === 'undefined') {
    window.setDRH = function(id, valor) {
        let el = document.getElementById(id);
        if (el) { el.innerText = valor; }
    };
}

// 2. GRÁFICOS FANTASMAS (Sem quebrar o sistema)
const fakeChart = { destroy: function(){}, update: function(){} };
if (typeof window.renderChartAbsMotivo === 'undefined') window.renderChartAbsMotivo = function() { return fakeChart; };
if (typeof window.renderChartAbsSetor === 'undefined') window.renderChartAbsSetor = function() { return fakeChart; };
if (typeof window.renderChartTurnover === 'undefined') window.renderChartTurnover = function() { return fakeChart; };
if (typeof window.renderChartPonto === 'undefined') window.renderChartPonto = function() { return fakeChart; };
if (typeof window.renderChartHoraExtra === 'undefined') window.renderChartHoraExtra = function() { return fakeChart; };

// 3. FREIO DO FIREBASE (Protege contra bloqueios)
if (typeof window.autoSaveData === 'function' && !window.autoSaveData_protegido) {
    const fOriginal = window.autoSaveData;
    let timer = null;
    window.autoSaveData = function(...args) {
        clearTimeout(timer);
        timer = setTimeout(() => { fOriginal.apply(this, args); }, 2500); 
    };
    window.autoSaveData_protegido = true;
}

// 4. MODAL EPI E QUADRO RH (Injeção Visual)
window.addEventListener('load', function() {
    // Injeta Modal EPI
    if (!document.getElementById('modalEntregaEpi')) {
        let divModal = document.createElement('div');
        divModal.innerHTML = `
        <div id="modalEntregaEpi" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999;align-items:center;justify-content:center;">
            <div style="background:#fff;width:90%;max-width:400px;border-radius:12px;padding:20px;box-shadow:0 10px 30px rgba(0,0,0,0.2);">
                <div style="display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid #27ae60;padding-bottom:10px;margin-bottom:15px;">
                    <h3 style="margin:0;color:#27ae60;"><i class="fas fa-box-open"></i> Entregar EPI</h3>
                    <button onclick="document.getElementById('modalEntregaEpi').style.display='none'" style="background:none;border:none;font-size:24px;cursor:pointer;">&times;</button>
                </div>
                <div id="modalEntregaEpiNome" style="background:#e8f8f5;border-radius:8px;padding:10px;margin-bottom:12px;font-weight:bold;color:#117a65;"></div>
                <div id="modalEntregaEpiItens" style="margin-bottom:16px;"></div>
                <input type="hidden" id="modalEntregaEpiId">
                <div style="display:flex;gap:10px;border-top:1px solid #eee;padding-top:14px;">
                    <button onclick="confirmarEntregaEpi('total')" style="flex:1;padding:10px;background:#27ae60;color:white;border:none;border-radius:6px;font-weight:bold;cursor:pointer;">Entregar Tudo</button>
                    <button onclick="confirmarEntregaEpi('parcial')" style="flex:1;padding:10px;background:#f39c12;color:white;border:none;border-radius:6px;font-weight:bold;cursor:pointer;">Entrega Parcial</button>
                </div>
            </div>
        </div>`;
        document.body.appendChild(divModal);
    }

    // Injeta Quadro Colaboradores
    let abaColab = document.getElementById('view-rh-colaboradores');
    if (abaColab && !document.getElementById('quadro-info-display')) {
        let quadroInfo = document.createElement('div');
        quadroInfo.id = 'quadro-info-display';
        quadroInfo.style.cssText = 'margin-bottom: 20px; font-size: 16px; background: #e8f8f5; padding: 15px; border-radius: 8px; border-left: 5px solid #117a65;';
        abaColab.insertBefore(quadroInfo, abaColab.children[1] || abaColab.firstChild);
    }
});

// 5. FUNÇÕES EPI
window.abrirModalEntregaEpi = function(docId) {
    try {
        let epi = (typeof dE !== 'undefined' ? dE : []).find(e => String(e.id) === String(docId));
        if(!epi) return;
        document.getElementById('modalEntregaEpiId').value = docId;
        document.getElementById('modalEntregaEpiNome').innerText = "Colaborador: " + (epi.nome || "");
        let htmlItens = "";
        let itensArray = epi.itens ? epi.itens.split(',').map(i => i.trim()).filter(i => i) : [];
        itensArray.forEach(item => {
            let jaEntregue = item.includes('[ENTREGUE]');
            let nomeItem = item.replace('[ENTREGUE]', '').trim();
            htmlItens += `<label style="display:flex;align-items:center;padding:10px;border:1px solid #eee;border-radius:6px;margin-bottom:5px;background:${jaEntregue ? '#f9f9f9' : '#fff'}">
                <input type="checkbox" class="chk-entrega-epi" value="${nomeItem}" ${jaEntregue ? 'checked disabled' : ''} style="margin-right:10px;transform:scale(1.2);">
                <span style="font-weight:bold;color:${jaEntregue ? '#aaa' : '#333'}; ${jaEntregue ? 'text-decoration:line-through;' : ''}">${nomeItem}</span>
                ${jaEntregue ? '<span style="margin-left:auto;font-size:11px;color:#27ae60;font-weight:bold;">Entregue</span>' : ''}
            </label>`;
        });
        document.getElementById('modalEntregaEpiItens').innerHTML = htmlItens;
        document.getElementById('modalEntregaEpi').style.display = 'flex';
    } catch(e) {}
};

window.confirmarEntregaEpi = function(tipo) {
    try {
        let docId = document.getElementById('modalEntregaEpiId').value;
        let epi = dE.find(e => String(e.id) === String(docId));
        if(!epi) return;
        let chks = document.querySelectorAll('.chk-entrega-epi');
        let itensAt = [], todosEntregues = true;
        chks.forEach(chk => {
            if(chk.disabled || chk.checked || tipo === 'total') itensAt.push(chk.value + " [ENTREGUE]");
            else { itensAt.push(chk.value); todosEntregues = false; }
        });
        epi.itens = itensAt.join(', ');
        epi.status = (tipo === 'total' || todosEntregues) ? 'Entregue' : 'Entrega Parcial';
        document.getElementById('modalEntregaEpi').style.display = 'none';
        if(typeof renderEpi === 'function') renderEpi();
        if(typeof window.saveToDB === 'function') window.saveToDB('epis', docId, epi).catch(e=>console.log("Banco pendente."));
    } catch(e) {}
};

// 6. MOTOR INTELIGENTE
setInterval(() => {
    try {
        let quadro = document.getElementById('quadro-info-display');
        if (quadro && typeof dRH !== 'undefined' && dRH.length > 0) {
            let ativos = dRH.filter(c => c.status && c.status.toLowerCase().includes('ativo'));
            let sups = ativos.filter(c => JSON.stringify(c).toLowerCase().includes('supervis')).length;
            let lids = ativos.filter(c => JSON.stringify(c).toLowerCase().includes('lider') || JSON.stringify(c).toLowerCase().includes('líder')).length;
            let ops = ativos.filter(c => JSON.stringify(c).toLowerCase().includes('operador')).length;
            quadro.innerHTML = `📊 <strong>${ativos.length} ativos</strong> | ${sups} supervisores | ${lids} líderes | ${ops} operadores`;
        }
        document.querySelectorAll('button').forEach(btn => {
            let txt = btn.innerText.trim().toLowerCase();
            let act = btn.getAttribute('onclick') || "";
            if (txt.includes('entregar') && !act.includes('abrirModalEntregaEpi')) {
                let id = act.match(/'([^']+)'/) || act.match(/"([^"]+)"/); 
                if (id && id[1]) btn.setAttribute('onclick', `abrirModalEntregaEpi('${id[1]}')`);
            }
        });
    } catch(e) {}
}, 1000);