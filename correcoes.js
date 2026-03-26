// ====================================================
// ARQUIVO SEPARADO DE CORREÇÕES (100% SEGURO) - V5.0
// ====================================================

// 0. VACINA CONTRA O ERRO DO DASHBOARD ("setDRH is not defined")
if (typeof window.setDRH === 'undefined') {
    window.setDRH = function(data) {
        if (data) window.dRH = data;
    };
}

// 1. FREIO DE EMERGÊNCIA DO FIREBASE (Protege contra bloqueios)
if (typeof window.autoSaveData === 'function' && !window.autoSaveData_protegido) {
    const funcaoAutoSaveOriginal = window.autoSaveData;
    let temporizadorAutoSave = null;
    window.autoSaveData = function(...args) {
        clearTimeout(temporizadorAutoSave);
        temporizadorAutoSave = setTimeout(() => { funcaoAutoSaveOriginal.apply(this, args); }, 2500); 
    };
    window.autoSaveData_protegido = true;
}

// 2. INJETA O MODAL E O QUADRO AUTOMATICAMENTE
window.addEventListener('load', function() {
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

    let abaColab = document.getElementById('view-rh-colaboradores');
    if (abaColab && !document.getElementById('quadro-info-display')) {
        let quadroInfo = document.createElement('div');
        quadroInfo.id = 'quadro-info-display';
        quadroInfo.style.cssText = 'margin-bottom: 20px; font-size: 16px; background: #e8f8f5; padding: 15px; border-radius: 8px; border-left: 5px solid #117a65;';
        abaColab.insertBefore(quadroInfo, abaColab.children[1] || abaColab.firstChild);
    }
});

// 3. LÓGICA DA NOVA JANELA DE EPI
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
        if(typeof window.saveToDB === 'function') window.saveToDB('epis', docId, epi).catch(e=>console.log("Aguardando Banco de Dados."));
    } catch(e) {}
};

// 4. MOTOR INTELIGENTE (Painel Dashboard + Quadro + Botões)
setInterval(() => {
    try {
        if (typeof dRH === 'undefined' || !dRH || dRH.length === 0) return;

        // A) Salva o Quadro da Aba de Colaboradores
        let quadro = document.getElementById('quadro-info-display');
        if (quadro) {
            let ativos = dRH.filter(c => c.status && c.status.toLowerCase().includes('ativo'));
            let sups = ativos.filter(c => JSON.stringify(c).toLowerCase().includes('supervis')).length;
            let lids = ativos.filter(c => JSON.stringify(c).toLowerCase().includes('lider') || JSON.stringify(c).toLowerCase().includes('líder')).length;
            let ops = ativos.filter(c => JSON.stringify(c).toLowerCase().includes('operador')).length;
            quadro.innerHTML = `📊 <strong>${ativos.length} ativos</strong> | ${sups} supervisores | ${lids} líderes | ${ops} operadores`;
        }

        // B) Salva o Dashboard Principal do RH
        let totalGeral = dRH.length;
        let totalFerias = dRH.filter(c => c.status && c.status.toLowerCase().includes('féri')).length;
        let totalEPIs = (typeof dE !== 'undefined' && dE) ? dE.filter(e => e.status === 'Pendente' || e.status === 'Solicitado' || e.status === 'Entrega Parcial').length : 0;

        let textosKPI = document.querySelectorAll('.kpi-title, div, span');
        textosKPI.forEach(el => {
            let texto = el.innerText.trim().toLowerCase();
            let valor = null;

            if (texto === 'total colaboradores' || texto === 'total de colaboradores') valor = totalGeral;
            else if (texto === 'em férias' || texto === 'em ferias') valor = totalFerias;
            else if (texto === 'pendências epi' || texto === 'pendencias epi') valor = totalEPIs;

            if (valor !== null) {
                let caixaNumero = el.nextElementSibling;
                if (!caixaNumero || (caixaNumero.innerText.trim() !== '0' && isNaN(parseInt(caixaNumero.innerText)))) {
                    caixaNumero = el.parentElement.querySelector('.kpi-val, [style*="font-size: 28px"], [id*="total"]');
                }
                if (caixaNumero) caixaNumero.innerText = valor;
            }
        });

        // C) Força o botão de EPI
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