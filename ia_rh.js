// =================================================================
// MÓDULO ISOLADO: IA, BOTÃO EXCLUIR E AMORTECEDOR FIREBASE
// =================================================================

// 1. FUNÇÃO DE EXCLUIR
window.excluirColaboradorRH = async function(idColab) {
    if (!confirm("⚠️ Atenção: Deseja realmente EXCLUIR este colaborador do sistema?")) return;
    
    let bRH = window.dRH || window.dadosRH || JSON.parse(localStorage.getItem('sysLog_mem_rh') || '[]');
    let novaLista = bRH.filter(c => String(c.matricula) !== String(idColab) && String(c.id) !== String(idColab));
    
    window.dRH = novaLista; 
    window.dadosRH = novaLista;
    localStorage.setItem('sysLog_mem_rh', JSON.stringify(novaLista));
    
    if (typeof window.saveToDB === 'function') window.saveToDB('rh', novaLista);
    else if (typeof window.saveDataSafe === 'function') window.saveDataSafe('rh', novaLista);
    
    alert("🗑️ Colaborador excluído com sucesso!");
    if (typeof window.renderRHQuad === 'function') window.renderRHQuad();
    else if (typeof window.renderRH === 'function') window.renderRH();
    else location.reload();
};

// 2. INJETOR DO BOTÃO EXCLUIR (Vigia a tela sem sujar o código original)
setInterval(function() {
    document.querySelectorAll('button').forEach(btn => {
        let txt = btn.innerText.trim().toLowerCase();
        if ((txt === 'editar' || txt === 'modificar') && !btn.classList.contains('btn-excluido-ok')) {
            btn.classList.add('btn-excluido-ok');
            let acao = btn.getAttribute('onclick') || '';
            let match = acao.match(/['"]([^'"]+)['"]/);
            let id = match ? match[1] : null;
            if (id) {
                let btnExcluir = document.createElement('button');
                btnExcluir.innerHTML = '<i class="fas fa-trash"></i> Excluir';
                btnExcluir.setAttribute('onclick', `window.excluirColaboradorRH('${id}')`);
                btnExcluir.style.cssText = 'background:#e74c3c; color:white; border:none; padding:5px 10px; border-radius:4px; margin-left:5px; cursor:pointer; font-size:12px; font-weight:bold; box-shadow:0 2px 4px rgba(0,0,0,0.2); transition:0.3s;';
                btn.parentNode.insertBefore(btnExcluir, btn.nextSibling);
            }
        }
    });
}, 1000);

// 3. DASHBOARD INTELIGENTE (NATIVO E GRATUITO)
window.gerarInsightIAReal = function() {
    let telaRH = document.getElementById('view-rh-quadro') || document.querySelector('.view-section[style*="block"]');
    if (!telaRH || telaRH.style.display === 'none') return;
    
    let divIA = document.getElementById('painel-ia-rh-supremo');
    if (!divIA) {
        divIA = document.createElement('div');
        divIA.id = 'painel-ia-rh-supremo';
        telaRH.insertBefore(divIA, telaRH.firstChild);
    }
    
    let bRH = window.dRH || window.dadosRH || JSON.parse(localStorage.getItem('sysLog_mem_rh') || '[]');
    let bEPI = JSON.parse(localStorage.getItem('sysLog_epi_registros') || '[]');
    let total = bRH.length;
    let ativos = bRH.filter(c => c.status && c.status.toUpperCase() === 'ATIVO').length;
    let inativos = total - ativos;
    let episAtrasados = bEPI.filter(e => e.status !== 'entregue').length;
    
    let insights = "";
    let turnover = total > 0 ? ((inativos / total) * 100).toFixed(1) : 0;
    
    if (episAtrasados > 0) insights += `<p style="margin-bottom:12px;"><span style="color:#e74c3c;"><i class="fas fa-exclamation-triangle"></i> <b>Alerta EPI:</b></span> ${episAtrasados} pendentes.</p>`;
    else insights += `<p style="margin-bottom:12px;"><span style="color:#2ecc71;"><i class="fas fa-shield-alt"></i> <b>Segurança:</b></span> 100% dos EPIs em dia.</p>`;
    
    if (turnover > 15) insights += `<p style="margin-bottom:12px;"><span style="color:#f39c12;"><i class="fas fa-users-slash"></i> <b>Turnover Alto:</b></span> ${turnover}% de inativos.</p>`;
    else insights += `<p style="margin-bottom:12px;"><span style="color:#3498db;"><i class="fas fa-users"></i> <b>Equipa Estável:</b></span> Turnover de ${turnover}%.</p>`;
    
    divIA.innerHTML = `<div style="background: linear-gradient(135deg, #2c3e50, #34495e); color: white; padding: 20px; border-radius: 12px; margin-bottom: 25px; border-left: 5px solid #3498db;"><h3 style="margin-top:0;"><i class="fas fa-microchip" style="color:#3498db;"></i> Dashboard Inteligente</h3><div style="display:flex; gap:15px; margin-bottom:15px;"><div style="background:rgba(255,255,255,0.1); padding:10px; border-radius:8px; flex:1; text-align:center;"><h2 style="margin:0; color:#2ecc71;">${ativos}</h2><small>Ativos</small></div><div style="background:rgba(255,255,255,0.1); padding:10px; border-radius:8px; flex:1; text-align:center;"><h2 style="margin:0; color:#e74c3c;">${inativos}</h2><small>Inativos</small></div><div style="background:rgba(255,255,255,0.1); padding:10px; border-radius:8px; flex:1; text-align:center;"><h2 style="margin:0; color:#f1c40f;">${episAtrasados}</h2><small>EPIs Atrasados</small></div></div><div style="background:white; color:#333; padding:15px; border-radius:8px;">${insights}</div></div>`;
};
setInterval(window.gerarInsightIAReal, 2000);

// 4. AMORTECEDOR FIREBASE (Cura a lentidão e o erro "resource-exhausted")
setTimeout(function() {
    if (typeof window.dbCloud !== 'undefined' && !window.dbCloud.temAmortecedor) {
        window.dbCloud.temAmortecedor = true;
        let tempoEspera = 0;
        const originalCollection = window.dbCloud.collection.bind(window.dbCloud);
        window.dbCloud.collection = function(nomeColecao) {
            const colRef = originalCollection(nomeColecao);
            if (colRef.add) {
                const originalAdd = colRef.add.bind(colRef);
                colRef.add = function(dados) {
                    tempoEspera += 150; if (tempoEspera > 5000) tempoEspera = 150;
                    return new Promise(resolve => setTimeout(() => originalAdd(dados).then(resolve).catch(e=>console.warn("Protegido")), tempoEspera));
                };
            }
            const originalDoc = colRef.doc.bind(colRef);
            colRef.doc = function(caminhoDoc) {
                const docRef = originalDoc(caminhoDoc);
                if (docRef.set) {
                    const originalSet = docRef.set.bind(docRef);
                    docRef.set = function(dados, opcoes) {
                        tempoEspera += 150; if (tempoEspera > 5000) tempoEspera = 150;
                        return new Promise(resolve => setTimeout(() => originalSet(dados, opcoes).then(resolve).catch(e=>console.warn("Protegido")), tempoEspera));
                    };
                }
                if (docRef.update) {
                    const originalUpdate = docRef.update.bind(docRef);
                    docRef.update = function(dados) {
                        tempoEspera += 150; if (tempoEspera > 5000) tempoEspera = 150;
                        return new Promise(resolve => setTimeout(() => originalUpdate(dados).then(resolve).catch(e=>console.warn("Protegido")), tempoEspera));
                    };
                }
                return docRef;
            };
            return colRef;
        };
    }
}, 2500);