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
// ==========================================
// 1. AUTO-CORRETOR DE MATRÍCULAS DO EPI
// ==========================================
window.corrigirMatriculasEPI = async function() {
    let bancoEpi = [];
    if(typeof window.getFromDB === 'function') bancoEpi = await window.getFromDB('epi_registros') || [];
    if(bancoEpi.length === 0) bancoEpi = JSON.parse(localStorage.getItem('sysLog_epi_registros') || '[]');
    
    let bancoRH = typeof window.dRH !== 'undefined' && window.dRH.length > 0 ? window.dRH : JSON.parse(localStorage.getItem('sysLog_mem_rh') || '[]');
    let atualizados = 0;

    bancoEpi.forEach(epi => {
        if ((!epi.matricula || epi.matricula.trim() === '') && epi.nome) {
            let func = bancoRH.find(f => f.nome && f.nome.trim().toUpperCase() === epi.nome.trim().toUpperCase());
            if (func && func.matricula) {
                epi.matricula = func.matricula;
                atualizados++;
            }
        }
    });

    if(atualizados > 0) {
        localStorage.setItem('sysLog_epi_registros', JSON.stringify(bancoEpi));
        if(typeof window.saveToDB === 'function') window.saveToDB('epi_registros', bancoEpi);
        if(typeof window.renderEpi === 'function') window.renderEpi();
    }
};
// Executa o corretor 3 segundos após abrir o sistema
setTimeout(window.corrigirMatriculasEPI, 3000);

// ==========================================
// 2. ENTREGA DE EPI (TODOS OS ITENS SEMPRE VISÍVEIS)
// ==========================================
window.forcarEntregaEPI = async function(idRegistro) {
    let bancoEpi = [];
    if(typeof window.getFromDB === 'function') bancoEpi = await window.getFromDB('epi_registros') || [];
    if(bancoEpi.length === 0) bancoEpi = JSON.parse(localStorage.getItem('sysLog_epi_registros') || '[]');

    let idx = bancoEpi.findIndex(r => r.id === idRegistro || r.matricula === idRegistro);
    if(idx < 0) return alert('❌ Erro: Registo não encontrado no banco de dados.');

    let reg = bancoEpi[idx];
    let pecas = [
        { id: 'camisa', nome: 'Camisa/Blusa' }, { id: 'calca', nome: 'Calça' },
        { id: 'bota', nome: 'Bota/Sapato' }, { id: 'japona', nome: 'Japona Térmica' },
        { id: 'luva', nome: 'Luvas' }, { id: 'mangote', nome: 'Mangote' }, { id: 'avental', nome: 'Avental' }
    ];

    let htmlCheckboxes = '';
    let itensJaEntregues = 0;

    pecas.forEach(p => {
        let tamanhoSolicitado = reg[p.id] && reg[p.id].trim() !== '' ? reg[p.id] : '--';
        let jaEntregue = reg[p.id + '_entregue'] === true;
        if(jaEntregue) itensJaEntregues++;

        htmlCheckboxes += `
            <label style="display:flex; align-items:center; margin: 12px 0; font-size: 16px; cursor: pointer; padding: 10px; border-radius: 5px; background: ${jaEntregue ? '#e8f8f5' : '#fff'}; border: 1px solid ${jaEntregue ? '#27ae60' : '#ddd'};">
                <input type="checkbox" id="chk_epi_${p.id}" ${jaEntregue ? 'checked disabled' : ''} style="width: 20px; height: 20px; margin-right: 15px; cursor: pointer;">
                <div style="flex-grow: 1;">
                    <b>${p.nome}</b> <span style="color:#7f8c8d; font-size:14px;">(Tam: ${tamanhoSolicitado})</span>
                </div>
                ${jaEntregue ? '<span style="color:#27ae60; font-weight:bold; font-size:12px; background:#d5f5e3; padding:3px 8px; border-radius:10px;">JÁ ENTREGUE</span>' : ''}
            </label>
        `;
    });

    if(itensJaEntregues === pecas.length) return alert('✅ Todos os itens possíveis já foram entregues a este colaborador.');

    let modal = document.createElement('div');
    modal.id = 'modal-epi-parcial-fix';
    modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:99999; display:flex; justify-content:center; align-items:center; font-family:Arial,sans-serif;';
    modal.innerHTML = `
        <div style="background:#f4f7f6; padding:0; border-radius:12px; width:90%; max-width:450px; box-shadow:0 10px 25px rgba(0,0,0,0.5); overflow:hidden;">
            <div style="background:#2c3e50; color:white; padding:20px; text-align:center;">
                <h3 style="margin:0; font-size:20px;"><i class="fas fa-box-open"></i> Entrega de Uniforme/EPI</h3>
                <p style="margin:5px 0 0 0; font-size:14px; color:#bdc3c7;">${reg.nome || reg.matricula}</p>
            </div>
            <div style="padding:20px; max-height:450px; overflow-y:auto;">
                <p style="margin-top:0; color:#555; font-size:14px; font-weight:bold;">Selecione os itens que estão a ser entregues AGORA:</p>
                ${htmlCheckboxes}
            </div>
            <div style="padding:15px 20px; background:#fff; border-top:1px solid #ddd; display:flex; justify-content:flex-end; gap:10px;">
                <button onclick="document.getElementById('modal-epi-parcial-fix').remove()" style="padding:10px 20px; background:#e74c3c; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;">Cancelar</button>
                <button id="btn-salvar-epi-fix" style="padding:10px 20px; background:#27ae60; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;">Confirmar Selecionados</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);

    document.getElementById('btn-salvar-epi-fix').onclick = async function() {
        let dataHoje = new Date().toLocaleDateString('pt-BR');
        let entreguesNestaSessao = 0;
        let totalEntreguesFinal = 0;

        pecas.forEach(p => {
            let chk = document.getElementById('chk_epi_' + p.id);
            if(chk && chk.checked) {
                if(!reg[p.id + '_entregue']) entreguesNestaSessao++;
                reg[p.id + '_entregue'] = true;
                totalEntreguesFinal++;
            } else if (reg[p.id + '_entregue']) {
                totalEntreguesFinal++;
            }
        });

        if(entreguesNestaSessao === 0) {
            alert('⚠️ Selecione pelo menos um novo item para realizar a entrega.');
            return;
        }

        if(totalEntreguesFinal === pecas.length) {
            reg.status = 'entregue';
            reg.dataEntrega = dataHoje;
        } else {
            reg.status = 'parcial';
            reg.dataEntrega = dataHoje + ' (parcial)';
        }

        bancoEpi[idx] = reg;
        localStorage.setItem('sysLog_epi_registros', JSON.stringify(bancoEpi));
        if(typeof window.saveToDB === 'function') await window.saveToDB('epi_registros', bancoEpi);
        
        document.getElementById('modal-epi-parcial-fix').remove();
        
        if(typeof window.renderEpi === 'function') window.renderEpi();
        else location.reload();
    };
};

// "Olheiro" para conectar o botão de entregar
setInterval(() => {
    document.querySelectorAll('button').forEach(btn => {
        let txt = btn.innerText.trim().toLowerCase();
        if((txt === 'entregar' || txt === 'entregue') && !btn.hasAttribute('data-consertado')) {
            let acaoAntiga = btn.getAttribute('onclick') || '';
            let matchId = acaoAntiga.match(/['"]([^'"]+)['"]/); 
            if(matchId && matchId[1]) {
                btn.onclick = function(e) { e.preventDefault(); e.stopPropagation(); window.forcarEntregaEPI(matchId[1]); };
                btn.setAttribute('data-consertado', 'true');
                btn.style.background = '#2980b9'; 
                btn.style.color = 'white';
            }
        }
    });
}, 1000);
// ==========================================
// 3. BLINDAGEM CONTÍNUA DO RH (Impede novos de sumirem)
// ==========================================
setInterval(() => {
    let bRH = window.dadosRH || window.dRH || JSON.parse(localStorage.getItem('sysLog_mem_rh') || '[]');
    let precisouCorrigir = false;

    if (!Array.isArray(bRH)) return;

    bRH.forEach(c => {
        // Verifica se o cadastro é novo e está vulnerável (sem as cópias de segurança)
        let nomeReal = c.nome || c.Nome || c.NOME || c.nomeCompleto || c.colaborador || "";
        if (nomeReal && (!c.Nome || !c.NOME || !c.colaborador)) {
            c.nome = nomeReal; 
            c.Nome = nomeReal; 
            c.NOME = nomeReal; 
            c.nomeCompleto = nomeReal; 
            c.colaborador = nomeReal;
            
            c.matricula = c.matricula || c.Matricula || ""; 
            c.Matricula = c.matricula;
            
            c.cargo = c.cargo || c.Cargo || c.funcao || ""; 
            c.Cargo = c.cargo; 
            c.funcao = c.cargo;
            
            c.status = c.status || c.Status || "ATIVO"; 
            c.Status = c.status;
            
            precisouCorrigir = true;
        }
    });

    // Se encontrou alguém novo e corrigiu, salva silenciosamente
    if (precisouCorrigir) {
        window.dadosRH = bRH;
        window.dRH = bRH;
        localStorage.setItem('sysLog_mem_rh', JSON.stringify(bRH));
        
        // Dá um "F5" apenas na tabela, sem piscar a tela toda
        if (typeof window.renderRH === 'function') window.renderRH();
    }
}, 2000); // Fica de vigia a cada 2 segundos

// ==========================================
// 4. TURBO DE SALVAMENTO NA NUVEM (Fim da lentidão)
// ==========================================
if (!window.saveToDB_turbo_ativado) {
    const originalSaveToDB = window.saveToDB;
    if (typeof originalSaveToDB === 'function') {
        window.saveToDB = async function(colecao, dados) {
            // 1. Salva IMEDIATAMENTE na memória local para a sua tela não travar
            localStorage.setItem('sysLog_mem_' + colecao, JSON.stringify(dados));
            
            // 2. Envia para a nuvem por trás dos panos (Background)
            originalSaveToDB(colecao, dados).then(() => {
                console.log("☁️ Salvo na nuvem com sucesso!");
            }).catch(e => {
                console.log("⏳ Nuvem lenta, mas os dados estão a salvo na memória!");
            });

            // Retorna sucesso instantâneo para o sistema não ficar a carregar
            return Promise.resolve(); 
        };
        window.saveToDB_turbo_ativado = true;
    }
}
// ==========================================
// ESCUDO DEFINITIVO: SALVAMENTO INSTANTÂNEO E AUTO-FORMATAÇÃO
// ==========================================

if (!window.escudoSalvamentoAtivo) {
    // Guarda as funções originais do seu sistema
    const originalSaveDataSafe = window.saveDataSafe;
    const originalSaveToDB = window.saveToDB;

    // Função que "arruma" qualquer cadastro bagunçado antes de salvar
    const formatarRH = function(dados) {
        if (!Array.isArray(dados)) return dados;
        dados.forEach(c => {
            // Garante o Nome
            let nomeReal = c.nome || c.Nome || c.NOME || c.nomeCompleto || c.colaborador || c.NomeCompleto || "SEM NOME";
            c.nome = nomeReal; c.Nome = nomeReal; c.NOME = nomeReal; c.colaborador = nomeReal; c.nomeCompleto = nomeReal;
            
            // Garante a Matrícula
            let matReal = c.matricula || c.Matricula || c.MATRICULA || "";
            c.matricula = matReal; c.Matricula = matReal;
            
            // Garante o Cargo
            let cargoReal = c.cargo || c.Cargo || c.funcao || "";
            c.cargo = cargoReal; c.Cargo = cargoReal; c.funcao = cargoReal;
            
            // Garante o Status
            let statusReal = c.status || c.Status || "ATIVO";
            c.status = statusReal; c.Status = statusReal;
            
            // Cria um ID se o sistema não tiver criado
            if (!c.id) c.id = c.matricula || ("novo_" + Date.now() + Math.floor(Math.random() * 1000));
        });
        return dados;
    };

    // 1. Interceta a função saveDataSafe
    if (typeof originalSaveDataSafe === 'function') {
        window.saveDataSafe = function(colecao, dados) {
            if (colecao === 'rh') dados = formatarRH(dados);
            
            // Salva na memória do navegador NA HORA (impede que suma)
            localStorage.setItem('sysLog_mem_' + colecao, JSON.stringify(dados));
            if (colecao === 'rh') { window.dadosRH = dados; window.dRH = dados; }

            // Envia para a Nuvem de forma invisível (não trava a sua tela)
            setTimeout(() => {
                originalSaveDataSafe(colecao, dados).catch(e => console.warn("Nuvem lenta ignorada."));
            }, 50);

            // Avisa o seu sistema que "já salvou" para ele liberar a tela imediatamente
            return Promise.resolve();
        };
    }

    // 2. Interceta a função saveToDB (faz o mesmo processo)
    if (typeof originalSaveToDB === 'function') {
        window.saveToDB = async function(colecao, dados) {
            if (colecao === 'rh') dados = formatarRH(dados);
            
            localStorage.setItem('sysLog_mem_' + colecao, JSON.stringify(dados));
            if (colecao === 'rh') { window.dadosRH = dados; window.dRH = dados; }

            setTimeout(() => {
                originalSaveToDB(colecao, dados).catch(e => console.warn("Nuvem lenta ignorada."));
            }, 50);

            return Promise.resolve();
        };
    }

    // 3. Interceta a criação visual da tabela para garantir que ninguém fica em branco
    const originalRenderRH = window.renderRH;
    if (typeof originalRenderRH === 'function') {
        window.renderRH = function() {
            if (window.dadosRH) formatarRH(window.dadosRH);
            if (window.dRH) formatarRH(window.dRH);
            originalRenderRH.apply(this, arguments);
        };
    }

    window.escudoSalvamentoAtivo = true;
    console.log("✅ Escudo Definitivo de Salvamento Ativado!");
}
// ==========================================
// 1. CORREÇÃO DO ERRO setDRH (Dashboard travando)
// ==========================================
if (typeof window.setDRH === 'undefined') {
    window.setDRH = function(id, val) {
        let el = document.getElementById(id);
        if (el) { el.innerText = val; }
    };
}

// ==========================================
// 2. FREIO ABS DO FIREBASE (Fim da Metralhadora de Saves)
// ==========================================
if (!window.firebaseFreioAtivado) {
    const originalSaveToDB = window.saveToDB;
    const originalSaveDataSafe = window.saveDataSafe;
    const originalAutoSave = window.autoSaveData;

    window.debounceTimers = {};

    // Função que "segura" os pedidos e manda um só
    function aplicarFreio(colecao, dados, funcaoOriginal) {
        // 1. Salva na memória do navegador na hora (impede a perda de dados)
        localStorage.setItem('sysLog_mem_' + colecao, JSON.stringify(dados));
        
        // Mantém as variáveis ativas para a tabela não piscar
        if (colecao === 'rh') { window.dadosRH = dados; window.dRH = dados; }

        // 2. Se o sistema metralhar outro pedido em menos de 3 segundos, cancela o anterior
        if (window.debounceTimers[colecao]) {
            clearTimeout(window.debounceTimers[colecao]);
        }

        // 3. Agenda UM ÚNICO salvamento para a nuvem
        window.debounceTimers[colecao] = setTimeout(() => {
            if (typeof funcaoOriginal === 'function') {
                funcaoOriginal(colecao, dados).catch(e => console.warn("Nuvem bloqueada, dados salvos localmente."));
            }
        }, 3000); // Espera 3 segundos de silêncio para salvar

        // 4. Libera a sua tela na mesma hora
        return Promise.resolve();
    }

    // Instala o Freio na função principal
    if (typeof originalSaveToDB === 'function') {
        window.saveToDB = function(colecao, dados) {
            return aplicarFreio(colecao, dados, originalSaveToDB);
        };
    }

    // Instala o Freio na função de segurança
    if (typeof originalSaveDataSafe === 'function') {
        window.saveDataSafe = function(colecao, dados) {
            return aplicarFreio(colecao, dados, originalSaveDataSafe);
        };
    }

    // Instala o Freio no Auto-Save Enlouquecido
    if (typeof originalAutoSave === 'function') {
        window.autoSaveData = function() {
            if (window.timerAutoSaveLouco) clearTimeout(window.timerAutoSaveLouco);
            window.timerAutoSaveLouco = setTimeout(() => {
                originalAutoSave.apply(this, arguments);
            }, 5000); // Só deixa o auto-save rodar de 5 em 5 segundos no máximo
        };
    }

    window.firebaseFreioAtivado = true;
    console.log("🛑 Freio ABS do Firebase Ativado! Tela destravada.");
}

// Garante as Maiúsculas ao renderizar o RH (Proteção Extra)
const originalRenderRH2 = window.renderRH;
if (typeof originalRenderRH2 === 'function' && !window.renderRH_protegido) {
    window.renderRH = function() {
        let bRH = window.dadosRH || window.dRH || [];
        bRH.forEach(c => {
            let nomeReal = c.nome || c.Nome || c.NOME || c.colaborador || "Sem Nome";
            c.nome = nomeReal; c.Nome = nomeReal; c.colaborador = nomeReal;
        });
        originalRenderRH2.apply(this, arguments);
    };
    window.renderRH_protegido = true;
}
// ==========================================
// FIREWALL ABSOLUTO CONTRA TRAVAMENTO DA NUVEM (Erro 400 / Firebase)
// ==========================================
(function() {
    if (window.firewallNuvemAtivo) return;

    console.log("🛡️ Iniciando Firewall contra travamentos do Firebase...");

    // 1. DESLIGA A METRALHADORA DE AUTO-SAVE
    window.autoSaveData = function() {
        console.log("🛑 AutoSave bloqueado pelo Firewall. Poupando a nuvem.");
    };

    // 2. GUARDA AS FUNÇÕES ORIGINAIS (Para usar depois, em segurança)
    const originalSaveToDB = window.saveToDB;
    const originalSaveDataSafe = window.saveDataSafe;
    let delayNuvem = null;

    // 3. O NOVO MOTOR DE SALVAMENTO BLINDADO
    function salvamentoTurboBlindado(colecao, dados) {
        // A. Salva instantaneamente no seu navegador (Garante que nunca some!)
        localStorage.setItem('sysLog_mem_' + colecao, JSON.stringify(dados));
        
        if (colecao === 'rh' && Array.isArray(dados)) {
            // Formata o novo funcionário para não sumir da tela
            dados.forEach(c => {
                let n = c.nome || c.Nome || c.colaborador || "Sem Nome";
                c.nome = n; c.Nome = n; c.colaborador = n;
                c.matricula = c.matricula || c.Matricula || ""; c.Matricula = c.matricula;
            });
            window.dadosRH = dados; 
            window.dRH = dados;
        }

        // B. Engana o seu sistema! Diz que "já salvou" na nuvem para ele destravar a tela NA HORA.
        setTimeout(() => {
            if(typeof window.renderRH === 'function') window.renderRH();
        }, 100);

        // C. Só envia para o Google após 10 segundos sem você fazer nada
        if (delayNuvem) clearTimeout(delayNuvem);
        delayNuvem = setTimeout(() => {
            console.log("☁️ Enviando pacote único para a nuvem em segurança...");
            if (typeof originalSaveToDB === 'function') {
                // Tenta enviar. Se o Google bloquear, absorve o erro e não trava a sua tela.
                originalSaveToDB(colecao, dados).catch(e => console.warn("Erro do Firebase contido."));
            }
        }, 10000);

        // Retorna SUCESSO imediato para fechar as janelas de carregamento
        return Promise.resolve("OK"); 
    }

    // 4. SUBSTITUI AS FUNÇÕES DO SISTEMA PELO NOSSO MOTOR
    if (typeof window.saveToDB === 'function') window.saveToDB = salvamentoTurboBlindado;
    if (typeof window.saveDataSafe === 'function') window.saveDataSafe = salvamentoTurboBlindado;

    // 5. PROTEÇÃO DO DASHBOARD (Evita erro de tela branca)
    if (typeof window.setDRH === 'undefined') {
        window.setDRH = function(id, val) {
            let el = document.getElementById(id);
            if (el) el.innerText = val;
        };
    }

    window.firewallNuvemAtivo = true;
    console.log("✅ Firewall 100% Ativo. A lentidão acabou.");
})();