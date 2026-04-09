// =========================================================================
// MÓDULO RH v164.1 — Recursos Humanos
// FIXES: (1) Botão Excluir, (2) Import PDF de Férias, (3) Dashboard RH
// =========================================================================

// ── Garante variáveis globais compartilhadas ──────────────────────────────
if (typeof window.dRH        === 'undefined') window.dRH        = [];
if (typeof window.rhData     === 'undefined') window.rhData     = {};
if (typeof window.pendenciasRH === 'undefined') window.pendenciasRH = [];

// ── Helper ────────────────────────────────────────────────────────────────
function ensureRhData(mat) {
    if (!window.rhData[mat]) {
        window.rhData[mat] = {
            feriasInicio: '', feriasFim: '', limiteFerias: '',
            periodoAberto: '', vendeu10Dias: false,
            banco: 0, faltas: [], statusManual: 'AUTO', ponto: []
        };
    }
}

// ── helpers localStorage ──────────────────────────────────────────────────
function _getRHList() {
    try { const r = localStorage.getItem('rh_colaboradores'); return r ? JSON.parse(r) : []; } catch(e) { return []; }
}
function _saveRHList(lista) {
    try { localStorage.setItem('rh_colaboradores', JSON.stringify(lista)); } catch(e) {}
    window.dRH = lista;
}

// =========================================================================
// FIX 1 — QUADRO DE COLABORADORES com botão EXCLUIR funcionando
// =========================================================================
window.renderRHQuad = function () {
    const tb = document.querySelector('#tbRH');
    if (!tb) return;

    // Fonte unificada: tenta dRH (Firebase), fallback localStorage
    let lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();
    if (!lista || lista.length === 0) {
        tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#999;padding:20px;">Nenhum colaborador registrado.</td></tr>';
        _updateKPIs([]);
        return;
    }

    let ativos = 0, lideres = 0, ops = 0, sups = 0;
    let rows = '';

    lista.forEach((r, index) => {
        const mat  = r['Matrícula'] || r['Matricula'] || r['MATRICULA'] || r['matricula'] || '-';
        const nome = r['Nome']      || r['NOME']      || r['nome']      || '-';
        if (nome === '-' && mat === '-') return;

        const func = (r['Função'] || r['Funcao'] || r['FUNÇÃO'] || r['funcao'] || '-').toUpperCase();
        let baseStatus = (r['Status'] || r['STATUS'] || r['status'] || 'ATIVO').toUpperCase();

        ensureRhData(mat);
        const rh = window.rhData[mat];
        let statusFinal = (rh.statusManual && rh.statusManual !== 'AUTO')
            ? rh.statusManual.toUpperCase() : baseStatus;

        if (statusFinal === 'ATIVO' && rh.feriasInicio && rh.feriasFim) {
            const today = new Date(); today.setHours(0, 0, 0, 0);
            const fIn  = new Date(rh.feriasInicio + 'T00:00:00');
            const fFim = new Date(rh.feriasFim    + 'T00:00:00');
            if (today >= fIn && today <= fFim) statusFinal = 'FÉRIAS';
        }

        if (statusFinal === 'INATIVO') return;

        ativos++;
        if      (func.includes('SUPERVISOR') || func.includes('SUPERVISORA')) sups++;
        else if (func.includes('LIDER')      || func.includes('LÍDER'))       lideres++;
        else if (func.includes('OPERADOR')   || func.includes('SEPARADOR'))   ops++;

        const statusMap = {
            'ATIVO':    { bg:'#e8f5e9', color:'#2e7d32', label:'ATIVO' },
            'FÉRIAS':   { bg:'#fff3cd', color:'#856404', label:'<i class="fas fa-umbrella-beach"></i> EM FÉRIAS' },
            'AFASTADO': { bg:'#fdf5e6', color:'#e67e22', label:'AFASTADO' },
        };
        let stBg = '#ffebeb', stColor = '#c0392b', stLabel = statusFinal;
        if (statusMap[statusFinal]) ({ bg: stBg, color: stColor, label: stLabel } = statusMap[statusFinal]);
        else if (statusFinal.includes('TRANSFERENCIA')) { stBg='#e0f7fa'; stColor='#00838f'; stLabel='<i class="fas fa-exchange-alt"></i> TRANSFERIDO'; }

        const celular = r['celular'] || r['Celular'] || r['CELULAR'] || '';
        const celDisplay = celular
            ? `<a href="https://wa.me/${celular.replace(/\D/g,'')}" target="_blank" style="color:#25d366;font-size:11px;"><i class="fab fa-whatsapp"></i> ${celular}</a>`
            : '\u2014';

        rows += `<tr>
            <td style="font-weight:bold;text-align:center;">${mat}</td>
            <td style="font-weight:600;">${nome}</td>
            <td style="color:#555;">${r['Função']||r['Funcao']||r['FUNÇÃO']||r['funcao']||'-'}</td>
            <td style="text-align:center;font-size:11px;">${celDisplay}</td>
            <td style="text-align:center;"><span style="background:${stBg};color:${stColor};padding:4px 10px;border-radius:12px;font-weight:900;font-size:10px;">${stLabel}</span></td>
            <td style="text-align:center;white-space:nowrap;">
                <button class="btn btn-blue" style="padding:4px 8px;font-size:11px;margin-right:3px;" onclick="abrirEdicaoRH(${index})">
                    <i class="fas fa-edit"></i> Editar
                </button>
                <button class="btn-del-colab" style="padding:4px 8px;font-size:11px;" onclick="excluirColaboradorRH(${index})">
                    <i class="fas fa-trash-alt"></i> Excluir
                </button>
            </td>
        </tr>`;
    });

    tb.innerHTML = rows || '<tr><td colspan="6" style="text-align:center;color:#999;padding:20px;">Nenhum colaborador ativo.</td></tr>';
    _updateKPIs({ ativos, lideres, ops, sups, total: lista.length });
};

function _updateKPIs(data) {
    if (Array.isArray(data)) {
        let a=0,l=0,o=0,s=0;
        data.forEach(c => {
            const st = (c.status||c.STATUS||c.statusManual||'ATIVO').toUpperCase();
            if (st === 'INATIVO') return;
            a++;
            const f = (c.funcao||c.FUNCAO||c['Função']||c['FUNÇÃO']||'').toUpperCase();
            if (f.includes('SUPERVISOR')) s++;
            else if (f.includes('LIDER')||f.includes('LÍDER')) l++;
            else if (f.includes('OPERADOR')||f.includes('SEPARADOR')) o++;
        });
        data = {ativos:a, lideres:l, ops:o, sups:s, total:data.length};
    }
    const set = (id, v) => { const el=document.getElementById(id); if(el) el.textContent=v; };
    set('rh-tot', data.ativos||0);
    set('rh-lid', data.lideres||0);
    set('rh-op',  data.ops||0);
    set('rh-sup', data.sups||0);
    const kpis = document.getElementById('rh-kpis');
    if (kpis && (data.total||0) > 0) kpis.style.display = 'grid';
}

// =========================================================================
// FIX 1B — EXCLUIR COLABORADOR (unifica dRH + localStorage)
// =========================================================================
window.excluirColaboradorRH = async function (index) {
    let lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();
    if (!lista || index < 0 || index >= lista.length) {
        alert('Colaborador não encontrado.'); return;
    }
    const colab = lista[index];
    const mat   = colab['Matrícula']||colab['Matricula']||colab['MATRICULA']||colab['matricula']||'';
    const nome  = colab['Nome']||colab['NOME']||colab['nome']||mat;

    const motivo = prompt(
        `Excluir: ${nome} (${mat})\n\nMotivo da saída:\n1 = Demissão\n2 = Transferência\n3 = Outro\n\nDigite o número:`, '1'
    );
    if (motivo === null) return;
    const motivoMap = { '1': 'Demissão', '2': 'Transferência', '3': 'Outro' };
    const motivoTexto = motivoMap[motivo] || 'Não informado';

    try {
        lista.splice(index, 1);
        window.dRH = lista;

        if (window.rhData && window.rhData[mat]) delete window.rhData[mat];

        if (window.saveToDB) {
            await window.saveToDB('dRH',    window.dRH);
            await window.saveToDB('rhData', window.rhData);
        }
        _saveRHList(lista);

        if (motivo === '1' || motivo === '2') _registrarPendencia(colab, motivoTexto);

        window.renderRHQuad();
        if (typeof renderFerias      === 'function') renderFerias();
        if (typeof renderBanco       === 'function') renderBanco();
        if (typeof renderAbs         === 'function') renderAbs();
        if (typeof renderPontoMensal === 'function') renderPontoMensal();

        if (typeof showToast === 'function')
            showToast(`${nome} removido. ${motivo==='1'||motivo==='2'?'Pendência de reposição criada.':''}`, 'warn');

    } catch (e) {
        alert('Erro ao excluir: ' + e.message);
    }
};

function _registrarPendencia(colab, motivo) {
    try {
        const arr = JSON.parse(localStorage.getItem('pendencias_reposicao') || '[]');
        arr.push({
            id: Date.now(),
            exColab:     colab['Nome']||colab['NOME']||colab['nome']||'N/A',
            matriculaEx: colab['Matrícula']||colab['Matricula']||colab['MATRICULA']||colab['matricula']||'',
            motivo, data: new Date().toLocaleDateString('pt-BR'),
            funcao: colab['Função']||colab['Funcao']||colab['FUNÇÃO']||colab['funcao']||''
        });
        localStorage.setItem('pendencias_reposicao', JSON.stringify(arr));
        if (typeof renderReposicoes === 'function') renderReposicoes();
    } catch(e) {}
}

// =========================================================================
// MODAIS
// =========================================================================
window.abrirModalRH = function () {
    document.getElementById('rhEditIndex').value = -1;
    ['rhMat','rhNome','rhFunc','rhAdm','rhCelular'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
    if (document.getElementById('rhFunc')) document.getElementById('rhFunc').value = 'OPERADOR';
    if (document.getElementById('rhStatus')) document.getElementById('rhStatus').value = 'AUTO';
    document.getElementById('modalRHTitle').innerHTML = '<i class="fas fa-user-plus"></i> Novo Colaborador';
    document.getElementById('modalRH').style.display = 'flex';
};

window.abrirEdicaoRH = function (index) {
    let lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();
    const c = lista[index];
    if (!c) return;
    const mat = c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';
    document.getElementById('rhEditIndex').value = index;
    if (document.getElementById('rhMat'))     document.getElementById('rhMat').value  = mat;
    if (document.getElementById('rhNome'))    document.getElementById('rhNome').value = c['Nome']||c['NOME']||c['nome']||'';
    if (document.getElementById('rhFunc'))    document.getElementById('rhFunc').value = c['Função']||c['Funcao']||c['FUNÇÃO']||c['funcao']||'';
    if (document.getElementById('rhCelular')) document.getElementById('rhCelular').value = c['celular']||c['Celular']||c['CELULAR']||'';

    let admRaw = c['Admissão']||c['Admissao']||c['admissao']||'';
    let admDate = '';
    if (admRaw && admRaw.includes('/')) {
        const pts = admRaw.split('/');
        if (pts.length === 3) admDate = `${pts[2]}-${pts[1].padStart(2,'0')}-${pts[0].padStart(2,'0')}`;
    } else if (admRaw && admRaw.includes('-')) admDate = admRaw;
    if (document.getElementById('rhAdm')) document.getElementById('rhAdm').value = admDate;

    ensureRhData(mat);
    if (document.getElementById('rhStatus')) document.getElementById('rhStatus').value = window.rhData[mat].statusManual || 'AUTO';
    document.getElementById('modalRHTitle').innerHTML = '<i class="fas fa-user-edit"></i> Editar Colaborador';
    document.getElementById('modalRH').style.display = 'flex';
};

window.fecharModalRH = function () { document.getElementById('modalRH').style.display = 'none'; };

window.salvarColaboradorRH = async function () {
    const idx     = parseInt(document.getElementById('rhEditIndex').value);
    const mat     = (document.getElementById('rhMat')?.value||'').trim();
    const nome    = (document.getElementById('rhNome')?.value||'').trim().toUpperCase();
    const func    = (document.getElementById('rhFunc')?.value||'').trim().toUpperCase();
    const adm     = document.getElementById('rhAdm')?.value||'';
    const status  = document.getElementById('rhStatus')?.value||'AUTO';
    const celular = (document.getElementById('rhCelular')?.value||'').trim();

    if (!mat || !nome) return alert('A Matrícula e o Nome são obrigatórios!');

    let admFormatada = adm;
    if (adm) { const d = new Date(adm+'T12:00:00'); admFormatada = d.toLocaleDateString('pt-BR'); }

    let lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();
    if (idx === -1) {
        lista.push({ 'Matrícula': mat, 'Nome': nome, 'Função': func, 'Admissão': admFormatada, 'Status': 'ATIVO', 'celular': celular });
    } else {
        lista[idx]['Matrícula'] = mat; lista[idx]['Nome'] = nome;
        lista[idx]['Função'] = func;   lista[idx]['Admissão'] = admFormatada;
        lista[idx]['celular'] = celular;
    }
    window.dRH = lista;
    _saveRHList(lista);

    ensureRhData(mat);
    window.rhData[mat].statusManual = status;

    if (window.saveToDB) {
        await window.saveToDB('dRH', window.dRH);
        await window.saveToDB('rhData', window.rhData);
    }

    window.fecharModalRH();
    window.renderRHQuad();
    if (typeof renderFerias      === 'function') renderFerias();
    if (typeof renderBanco       === 'function') renderBanco();
    if (typeof renderAbs         === 'function') renderAbs();
    if (typeof renderPontoMensal === 'function') renderPontoMensal();
    if (typeof showToast === 'function') showToast('Colaborador salvo!', 'success');
};

// =========================================================================
// RENDERIZAR FÉRIAS
// =========================================================================
window.renderFerias = function () {
    const tb = document.getElementById('tbFerias');
    if (!tb) return;
    let lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();
    if (!lista.length) return;
    tb.innerHTML = '';

    lista.forEach(r => {
        const mat  = r['Matrícula']||r['Matricula']||r['MATRICULA']||r['matricula']||'-';
        const nome = r['Nome']||r['NOME']||r['nome']||'-';
        const adm  = r['Admissão']||r['Admissao']||r['ADMISSÃO']||r['admissao']||'-';
        ensureRhData(mat);
        const rh = window.rhData[mat];
        const statusManual = rh.statusManual || 'AUTO';
        const baseStatus   = (r['Status']||r['STATUS']||r['status']||'-').toUpperCase();

        if (nome==='-' || (baseStatus==='INATIVO'&&statusManual!=='ATIVO') || statusManual==='INATIVO' || statusManual==='TRANSFERENCIA DE TURNO') return;

        let statusBadge = '<span style="background:#3498db;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">AQUISITIVO</span>';
        const dtIni = rh.feriasInicio||'', dtFim = rh.feriasFim||'';
        const hasMarcacao = (dtIni && dtFim);

        if (hasMarcacao) {
            const today=new Date(); today.setHours(0,0,0,0);
            const dI=new Date(dtIni+'T00:00:00'), dF=new Date(dtFim+'T00:00:00');
            if      (today>=dI&&today<=dF) statusBadge='<span style="background:#8e44ad;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">EM FÉRIAS</span>';
            else if (dI>today)             statusBadge='<span style="background:#f1c40f;color:#333;padding:3px 6px;border-radius:4px;font-weight:bold;">MARCADA</span>';
            else                           statusBadge='<span style="background:#7f8c8d;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">GOZADAS</span>';
        } else if (rh.limiteFerias) {
            const today=new Date(); today.setHours(0,0,0,0);
            const lim=new Date(rh.limiteFerias+'T00:00:00');
            const d=Math.floor((lim-today)/86400000);
            if      (d<0)   statusBadge='<span style="background:#c0392b;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">VENCIDA</span>';
            else if (d<=90) statusBadge='<span style="background:#f39c12;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">VENCENDO</span>';
            else            statusBadge='<span style="background:#27ae60;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">NO PRAZO</span>';
        }

        let limitDisplay = '-';
        if (rh.limiteFerias) {
            const ls = rh.limiteFerias.split('-').reverse().join('/');
            limitDisplay = `<strong style="color:var(--dark);font-size:12px;">${ls}</strong>`;
            const today=new Date(); today.setHours(0,0,0,0);
            if (new Date(rh.limiteFerias+'T00:00:00')<today && !hasMarcacao)
                limitDisplay=`<strong style="color:#c0392b;font-size:12px;">${ls}</strong>`;
        }

        const periodoAberto = rh.periodoAberto
            ? `<strong style="color:var(--blue);font-size:11px;">${rh.periodoAberto}</strong>`
            : '<span style="color:#ccc;">-</span>';

        tb.innerHTML += `<tr>
            <td style="font-weight:bold;font-size:12px;">${mat}</td>
            <td style="text-align:left;font-size:12px;">${nome}</td>
            <td style="font-size:11px;color:#555;">${adm}</td>
            <td>${periodoAberto}</td>
            <td>${limitDisplay}</td>
            <td>${statusBadge}</td>
            <td><input type="date" value="${dtIni}" onchange="updateRHData('${mat}','feriasInicio',this.value,this);window.renderRHQuad();window.renderFerias();" style="padding:4px;border:1px solid #ccc;border-radius:4px;font-weight:bold;font-size:10px;"></td>
            <td><input type="date" value="${dtFim}" onchange="updateRHData('${mat}','feriasFim',this.value,this);window.renderRHQuad();window.renderFerias();" style="padding:4px;border:1px solid #ccc;border-radius:4px;font-weight:bold;font-size:10px;"></td>
            <td><label style="cursor:pointer;font-weight:800;color:var(--orange);font-size:10px;display:flex;align-items:center;gap:5px;justify-content:center;"><input type="checkbox" ${rh.vendeu10Dias?'checked':''} onchange="updateRHData('${mat}','vendeu10Dias',this.checked,this)"> Vendeu 10 Dias</label></td>
            <td><button class="btn-sm" style="background:#25D366;color:white;padding:5px;border-radius:20px;" onclick="sendWhatsAppLink('${mat}','${nome}')" title="Link Individual"><i class="fab fa-whatsapp"></i></button></td>
        </tr>`;
    });
};

window.updateRHData = async function (mat, field, val, inputEl) {
    ensureRhData(mat);
    window.rhData[mat][field] = val;
    if (window.saveToDB) window.saveToDB('rhData', window.rhData);
    if (inputEl && inputEl.type !== 'checkbox') {
        inputEl.style.backgroundColor = '#d4edda';
        setTimeout(() => { inputEl.style.backgroundColor = ''; }, 1000);
    }
};

// =========================================================================
// RENDERIZAR BANCO DE HORAS
// =========================================================================
window.renderBanco = function () {
    const tb = document.getElementById('tbBanco');
    if (!tb) return;
    let lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();
    if (!lista.length) return;
    tb.innerHTML = '';
    let totalHoras = 0;
    lista.forEach(r => {
        const mat    = r['Matrícula']||r['Matricula']||r['MATRICULA']||r['matricula']||'-';
        const nome   = r['Nome']||r['NOME']||r['nome']||'-';
        const status = (r['Status']||r['STATUS']||r['status']||'-').toUpperCase();
        if (nome==='-'||status==='INATIVO') return;
        ensureRhData(mat);
        const h=window.rhData[mat].banco||0; totalHoras+=h;
        const color = h>0?'var(--green)':(h<0?'var(--red)':'#555');
        tb.innerHTML += `<tr><td style="font-weight:bold;">${mat}</td><td style="text-align:left;">${nome}</td><td><strong style="color:${color};font-size:16px;">${formatDecimalToTime(h)}</strong></td></tr>`;
    });
    if (typeof safeUpdate==='function') safeUpdate('bh-global', formatDecimalToTime(totalHoras));
};

// =========================================================================
// FIX 2 — IMPORT ESCALA DE FÉRIAS: PDF + Excel
// =========================================================================
document.addEventListener('DOMContentLoaded', function () {

    const fRH = document.getElementById('f_rh');
    if (fRH) {
        fRH.addEventListener('change', function () {
            if (typeof XLSX === 'undefined' || !this.files[0]) return;
            document.getElementById('loading').style.display = 'flex';
            const reader = new FileReader();
            reader.onload = async e => {
                try {
                    const wb=XLSX.read(e.target.result,{type:'array'});
                    const sheet=wb.Sheets[wb.SheetNames[0]];
                    const raw=XLSX.utils.sheet_to_json(sheet,{header:1});
                    let headerRow=0;
                    for(let i=0;i<Math.min(20,raw.length);i++){
                        const s=(raw[i]||[]).join('').toUpperCase();
                        if(s.includes('NOME')&&(s.includes('MATRICULA')||s.includes('MATRÍCULA'))){headerRow=i;break;}
                    }
                    window.dRH=XLSX.utils.sheet_to_json(sheet,{range:headerRow,raw:false});
                    _saveRHList(window.dRH);
                    if(window.saveToDB) await window.saveToDB('dRH',window.dRH);
                    window.renderRHQuad(); window.renderFerias(); window.renderBanco();
                    if(typeof renderAbs==='function') renderAbs();
                    if(typeof renderPontoMensal==='function') renderPontoMensal();
                    alert('Quadro importado!');
                } catch(err){alert('Erro: '+err);}
                document.getElementById('loading').style.display='none';
            };
            reader.readAsArrayBuffer(this.files[0]);
        });
    }

    const fBanco = document.getElementById('f_banco');
    if (fBanco) {
        fBanco.addEventListener('change', function () {
            if (typeof XLSX==='undefined'||!this.files[0]) return;
            document.getElementById('loading').style.display='flex';
            const reader=new FileReader();
            reader.onload = async e => {
                try {
                    const wb=XLSX.read(e.target.result,{type:'array'});
                    const raw=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{raw:false});
                    raw.forEach(row => {
                        let mat=row['Matrícula']||row['Matricula']||row['MATRICULA'];
                        const nm=row['Nome']||row['NOME']||row['NOME COMPLETO'];
                        const saldo=row['Total Banco']||row['TOTAL BANCO']||row['Saldo']||row['SALDO']||row['Horas']||row['Banco'];
                        if(!mat&&nm){const emp=window.dRH.find(e=>String(e['Nome']||e['NOME']).toUpperCase().trim()===String(nm).toUpperCase().trim());if(emp)mat=emp['Matrícula']||emp['Matricula'];}
                        if(mat&&saldo!==undefined){
                            ensureRhData(mat);
                            let v=String(saldo).trim(),num=0;
                            if(v.includes(':')){let sign=1;if(v.startsWith('-')){sign=-1;v=v.substring(1);}else if(v.startsWith('+'))v=v.substring(1);const pts=v.split(':');num=sign*((parseInt(pts[0])||0)+((parseInt(pts[1])||0)/60));}
                            else{num=parseFloat(v.replace(',','.').replace(/[^0-9.-]/g,''))||0;}
                            window.rhData[mat].banco=parseFloat(num.toFixed(2));
                        }
                    });
                    if(window.saveToDB) await window.saveToDB('rhData',window.rhData);
                    window.renderBanco(); alert('Banco de Horas importado!');
                } catch(err){alert('Erro: '+err);}
                document.getElementById('loading').style.display='none'; this.value='';
            };
            reader.readAsArrayBuffer(this.files[0]);
        });
    }

    // ── Import Férias: PDF ou Excel ───────────────────────────────────────
    const fFerias = document.getElementById('f_ferias');
    if (fFerias) {
        fFerias.addEventListener('change', function () {
            if (!this.files[0]) return;
            const file  = this.files[0];
            const isPDF = file.name.toLowerCase().endsWith('.pdf') || file.type === 'application/pdf';
            document.getElementById('loading').style.display = 'flex';
            if (document.getElementById('loadMsg')) document.getElementById('loadMsg').innerText = 'Lendo escala de férias...';

            const reader = new FileReader();
            reader.onload = async e => {
                try {
                    if (isPDF) await _importarEscalaFeiasPDF(e.target.result);
                    else       await _importarEscalaFeriasExcel(e.target.result);
                } catch(err) { alert('Erro ao importar: '+err.message); console.error(err); }
                document.getElementById('loading').style.display = 'none';
                this.value = '';
            };
            reader.readAsArrayBuffer(file);
        });
    }
});

// ── PARSER PDF ────────────────────────────────────────────────────────────
async function _importarEscalaFeiasPDF(arrayBuffer) {
    if (typeof pdfjsLib === 'undefined')
        throw new Error('pdf.js não carregado. Verifique os scripts no index.html.');

    pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';

    for (let p = 1; p <= pdf.numPages; p++) {
        const page    = await pdf.getPage(p);
        const content = await page.getTextContent();
        // Reconstrói linhas usando posição Y
        const items   = content.items;
        let lastY = null, lineStr = '';
        // pdf.js retorna items em ordem de renderização — precisamos agrupar por Y
        const byY = {};
        items.forEach(item => {
            const y = Math.round(item.transform[5]);
            if (!byY[y]) byY[y] = '';
            byY[y] += item.str + ' ';
        });
        // Ordena linhas por Y decrescente (pdf.js: Y=0 é base da página)
        Object.keys(byY).sort((a,b)=>b-a).forEach(y => {
            fullText += byY[y].trim() + '\n';
        });
    }

    return _processarTextoEscalaFerias(fullText);
}

async function _importarEscalaFeriasExcel(arrayBuffer) {
    if (typeof XLSX === 'undefined') throw new Error('Biblioteca Excel não carregada.');
    const wb   = XLSX.read(arrayBuffer, { type: 'array' });
    const raw  = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    const text = raw.map(r => (r||[]).map(c => String(c).trim()).join(' ')).join('\n');
    return _processarTextoEscalaFerias(text);
}

// ── PROCESSADOR CENTRAL DO TEXTO DA ESCALA ───────────────────────────────
// Padrão de linha do PDF Venancio:
// 0013108 FABIO MARTINS  09/09/2013 09/09/2024 DE 08/09/2025  10/06/2026 DE 29/06/2026 D  ...  09/07/2026
function _processarTextoEscalaFerias(fullText) {
    const lista = (window.dRH && window.dRH.length) ? window.dRH : _getRHList();

    const blocks = {};
    const lines  = fullText.split('\n');

    for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line || line.length < 5) continue;

        const matchMat = line.match(/^(\d{7})\b/);
        if (matchMat) {
            const key = matchMat[1];
            if (!blocks[key]) blocks[key] = '';
            blocks[key] += ' ' + line;
        } else {
            // Complemento de linha (nome muito longo quebrando ou datas extras)
            // Associa à última matrícula encontrada antes desta linha no texto
        }
    }

    // Segunda passagem: concatena linhas de continuação (sem matrícula) à última chave
    let lastKey = null;
    for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) continue;
        if (/^\d{7}\b/.test(line)) {
            lastKey = line.match(/^(\d{7})\b/)[1];
        } else if (lastKey && /\d{2}\/\d{2}\/\d{4}/.test(line)) {
            blocks[lastKey] = (blocks[lastKey]||'') + ' ' + line;
        }
    }

    let count = 0;

    lista.forEach(emp => {
        const matRh    = String(emp['Matrícula']||emp['Matricula']||emp['MATRICULA']||emp['matricula']||'').trim();
        const numMatRh = parseInt(matRh, 10);
        if (!matRh || isNaN(numMatRh)) return;

        let blockText = null;
        for (const key in blocks) {
            if (parseInt(key, 10) === numMatRh) { blockText = blocks[key]; break; }
        }
        if (!blockText) return;

        ensureRhData(matRh);

        // Extrai todas as datas do bloco
        const todasDatas = [];
        let m;
        const reDatas = /\b(\d{2}\/\d{2}\/\d{4})\b/g;
        while ((m = reDatas.exec(blockText)) !== null) todasDatas.push(m[1]);
        if (todasDatas.length === 0) return;

        // Última data = Limite para Início
        const ultimaData = todasDatas[todasDatas.length - 1];
        window.rhData[matRh].limiteFerias = ultimaData.split('/').reverse().join('-');

        // Detecta período aquisitivo (~365 dias)
        for (let i = 0; i < todasDatas.length - 1; i++) {
            const [d1, d2] = [todasDatas[i], todasDatas[i+1]].map(d => {
                const p=d.split('/'); return new Date(p[2],p[1]-1,p[0]);
            });
            const diff = Math.abs((d2-d1)/86400000);
            if (diff >= 355 && diff <= 375) {
                const dIn  = d1<d2 ? todasDatas[i] : todasDatas[i+1];
                const dFim = d1<d2 ? todasDatas[i+1] : todasDatas[i];
                window.rhData[matRh].periodoAberto = `${dIn} a ${dFim}`;
                break;
            }
        }

        // Detecta férias marcadas (14-32 dias)
        for (let i = 0; i < todasDatas.length - 1; i++) {
            const [d1, d2] = [todasDatas[i], todasDatas[i+1]].map(d => {
                const p=d.split('/'); return new Date(p[2],p[1]-1,p[0]);
            });
            const diff = Math.abs((d2-d1)/86400000);
            if (diff >= 14 && diff <= 32) {
                window.rhData[matRh].feriasInicio = (d1<d2 ? todasDatas[i] : todasDatas[i+1]).split('/').reverse().join('-');
                window.rhData[matRh].feriasFim    = (d1<d2 ? todasDatas[i+1] : todasDatas[i]).split('/').reverse().join('-');
                window.rhData[matRh].statusManual = 'AUTO';
                break;
            }
        }

        count++;
    });

    (async () => {
        if (window.saveToDB) await window.saveToDB('rhData', window.rhData);
        window.renderFerias();
        window.renderRHQuad();
        if (count === 0)
            alert('⚠️ Nenhuma correspondência encontrada.\n\nDica: importe o quadro de colaboradores primeiro.');
        else
            alert(`✅ Escala lida com sucesso!\n${count} colaboradores atualizados.`);
    })();
}

// =========================================================================
// FIX 3 — DASHBOARD RH completo e autônomo
// =========================================================================
window.renderDashboardRH = async function () {
    const btn = document.querySelector('[onclick*="renderDashboardRH"]');
    if (btn) { btn.disabled=true; btn.innerHTML='<i class="fas fa-circle-notch fa-spin"></i> Atualizando...'; }

    const set = (id, val, isHTML=false) => {
        const el = document.getElementById(id);
        if (!el) return;
        if (isHTML) el.innerHTML = '<strong><i class="fas fa-robot"></i> Análise IA:</strong> ' + val;
        else el.textContent = val;
    };

    try {
        const ano = new Date().getFullYear();
        set('drh-ano-label', ano);

        // Carrega dados com fallback
        let lista = [], rh = {};
        try {
            const [dRHfb, rhDataFb] = await Promise.all([
                window.getFromDB ? window.getFromDB('dRH').catch(()=>null) : Promise.resolve(null),
                window.getFromDB ? window.getFromDB('rhData').catch(()=>null) : Promise.resolve(null),
            ]);
            lista = (dRHfb&&dRHfb.length) ? dRHfb : ((window.dRH&&window.dRH.length)?window.dRH:_getRHList());
            rh    = rhDataFb || window.rhData || {};
        } catch(e) {
            lista = (window.dRH&&window.dRH.length) ? window.dRH : _getRHList();
            rh    = window.rhData || {};
        }

        if (!lista.length) {
            set('ai-colab-insight','⚠️ Nenhum colaborador carregado. Importe o quadro de colaboradores primeiro.',true);
            return;
        }

        const hoje = new Date(); hoje.setHours(0,0,0,0);

        // Colaboradores
        let ativos=0,inativos=0,entradasAno=0,supCount=0,liderCount=0,opCount=0;
        lista.forEach(c => {
            const mat  = c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';
            const func = (c['Função']||c['Funcao']||c['FUNÇÃO']||c['funcao']||'').toUpperCase();
            const s    = (c['Status']||c['STATUS']||c['status']||'ATIVO').toUpperCase();
            const rhM  = rh[mat]||{};
            const sf   = (rhM.statusManual&&rhM.statusManual!=='AUTO') ? rhM.statusManual.toUpperCase() : s;
            if(sf==='INATIVO'){inativos++;return;}
            ativos++;
            if(func.includes('SUPERVISOR'))supCount++;
            else if(func.includes('LIDER')||func.includes('LÍDER'))liderCount++;
            else opCount++;
            const admRaw=c['Admissão']||c['Admissao']||c['ADMISSÃO']||c['admissao']||'';
            let admY=0;
            if(admRaw.includes('/'))admY=parseInt(admRaw.split('/')[2]||0);
            else if(admRaw.includes('-'))admY=parseInt(admRaw.split('-')[0]||0);
            if(admY===ano)entradasAno++;
        });
        const turnover=inativos>0?((inativos/Math.max(1,ativos+inativos))*100).toFixed(1):'0.0';
        const pendRep=(typeof window.pendenciasRH!=='undefined')?window.pendenciasRH.length:0;
        set('drh-entradas',entradasAno); set('drh-saidas',inativos);
        set('drh-turnover',turnover+'%'); set('drh-ativos',ativos); set('drh-reposicoes',pendRep);
        let insC=`📊 <strong>${ativos} ativos</strong> | ${supCount} sup | ${liderCount} líd | ${opCount} op. `;
        insC+=parseFloat(turnover)>15?`🚨 Turnover elevado (${turnover}%).`:parseFloat(turnover)>8?`⚡ Moderado (${turnover}%).`:`✅ Controlado (${turnover}%).`;
        if(pendRep>0)insC+=` 🔴 <strong>${pendRep} vaga(s) em reposição.</strong>`;
        set('ai-colab-insight',insC,true);

        // Férias
        let fAtivas=0,fVencidas=0,em30=0;
        lista.forEach(c=>{
            const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';
            const r=rh[mat]||{};
            if(r.feriasInicio&&r.feriasFim){const fi=new Date(r.feriasInicio+'T00:00:00'),ff=new Date(r.feriasFim+'T00:00:00');if(hoje>=fi&&hoje<=ff)fAtivas++;}
            if(r.limiteFerias){const lim=new Date(r.limiteFerias+'T00:00:00'),d=(lim-hoje)/86400000;if(d<0&&!r.feriasInicio)fVencidas++;else if(d>=0&&d<=30&&!r.feriasInicio)em30++;}
        });
        let pendFer=0;
        try{if(typeof dbCloud!=='undefined'){const sn=await dbCloud.collection('logistica_ferias_inbox').where('status','==','PENDENTE').get();pendFer=sn.size;}}catch(e){}
        set('drh-ferias-ativas',fAtivas); set('drh-ferias-vencidas',fVencidas);
        set('drh-ferias-a-vencer',em30); set('drh-ferias-pendentes',pendFer);
        let insFer='';
        if(fVencidas>0)insFer+=`🚨 <strong>${fVencidas} vencidas</strong> — risco legal. `;
        if(em30>0)insFer+=`⏰ ${em30} a vencer em 30 dias. `;
        if(fAtivas>0)insFer+=`🌴 ${fAtivas} em gozo. `;
        if(pendFer>0)insFer+=`📬 ${pendFer} pendente(s). `;
        if(!insFer)insFer='✅ Sem irregularidades.';
        set('ai-ferias-insight',insFer,true);

        // Banco de Horas
        let totPos=0,totNeg=0,qtdPos=0,qtdNeg=0;
        Object.keys(rh).forEach(mat=>{const bh=typeof rh[mat].banco==='number'?rh[mat].banco:0;if(bh>0){totPos+=bh;qtdPos++;}else if(bh<0){totNeg+=Math.abs(bh);qtdNeg++;}});
        const fmtBH=h=>{const a=Math.abs(h);return(h<0?'-':'+')+String(Math.floor(a)).padStart(2,'0')+':'+String(Math.round((a-Math.floor(a))*60)).padStart(2,'0');};
        set('drh-bh-positivo',fmtBH(totPos)); set('drh-bh-negativo',fmtBH(-totNeg));
        set('drh-bh-qtd-pos',qtdPos); set('drh-bh-qtd-neg',qtdNeg);
        set('ai-bh-insight',(totPos>100?'⚠️ Saldo positivo elevado. ':qtdNeg>ativos*0.2?'🔴 +20% com saldo negativo. ':'')+'✅ OK',true);

        // Absenteísmo
        let totalAbs=0,diasAbs=0,motCount={},pesCount={},absMeses={};
        lista.forEach(c=>{
            const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';
            const nome=c['Nome']||c['NOME']||c['nome']||mat;
            ((rh[mat]||{}).faltas||[]).forEach(f=>{totalAbs++;const dias=parseInt(f.dias)||1;diasAbs+=dias;motCount[f.motivo||'Outro']=(motCount[f.motivo||'Outro']||0)+dias;pesCount[nome]=(pesCount[nome]||0)+dias;const mk=(f.data||'').substring(0,7)||'N/A';absMeses[mk]=(absMeses[mk]||0)+dias;});
        });
        const topMot=Object.keys(motCount).sort((a,b)=>motCount[b]-motCount[a])[0]||'—';
        const topPes=Object.keys(pesCount).sort((a,b)=>pesCount[b]-pesCount[a])[0]||'—';
        set('drh-abs-total',totalAbs); set('drh-abs-dias',diasAbs);
        set('drh-abs-motivo',topMot.length>22?topMot.slice(0,20)+'...':topMot);
        set('drh-abs-pessoa',topPes.length>22?topPes.slice(0,20)+'...':topPes);
        if(typeof renderChartAbsMotivo==='function')renderChartAbsMotivo(motCount);
        if(typeof renderChartAbsMensal2==='function')renderChartAbsMensal2(absMeses);
        const taxa=ativos>0?((diasAbs/(ativos*22))*100).toFixed(1):0;
        set('ai-abs-insight',`📊 Taxa: <strong>${taxa}%</strong>. `+(parseFloat(taxa)>4?`🚨 Acima do limite. Motivo: <strong>${topMot}</strong>.`:parseFloat(taxa)>2?'⚠️ Moderada.':'✅ Ok.'),true);

        // Ponto
        let pTotal=0,pFaltas=0,pMarcacoes=0,pOk=0,pontoTipos={};
        Object.keys(rh).forEach(mat=>{((rh[mat]||{}).ponto||[]).forEach(p=>{pTotal++;const tp=(p.tipo||'').toLowerCase(),ac=(p.acao||'').toLowerCase();if(tp.includes('falta sem'))pFaltas++;if(tp.includes('marcação')||tp.includes('marcacao'))pMarcacoes++;if(ac.includes('justificado')||ac.includes('arquivado'))pOk++;pontoTipos[p.tipo||'Outro']=(pontoTipos[p.tipo||'Outro']||0)+1;});});
        set('drh-ponto-total',pTotal); set('drh-ponto-faltas',pFaltas);
        set('drh-ponto-marcacoes',pMarcacoes); set('drh-ponto-ok',pOk);
        if(typeof renderChartPontoTipo==='function')renderChartPontoTipo(pontoTipos);
        set('ai-ponto-insight',pTotal>0?`🗓 <strong>${pTotal} ocorrência(s)</strong>.`+(pFaltas>5?` ⚠️ Faltas: ${pFaltas}.`:''):'Nenhuma ocorrência.',true);

        // EPI
        let epiTotal=0,epiEntg=0;
        try{if(typeof dbCloud!=='undefined'){const er=await dbCloud.collection('logistica_epi_registros').get();er.forEach(d=>{epiTotal++;if(d.data().status==='entregue')epiEntg++;});}}catch(e){}
        set('drh-epi-total',epiTotal||'—'); set('drh-epi-entregues',epiEntg||'—'); set('drh-epi-pendentes',epiTotal?(epiTotal-epiEntg):'—');

    } catch(e) {
        console.error('Dashboard RH:', e);
    } finally {
        if(btn){btn.disabled=false;btn.innerHTML='<i class="fas fa-sync-alt"></i> Atualizar Painel';}
    }
};

// ── Inicialização ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function () {
    setTimeout(async () => {
        try {
            let lista = [];
            if (window.getFromDB) {
                lista = (await window.getFromDB('dRH').catch(()=>null)) || [];
                const rhDb = (await window.getFromDB('rhData').catch(()=>null)) || {};
                if (lista.length) {
                    window.dRH    = lista;
                    window.rhData = Object.assign({}, rhDb, window.rhData);
                    _saveRHList(lista);
                    window.renderRHQuad();
                    window.renderFerias();
                    window.renderBanco();
                    if (typeof renderAbs==='function') renderAbs();
                    return;
                }
            }
            lista = _getRHList();
            if (lista.length) { window.dRH=lista; window.renderRHQuad(); }
        } catch(e) { console.warn('RH init:', e); }
    }, 1500);
});
