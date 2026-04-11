// =========================================================================
// MÓDULO RH v164.2 — FIXES DEFINITIVOS
// (1) Excluir colaborador — sobrescreve TODAS as versões anteriores
// (2) Import PDF de Férias — usa pdfjsLib (variável correta do CDN)
// (3) Dashboard RH — usa setDRH local, não depende do index.html
// =========================================================================

// ── Variáveis globais ─────────────────────────────────────────────────────
if (typeof window.dRH        === 'undefined') window.dRH        = [];
if (typeof window.rhData     === 'undefined') window.rhData     = {};
if (typeof window.pendenciasRH === 'undefined') window.pendenciasRH = [];

// ── Helpers localStorage ──────────────────────────────────────────────────
function _getRHList() {
    try { const r = localStorage.getItem('rh_colaboradores'); return r ? JSON.parse(r) : []; } catch(e) { return []; }
}
function _saveRHList(lista) {
    try { localStorage.setItem('rh_colaboradores', JSON.stringify(lista)); } catch(e) {}
    window.dRH = lista;
}
function ensureRhData(mat) {
    if (!window.rhData[mat]) {
        window.rhData[mat] = { feriasInicio:'', feriasFim:'', limiteFerias:'', periodoAberto:'',
            vendeu10Dias:false, banco:0, faltas:[], statusManual:'AUTO', ponto:[] };
    }
}
// ── setDRH LOCAL (não depende do patch do index.html) ─────────────────────
function _setDRH(id, val, isHTML) {
    const el = document.getElementById(id);
    if (!el) return;
    if (isHTML) el.innerHTML = '<strong><i class="fas fa-robot"></i> Análise IA:</strong> ' + val;
    else el.textContent = val;
}
// Expõe globalmente para o logica.js e index.html
window.setDRH = _setDRH;

// ── Fonte de dados unificada ──────────────────────────────────────────────
function _getLista() {
    if (window.dRH && window.dRH.length) return window.dRH;
    const ls = _getRHList();
    if (ls.length) { window.dRH = ls; }
    return window.dRH || [];
}

// =========================================================================
// QUADRO DE COLABORADORES
// =========================================================================
window.renderRHQuad = function () {
    const tb = document.querySelector('#tbRH');
    if (!tb) return;
    const lista = _getLista();

    if (!lista.length) {
        tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#999;padding:20px;">Nenhum colaborador registrado. Importe um arquivo Excel.</td></tr>';
        _atualizarKPIs([]);
        return;
    }

    let ativos=0, lideres=0, ops=0, sups=0, rows='';

    lista.forEach((r, index) => {
        const mat  = r['Matrícula']||r['Matricula']||r['MATRICULA']||r['matricula']||'-';
        const nome = r['Nome']||r['NOME']||r['nome']||'-';
        if (nome==='-' && mat==='-') return;

        const func = (r['Função']||r['Funcao']||r['FUNÇÃO']||r['funcao']||'-').toUpperCase();
        let baseSt = (r['Status']||r['STATUS']||r['status']||'ATIVO').toUpperCase();
        ensureRhData(mat);
        const rh = window.rhData[mat];
        let sf = (rh.statusManual && rh.statusManual!=='AUTO') ? rh.statusManual.toUpperCase() : baseSt;

        if (sf==='ATIVO' && rh.feriasInicio && rh.feriasFim) {
            const today=new Date(); today.setHours(0,0,0,0);
            if (today>=new Date(rh.feriasInicio+'T00:00:00') && today<=new Date(rh.feriasFim+'T00:00:00')) sf='FÉRIAS';
        }
        if (sf==='INATIVO') return;

        ativos++;
        if (func.includes('SUPERVISOR')||func.includes('SUPERVISORA')) sups++;
        else if (func.includes('LIDER')||func.includes('LÍDER')) lideres++;
        else if (func.includes('OPERADOR')||func.includes('SEPARADOR')) ops++;

        const statusColors = {
            'ATIVO':    {bg:'#e8f5e9',color:'#2e7d32',lbl:'ATIVO'},
            'FÉRIAS':   {bg:'#fff3cd',color:'#856404',lbl:'<i class="fas fa-umbrella-beach"></i> EM FÉRIAS'},
            'AFASTADO': {bg:'#fdf5e6',color:'#e67e22',lbl:'AFASTADO'},
        };
        let stBg='#ffebeb', stColor='#c0392b', stLbl=sf;
        if (statusColors[sf]) ({bg:stBg,color:stColor,lbl:stLbl}=statusColors[sf]);
        else if (sf.includes('TRANSFERENCIA')) {stBg='#e0f7fa';stColor='#00838f';stLbl='<i class="fas fa-exchange-alt"></i> TRANSFERIDO';}

        const cel = r['celular']||r['Celular']||r['CELULAR']||'';
        const celHtml = cel ? `<a href="https://wa.me/${cel.replace(/\D/g,'')}" target="_blank" style="color:#25d366;font-size:11px;"><i class="fab fa-whatsapp"></i> ${cel}</a>` : '—';

        rows += `<tr>
            <td style="font-weight:bold;text-align:center;">${mat}</td>
            <td style="font-weight:600;">${nome}</td>
            <td style="color:#555;">${r['Função']||r['Funcao']||r['FUNÇÃO']||r['funcao']||'-'}</td>
            <td style="text-align:center;font-size:11px;">${celHtml}</td>
            <td style="text-align:center;"><span style="background:${stBg};color:${stColor};padding:4px 10px;border-radius:12px;font-weight:900;font-size:10px;">${stLbl}</span></td>
            <td style="text-align:center;white-space:nowrap;">
                <button class="btn btn-blue" style="padding:4px 8px;font-size:11px;margin-right:3px;" onclick="abrirEdicaoRH(${index})"><i class="fas fa-edit"></i> Editar</button>
                <button class="btn-del-colab" style="padding:4px 8px;font-size:11px;background:#c0392b;color:white;border:none;border-radius:4px;cursor:pointer;" onclick="excluirColaboradorRH('${mat}')"><i class="fas fa-trash-alt"></i> Excluir</button>
            </td>
        </tr>`;
    });

    tb.innerHTML = rows || '<tr><td colspan="6" style="text-align:center;color:#999;padding:20px;">Nenhum ativo.</td></tr>';
    _atualizarKPIs({ativos,lideres,ops,sups,total:lista.length});
};

function _atualizarKPIs(data) {
    if (Array.isArray(data)) {
        let a=0,l=0,o=0,s=0;
        data.forEach(c=>{const st=(c.status||c.STATUS||c.statusManual||'ATIVO').toUpperCase();if(st==='INATIVO')return;a++;const f=(c.funcao||c.FUNCAO||c['Função']||c['FUNÇÃO']||'').toUpperCase();if(f.includes('SUPERVISOR'))s++;else if(f.includes('LIDER')||f.includes('LÍDER'))l++;else if(f.includes('OPERADOR')||f.includes('SEPARADOR'))o++;});
        data={ativos:a,lideres:l,ops:o,sups:s,total:data.length};
    }
    const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
    set('rh-tot',data.ativos||0); set('rh-lid',data.lideres||0); set('rh-op',data.ops||0); set('rh-sup',data.sups||0);
    const kpis=document.getElementById('rh-kpis');
    if(kpis&&(data.total||0)>0)kpis.style.display='grid';
}

// =====================================================
// EXCLUIR COLABORADOR DEFINITIVO (CORRIGIDO)
// =====================================================
// =========================================================================
// MODAIS
// =========================================================================
window.abrirModalRH = function () {
    document.getElementById('rhEditIndex').value=-1;
    ['rhMat','rhNome','rhAdm','rhCelular'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
    if(document.getElementById('rhFunc'))   document.getElementById('rhFunc').value='OPERADOR';
    if(document.getElementById('rhStatus')) document.getElementById('rhStatus').value='AUTO';
    document.getElementById('modalRHTitle').innerHTML='<i class="fas fa-user-plus"></i> Novo Colaborador';
    document.getElementById('modalRH').style.display='flex';
};
window.abrirEdicaoRH = function (index) {
    const lista=_getLista(); const c=lista[index]; if(!c)return;
    const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';
    document.getElementById('rhEditIndex').value=index;
    if(document.getElementById('rhMat'))     document.getElementById('rhMat').value=mat;
    if(document.getElementById('rhNome'))    document.getElementById('rhNome').value=c['Nome']||c['NOME']||c['nome']||'';
    if(document.getElementById('rhFunc'))    document.getElementById('rhFunc').value=c['Função']||c['Funcao']||c['FUNÇÃO']||c['funcao']||'';
    if(document.getElementById('rhCelular')) document.getElementById('rhCelular').value=c['celular']||c['Celular']||c['CELULAR']||'';
    let admRaw=c['Admissão']||c['Admissao']||c['admissao']||''; let admDate='';
    if(admRaw.includes('/')){const p=admRaw.split('/');if(p.length===3)admDate=`${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;}
    else if(admRaw.includes('-'))admDate=admRaw;
    if(document.getElementById('rhAdm'))document.getElementById('rhAdm').value=admDate;
    ensureRhData(mat);
    if(document.getElementById('rhStatus'))document.getElementById('rhStatus').value=window.rhData[mat].statusManual||'AUTO';
    document.getElementById('modalRHTitle').innerHTML='<i class="fas fa-user-edit"></i> Editar Colaborador';
    document.getElementById('modalRH').style.display='flex';
};
window.fecharModalRH = function() { document.getElementById('modalRH').style.display='none'; };
window.salvarColaboradorRH = async function () {
    const idx=parseInt(document.getElementById('rhEditIndex').value);
    const mat=(document.getElementById('rhMat')?.value||'').trim();
    const nome=(document.getElementById('rhNome')?.value||'').trim().toUpperCase();
    const func=(document.getElementById('rhFunc')?.value||'').trim().toUpperCase();
    const adm=document.getElementById('rhAdm')?.value||'';
    const status=document.getElementById('rhStatus')?.value||'AUTO';
    const celular=(document.getElementById('rhCelular')?.value||'').trim();
    if(!mat||!nome)return alert('Matrícula e Nome são obrigatórios!');
    let admF=adm; if(adm){const d=new Date(adm+'T12:00:00');admF=d.toLocaleDateString('pt-BR');}
    const lista=_getLista();
    if(idx===-1){lista.push({'Matrícula':mat,'Nome':nome,'Função':func,'Admissão':admF,'Status':'ATIVO','celular':celular});}
    else{lista[idx]['Matrícula']=mat;lista[idx]['Nome']=nome;lista[idx]['Função']=func;lista[idx]['Admissão']=admF;lista[idx]['celular']=celular;}
    window.dRH=lista; _saveRHList(lista);
    ensureRhData(mat); window.rhData[mat].statusManual=status;
    if(window.saveToDB){window.saveToDB('dRH',window.dRH);window.saveToDB('rhData',window.rhData);}
    window.fecharModalRH(); window.renderRHQuad();
    if(typeof window.renderFerias==='function')window.renderFerias();
    if(typeof renderBanco==='function')renderBanco();
    if(typeof showToast==='function')showToast('Colaborador salvo!','success');
};

// =========================================================================
// RENDERIZAR FÉRIAS / BANCO
// =========================================================================
window.renderFerias = function () {
    const tb=document.getElementById('tbFerias'); if(!tb)return;
    const lista=_getLista(); if(!lista.length)return;
    tb.innerHTML='';
    lista.forEach(r=>{
        const mat=r['Matrícula']||r['Matricula']||r['MATRICULA']||r['matricula']||'-';
        const nome=r['Nome']||r['NOME']||r['nome']||'-';
        const adm=r['Admissão']||r['Admissao']||r['ADMISSÃO']||r['admissao']||'-';
        ensureRhData(mat);
        const rh=window.rhData[mat];
        const sm=rh.statusManual||'AUTO', bs=(r['Status']||r['STATUS']||r['status']||'-').toUpperCase();
        if(nome==='-'||(bs==='INATIVO'&&sm!=='ATIVO')||sm==='INATIVO'||sm==='TRANSFERENCIA DE TURNO')return;
        let badge='<span style="background:#3498db;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">AQUISITIVO</span>';
        const dtI=rh.feriasInicio||'',dtF=rh.feriasFim||'',hasM=(dtI&&dtF);
        if(hasM){const today=new Date();today.setHours(0,0,0,0);const dI=new Date(dtI+'T00:00:00'),dF=new Date(dtF+'T00:00:00');if(today>=dI&&today<=dF)badge='<span style="background:#8e44ad;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">EM FÉRIAS</span>';else if(dI>today)badge='<span style="background:#f1c40f;color:#333;padding:3px 6px;border-radius:4px;font-weight:bold;">MARCADA</span>';else badge='<span style="background:#7f8c8d;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">GOZADAS</span>';}
        else if(rh.limiteFerias){const today=new Date();today.setHours(0,0,0,0);const lim=new Date(rh.limiteFerias+'T00:00:00'),d=Math.floor((lim-today)/86400000);if(d<0)badge='<span style="background:#c0392b;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">VENCIDA</span>';else if(d<=90)badge='<span style="background:#f39c12;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">VENCENDO</span>';else badge='<span style="background:#27ae60;color:white;padding:3px 6px;border-radius:4px;font-weight:bold;">NO PRAZO</span>';}
        let limDisp='-';
        if(rh.limiteFerias){const ls=rh.limiteFerias.split('-').reverse().join('/');const today=new Date();today.setHours(0,0,0,0);limDisp=new Date(rh.limiteFerias+'T00:00:00')<today&&!hasM?`<strong style="color:#c0392b;font-size:12px;">${ls}</strong>`:`<strong style="color:var(--dark);font-size:12px;">${ls}</strong>`;}
        const pAberto=rh.periodoAberto?`<strong style="color:var(--blue);font-size:11px;">${rh.periodoAberto}</strong>`:'<span style="color:#ccc;">-</span>';
        tb.innerHTML+=`<tr><td style="font-weight:bold;font-size:12px;">${mat}</td><td style="text-align:left;font-size:12px;">${nome}</td><td style="font-size:11px;color:#555;">${adm}</td><td>${pAberto}</td><td>${limDisp}</td><td>${badge}</td><td><input type="date" value="${dtI}" onchange="updateRHData('${mat}','feriasInicio',this.value,this);window.renderRHQuad();window.renderFerias();" style="padding:4px;border:1px solid #ccc;border-radius:4px;font-size:10px;"></td><td><input type="date" value="${dtF}" onchange="updateRHData('${mat}','feriasFim',this.value,this);window.renderRHQuad();window.renderFerias();" style="padding:4px;border:1px solid #ccc;border-radius:4px;font-size:10px;"></td><td><label style="cursor:pointer;font-weight:800;color:var(--orange);font-size:10px;display:flex;align-items:center;gap:5px;justify-content:center;"><input type="checkbox" ${rh.vendeu10Dias?'checked':''} onchange="updateRHData('${mat}','vendeu10Dias',this.checked,this)"> Vendeu 10 Dias</label></td><td><button class="btn-sm" style="background:#25D366;color:white;padding:5px;border-radius:20px;" onclick="sendWhatsAppLink('${mat}','${nome}')"><i class="fab fa-whatsapp"></i></button></td></tr>`;
    });
};
// =========================================================================
// RELATÓRIO DE FÉRIAS EM PDF (para impressão)
// =========================================================================
window.gerarRelatorioPDF = function () {
    const lista = _getLista();
    if (!lista.length) { alert('Nenhum colaborador para gerar relatório.'); return; }

    const hoje   = new Date();
    const dHoje  = hoje.toLocaleDateString('pt-BR');
    const ano    = hoje.getFullYear();

    // Monta linhas da tabela — apenas MARCADA e EM FÉRIAS
    let linhas = '';
    let seq = 1;
    const hj = new Date(); hj.setHours(0,0,0,0);

    lista.forEach(r => {
        const mat  = r['Matrícula']||r['Matricula']||r['MATRICULA']||r['matricula']||'-';
        const nome = r['Nome']||r['NOME']||r['nome']||'-';
        const func = r['Função']||r['Funcao']||r['FUNÇÃO']||r['funcao']||'-';
        const st   = (r['Status']||r['STATUS']||r['status']||'ATIVO').toUpperCase();
        if (nome==='-' || st==='INATIVO') return;

        ensureRhData(mat);
        const rh = window.rhData[mat] || {};
        const sm = rh.statusManual || 'AUTO';
        if (sm === 'INATIVO' || sm === 'TRANSFERENCIA DE TURNO') return;

        // ── Filtro: somente quem tem férias MARCADAS ou EM GOZO ──────────
        // Exige que inicio E fim estejam preenchidos
        if (!rh.feriasInicio || !rh.feriasFim) return;

        const dI = new Date(rh.feriasInicio+'T00:00:00');
        const dF = new Date(rh.feriasFim+'T00:00:00');

        let stLabel;
        if (hj >= dI && hj <= dF) {
            stLabel = 'EM FÉRIAS';       // em gozo hoje
        } else if (dI > hj) {
            stLabel = 'MARCADA';         // marcada para o futuro
        } else {
            return; // já gozadas — não entra no relatório
        }
        // ─────────────────────────────────────────────────────────────────

        const ini = formatDateBR(rh.feriasInicio);
        const fim = formatDateBR(rh.feriasFim);
        const lim = rh.limiteFerias ? formatDateBR(rh.limiteFerias) : '-';
        const per = rh.periodoAberto || '-';
        const v10 = rh.vendeu10Dias ? 'Sim' : 'Não';

        // Cor da linha por status
        const rowBg  = stLabel === 'EM FÉRIAS' ? '#e8f8f5' : (seq%2===0 ? '#f9f9f9' : '#ffffff');
        const stColor= stLabel === 'EM FÉRIAS' ? '#0f6e56' : '#185FA5';

        linhas += `<tr style="background:${rowBg};">
            <td style="text-align:center;">${seq++}</td>
            <td style="text-align:center;font-weight:bold;">${mat}</td>
            <td>${nome}</td>
            <td>${func}</td>
            <td style="font-size:10px;">${per}</td>
            <td style="text-align:center;">${lim}</td>
            <td style="text-align:center;font-weight:bold;">${ini}</td>
            <td style="text-align:center;font-weight:bold;">${fim}</td>
            <td style="text-align:center;">${v10}</td>
            <td style="text-align:center;font-size:10px;font-weight:bold;color:${stColor};">${stLabel}</td>
        </tr>`;
    });

    if (!linhas) {
        alert('Nenhuma férias marcada ou em andamento encontrada para gerar o relatório.\n\nImporte a escala de férias primeiro.');
        return;
    }

    const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório de Férias — ${ano}</title>
<style>
  @page { size: A4 landscape; margin: 15mm 12mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #222; background: #fff; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #333; padding-bottom: 8px; margin-bottom: 10px; }
  .header-left h1 { font-size: 16px; font-weight: bold; color: #1a237e; }
  .header-left p  { font-size: 10px; color: #555; margin-top: 3px; }
  .header-right   { text-align: right; font-size: 10px; color: #555; }
  table { width: 100%; border-collapse: collapse; margin-top: 6px; }
  th { background: #1a237e; color: #fff; padding: 6px 5px; text-align: center; font-size: 10px; font-weight: bold; border: 1px solid #1a237e; }
  td { padding: 5px 4px; border: 1px solid #ddd; font-size: 10px; vertical-align: middle; }
  tr:hover { background: #e8eaf6 !important; }
  .footer { margin-top: 18px; border-top: 1px solid #ccc; padding-top: 10px; }
  .footer-obs { font-size: 10px; color: #333; line-height: 1.6; }
  .footer-obs .label { font-weight: bold; font-size: 10px; color: #c0392b; text-transform: uppercase; }
  .footer-aviso { margin-top: 10px; background: #fff8e1; border: 1px solid #f9a825; border-radius: 4px; padding: 8px 12px; font-size: 10px; line-height: 1.6; }
  .footer-aviso .label { font-weight: bold; color: #e65100; }
  .assinaturas { margin-top: 30px; display: flex; gap: 40px; }
  .assinatura-box { flex: 1; text-align: center; }
  .assinatura-box .linha { border-top: 1px solid #333; margin-bottom: 4px; }
  .assinatura-box p { font-size: 10px; color: #555; }
  @media print {
    .no-print { display: none; }
    tr { page-break-inside: avoid; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>📋 Relatório de Escala de Férias — ${ano}</h1>
    <p>Módulo Recursos Humanos · Gerado em: ${dHoje}</p>
  </div>
  <div class="header-right">
    <strong>Data de emissão:</strong> ${dHoje}<br>
    <strong>Exercício:</strong> ${ano}
  </div>
</div>

<table>
  <thead>
    <tr>
      <th style="width:30px;">#</th>
      <th style="width:70px;">Matrícula</th>
      <th style="width:18%;">Nome</th>
      <th style="width:14%;">Função</th>
      <th style="width:16%;">Período Aquisitivo</th>
      <th style="width:8%;">Limite</th>
      <th style="width:8%;">Início Férias</th>
      <th style="width:8%;">Fim Férias</th>
      <th style="width:7%;">Vendeu 10d</th>
      <th style="width:9%;">Situação</th>
    </tr>
  </thead>
  <tbody>
    ${linhas || '<tr><td colspan="10" style="text-align:center;padding:12px;color:#999;">Nenhum colaborador ativo.</td></tr>'}
  </tbody>
</table>

<div class="footer">
  <div class="footer-obs">
    <span class="label">⚠ Observação:</span>
    As férias poderão sofrer alterações até a data limite.
  </div>

  <div class="footer-aviso">
    <span class="label">📱 Aviso aos Colaboradores:</span><br>
    Acompanhe suas férias pelo aplicativo <strong>Meu RH</strong>. No prazo de até <strong>40 dias antes do início das férias</strong>,
    verifique se elas foram confirmadas no sistema. Caso não estejam marcadas, procure imediatamente
    sua <strong>liderança</strong> para regularização.
  </div>

  <div class="assinaturas">
    <div class="assinatura-box">
      <div class="linha"></div>
      <p>Responsável RH</p>
    </div>
    <div class="assinatura-box">
      <div class="linha"></div>
      <p>Gestor de Operações</p>
    </div>
    <div class="assinatura-box">
      <div class="linha"></div>
      <p>Diretoria</p>
    </div>
  </div>
</div>

<script>
  window.onload = function() { window.print(); };
<\/script>
</body>
</html>`;

    const blob = new Blob([html], { type: 'text/html; charset=utf-8' });
    const url  = URL.createObjectURL(blob);
    const win  = window.open(url, '_blank');
    if (!win) alert('Permita pop-ups para gerar o relatório em PDF.');
};

window.updateRHData = async function(mat,field,val,inputEl){ensureRhData(mat);window.rhData[mat][field]=val;if(window.saveToDB)window.saveToDB('rhData',window.rhData);if(inputEl&&inputEl.type!=='checkbox'){inputEl.style.backgroundColor='#d4edda';setTimeout(()=>{inputEl.style.backgroundColor='';},1000);}};
window.renderBanco = function(){const tb=document.getElementById('tbBanco');if(!tb)return;const lista=_getLista();if(!lista.length)return;tb.innerHTML='';let tot=0;lista.forEach(r=>{const mat=r['Matrícula']||r['Matricula']||r['MATRICULA']||r['matricula']||'-';const nome=r['Nome']||r['NOME']||r['nome']||'-';const st=(r['Status']||r['STATUS']||r['status']||'-').toUpperCase();if(nome==='-'||st==='INATIVO')return;ensureRhData(mat);const h=window.rhData[mat].banco||0;tot+=h;const c=h>0?'var(--green)':(h<0?'var(--red)':'#555');tb.innerHTML+=`<tr><td style="font-weight:bold;">${mat}</td><td>${nome}</td><td><strong style="color:${c};font-size:16px;">${formatDecimalToTime(h)}</strong></td></tr>`;});if(typeof safeUpdate==='function')safeUpdate('bh-global',formatDecimalToTime(tot));};

// =====================================================
// NOVO LEITOR E IMPORTADOR DE PDF DE FÉRIAS
// =====================================================
// ── PROCESSADOR CENTRAL ───────────────────────────────────────────────────
function _processarEscalaFerias(texto) {
    const lista = _getLista();
    const blocks = {};
    let lastKey  = null;

    texto.split('\n').forEach(rawLine => {
        const line = rawLine.trim();
        if (!line || line.length < 5) return;
        const matchMat = line.match(/^(\d{7})\b/);
        if (matchMat) {
            lastKey = matchMat[1];
            blocks[lastKey] = (blocks[lastKey]||'') + ' ' + line;
        } else if (lastKey && /\d{2}\/\d{2}\/\d{4}/.test(line)) {
            blocks[lastKey] += ' ' + line;
        }
    });

    let count = 0;

    lista.forEach(emp => {
        const matRh    = String(emp['Matrícula']||emp['Matricula']||emp['MATRICULA']||emp['matricula']||'').trim();
        const numMatRh = parseInt(matRh, 10);
        if (!matRh || isNaN(numMatRh)) return;

        let blockText = null;
        for (const key in blocks) {
            if (parseInt(key,10) === numMatRh) { blockText = blocks[key]; break; }
        }
        if (!blockText) return;

        ensureRhData(matRh);
        const todasDatas = [];
        let m;
        const re = /\b(\d{2}\/\d{2}\/\d{4})\b/g;
        while ((m=re.exec(blockText))!==null) todasDatas.push(m[1]);
        if (!todasDatas.length) return;

        // Última data = Limite para Início
        window.rhData[matRh].limiteFerias = todasDatas[todasDatas.length-1].split('/').reverse().join('-');

        // Detecta período aquisitivo (~365 dias)
        for (let i=0;i<todasDatas.length-1;i++) {
            const [d1,d2]=[todasDatas[i],todasDatas[i+1]].map(d=>{const p=d.split('/');return new Date(p[2],p[1]-1,p[0]);});
            const diff=Math.abs((d2-d1)/86400000);
            if(diff>=355&&diff<=375){
                const ini=d1<d2?todasDatas[i]:todasDatas[i+1];
                const fim=d1<d2?todasDatas[i+1]:todasDatas[i];
                window.rhData[matRh].periodoAberto=`${ini} a ${fim}`;
                break;
            }
        }

        // Detecta férias marcadas (14-32 dias)
        for (let i=0;i<todasDatas.length-1;i++) {
            const [d1,d2]=[todasDatas[i],todasDatas[i+1]].map(d=>{const p=d.split('/');return new Date(p[2],p[1]-1,p[0]);});
            const diff=Math.abs((d2-d1)/86400000);
            if(diff>=14&&diff<=32){
                window.rhData[matRh].feriasInicio=(d1<d2?todasDatas[i]:todasDatas[i+1]).split('/').reverse().join('-');
                window.rhData[matRh].feriasFim   =(d1<d2?todasDatas[i+1]:todasDatas[i]).split('/').reverse().join('-');
                window.rhData[matRh].statusManual='AUTO';
                break;
            }
        }
        count++;
    });

    (async()=>{
        if(window.saveToDB) await window.saveToDB('rhData',window.rhData);
        window.renderFerias();
        window.renderRHQuad();
        if(count===0) alert('⚠️ Nenhuma correspondência encontrada.\n\nDica: importe o quadro de colaboradores primeiro (aba Colaboradores → Importar Excel).');
        else alert(`✅ Escala lida!\n${count} colaboradores atualizados.`);
    })();
}

// =========================================================================
// FIX 3 — DASHBOARD RH com análise via API Claude (IA real)
// =========================================================================
window.renderDashboardRH = async function () {
    const btn=document.querySelector('[onclick*="renderDashboardRH"]');
    if(btn){btn.disabled=true;btn.innerHTML='<i class="fas fa-circle-notch fa-spin"></i> Atualizando...';}
    try {
        const ano=new Date().getFullYear();
        _setDRH('drh-ano-label',ano);

        let lista=[], rh={};
        try {
            const [dRHfb,rhDataFb]=await Promise.all([
                window.getFromDB?window.getFromDB('dRH').catch(()=>null):Promise.resolve(null),
                window.getFromDB?window.getFromDB('rhData').catch(()=>null):Promise.resolve(null),
            ]);
            lista=(dRHfb&&dRHfb.length)?dRHfb:_getLista();
            rh=rhDataFb||window.rhData||{};
        }catch(e){lista=_getLista();rh=window.rhData||{};}

        if(!lista.length){_setDRH('ai-colab-insight','⚠️ Importe o quadro de colaboradores primeiro.',true);return;}

        const hoje=new Date();hoje.setHours(0,0,0,0);
        let ativos=0,inativos=0,entradas=0,sups=0,lids=0,ops=0;
        lista.forEach(c=>{
            const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';
            const func=(c['Função']||c['Funcao']||c['FUNÇÃO']||c['funcao']||'').toUpperCase();
            const s=(c['Status']||c['STATUS']||c['status']||'ATIVO').toUpperCase();
            const rhM=rh[mat]||{};
            const sf=(rhM.statusManual&&rhM.statusManual!=='AUTO')?rhM.statusManual.toUpperCase():s;
            if(sf==='INATIVO'){inativos++;return;}
            ativos++;
            if(func.includes('SUPERVISOR'))sups++;else if(func.includes('LIDER')||func.includes('LÍDER'))lids++;else ops++;
            const admRaw=c['Admissão']||c['Admissao']||c['ADMISSÃO']||c['admissao']||'';
            let admY=0;if(admRaw.includes('/'))admY=parseInt(admRaw.split('/')[2]||0);else if(admRaw.includes('-'))admY=parseInt(admRaw.split('-')[0]||0);
            if(admY===ano)entradas++;
        });
        const turnover=inativos>0?((inativos/Math.max(1,ativos+inativos))*100).toFixed(1):'0.0';
        const pendRep=(typeof window.pendenciasRH!=='undefined')?window.pendenciasRH.length:0;
        _setDRH('drh-entradas',entradas);_setDRH('drh-saidas',inativos);_setDRH('drh-turnover',turnover+'%');_setDRH('drh-ativos',ativos);_setDRH('drh-reposicoes',pendRep);

        // Férias
        let fAtv=0,fVenc=0,em30=0;
        lista.forEach(c=>{const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';const r=rh[mat]||{};if(r.feriasInicio&&r.feriasFim){const fi=new Date(r.feriasInicio+'T00:00:00'),ff=new Date(r.feriasFim+'T00:00:00');if(hoje>=fi&&hoje<=ff)fAtv++;}if(r.limiteFerias){const lim=new Date(r.limiteFerias+'T00:00:00'),d=(lim-hoje)/86400000;if(d<0&&!r.feriasInicio)fVenc++;else if(d>=0&&d<=30&&!r.feriasInicio)em30++;}});
        let pendFer=0;try{if(typeof dbCloud!=='undefined'){const sn=await dbCloud.collection('logistica_ferias_inbox').where('status','==','PENDENTE').get();pendFer=sn.size;}}catch(e){}
        _setDRH('drh-ferias-ativas',fAtv);_setDRH('drh-ferias-vencidas',fVenc);_setDRH('drh-ferias-a-vencer',em30);_setDRH('drh-ferias-pendentes',pendFer);

        // Banco de Horas
        let totPos=0,totNeg=0,qtdPos=0,qtdNeg=0;
        Object.keys(rh).forEach(mat=>{const bh=typeof rh[mat].banco==='number'?rh[mat].banco:0;if(bh>0){totPos+=bh;qtdPos++;}else if(bh<0){totNeg+=Math.abs(bh);qtdNeg++;}});
        const fmtBH=h=>{const a=Math.abs(h);return(h<0?'-':'+')+String(Math.floor(a)).padStart(2,'0')+':'+String(Math.round((a-Math.floor(a))*60)).padStart(2,'0');};
        _setDRH('drh-bh-positivo',fmtBH(totPos));_setDRH('drh-bh-negativo',fmtBH(-totNeg));_setDRH('drh-bh-qtd-pos',qtdPos);_setDRH('drh-bh-qtd-neg',qtdNeg);

        // Absenteísmo
        let totAbs=0,diasAbs=0,motCount={},pesCount={},absMeses={};
        lista.forEach(c=>{const mat=c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'';const nome=c['Nome']||c['NOME']||c['nome']||mat;((rh[mat]||{}).faltas||[]).forEach(f=>{totAbs++;const dias=parseInt(f.dias)||1;diasAbs+=dias;motCount[f.motivo||'Outro']=(motCount[f.motivo||'Outro']||0)+dias;pesCount[nome]=(pesCount[nome]||0)+dias;const mk=(f.data||'').substring(0,7)||'N/A';absMeses[mk]=(absMeses[mk]||0)+dias;});});
        const topMot=Object.keys(motCount).sort((a,b)=>motCount[b]-motCount[a])[0]||'—';
        const topPes=Object.keys(pesCount).sort((a,b)=>pesCount[b]-pesCount[a])[0]||'—';
        _setDRH('drh-abs-total',totAbs);_setDRH('drh-abs-dias',diasAbs);_setDRH('drh-abs-motivo',topMot.length>22?topMot.slice(0,20)+'...':topMot);_setDRH('drh-abs-pessoa',topPes.length>22?topPes.slice(0,20)+'...':topPes);
        if(typeof renderChartAbsMotivo==='function')renderChartAbsMotivo(motCount);
        if(typeof renderChartAbsMensal2==='function')renderChartAbsMensal2(absMeses);

        const taxa=ativos>0?((diasAbs/(ativos*22))*100).toFixed(1):0;

        // Ponto
        let pTot=0,pFlt=0,pMrc=0,pOk=0,pontoTipos={};
        Object.keys(rh).forEach(mat=>{((rh[mat]||{}).ponto||[]).forEach(p=>{pTot++;const tp=(p.tipo||'').toLowerCase(),ac=(p.acao||'').toLowerCase();if(tp.includes('falta sem'))pFlt++;if(tp.includes('marcação')||tp.includes('marcacao'))pMrc++;if(ac.includes('justificado')||ac.includes('arquivado'))pOk++;pontoTipos[p.tipo||'Outro']=(pontoTipos[p.tipo||'Outro']||0)+1;});});
        _setDRH('drh-ponto-total',pTot);_setDRH('drh-ponto-faltas',pFlt);_setDRH('drh-ponto-marcacoes',pMrc);_setDRH('drh-ponto-ok',pOk);
        if(typeof renderChartPontoTipo==='function')renderChartPontoTipo(pontoTipos);

        // EPI
        let epiTot=0,epiEntg=0;
        try{if(typeof dbCloud!=='undefined'){const er=await dbCloud.collection('logistica_epi_registros').get();er.forEach(d=>{epiTot++;if(d.data().status==='entregue')epiEntg++;});}}catch(e){}
        _setDRH('drh-epi-total',epiTot||'—');_setDRH('drh-epi-entregues',epiEntg||'—');_setDRH('drh-epi-pendentes',epiTot?(epiTot-epiEntg):'—');

        // ── Análise IA via API Claude ─────────────────────────────────────
        _setDRH('ai-colab-insight','<i class="fas fa-circle-notch fa-spin"></i> Analisando com IA…',true);
        _setDRH('ai-ferias-insight','<i class="fas fa-circle-notch fa-spin"></i> Analisando com IA…',true);
        _setDRH('ai-bh-insight','<i class="fas fa-circle-notch fa-spin"></i> Analisando com IA…',true);
        _setDRH('ai-abs-insight','<i class="fas fa-circle-notch fa-spin"></i> Analisando com IA…',true);
        _setDRH('ai-ponto-insight','<i class="fas fa-circle-notch fa-spin"></i> Analisando com IA…',true);

        try {
            const dadosRH = {
                colaboradores: { ativos, inativos, entradas, sups, lids, ops, turnover: parseFloat(turnover), pendentesReposicao: pendRep },
                ferias: { emGozo: fAtv, vencidas: fVenc, aVencer30dias: em30, pendentes: pendFer },
                bancoHoras: { totalPositivo: totPos.toFixed(1), totalNegativo: totNeg.toFixed(1), qtdPositivo: qtdPos, qtdNegativo: qtdNeg },
                absenteismo: { totalOcorrencias: totAbs, totalDias: diasAbs, taxaPercent: parseFloat(taxa), principalMotivo: topMot, colaboradorMaisAusente: topPes },
                ponto: { totalOcorrencias: pTot, faltasSemJust: pFlt, marcacoesErradas: pMrc, resolvidos: pOk }
            };

            const resp = await fetch('https://api.anthropic.com/v1/messages', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    model: 'claude-sonnet-4-20250514',
                    max_tokens: 1000,
                    messages: [{
                        role: 'user',
                        content: `Você é um analista de RH. Analise os dados abaixo e gere insights curtos e objetivos em português (máximo 2 frases por seção). Use emojis moderadamente. Seja direto e prático.

Dados: ${JSON.stringify(dadosRH)}

Responda SOMENTE em JSON com as chaves: colaboradores, ferias, bancoHoras, absenteismo, ponto.
Cada valor deve ser HTML inline simples (sem tags block). Exemplo:
{"colaboradores":"✅ Equipe estável com ${ativos} ativos. Turnover de ${turnover}% está dentro do esperado.", "ferias":"...", "bancoHoras":"...", "absenteismo":"...", "ponto":"..."}`
                    }]
                })
            });

            if (resp.ok) {
                const data = await resp.json();
                const txt  = (data.content||[]).map(b=>b.text||'').join('');
                const clean = txt.replace(/```json|```/g,'').trim();
                const insights = JSON.parse(clean);
                _setDRH('ai-colab-insight', insights.colaboradores || '—', true);
                _setDRH('ai-ferias-insight', insights.ferias || '—', true);
                _setDRH('ai-bh-insight', insights.bancoHoras || '—', true);
                _setDRH('ai-abs-insight', insights.absenteismo || '—', true);
                _setDRH('ai-ponto-insight', insights.ponto || '—', true);
            } else {
                throw new Error('HTTP ' + resp.status);
            }
        } catch(iaErr) {
            console.warn('[Dashboard RH] IA indisponível, usando análise local:', iaErr);
            // Fallback: análise local sem IA
            let insC=`📊 <strong>${ativos} ativos</strong> | ${sups} sup | ${lids} líd | ${ops} op. `;
            insC+=parseFloat(turnover)>15?`🚨 Turnover elevado (${turnover}%).`:parseFloat(turnover)>8?`⚡ Moderado (${turnover}%).`:`✅ Controlado (${turnover}%).`;
            if(pendRep>0)insC+=` 🔴 <strong>${pendRep} vaga(s) em reposição.</strong>`;
            _setDRH('ai-colab-insight',insC,true);

            let insFer='';if(fVenc>0)insFer+=`🚨 <strong>${fVenc} vencidas</strong> — risco legal. `;if(em30>0)insFer+=`⏰ ${em30} a vencer em 30 dias. `;if(fAtv>0)insFer+=`🌴 ${fAtv} em gozo. `;if(pendFer>0)insFer+=`📬 ${pendFer} pendente(s). `;if(!insFer)insFer='✅ Sem irregularidades.';
            _setDRH('ai-ferias-insight',insFer,true);

            _setDRH('ai-bh-insight',(totPos>100?'⚠️ Saldo positivo elevado. ':qtdNeg>ativos*0.2?'🔴 +20% com saldo negativo. ':'')+'✅ OK',true);
            _setDRH('ai-abs-insight',`📊 Taxa: <strong>${taxa}%</strong>. `+(parseFloat(taxa)>4?`🚨 Acima do limite. Motivo: <strong>${topMot}</strong>.`:parseFloat(taxa)>2?'⚠️ Moderada.':'✅ Ok.'),true);
            _setDRH('ai-ponto-insight',pTot>0?`🗓 <strong>${pTot} ocorrência(s)</strong>.`+(pFlt>5?` ⚠️ Faltas: ${pFlt}.`:''):'Nenhuma ocorrência.',true);
        }

    }catch(e){console.error('Dashboard RH:',e);_setDRH('ai-colab-insight','❌ Erro ao carregar: '+e.message,true);}
    finally{if(btn){btn.disabled=false;btn.innerHTML='<i class="fas fa-sync-alt"></i> Atualizar Painel';}}
};
// ── Inicialização (carrega dados na abertura do app) ──────────────────────
document.addEventListener('DOMContentLoaded', function () {
    setTimeout(async () => {
        try {
            let lista=[];
            if(window.getFromDB){
                lista=(await window.getFromDB('dRH').catch(()=>null))||[];
                const rhDb=(await window.getFromDB('rhData').catch(()=>null))||{};
                if(lista.length){
                    window.dRH=lista;
                    window.rhData=Object.assign({},rhDb,window.rhData);
                    _saveRHList(lista);
                    window.renderRHQuad();
                    window.renderFerias();
                    window.renderBanco();
                    if(typeof renderAbs==='function')renderAbs();
                    return;
                }
            }
            lista=_getRHList();
            if(lista.length){window.dRH=lista;window.renderRHQuad();}
        }catch(e){console.warn('RH init:',e);}
    }, 1800);
});

// =========================================================================
// EXCLUIR COLABORADOR — busca por matrícula (evita índice stale após re-render)
// =========================================================================
window.excluirColaboradorRH = async function (indexOuMat) {
    const lista = _getLista();
    if (!lista || !lista.length) { alert('Colaborador não encontrado.'); return; }

    // Aceita tanto índice numérico quanto string de matrícula
    let colab, realIndex;
    if (typeof indexOuMat === 'string') {
        realIndex = lista.findIndex(c =>
            (c['Matrícula']||c['Matricula']||c['MATRICULA']||c['matricula']||'') === indexOuMat);
    } else {
        // Re-valida pelo índice atual da lista em memória
        realIndex = indexOuMat;
    }

    if (realIndex < 0 || realIndex >= lista.length) {
        alert('Colaborador não encontrado. Recarregue a página e tente novamente.'); return;
    }
    colab = lista[realIndex];

    const mat     = colab['Matrícula']||colab['Matricula']||colab['MATRICULA']||colab['matricula']||'';
    const nome    = colab['Nome']||colab['NOME']||colab['nome']||mat;
    const funcao  = colab['Função']||colab['Funcao']||colab['FUNÇÃO']||colab['funcao']||'';

    const motivo = prompt(
        `Excluir: ${nome} (${mat})\n\nMotivo da saída:\n1 = Demissão\n2 = Transferência\n3 = Outro\n\nDigite:`, '1'
    );
    if (motivo === null) return; // cancelado

    const motivoTexto = ({'1':'Demissão','2':'Transferência','3':'Outro'})[motivo] || 'Não informado';

    // Remove da lista em memória (usa realIndex para garantir remoção correta)
    lista.splice(realIndex, 1);
    window.dRH = lista;

    // Remove dados de rhData
    if (window.rhData && window.rhData[mat]) delete window.rhData[mat];

    // Persiste no localStorage IMEDIATAMENTE
    try { localStorage.setItem('rh_colaboradores', JSON.stringify(window.dRH)); } catch(e) {}

    // Persiste no Firebase (assíncrono — não bloqueia a UI)
    if (window.saveToDB) {
        window.saveToDB('dRH',    window.dRH).catch(e => console.warn('saveToDB dRH:', e));
        window.saveToDB('rhData', window.rhData).catch(e => console.warn('saveToDB rhData:', e));
    }

    // Registra pendência de reposição
    if (motivo === '1' || motivo === '2') {
        try {
            const arr = JSON.parse(localStorage.getItem('pendencias_reposicao') || '[]');
            arr.push({ id: Date.now(), exColab: nome, matriculaEx: mat, motivo: motivoTexto,
                       data: new Date().toLocaleDateString('pt-BR'), funcao: funcao });
            localStorage.setItem('pendencias_reposicao', JSON.stringify(arr));
            if (typeof renderReposicoes === 'function') renderReposicoes();
        } catch(e) {}
    }

    // Atualiza UI imediatamente — SEM reload
    window.renderRHQuad();
    if (typeof window.renderFerias      === 'function') window.renderFerias();
    if (typeof window.renderBanco       === 'function') window.renderBanco();
    if (typeof renderAbs                === 'function') renderAbs();
    if (typeof renderPontoMensal        === 'function') renderPontoMensal();

    if (typeof showToast === 'function')
        showToast(`${nome} removido do quadro.`, 'warn');
};

// =========================================================================
// LEITOR DE ESCALA DE FÉRIAS — PDF (correto) ou Excel
// Detecta: limiteFerias, periodoAberto, feriasInicio, feriasFim
// =========================================================================
(function () {
    // Substitui o listener do f_ferias definido anteriormente pelo correto
    document.addEventListener('DOMContentLoaded', function () {
        const fFerias = document.getElementById('f_ferias');
        if (!fFerias) return;

        // Remove listeners antigos clonando o elemento
        const clone = fFerias.cloneNode(true);
        fFerias.parentNode.replaceChild(clone, fFerias);

        clone.addEventListener('change', function () {
            if (!this.files[0]) return;
            const file  = this.files[0];
            const isPDF = file.name.toLowerCase().endsWith('.pdf') || file.type === 'application/pdf';

            document.getElementById('loading').style.display = 'flex';
            if (document.getElementById('loadMsg'))
                document.getElementById('loadMsg').innerText = 'Lendo escala de férias...';

            const rd = new FileReader();
            rd.onload = async e => {
                try {
                    if (isPDF) await _lerPDF_Ferias(e.target.result);
                    else       await _lerExcel_Ferias(e.target.result);
                } catch (err) {
                    alert('Erro ao importar escala:\n' + err.message);
                    console.error(err);
                }
                document.getElementById('loading').style.display = 'none';
                this.value = '';
            };
            rd.readAsArrayBuffer(file);
        });
    });
})();

// ── PDF: extrai texto página por página agrupando por Y ──────────────────
async function _lerPDF_Ferias(arrayBuffer) {
    const pdfLib = window.pdfjsLib || window.pdfjs;
    if (!pdfLib) throw new Error('pdf.js não carregado no index.html.');

    pdfLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    const pdf = await pdfLib.getDocument({ data: arrayBuffer }).promise;
    let texto = '';

    for (let p = 1; p <= pdf.numPages; p++) {
        const page    = await pdf.getPage(p);
        const content = await page.getTextContent();

        // Agrupa items por linha (posição Y arredondada)
        const byY = {};
        content.items.forEach(item => {
            const y = Math.round(item.transform[5]);
            byY[y]  = (byY[y] || '') + item.str + ' ';
        });
        // Y descendente = ordem natural de leitura (topo → base)
        Object.keys(byY).map(Number).sort((a, b) => b - a).forEach(y => {
            texto += byY[y].trim() + '\n';
        });
    }

    _processar_EscalaFerias(texto);
}

// ── Excel: converte para texto e processa igual ao PDF ───────────────────
async function _lerExcel_Ferias(arrayBuffer) {
    if (typeof XLSX === 'undefined') throw new Error('Biblioteca Excel não carregada.');
    const wb   = XLSX.read(arrayBuffer, { type: 'array' });
    const raw  = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    const text = raw.map(r => (r || []).map(c => String(c).trim()).join(' ')).join('\n');
    _processar_EscalaFerias(text);
}

// ── PROCESSAMENTO CENTRAL: monta blocos por matrícula e extrai datas ─────
function _processar_EscalaFerias(texto) {
    const lista   = _getLista();
    const blocks  = {};
    let   lastKey = null;

    // Cada linha que começa com 7 dígitos = nova entrada
    texto.split('\n').forEach(rawLine => {
        const line = rawLine.trim();
        if (!line || line.length < 5) return;
        const m = line.match(/^(\d{7})\b/);
        if (m) {
            lastKey = m[1];
            blocks[lastKey] = (blocks[lastKey] || '') + ' ' + line;
        } else if (lastKey && /\d{2}\/\d{2}\/\d{4}/.test(line)) {
            // linha de continuação (nome longo ou datas extras)
            blocks[lastKey] += ' ' + line;
        }
    });

    let count = 0;

    lista.forEach(emp => {
        const matRh    = String(emp['Matrícula']||emp['Matricula']||emp['MATRICULA']||emp['matricula']||'').trim();
        const numMatRh = parseInt(matRh, 10);
        if (!matRh || isNaN(numMatRh)) return;

        let bloco = null;
        for (const key in blocks) {
            if (parseInt(key, 10) === numMatRh) { bloco = blocks[key]; break; }
        }
        if (!bloco) return;

        ensureRhData(matRh);
        const rh_entry = window.rhData[matRh];

        // Extrai todas as datas DD/MM/AAAA do bloco
        const datas = [];
        let m;
        const re = /\b(\d{2}\/\d{2}\/\d{4})\b/g;
        while ((m = re.exec(bloco)) !== null) datas.push(m[1]);
        if (!datas.length) return;

        // 1. LIMITE PARA INÍCIO = última data do bloco
        rh_entry.limiteFerias = datas[datas.length - 1].split('/').reverse().join('-');

        // 2. PERÍODO AQUISITIVO = par de datas com ~365 dias de diferença
        for (let i = 0; i < datas.length - 1; i++) {
            const [d1, d2] = [datas[i], datas[i+1]].map(d => {
                const p = d.split('/'); return new Date(p[2], p[1]-1, p[0]);
            });
            const diff = Math.abs((d2 - d1) / 86400000);
            if (diff >= 355 && diff <= 375) {
                const ini = d1 < d2 ? datas[i] : datas[i+1];
                const fim = d1 < d2 ? datas[i+1] : datas[i];
                rh_entry.periodoAberto = `${ini} a ${fim}`;
                break;
            }
        }

        // 3. FÉRIAS MARCADAS = par com 14 a 35 dias de diferença
        //    Usa a primeira marcação futura (ou a mais recente se passada)
        const hoje = new Date(); hoje.setHours(0,0,0,0);
        let melhorMarcacao = null;

        for (let i = 0; i < datas.length - 1; i++) {
            const [d1, d2] = [datas[i], datas[i+1]].map(d => {
                const p = d.split('/'); return new Date(p[2], p[1]-1, p[0]);
            });
            const diff = Math.abs((d2 - d1) / 86400000);
            if (diff >= 14 && diff <= 35) {
                const ini = d1 < d2 ? d1 : d2;
                const fim = d1 < d2 ? d2 : d1;
                const iniStr = (d1 < d2 ? datas[i] : datas[i+1]).split('/').reverse().join('-');
                const fimStr = (d1 < d2 ? datas[i+1] : datas[i]).split('/').reverse().join('-');

                // Prefere marcação que ainda está vigente ou futura
                if (!melhorMarcacao || ini >= hoje || (ini < hoje && fim >= hoje)) {
                    melhorMarcacao = { iniStr, fimStr, ini, fim };
                }
            }
        }

        if (melhorMarcacao) {
            rh_entry.feriasInicio  = melhorMarcacao.iniStr;
            rh_entry.feriasFim     = melhorMarcacao.fimStr;
            rh_entry.statusManual  = 'AUTO'; // deixa o sistema calcular se está em férias
        }

        count++;
    });

    // Persiste e atualiza tela
    (async () => {
        if (window.saveToDB) await window.saveToDB('rhData', window.rhData);
        window.renderFerias();
        window.renderRHQuad();

        if (count === 0) {
            alert('⚠️ Nenhuma correspondência encontrada.\n\nDica: importe o quadro de colaboradores primeiro (aba Colaboradores → Importar Excel).');
        } else {
            alert(`✅ Escala lida!\n${count} colaboradores atualizados.\n\nColaboradores em férias hoje aparecerão com badge "EM FÉRIAS" na tabela.`);
        }
    })();
}
