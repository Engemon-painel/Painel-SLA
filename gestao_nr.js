// ======================================================================
// GESTÃO DE NR — módulo do Painel do ESER
// ----------------------------------------------------------------------
// Este arquivo cria sozinho o item "🦺 Gestão de NR" no menu lateral
// (logo abaixo de Gestão de Frotas) e a página correspondente.
// Os dados vêm de dados_nr.json (mesma pasta do index.html).
// No index.html basta UMA linha, logo antes de </body>:
//   <script src="gestao_nr.js"></script>
// ======================================================================
(function(){

  const NR_STATUS = [
    { k:'OK',             label:'Em dia',         cor:'#059669', bg:'rgba(5,150,105,.14)',   txt:'#047857' },
    { k:'SEM ASSINATURA', label:'Sem assinatura', cor:'#d97706', bg:'rgba(217,119,6,.16)',   txt:'#b45309' },
    { k:'DIVERGÊNCIA',    label:'Divergência',    cor:'#5e34b5', bg:'rgba(94,52,181,.14)',   txt:'#5e34b5' },
    { k:'PENDENTE',       label:'Pendente',       cor:'#dc2626', bg:'rgba(220,38,38,.14)',   txt:'#b91c1c' },
    { k:'EM ANDAMENTO',   label:'Em andamento',   cor:'#2563eb', bg:'rgba(37,99,235,.14)',   txt:'#1d4ed8' },
    { k:'AVALIAR',        label:'Avaliar',        cor:'#0e7490', bg:'rgba(8,145,178,.12)',   txt:'#0e7490' },
    { k:'N/A',            label:'Não se aplica',  cor:'#94a3b8', bg:'rgba(107,117,144,.12)', txt:'#6b7590' },
    { k:'SEM REGISTRO',   label:'Sem registro',   cor:'#d1d5db', bg:'transparent',           txt:'#9ca3af' }
  ];
  const ST = {}; NR_STATUS.forEach(s => ST[s.k] = s);
  // Contam na % de conformidade (AVALIAR, N/A e SEM REGISTRO ficam de fora).
  const APLICAVEIS = new Set(['OK','SEM ASSINATURA','DIVERGÊNCIA','PENDENTE','EM ANDAMENTO']);
  // O que é "pendência" de verdade (precisa de ação).
  const PROBLEMA = ['PENDENTE','SEM ASSINATURA','DIVERGÊNCIA','EM ANDAMENTO'];

  let NR = null, nrErro = null;
  const filtro = { coord:'', sup:'', area:'', situacao:'', q:'', curso:'', status:'' };
  let chartCurso = null, chartSup = null;

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const normMat = m => String(m ?? '').replace(/\D/g,'').replace(/^0+/,'');
  const dataBr = iso => { if(!iso) return ''; const p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0].slice(2); };
  const fmt = n => Number(n).toLocaleString('pt-BR');

  // Função (cargo) vem da lista de colaboradores do dados.json, cruzando pela matrícula.
  function funcaoPorMatricula(){
    const mapa = {};
    const lista = (typeof COLABORADORES !== 'undefined' && Array.isArray(COLABORADORES)) ? COLABORADORES : [];
    lista.forEach(c => { if(c && c.matricula) mapa[normMat(c.matricula)] = c.descFuncao || ''; });
    return mapa;
  }

  // ---------------------------------------------------------------- estrutura
  const CSS = `
    #pagina-nr .nr-filtros{ margin:10px 0 16px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
    #pagina-nr .nr-lbl{ font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em; font-family:'JetBrains Mono', monospace; }
    #pagina-nr select{ font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); font-family:'JetBrains Mono', monospace; }
    #pagina-nr .kpi{ cursor:default; }
    #nrKpiRow{ grid-template-columns: repeat(auto-fit, minmax(135px,1fr)); }
    #pagina-nr .nr-cell{ display:inline-block; min-width:64px; padding:3px 6px; border-radius:6px; font-family:'JetBrains Mono', monospace; font-size:10.5px; font-weight:700; text-align:center; white-space:nowrap; }
    #nrTable th, #nrTable td{ white-space:nowrap; }
    #nrTable td{ padding:6px 5px; }
    #nrTable th.nr-curso{ text-align:center; }
    #nrTable td.nr-curso{ text-align:center; }
    #nrTable .nr-fixa{ position:sticky; left:0; background:#eef1f5; z-index:1; }
    #nrTable thead th{ position:sticky; top:0; background:#eef1f5; z-index:2; }
    #nrTable thead th.nr-fixa{ z-index:3; }
    #pagina-nr .nr-legenda{ display:flex; gap:10px; flex-wrap:wrap; margin:4px 0 12px; }
    #pagina-nr .nr-legenda span{ font-size:11px; display:flex; align-items:center; gap:5px; color:var(--muted); }
    #pagina-nr .nr-legenda i{ width:10px; height:10px; border-radius:3px; display:inline-block; }
  `;

  const TEMPLATE = `
    <div class="panel">
      <h2>Gestão de NR — Treinamentos e Documentos</h2>
      <div class="hint mono" id="nrInfo">Carregando dados_nr.json...</div>

      <div class="nr-filtros">
        <label class="nr-lbl" for="nrCoord">Coordenador</label>
        <select id="nrCoord" style="min-width:200px;"></select>
        <label class="nr-lbl" for="nrSup" style="margin-left:8px;">Supervisor</label>
        <select id="nrSup" style="min-width:220px;"></select>
        <label class="nr-lbl" for="nrArea" style="margin-left:8px;">Área</label>
        <select id="nrArea" style="min-width:150px;"></select>
        <label class="nr-lbl" for="nrSituacao" style="margin-left:8px;">Situação</label>
        <select id="nrSituacao" style="min-width:130px;"></select>
        <div class="search" style="margin-left:auto;" title="Buscar por nome, matrícula ou função">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
          <input id="nrBusca" type="search" placeholder="Buscar por nome, matrícula ou função..." aria-label="Buscar colaborador" style="min-width:220px;" />
        </div>
      </div>

      <div class="kpi-row" id="nrKpiRow"></div>

      <div class="grid" style="grid-template-columns: 1.3fr 1fr;">
        <div class="panel">
          <h2>Status por curso</h2>
          <div class="hint">Clique numa barra para ver quem está naquele status, na matriz abaixo.</div>
          <div id="nrCursoWrap" style="position:relative; height:470px;"><canvas id="nrCursoChart"></canvas></div>
        </div>
        <div class="panel">
          <h2>Pendências por supervisor</h2>
          <div class="hint">Soma de cursos pendentes, sem assinatura, com divergência ou em andamento. Clique para filtrar.</div>
          <div id="nrSupWrap" style="position:relative; height:470px;"><canvas id="nrSupChart"></canvas></div>
        </div>
      </div>

      <div class="panel" id="nrMatrizPanel">
        <h2>Matriz colaborador × curso</h2>
        <div class="nr-filtros">
          <label class="nr-lbl" for="nrCurso">Curso</label>
          <select id="nrCurso" style="min-width:140px;"></select>
          <label class="nr-lbl" for="nrStatus" style="margin-left:8px;">Status</label>
          <select id="nrStatus" style="min-width:200px;"></select>
          <span id="nrContagem" class="mono" style="font-size:11px; color:var(--muted);"></span>
          <button type="button" class="link-base" style="margin-left:auto; cursor:pointer; border:1px solid var(--accent); font-family:'JetBrains Mono', monospace;" id="nrExportar">⬇ Exportar (CSV)</button>
        </div>
        <div class="nr-legenda" id="nrLegenda"></div>
        <div style="overflow:auto; max-height:70vh; border-radius:8px;">
          <table id="nrTable">
            <thead id="nrThead"></thead>
            <tbody id="nrTbody"></tbody>
          </table>
          <div id="nrVazio" class="no-results" style="display:none">Nenhum colaborador encontrado.</div>
        </div>
      </div>
    </div>
  `;

  function montarEstrutura(){
    if(document.getElementById('pagina-nr')) return;

    const st = document.createElement('style');
    st.textContent = CSS;
    document.head.appendChild(st);

    const nav = document.createElement('div');
    nav.className = 'nav-item';
    nav.id = 'nav-nr';
    nav.innerHTML = '<span class="nav-icon">🦺</span> Gestão de NR';
    nav.onclick = () => mostrarPagina('nr');
    const navFrota = document.getElementById('nav-frota');
    if(navFrota) navFrota.after(nav);
    else document.querySelector('.sidebar')?.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina';
    pg.id = 'pagina-nr';
    pg.style.display = 'none';
    pg.innerHTML = TEMPLATE;
    document.querySelector('main.main-content')?.appendChild(pg);

    // registra a página na navegação existente do painel
    if(typeof PAGINAS !== 'undefined' && PAGINAS.indexOf('nr') === -1) PAGINAS.push('nr');
    const mostrarOriginal = window.mostrarPagina;
    window.mostrarPagina = function(nome){
      mostrarOriginal(nome);
      if(nome === 'nr') renderNr();
    };

    const liga = (id, campo) => document.getElementById(id).addEventListener('change', e => { filtro[campo] = e.target.value; renderNr(); });
    liga('nrCoord','coord'); liga('nrSup','sup'); liga('nrArea','area'); liga('nrSituacao','situacao');
    document.getElementById('nrCurso').addEventListener('change', e => { filtro.curso = e.target.value; renderTabela(); });
    document.getElementById('nrStatus').addEventListener('change', e => { filtro.status = e.target.value; renderTabela(); });
    let t;
    document.getElementById('nrBusca').addEventListener('input', e => {
      clearTimeout(t); t = setTimeout(() => { filtro.q = e.target.value.trim().toLowerCase(); renderNr(); }, 200);
    });
    document.getElementById('nrExportar').addEventListener('click', exportarCsv);

    document.getElementById('nrLegenda').innerHTML = NR_STATUS.map(s =>
      `<span><i style="background:${s.k === 'SEM REGISTRO' ? '#e5e7eb' : s.cor};"></i>${s.label}</span>`).join('');
  }

  // ---------------------------------------------------------------- dados
  async function carregarNr(){
    try{
      const r = await fetch('dados_nr.json?_=' + Date.now());
      if(!r.ok) throw new Error('HTTP ' + r.status);
      NR = await r.json();
      nrErro = null;
      popularFiltros();
    }catch(e){
      nrErro = e.message;
      console.error('[Gestão de NR]', e);
    }
    const pg = document.getElementById('pagina-nr');
    if(pg && pg.style.display !== 'none') renderNr();
  }

  function popularFiltros(){
    const cols = NR.colaboradores || [];
    const unicos = campo => Array.from(new Set(cols.map(c => c[campo]).filter(Boolean))).sort((a,b) => a.localeCompare(b,'pt-BR'));
    const preencher = (id, valores, rotuloTodos, campo) => {
      const sel = document.getElementById(id);
      if(filtro[campo] && valores.indexOf(filtro[campo]) === -1) filtro[campo] = '';
      sel.innerHTML = `<option value="">${rotuloTodos}</option>` + valores.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
      sel.value = filtro[campo];
    };
    preencher('nrCoord', unicos('coordenador'), 'Todos', 'coord');
    preencher('nrSup', unicos('supervisor'), 'Todos', 'sup');
    preencher('nrArea', unicos('area'), 'Todas', 'area');
    preencher('nrSituacao', unicos('situacao_nr'), 'Todas', 'situacao');
    preencher('nrCurso', NR.cursos || [], 'Todos os cursos', 'curso');

    const selStatus = document.getElementById('nrStatus');
    selStatus.innerHTML = '<option value="">Todos</option>' +
      '<option value="__PROB">Com alguma pendência</option>' +
      '<option value="__VAZIO">Sem nenhum registro</option>' +
      NR_STATUS.map(s => `<option value="${esc(s.k)}">${esc(s.label)}</option>`).join('');
    selStatus.value = filtro.status;
  }

  // filtros do topo (valem para KPIs, gráficos e matriz)
  function baseFiltrada(){
    const funcoes = funcaoPorMatricula();
    return (NR.colaboradores || []).map(c => Object.assign({}, c, { funcao: funcoes[normMat(c.matricula)] || '' }))
      .filter(c => {
        if(filtro.coord && c.coordenador !== filtro.coord) return false;
        if(filtro.sup && c.supervisor !== filtro.sup) return false;
        if(filtro.area && c.area !== filtro.area) return false;
        if(filtro.situacao && c.situacao_nr !== filtro.situacao) return false;
        if(filtro.q){
          const hay = [c.nome, c.matricula, c.funcao, c.supervisor].join(' ').toLowerCase();
          if(!hay.includes(filtro.q)) return false;
        }
        return true;
      });
  }

  const semNenhumRegistro = c => (NR.cursos || []).every(k => (c.cursos[k] || {}).status === 'SEM REGISTRO');

  // filtros da matriz (curso + status) — só afetam a tabela
  function linhasTabela(base){
    const cursos = filtro.curso ? [filtro.curso] : (NR.cursos || []);
    return base.filter(c => {
      if(!filtro.status) return true;
      if(filtro.status === '__VAZIO') return semNenhumRegistro(c);
      if(filtro.status === '__PROB') return cursos.some(k => PROBLEMA.indexOf((c.cursos[k] || {}).status) !== -1);
      return cursos.some(k => (c.cursos[k] || {}).status === filtro.status);
    });
  }

  // ---------------------------------------------------------------- render
  function renderNr(){
    const info = document.getElementById('nrInfo');
    if(!info) return;
    if(nrErro){
      info.textContent = 'Não foi possível carregar dados_nr.json (' + nrErro + '). Confira se o arquivo está na mesma pasta do index.html.';
      info.style.color = '#dc2626';
      return;
    }
    if(!NR){ info.textContent = 'Carregando dados_nr.json...'; return; }
    info.style.color = '';
    info.textContent = 'Base NR atualizada em ' + dataBr(NR.atualizado_em).replace(/\/(\d{2})$/, '/20$1') +
      ' · ' + fmt((NR.colaboradores || []).length) + ' colaboradores na base';

    const base = baseFiltrada();
    renderKpis(base);
    renderChartCurso(base);
    renderChartSup(base);
    renderTabela();
  }

  function renderKpis(base){
    const cursos = NR.cursos || [];
    const cont = {}; NR_STATUS.forEach(s => cont[s.k] = 0);
    let aplic = 0;
    base.forEach(c => cursos.forEach(k => {
      const s = (c.cursos[k] || {}).status || 'SEM REGISTRO';
      cont[s] = (cont[s] || 0) + 1;
      if(APLICAVEIS.has(s)) aplic++;
    }));
    const n = base.length;
    const concluidos = base.filter(c => String(c.situacao_nr).toUpperCase().startsWith('CONCLU')).length;
    const pctConf = aplic > 0 ? cont['OK'] / aplic : 0;
    const colabPend = base.filter(c => cursos.some(k => PROBLEMA.indexOf((c.cursos[k] || {}).status) !== -1)).length;
    const vazios = base.filter(semNenhumRegistro);
    const vaziosTec = vazios.filter(c => String(c.area).trim().toUpperCase() === 'TÉCNICO').length;
    const corConf = pctConf >= 0.85 ? '#059669' : (pctConf >= 0.7 ? '#d97706' : '#dc2626');
    const pct = (a, b) => b > 0 ? (a / b * 100).toFixed(0) + '%' : '0%';

    document.getElementById('nrKpiRow').innerHTML = `
      <div class="kpi"><div class="label">👥 Colaboradores</div><div class="value" style="color:var(--accent)">${fmt(n)}</div><div class="delta" style="color:var(--muted)">${fmt(colabPend)} com alguma pendência</div></div>
      <div class="kpi"><div class="label">✅ Situação concluída</div><div class="value" style="color:#059669">${fmt(concluidos)}</div><div class="delta" style="color:var(--muted)">${pct(concluidos, n)} dos colaboradores</div></div>
      <div class="kpi"><div class="label">📈 % Conformidade</div><div class="value" style="color:${corConf}">${(pctConf*100).toFixed(1)}%</div><div class="delta" style="color:var(--muted)">${fmt(cont['OK'])} em dia de ${fmt(aplic)} aplicáveis</div></div>
      <div class="kpi"><div class="label">⏳ Pendentes</div><div class="value" style="color:#dc2626">${fmt(cont['PENDENTE'])}</div><div class="delta" style="color:var(--muted)">cursos/documentos</div></div>
      <div class="kpi"><div class="label">✍️ Sem assinatura</div><div class="value" style="color:#d97706">${fmt(cont['SEM ASSINATURA'])}</div><div class="delta" style="color:var(--muted)">cursos/documentos</div></div>
      <div class="kpi"><div class="label">⚠️ Divergências</div><div class="value" style="color:#5e34b5">${fmt(cont['DIVERGÊNCIA'])}</div><div class="delta" style="color:var(--muted)">data errada / divergente</div></div>
      <div class="kpi"><div class="label">📭 Sem nenhum registro</div><div class="value">${fmt(vazios.length)}</div><div class="delta" style="color:${vaziosTec > 0 ? '#dc2626' : 'var(--muted)'}">${fmt(vaziosTec)} técnico${vaziosTec === 1 ? '' : 's'}</div></div>
    `;
  }

  function irParaMatriz(curso, status){
    filtro.curso = curso; filtro.status = status;
    document.getElementById('nrCurso').value = curso;
    document.getElementById('nrStatus').value = status;
    renderTabela();
    document.getElementById('nrMatrizPanel').scrollIntoView({ behavior:'smooth', block:'start' });
  }

  function renderChartCurso(base){
    const canvas = document.getElementById('nrCursoChart');
    if(!canvas || typeof Chart === 'undefined') return;
    const cursos = NR.cursos || [];
    const datasets = NR_STATUS.map(s => ({
      label: s.label,
      data: cursos.map(k => base.filter(c => (c.cursos[k] || {}).status === s.k).length),
      backgroundColor: s.k === 'SEM REGISTRO' ? '#e5e7eb' : s.cor,
      borderWidth: 0,
      _status: s.k,
      datalabels: {
        display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) >= 3,
        color: (s.k === 'SEM REGISTRO' || s.k === 'N/A') ? '#374151' : '#ffffff',
        anchor:'center', align:'center', font:{ size:10, family:'JetBrains Mono', weight:'700' }
      }
    }));
    if(chartCurso) chartCurso.destroy();
    chartCurso = new Chart(canvas, {
      type: 'bar',
      data: { labels: cursos, datasets },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        onClick: (evt, els) => {
          if(!els.length) return;
          const el = els[0];
          irParaMatriz(cursos[el.index], chartCurso.data.datasets[el.datasetIndex]._status);
        },
        onHover: (evt, els) => { evt.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        plugins: {
          legend: { position:'bottom', labels:{ color:'#6b7590', font:{ family:'Inter', size:10 }, boxWidth:10 } },
          datalabels: { display:true }
        },
        scales: {
          x: { stacked:true, beginAtZero:true, ticks:{ color:'#1b2440', font:{ weight:'bold' } }, grid:{ display:false } },
          y: { stacked:true, ticks:{ color:'#1b2440', font:{ size:11, weight:'bold' } }, grid:{ display:false } }
        }
      }
    });
  }

  function renderChartSup(base){
    const canvas = document.getElementById('nrSupChart');
    if(!canvas || typeof Chart === 'undefined') return;
    const cursos = NR.cursos || [];
    const porSup = {};
    base.forEach(c => {
      const sup = c.supervisor || 'Sem supervisor';
      if(!porSup[sup]){ porSup[sup] = { total:0 }; PROBLEMA.forEach(p => porSup[sup][p] = 0); }
      cursos.forEach(k => {
        const s = (c.cursos[k] || {}).status;
        if(PROBLEMA.indexOf(s) !== -1){ porSup[sup][s]++; porSup[sup].total++; }
      });
    });
    const sups = Object.keys(porSup).filter(s => porSup[s].total > 0).sort((a,b) => porSup[b].total - porSup[a].total);
    const wrap = document.getElementById('nrSupWrap');
    if(wrap) wrap.style.height = Math.max(260, sups.length * 30 + 90) + 'px';

    const datasets = PROBLEMA.map(k => ({
      label: ST[k].label,
      data: sups.map(s => porSup[s][k]),
      backgroundColor: ST[k].cor, borderWidth:0,
      datalabels: {
        display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) >= 3,
        color:'#ffffff', anchor:'center', align:'center', font:{ size:10, family:'JetBrains Mono', weight:'700' }
      }
    }));
    datasets.push({
      label:'Total', data: sups.map(() => 0), backgroundColor:'rgba(0,0,0,0)',
      datalabels: {
        anchor:'end', align:'right', offset:4, color:'#000000',
        font:{ size:11, family:'JetBrains Mono', weight:'700' },
        formatter: (v, ctx) => porSup[sups[ctx.dataIndex]].total
      }
    });

    if(chartSup) chartSup.destroy();
    chartSup = new Chart(canvas, {
      type:'bar',
      data:{ labels: sups, datasets },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options:{
        indexAxis:'y', responsive:true, maintainAspectRatio:false,
        layout:{ padding:{ right:30 } },
        onClick: (evt, els) => {
          if(!els.length) return;
          const sup = sups[els[0].index];
          filtro.sup = (filtro.sup === sup) ? '' : sup;   // clicar de novo limpa
          document.getElementById('nrSup').value = filtro.sup;
          renderNr();
        },
        onHover: (evt, els) => { evt.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        plugins:{
          legend:{ position:'bottom', labels:{ color:'#6b7590', font:{ family:'Inter', size:10 }, boxWidth:10, filter: it => it.text !== 'Total' } },
          datalabels:{ display:true }
        },
        scales:{
          x:{ stacked:true, beginAtZero:true, ticks:{ color:'#1b2440', font:{ weight:'bold' } }, grid:{ display:false } },
          y:{ stacked:true, ticks:{ color:'#1b2440', font:{ size:11, weight:'bold' } }, grid:{ display:false } }
        }
      }
    });
  }

  function textoCelula(r){
    const s = r.status;
    if(s === 'OK') return r.data ? dataBr(r.data) : 'OK';
    if(s === 'SEM ASSINATURA') return '✍ ' + (r.data ? dataBr(r.data) : 'assinar');
    if(s === 'DIVERGÊNCIA') return '≠ ' + (r.data ? dataBr(r.data) : 'data');
    if(s === 'PENDENTE') return 'Pendente';
    if(s === 'EM ANDAMENTO' || s === 'AVALIAR'){
      const o = String(r.obs || s).toLowerCase();
      return o.charAt(0).toUpperCase() + o.slice(1);
    }
    if(s === 'N/A') return 'N/A';
    return '—';
  }

  function renderTabela(){
    if(!NR) return;
    const cursos = NR.cursos || [];
    const linhas = linhasTabela(baseFiltrada());

    document.getElementById('nrThead').innerHTML = `<tr>
      <th>Matrícula</th><th class="nr-fixa">Nome</th><th>Função</th><th>Supervisor</th><th>Situação</th>
      ${cursos.map(k => `<th class="nr-curso"${k === filtro.curso ? ' style="color:var(--accent);"' : ''}>${esc(k)}</th>`).join('')}
    </tr>`;

    document.getElementById('nrContagem').textContent = fmt(linhas.length) + ' colaborador' + (linhas.length === 1 ? '' : 'es');
    const vazio = document.getElementById('nrVazio');
    const tbody = document.getElementById('nrTbody');
    if(!linhas.length){ tbody.innerHTML = ''; vazio.style.display = ''; return; }
    vazio.style.display = 'none';

    tbody.innerHTML = linhas.map(c => {
      const sit = String(c.situacao_nr || '');
      const pillSit = sit.toUpperCase().startsWith('CONCLU') ? `<span class="pill good">${esc(sit)}</span>` : `<span class="pill bad">${esc(sit || '—')}</span>`;
      const celulas = cursos.map(k => {
        const r = c.cursos[k] || { status:'SEM REGISTRO' };
        const s = ST[r.status] || ST['SEM REGISTRO'];
        const destaque = filtro.curso && k !== filtro.curso ? 'opacity:.45;' : '';
        const titulo = s.label + (r.obs && r.obs !== r.status ? ' — ' + r.obs : '') + (r.data ? ' (' + dataBr(r.data) + ')' : '');
        return `<td class="nr-curso"><span class="nr-cell" title="${esc(k + ': ' + titulo)}" style="background:${s.bg}; color:${s.txt}; ${destaque}">${esc(textoCelula(r))}</span></td>`;
      }).join('');
      return `<tr>
        <td class="mono">${esc(c.matricula)}</td>
        <td class="nr-fixa" style="font-weight:600;">${esc(c.nome)}</td>
        <td style="font-size:12px;">${esc(c.funcao || '—')}</td>
        <td style="font-size:12px;">${esc(c.supervisor || '—')}</td>
        <td>${pillSit}</td>
        ${celulas}
      </tr>`;
    }).join('');
  }

  function exportarCsv(){
    if(!NR) return;
    const cursos = NR.cursos || [];
    const linhas = linhasTabela(baseFiltrada());
    if(!linhas.length){ alert('Nenhum colaborador para exportar com os filtros atuais.'); return; }
    const csvEsc = v => { const s = String(v ?? ''); return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s; };
    const cab = ['Matricula','Nome','Funcao','Area','Supervisor','Coordenador','Situacao NR'].concat(cursos);
    const out = [cab.map(csvEsc).join(';')];
    linhas.forEach(c => {
      out.push([c.matricula, c.nome, c.funcao, c.area, c.supervisor, c.coordenador, c.situacao_nr]
        .concat(cursos.map(k => {
          const r = c.cursos[k] || { status:'SEM REGISTRO' };
          return (ST[r.status] || ST['SEM REGISTRO']).label + (r.data ? ' ' + dataBr(r.data) : '');
        })).map(csvEsc).join(';'));
    });
    const blob = new Blob(['\uFEFF' + out.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'Gestao_NR.csv'; a.style.display = 'none';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  // ---------------------------------------------------------------- início
  function iniciar(){
    montarEstrutura();
    carregarNr();
    setInterval(carregarNr, 5 * 60 * 1000);
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();

})();
