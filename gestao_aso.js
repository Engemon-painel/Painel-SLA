// ======================================================================
// EXAMES PERIÓDICOS (ASO) — módulo do Painel do ESER
// ----------------------------------------------------------------------
// Cria o item "🩺 Exames Periódicos" no menu (abaixo de Gestão de NR)
// e a página. Dados: base_nr.json (automático) ou dados_aso.json, via nr_conversor.js,
// gerado da aba "Preenchimento" da Base_NR.xlsx — SEM CPF.
// No index.html: <script src="gestao_aso.js"></script> antes de </body>
// ======================================================================
(function(){

  // Situação calculada pela data de HOJE (atualiza sozinha todo dia):
  const CATS = [
    { k:'VENCIDO',   label:'Vencido',             cor:'#7f1d1d', icone:'⛔' },
    { k:'ATE7',      label:'Vence em até 7 dias', cor:'#dc2626', icone:'🔴' },
    { k:'ATE30',     label:'Vence em 8 a 30 dias', cor:'#d97706', icone:'🟠' },
    { k:'PRAZO',     label:'Mais de 30 dias',     cor:'#0e7c86', icone:'🟢' },
    { k:'AGENDADO',  label:'Exame agendado',      cor:'#2563eb', icone:'📅' },
    { k:'REALIZADO', label:'Exame realizado',     cor:'#059669', icone:'✅' }
  ];
  const CAT = {}; CATS.forEach(c => CAT[c.k] = c);

  let ASO = null, asoErro = null;
  const filtro = { sup:'', area:'', cat:'', q:'' };
  let chartSemana = null, chartSup = null;

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const normMat = m => String(m ?? '').replace(/\D/g,'').replace(/^0+/,'');
  const fmt = n => Number(n).toLocaleString('pt-BR');
  const dataBr = iso => { if(!iso) return ''; const p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; };
  const paraData = iso => { if(!iso) return null; const p = iso.split('-'); return new Date(+p[0], +p[1]-1, +p[2]); };
  function hoje0(){ const h = new Date(); return new Date(h.getFullYear(), h.getMonth(), h.getDate()); }

  function diasParaVencer(e){
    const v = paraData(e.vencimento);
    return v ? Math.round((v - hoje0()) / 86400000) : null;
  }
  function categoria(e){
    const st = String(e.status || '').toUpperCase();
    const dEx = paraData(e.data_exame);
    if(st.includes('REALIZ') || (dEx && dEx <= hoje0())) return 'REALIZADO';
    if(st.includes('AGEND') || (dEx && dEx > hoje0())) return 'AGENDADO';
    const d = diasParaVencer(e);
    if(d === null) return 'PRAZO';
    if(d < 0) return 'VENCIDO';
    if(d <= 7) return 'ATE7';
    if(d <= 30) return 'ATE30';
    return 'PRAZO';
  }

  // ---------------------------------------------------------------- estrutura
  const CSS = `
    #pagina-aso .aso-filtros{ margin:10px 0 16px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
    #pagina-aso .aso-lbl{ font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em; font-family:'JetBrains Mono', monospace; }
    #pagina-aso select{ font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); font-family:'JetBrains Mono', monospace; }
    #asoKpiRow{ grid-template-columns: repeat(auto-fit, minmax(135px,1fr)); }
    #asoKpiRow .kpi{ cursor:pointer; }
    #asoTable th, #asoTable td{ white-space:nowrap; }
    #asoTable td{ padding:8px 6px; }
    #asoTable thead th{ position:sticky; top:0; background:#eef1f5; z-index:2; }
  `;

  const TEMPLATE = `
    <div class="panel">
      <h2>Exames Periódicos (ASO)</h2>
      <div class="hint mono" id="asoInfo">Carregando dados de ASO...</div>

      <div class="aso-filtros">
        <label class="aso-lbl" for="asoSup">Supervisor</label>
        <select id="asoSup" style="min-width:220px;"></select>
        <label class="aso-lbl" for="asoArea" style="margin-left:8px;">Área</label>
        <select id="asoArea" style="min-width:160px;"></select>
        <label class="aso-lbl" for="asoCat" style="margin-left:8px;">Situação</label>
        <select id="asoCat" style="min-width:190px;"></select>
      </div>

      <div class="kpi-row" id="asoKpiRow"></div>

      <div class="grid" style="grid-template-columns: 1.2fr 1fr;">
        <div class="panel">
          <h2>Vencimentos por semana</h2>
          <div class="hint">Quantos ASOs vencem em cada semana — ajuda a planejar os agendamentos.</div>
          <div style="position:relative; height:300px;"><canvas id="asoSemanaChart"></canvas></div>
        </div>
        <div class="panel">
          <h2>Por supervisor</h2>
          <div class="hint">Cada colaborador conta uma vez, pela situação do ASO dele. Clique para filtrar.</div>
          <div id="asoSupWrap" style="position:relative; height:300px;"><canvas id="asoSupChart"></canvas></div>
        </div>
      </div>

      <div class="panel">
        <h2>Colaboradores</h2>
        <div class="aso-filtros">
          <div class="search" title="Digite a matrícula (número) ou parte do nome">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
            <input id="asoBusca" type="search" placeholder="Matrícula ou nome..." aria-label="Buscar colaborador" style="min-width:230px;" />
          </div>
          <span id="asoContagem" class="mono" style="font-size:11px; color:var(--muted);"></span>
          <button type="button" class="link-base" id="asoExportar" style="margin-left:auto; cursor:pointer; border:1px solid var(--accent); font-family:'JetBrains Mono', monospace;">⬇ Exportar (CSV)</button>
        </div>
        <div style="overflow:auto; max-height:70vh; border-radius:8px;">
          <table id="asoTable">
            <thead><tr>
              <th>Matrícula</th><th>Nome</th><th>Função</th><th>Área</th><th>Supervisor</th>
              <th class="num">Vencimento</th><th class="num">Dias p/ vencer</th><th class="num">Data do exame</th><th>Situação</th><th>Observações</th>
            </tr></thead>
            <tbody id="asoTbody"></tbody>
          </table>
          <div id="asoVazio" class="no-results" style="display:none">Nenhum colaborador encontrado.</div>
        </div>
      </div>
    </div>
  `;

  function montarEstrutura(){
    if(document.getElementById('pagina-aso')) return;
    const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);

    const nav = document.createElement('div');
    nav.className = 'nav-item'; nav.id = 'nav-aso';
    nav.innerHTML = '<span class="nav-icon">🩺</span> Exames Periódicos';
    nav.onclick = () => mostrarPagina('aso');
    const ancora = document.getElementById('nav-nr') || document.getElementById('nav-frota');
    if(ancora) ancora.after(nav); else document.querySelector('.sidebar')?.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina'; pg.id = 'pagina-aso'; pg.style.display = 'none';
    pg.innerHTML = TEMPLATE;
    document.querySelector('main.main-content')?.appendChild(pg);

    if(typeof PAGINAS !== 'undefined' && PAGINAS.indexOf('aso') === -1) PAGINAS.push('aso');
    const mostrarOriginal = window.mostrarPagina;
    window.mostrarPagina = function(nome){
      mostrarOriginal(nome);
      if(nome === 'aso') renderAso();
    };

    const liga = (id, campo) => document.getElementById(id).addEventListener('change', e => { filtro[campo] = e.target.value; renderAso(); });
    liga('asoSup','sup'); liga('asoArea','area'); liga('asoCat','cat');
    let t;
    document.getElementById('asoBusca').addEventListener('input', e => {
      clearTimeout(t); t = setTimeout(() => { filtro.q = e.target.value.trim().toLowerCase(); renderTabela(); }, 200);
    });
    document.getElementById('asoExportar').addEventListener('click', exportarCsv);
  }

  // ---------------------------------------------------------------- dados
  async function carregarAso(){
    try{
      const dados = await NRConversor.carregar();
      if(!dados.aso) throw new Error('dados de ASO não encontrados');
      ASO = dados.aso; asoErro = null;
      popularFiltros();
    }catch(e){ asoErro = e.message; console.error('[Exames Periódicos]', e); }
    const pg = document.getElementById('pagina-aso');
    if(pg && pg.style.display !== 'none') renderAso();
  }

  function popularFiltros(){
    const lista = ASO.exames || [];
    const unicos = campo => Array.from(new Set(lista.map(e => e[campo]).filter(Boolean))).sort((a,b) => a.localeCompare(b,'pt-BR'));
    const preencher = (id, valores, campo) => {
      const sel = document.getElementById(id);
      if(filtro[campo] && valores.indexOf(filtro[campo]) === -1) filtro[campo] = '';
      sel.innerHTML = '<option value="">Todos</option>' + valores.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
      sel.value = filtro[campo];
    };
    preencher('asoSup', unicos('supervisor'), 'sup');
    preencher('asoArea', unicos('area'), 'area');
    const selCat = document.getElementById('asoCat');
    selCat.innerHTML = '<option value="">Todas</option>' + CATS.map(c => `<option value="${c.k}">${c.icone} ${c.label}</option>`).join('');
    selCat.value = filtro.cat;
  }

  function baseFiltrada(ignorarCat){
    return (ASO.exames || []).map(e => Object.assign({}, e, { cat: categoria(e), dias: diasParaVencer(e) }))
      .filter(e => {
        if(filtro.sup && e.supervisor !== filtro.sup) return false;
        if(filtro.area && e.area !== filtro.area) return false;
        if(!ignorarCat && filtro.cat && e.cat !== filtro.cat) return false;
        return true;
      });
  }

  function linhasTabela(){
    return baseFiltrada().filter(e => {
      if(!filtro.q) return true;
      if(/^\d+$/.test(filtro.q)) return normMat(e.matricula).startsWith(filtro.q.replace(/^0+/, ''));
      return [e.nome, e.funcao, e.supervisor].join(' ').toLowerCase().includes(filtro.q);
    }).sort((a,b) => String(a.vencimento || '9999').localeCompare(String(b.vencimento || '9999')) || a.nome.localeCompare(b.nome,'pt-BR'));
  }

  // ---------------------------------------------------------------- render
  function renderAso(){
    const info = document.getElementById('asoInfo');
    if(!info) return;
    if(asoErro){
      info.textContent = 'Não foi possível carregar os dados de ASO (' + asoErro + '). Confira se base_nr.json ou dados_aso.json está na mesma pasta do index.html.';
      info.style.color = '#dc2626'; return;
    }
    if(!ASO){ info.textContent = 'Carregando dados de ASO...'; return; }
    info.style.color = '';
    info.textContent = 'Base atualizada em ' + dataBr(ASO.atualizado_em) + ' · ' + fmt((ASO.exames || []).length) +
      ' colaboradores · situação calculada com a data de hoje (' + hoje0().toLocaleDateString('pt-BR') + ')';
    const base = baseFiltrada();
    renderKpis();
    renderChartSemana(base);
    renderChartSup(base);
    renderTabela();
  }

  function renderKpis(){
    const todos = baseFiltrada(true);
    const cont = {}; CATS.forEach(c => cont[c.k] = 0);
    todos.forEach(e => cont[e.cat]++);
    const n = todos.length;
    const card = (k, icone, titulo, valor, cor) => `
      <div class="kpi" data-cat="${k}" style="${filtro.cat === k ? 'box-shadow:0 0 0 2px ' + cor + ' inset;' : ''}">
        <div class="label">${icone} ${titulo}</div>
        <div class="value" style="color:${cor}">${fmt(valor)}</div>
        <div class="delta" style="color:var(--muted)">${k ? (n ? (valor / n * 100).toFixed(0) + '% do total' : '—') : 'clique num card para filtrar'}</div>
      </div>`;
    const row = document.getElementById('asoKpiRow');
    row.innerHTML = card('', '🩺', 'Colaboradores', n, 'var(--accent)') +
      CATS.map(c => card(c.k, c.icone, c.label, cont[c.k], c.cor)).join('');
    row.querySelectorAll('.kpi').forEach(el => el.onclick = () => {
      const k = el.dataset.cat;
      filtro.cat = (filtro.cat === k) ? '' : k;
      document.getElementById('asoCat').value = filtro.cat;
      renderAso();
    });
  }

  function inicioSemana(d){ const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; } // segunda-feira

  function renderChartSemana(base){
    const canvas = document.getElementById('asoSemanaChart');
    if(!canvas || typeof Chart === 'undefined') return;
    const porSemana = {};
    base.forEach(e => {
      const v = paraData(e.vencimento); if(!v) return;
      const s = inicioSemana(v); const k = s.toISOString().slice(0,10);
      if(!porSemana[k]){ porSemana[k] = { ini:s }; CATS.forEach(c => porSemana[k][c.k] = 0); }
      porSemana[k][e.cat]++;
    });
    const chaves = Object.keys(porSemana).sort();
    const rotulos = chaves.map(k => {
      const i = porSemana[k].ini, f = new Date(i); f.setDate(f.getDate() + 6);
      const dm = d => String(d.getDate()).padStart(2,'0') + '/' + String(d.getMonth()+1).padStart(2,'0');
      return dm(i) + ' a ' + dm(f);
    });
    const datasets = CATS.map(c => ({
      label: c.label, data: chaves.map(k => porSemana[k][c.k]), backgroundColor: c.cor, borderRadius: 3,
      datalabels: { display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) > 0, color:'#ffffff', anchor:'center', align:'center', font:{ size:10, family:'JetBrains Mono', weight:'700' } }
    }));
    if(chartSemana) chartSemana.destroy();
    chartSemana = new Chart(canvas, {
      type:'bar', data:{ labels: rotulos, datasets },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options:{
        responsive:true, maintainAspectRatio:false,
        plugins:{
          legend:{ position:'bottom', labels:{ color:'#6b7590', font:{ family:'Inter', size:10 }, boxWidth:10, filter:(it, data) => data.datasets[it.datasetIndex].data.some(v => v > 0) } },
          datalabels:{ display:true }
        },
        scales:{
          x:{ stacked:true, ticks:{ color:'#1b2440', font:{ size:10, weight:'bold' } }, grid:{ display:false } },
          y:{ stacked:true, beginAtZero:true, ticks:{ color:'#1b2440', precision:0 }, grid:{ display:false } }
        }
      }
    });
  }

  function renderChartSup(base){
    const canvas = document.getElementById('asoSupChart');
    if(!canvas || typeof Chart === 'undefined') return;
    const porSup = {};
    base.forEach(e => {
      const s = e.supervisor || 'Sem supervisor';
      if(!porSup[s]){ porSup[s] = { total:0 }; CATS.forEach(c => porSup[s][c.k] = 0); }
      porSup[s][e.cat]++; porSup[s].total++;
    });
    const sups = Object.keys(porSup).sort((a,b) => porSup[b].total - porSup[a].total);
    const wrap = document.getElementById('asoSupWrap');
    if(wrap) wrap.style.height = Math.max(260, sups.length * 30 + 80) + 'px';
    const datasets = CATS.map(c => ({
      label: c.label, _cat: c.k, data: sups.map(s => porSup[s][c.k]), backgroundColor: c.cor, borderWidth:0,
      datalabels: { display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) > 0, color:'#ffffff', anchor:'center', align:'center', font:{ size:10, family:'JetBrains Mono', weight:'700' } }
    }));
    datasets.push({
      label:'Total', data: sups.map(() => 0), backgroundColor:'rgba(0,0,0,0)',
      datalabels:{ anchor:'end', align:'right', offset:4, color:'#000000', font:{ size:11, family:'JetBrains Mono', weight:'700' }, formatter:(v, ctx) => porSup[sups[ctx.dataIndex]].total }
    });
    if(chartSup) chartSup.destroy();
    chartSup = new Chart(canvas, {
      type:'bar', data:{ labels: sups, datasets },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options:{
        indexAxis:'y', responsive:true, maintainAspectRatio:false, maxBarThickness:24,
        layout:{ padding:{ right:30 } },
        onClick:(evt, els) => {
          if(!els.length) return;
          const sup = sups[els[0].index];
          filtro.sup = (filtro.sup === sup) ? '' : sup;
          document.getElementById('asoSup').value = filtro.sup;
          renderAso();
        },
        onHover:(evt, els) => { evt.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        plugins:{
          legend:{ position:'bottom', labels:{ color:'#6b7590', font:{ family:'Inter', size:10 }, boxWidth:10, filter:(it, data) => it.text !== 'Total' && data.datasets[it.datasetIndex].data.some(v => v > 0) } },
          datalabels:{ display:true }
        },
        scales:{
          x:{ stacked:true, beginAtZero:true, ticks:{ color:'#1b2440', precision:0 }, grid:{ display:false } },
          y:{ stacked:true, ticks:{ color:'#1b2440', font:{ size:11, weight:'bold' } }, grid:{ display:false } }
        }
      }
    });
  }

  function renderTabela(){
    if(!ASO) return;
    const linhas = linhasTabela();
    document.getElementById('asoContagem').textContent = fmt(linhas.length) + ' colaborador' + (linhas.length === 1 ? '' : 'es');
    const tbody = document.getElementById('asoTbody'), vazio = document.getElementById('asoVazio');
    if(!linhas.length){ tbody.innerHTML = ''; vazio.style.display = ''; return; }
    vazio.style.display = 'none';
    tbody.innerHTML = linhas.map(e => {
      const c = CAT[e.cat];
      const urgente = e.cat === 'VENCIDO' || e.cat === 'ATE7';
      const diasTxt = (e.cat === 'REALIZADO' || e.dias === null) ? '—' : (e.dias < 0 ? 'vencido há ' + Math.abs(e.dias) + 'd' : e.dias + 'd');
      return `<tr${urgente ? ' style="background:rgba(220,38,38,.07);"' : ''}>
        <td class="mono">${esc(e.matricula || '—')}</td>
        <td style="font-weight:600;">${esc(e.nome)}</td>
        <td style="font-size:12px;">${esc(e.funcao || '—')}</td>
        <td style="font-size:12px;">${esc(e.area || '—')}</td>
        <td style="font-size:12px;">${esc(e.supervisor || '—')}</td>
        <td class="num mono">${esc(dataBr(e.vencimento) || '—')}</td>
        <td class="num mono" style="color:${c.cor}; font-weight:700;">${esc(diasTxt)}</td>
        <td class="num mono">${esc(dataBr(e.data_exame) || '—')}</td>
        <td><span class="pill" style="background:${c.cor}22; color:${c.cor};">${c.icone} ${esc(c.label)}</span></td>
        <td style="font-size:12px; white-space:normal; min-width:160px;">${esc(e.obs || '')}</td>
      </tr>`;
    }).join('');
  }

  function exportarCsv(){
    if(!ASO) return;
    const linhas = linhasTabela();
    if(!linhas.length){ alert('Nenhum colaborador para exportar com os filtros atuais.'); return; }
    const csvEsc = v => { const s = String(v ?? ''); return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s; };
    const out = [['Matricula','Nome','Funcao','Area','Supervisor','Vencimento ASO','Dias p/ vencer','Data do Exame','Situacao','Observacoes'].join(';')];
    linhas.forEach(e => out.push([e.matricula, e.nome, e.funcao, e.area, e.supervisor, dataBr(e.vencimento), e.dias, dataBr(e.data_exame), CAT[e.cat].label, e.obs].map(csvEsc).join(';')));
    const blob = new Blob(['\uFEFF' + out.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'Exames_Periodicos_ASO.csv'; a.style.display = 'none';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  // ---------------------------------------------------------------- início
  function iniciar(){
    montarEstrutura();
    carregarAso();
    setInterval(carregarAso, 5 * 60 * 1000);
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();

})();
