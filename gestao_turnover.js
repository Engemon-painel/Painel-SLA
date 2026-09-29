// =====================================================================
// gestao_turnover.js — página "Análise de Turnover" do Painel do ESER
//
// Lê o campo "turnover" do dados.json (vem da aba "Base" da planilha
// "Analise de Turnover", via script auxiliar Turnover_LerBase + Power
// Automate + item 41 do script principal).
//
// Instalação: no index.html, junto dos outros módulos, acrescente
//   <script src="gestao_turnover.js"></script>
// Não precisa mexer em mais nada — este arquivo cria sozinho o item no
// menu lateral e a página.
// =====================================================================
(function(){
  const C = (typeof cor !== 'undefined') ? cor : {
    good:'#059669', warn:'#d97706', bad:'#dc2626', accent:'#0e7c86', muted:'#6b7590', line:'#e1e5f0', text:'#1b2440'
  };
  const PALETA = ['#0e7c86','#2563eb','#d97706','#dc2626','#5e34b5','#059669','#be185d','#a16207','#0e7490','#4d7c0f'];
  const MES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

  let TURNOVER = [];
  let periodo = '12m';      // '12m' | 'ano' | 'todos'
  let supervisorFiltro = '';
  let tipoFiltro = '';
  let busca = '';
  const charts = {};

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const norm = s => String(s ?? '').replace(/\s+/g,' ').trim();
  const pad2 = n => n < 10 ? '0' + n : String(n);
  function dataIso(s){
    const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  }
  function dataBr(s){
    const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[3] + '/' + m[2] + '/' + m[1] : '—';
  }
  const selStyle = 'font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); min-width:180px;';
  const lblStyle = 'font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em;';

  // ---------- Menu + página ----------
  function montarEstrutura(){
    const sidebar = document.querySelector('.sidebar');
    const main = document.querySelector('.main-content');
    if(!sidebar || !main || document.getElementById('nav-turnover')) return;

    const nav = document.createElement('div');
    nav.className = 'nav-item';
    nav.id = 'nav-turnover';
    nav.innerHTML = '<span class="nav-icon">🔄</span> Turnover';
    nav.onclick = abrirTurnover;
    const ref = document.getElementById('nav-organograma');
    if(ref && ref.parentNode === sidebar) sidebar.insertBefore(nav, ref.nextSibling);
    else sidebar.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina';
    pg.id = 'pagina-turnover';
    pg.style.display = 'none';
    pg.innerHTML = `
      <div class="panel">
        <h2>Análise de Turnover</h2>
        <div id="toFonte" class="mono" style="font-size:11px; color:var(--muted); margin:2px 0 12px;"></div>

        <div style="margin:0 0 14px; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
          <label class="mono" style="${lblStyle}">Período</label>
          <select id="toPeriodo" class="mono" style="${selStyle}">
            <option value="12m">Últimos 12 meses</option>
            <option value="ano">Ano atual</option>
            <option value="todos">Todo o histórico</option>
          </select>
          <label class="mono" style="${lblStyle} margin-left:8px;">Supervisor</label>
          <select id="toSupervisor" class="mono" style="${selStyle}"><option value="">Todos</option></select>
          <label class="mono" style="${lblStyle} margin-left:8px;">Tipo</label>
          <select id="toTipo" class="mono" style="${selStyle}"><option value="">Todos</option></select>
          <button type="button" class="link-base" style="margin-left:auto; cursor:pointer; border:1px solid var(--accent); font-family:'JetBrains Mono', monospace;" id="toExportar">⬇ Exportar (CSV)</button>
        </div>

        <div class="kpi-row" id="toKpiRow" style="margin-bottom:16px;"></div>

        <div class="panel" style="margin-bottom:16px;">
          <h2>Pareto — Causas de Desligamento</h2>
          <div class="hint">Barras = quantidade por causa · linha = % acumulado · tracejado = corte de 80%</div>
          <div id="toParetoWrap" style="position:relative; height:340px;"><canvas id="toParetoChart"></canvas></div>
        </div>

        <div class="grid" style="grid-template-columns: 1fr 1.4fr; margin-bottom:16px;">
          <div class="panel">
            <h2>Tipo de Desligamento</h2>
            <div style="position:relative; height:280px;"><canvas id="toTipoChart"></canvas></div>
          </div>
          <div class="panel">
            <h2>Desligamentos por Mês</h2>
            <div style="position:relative; height:280px;"><canvas id="toMesChart"></canvas></div>
          </div>
        </div>

        <div class="grid" style="grid-template-columns: 1fr 1fr; margin-bottom:16px;">
          <div class="panel">
            <h2>Desligamentos por Supervisor</h2>
            <div id="toSupWrap" style="position:relative; height:300px;"><canvas id="toSupChart"></canvas></div>
          </div>
          <div class="panel">
            <h2>Tempo de Casa no Desligamento</h2>
            <div style="position:relative; height:300px;"><canvas id="toCasaChart"></canvas></div>
          </div>
        </div>

        <div class="panel" style="margin-bottom:16px;">
          <h2>Causas — o que entra em cada uma</h2>
          <table>
            <thead><tr><th>Causa</th><th class="num">Qtd</th><th class="num">%</th><th class="num">% Acum.</th><th>O que entra nessa causa</th></tr></thead>
            <tbody id="toCausasBody"></tbody>
          </table>
        </div>

        <div class="panel">
          <h2>Desligados</h2>
          <div style="display:flex; gap:8px; margin-bottom:10px; align-items:center; flex-wrap:wrap;">
            <div class="search">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
              <input id="toBusca" type="search" placeholder="Buscar por nome, supervisor, causa..." aria-label="Buscar desligados" />
            </div>
            <span id="toContagem" class="mono" style="font-size:11px; color:var(--muted); margin-left:auto;"></span>
          </div>
          <div style="overflow:auto; border-radius:8px;">
            <table>
              <thead><tr><th>Dt. Demissão</th><th>Data Admis.</th><th class="num">Tempo de casa</th><th>Nome</th><th>Supervisor</th><th>Tipo</th><th>Causa Pareto</th></tr></thead>
              <tbody id="toTabelaBody"></tbody>
            </table>
          </div>
        </div>
      </div>`;
    main.appendChild(pg);

    document.getElementById('toPeriodo').onchange = e => { periodo = e.target.value; render(); };
    document.getElementById('toSupervisor').onchange = e => { supervisorFiltro = e.target.value; render(); };
    document.getElementById('toTipo').onchange = e => { tipoFiltro = e.target.value; render(); };
    document.getElementById('toExportar').onclick = exportarCsv;
    let t;
    document.getElementById('toBusca').addEventListener('input', e => {
      clearTimeout(t);
      t = setTimeout(() => { busca = e.target.value.toLowerCase().trim(); renderTabela(filtrados()); }, 200);
    });
  }

  function abrirTurnover(){
    document.querySelectorAll('[id^="pagina-"]').forEach(p => { p.style.display = 'none'; });
    document.querySelectorAll('.sidebar .nav-item').forEach(n => n.classList.remove('active'));
    document.getElementById('pagina-turnover').style.display = '';
    document.getElementById('nav-turnover').classList.add('active');
    render();
  }

  // Quando o usuário clica em outra página do menu, esconde a de Turnover.
  if(typeof window.mostrarPagina === 'function'){
    const original = window.mostrarPagina;
    window.mostrarPagina = function(){
      const pg = document.getElementById('pagina-turnover');
      if(pg) pg.style.display = 'none';
      const nav = document.getElementById('nav-turnover');
      if(nav) nav.classList.remove('active');
      return original.apply(this, arguments);
    };
  }

  // ---------- Dados ----------
  function tempoCasaMeses(r){
    const a = dataIso(r.dataAdmis), d = dataIso(r.dataDemissao);
    if(!a || !d || d < a) return null;
    return (d - a) / (86400000 * 30.44);
  }

  // Dias corridos entre admissão e demissão (null se faltar data).
  function tempoCasaDias(r){
    const a = dataIso(r.dataAdmis), d = dataIso(r.dataDemissao);
    if(!a || !d || d < a) return null;
    return Math.round((d - a) / 86400000);
  }

  // Texto amigável a partir de uma quantidade de dias:
  //   menos de 1 mês  -> "22 dias"
  //   menos de 1 ano  -> "2 meses e 5 dias"
  //   1 ano ou mais   -> "1 ano e 3 meses"
  function formatarDias(dias){
    if(dias === null || dias === undefined || isNaN(dias)) return '—';
    dias = Math.round(dias);
    const pl = (n, s, p) => n + ' ' + (n === 1 ? s : p);
    if(dias < 30) return pl(dias, 'dia', 'dias');
    if(dias < 365){
      const meses = Math.floor(dias / 30.44);
      const resto = Math.round(dias - meses * 30.44);
      return resto > 0 ? pl(meses, 'mês', 'meses') + ' e ' + pl(resto, 'dia', 'dias') : pl(meses, 'mês', 'meses');
    }
    const anos = Math.floor(dias / 365.25);
    const meses = Math.floor((dias - anos * 365.25) / 30.44);
    return meses > 0 ? pl(anos, 'ano', 'anos') + ' e ' + pl(meses, 'mês', 'meses') : pl(anos, 'ano', 'anos');
  }

  function filtrados(){
    const hoje = new Date();
    const inicio12 = new Date(hoje.getFullYear(), hoje.getMonth() - 11, 1);
    return TURNOVER.filter(r => {
      if(periodo !== 'todos'){
        const d = dataIso(r.dataDemissao);
        if(!d) return false;
        if(periodo === '12m' && d < inicio12) return false;
        if(periodo === 'ano' && d.getFullYear() !== hoje.getFullYear()) return false;
      }
      if(supervisorFiltro && norm(r.supervisor) !== supervisorFiltro) return false;
      if(tipoFiltro && (norm(r.tipoDesligamento) || 'Não informado') !== tipoFiltro) return false;
      return true;
    });
  }

  function contar(itens, campo){
    const m = {};
    itens.forEach(r => { const k = norm(r[campo]) || 'Não informado'; m[k] = (m[k] || 0) + 1; });
    return Object.entries(m).sort((a,b) => b[1] - a[1]);
  }

  function novoChart(id, cfg){
    const el = document.getElementById(id);
    if(!el) return;
    if(charts[id]) charts[id].destroy();
    if(typeof ChartDataLabels !== 'undefined') cfg.plugins = [ChartDataLabels];
    charts[id] = new Chart(el, cfg);
  }

  const escalas = (horizontal) => ({
    x: { ticks:{color:C.text, font:{weight:'bold', size:10}}, grid:{display:false}, beginAtZero:true },
    y: { ticks:{color:C.text, font:{weight:'bold', size: horizontal ? 11 : 10}}, grid:{display:false}, beginAtZero:true }
  });
  const rotulo = { color:'#000000', anchor:'end', align:'top', font:{size:10, family:'JetBrains Mono', weight:'700'}, formatter: v => v > 0 ? v : '' };

  // ---------- Render ----------
  function popularFiltros(){
    const sSup = document.getElementById('toSupervisor');
    const sTipo = document.getElementById('toTipo');
    const sups = Array.from(new Set(TURNOVER.map(r => norm(r.supervisor)).filter(Boolean))).sort((a,b)=>a.localeCompare(b,'pt-BR'));
    const tipos = Array.from(new Set(TURNOVER.map(r => norm(r.tipoDesligamento) || 'Não informado'))).sort((a,b)=>a.localeCompare(b,'pt-BR'));
    sSup.innerHTML = '<option value="">Todos</option>' + sups.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
    sTipo.innerHTML = '<option value="">Todos</option>' + tipos.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
    sSup.value = supervisorFiltro;
    sTipo.value = tipoFiltro;
  }

  function render(){
    const pg = document.getElementById('pagina-turnover');
    if(!pg || pg.style.display === 'none') return;

    const fonte = document.getElementById('toFonte');
    if(!TURNOVER.length){
      fonte.textContent = 'Sem dados de turnover no dados.json — confira o script auxiliar Turnover_LerBase e o parâmetro turnoverJson no Power Automate.';
      fonte.style.color = C.bad;
    } else {
      fonte.textContent = TURNOVER.length.toLocaleString('pt-BR') + ' desligamentos na base (aba "Base" da Análise de Turnover)';
      fonte.style.color = '';
    }

    popularFiltros();
    const itens = filtrados();
    const total = itens.length;

    // KPIs
    const porTipo = contar(itens, 'tipoDesligamento');
    const casas = itens.map(tempoCasaDias).filter(v => v !== null);
    const mediaCasa = casas.length ? casas.reduce((a,b)=>a+b,0) / casas.length : null;
    const ate90 = itens.filter(r => { const m = tempoCasaMeses(r); return m !== null && m < 3; }).length;
    const causas = contar(itens, 'causaPareto');
    const topCausa = causas[0];

    document.getElementById('toKpiRow').innerHTML =
      `<div class="kpi"><div class="label">🔄 Desligamentos</div><div class="value" style="color:${C.accent}">${total.toLocaleString('pt-BR')}</div></div>` +
      porTipo.slice(0, 3).map(([k, q], i) => `<div class="kpi"><div class="label">${esc(k)}</div><div class="value" style="color:${PALETA[(i+1) % PALETA.length]}">${q}</div><div class="delta">${total ? (q/total*100).toFixed(1) : 0}% do total</div></div>`).join('') +
      `<div class="kpi"><div class="label">⏳ Tempo médio de casa</div><div class="value" style="font-size:20px;">${formatarDias(mediaCasa)}</div><div class="delta">entre admissão e demissão</div></div>` +
      `<div class="kpi" ${ate90 ? `style="border-color:${C.bad}; box-shadow:0 0 0 2px ${C.bad} inset;"` : ''}><div class="label">⚠️ Saíram em até 90 dias</div><div class="value" style="color:${ate90 ? C.bad : C.good}">${ate90}</div><div class="delta">${total ? (ate90/total*100).toFixed(1) : 0}% do total</div></div>` +
      (topCausa ? `<div class="kpi"><div class="label">🥇 Principal causa</div><div class="value" style="font-size:16px;">${esc(topCausa[0])}</div><div class="delta">${topCausa[1]} (${(topCausa[1]/total*100).toFixed(1)}%)</div></div>` : '');

    // Pareto
    let acum = 0;
    const acumPct = causas.map(([,q]) => { acum += q; return total ? acum / total * 100 : 0; });
    const cortes80 = acumPct.findIndex(v => v >= 80);
    document.getElementById('toParetoWrap').style.height = Math.max(300, Math.min(520, 260 + causas.length * 8)) + 'px';
    novoChart('toParetoChart', {
      data: {
        labels: causas.map(c => c[0]),
        datasets: [
          { type:'bar', label:'Desligamentos', data: causas.map(c => c[1]), order:2, yAxisID:'y', borderRadius:4,
            backgroundColor: causas.map((_, i) => (cortes80 === -1 || i <= cortes80) ? C.bad : 'rgba(107,117,144,.6)'),
            datalabels: rotulo },
          { type:'line', label:'% Acumulado', data: acumPct, order:1, yAxisID:'y1', borderColor:C.accent, backgroundColor:C.accent, pointRadius:4, tension:.2,
            datalabels:{ color:C.accent, align:'top', offset:6, font:{size:9, family:'JetBrains Mono', weight:'700'}, formatter: v => v.toFixed(0) + '%' } },
          { type:'line', label:'80%', data: causas.map(() => 80), order:1, yAxisID:'y1', borderColor:'#000000', borderDash:[5,5], pointRadius:0, datalabels:{display:false} }
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ labels:{ color:C.muted, font:{family:'Inter', size:11} } }, datalabels:{display:true} },
        scales:{
          x:{ ticks:{color:C.text, font:{size:10, weight:'bold'}, maxRotation:45, minRotation:0}, grid:{display:false} },
          y:{ ticks:{color:C.text, font:{weight:'bold'}}, grid:{display:false}, beginAtZero:true },
          y1:{ position:'right', min:0, max:110, ticks:{color:C.text, font:{weight:'bold'}, callback: v => v <= 100 ? v + '%' : ''}, grid:{display:false} }
        }
      }
    });

    // Tipo de desligamento
    novoChart('toTipoChart', {
      type:'doughnut',
      data:{ labels: porTipo.map(t => t[0]), datasets:[{ data: porTipo.map(t => t[1]), backgroundColor: porTipo.map((_, i) => PALETA[i % PALETA.length]), borderWidth:0,
        datalabels:{ color:'#ffffff', font:{size:11, family:'JetBrains Mono', weight:'700'}, formatter: v => total ? Math.round(v/total*100) + '%' : '' } }] },
      options:{ responsive:true, maintainAspectRatio:false, cutout:'58%',
        plugins:{ legend:{ position:'bottom', labels:{ color:C.muted, font:{family:'Inter', size:10}, boxWidth:10 } }, datalabels:{display:true} } }
    });

    // Por mês (empilhado por tipo)
    const chavesMes = [];
    if(periodo === 'todos'){
      const ks = Array.from(new Set(itens.map(r => (r.dataDemissao || '').slice(0,7)).filter(Boolean))).sort();
      ks.forEach(k => chavesMes.push(k));
    } else {
      const hoje = new Date();
      const n = periodo === '12m' ? 12 : hoje.getMonth() + 1;
      for(let i = n - 1; i >= 0; i--){
        const d = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
        chavesMes.push(d.getFullYear() + '-' + pad2(d.getMonth() + 1));
      }
    }
    const tiposMes = porTipo.map(t => t[0]);
    novoChart('toMesChart', {
      type:'bar',
      data:{
        labels: chavesMes.map(k => MES_PT[Number(k.slice(5,7)) - 1] + '/' + k.slice(2,4)),
        datasets: tiposMes.map((tipo, i) => ({
          label: tipo, backgroundColor: PALETA[i % PALETA.length], borderRadius:3,
          data: chavesMes.map(k => itens.filter(r => (r.dataDemissao || '').slice(0,7) === k && (norm(r.tipoDesligamento) || 'Não informado') === tipo).length),
          datalabels:{ display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) > 0, color:'#ffffff', font:{size:9, family:'JetBrains Mono', weight:'700'} }
        }))
      },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ labels:{ color:C.muted, font:{family:'Inter', size:10} } }, datalabels:{display:true} },
        scales:{ x:{ stacked:true, ticks:{color:C.text, font:{size:10, weight:'bold'}}, grid:{display:false} },
                 y:{ stacked:true, ticks:{color:C.text, font:{weight:'bold'}}, grid:{display:false}, beginAtZero:true } } }
    });

    // Por supervisor
    const porSup = contar(itens, 'supervisor');
    document.getElementById('toSupWrap').style.height = Math.max(260, porSup.length * 26 + 40) + 'px';
    novoChart('toSupChart', {
      type:'bar',
      data:{ labels: porSup.map(s => s[0]), datasets:[{ label:'Desligamentos', data: porSup.map(s => s[1]), backgroundColor:'rgba(37,99,235,.75)', borderRadius:4,
        datalabels:{ color:'#000000', anchor:'end', align:'right', font:{size:10, family:'JetBrains Mono', weight:'700'} } }] },
      options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, layout:{ padding:{ right:24 } },
        plugins:{ legend:{display:false}, datalabels:{display:true} }, scales: escalas(true) }
    });

    // Tempo de casa (faixas)
    const faixas = [
      { nome:'Até 3 meses', min:0, max:3 }, { nome:'3 a 6 meses', min:3, max:6 },
      { nome:'6 a 12 meses', min:6, max:12 }, { nome:'1 a 2 anos', min:12, max:24 },
      { nome:'Mais de 2 anos', min:24, max:Infinity }
    ];
    const qtdFaixa = faixas.map(f => itens.filter(r => { const m = tempoCasaMeses(r); return m !== null && m >= f.min && m < f.max; }).length);
    const semData = itens.filter(r => tempoCasaMeses(r) === null).length;
    const labelsFaixa = faixas.map(f => f.nome).concat(semData ? ['Sem data'] : []);
    novoChart('toCasaChart', {
      type:'bar',
      data:{ labels: labelsFaixa, datasets:[{ label:'Desligamentos', data: qtdFaixa.concat(semData ? [semData] : []),
        backgroundColor: [C.bad, C.warn, '#2563eb', C.accent, C.good, '#9ca3af'], borderRadius:4, datalabels: rotulo }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, datalabels:{display:true} }, scales: escalas(false) }
    });

    // Tabela de causas
    const oQueEntraPorCausa = {};
    itens.forEach(r => {
      const k = norm(r.causaPareto) || 'Não informado';
      const t = norm(r.oQueEntra);
      if(!t) return;
      if(!oQueEntraPorCausa[k]) oQueEntraPorCausa[k] = {};
      oQueEntraPorCausa[k][t] = (oQueEntraPorCausa[k][t] || 0) + 1;
    });
    let ac = 0;
    document.getElementById('toCausasBody').innerHTML = causas.length ? causas.map(([k, q]) => {
      ac += q;
      const textos = Object.entries(oQueEntraPorCausa[k] || {}).sort((a,b) => b[1] - a[1]).map(e => e[0]);
      const pctAc = total ? ac / total * 100 : 0;
      return `<tr>
        <td style="font-weight:600;">${esc(k)}</td>
        <td class="num">${q}</td>
        <td class="num">${total ? (q/total*100).toFixed(1) : '0.0'}%</td>
        <td class="num"><span class="pill ${pctAc <= 80 ? 'bad' : 'warn'}">${pctAc.toFixed(1)}%</span></td>
        <td style="font-size:12px; color:${C.text};">${textos.length ? textos.map(esc).join('<br>') : '—'}</td>
      </tr>`;
    }).join('') : `<tr><td colspan="5" style="text-align:center; color:${C.muted};">Nenhum desligamento no período</td></tr>`;

    renderTabela(itens);
  }

  function renderTabela(itens){
    const body = document.getElementById('toTabelaBody');
    if(!body) return;
    const lista = itens.filter(r => {
      if(!busca) return true;
      return [r.nome, r.matricula, r.supervisor, r.causaPareto, r.tipoDesligamento, r.oQueEntra].map(x => String(x || '').toLowerCase()).join(' ').includes(busca);
    }).sort((a,b) => String(b.dataDemissao || '').localeCompare(String(a.dataDemissao || '')));
    document.getElementById('toContagem').textContent = lista.length.toLocaleString('pt-BR') + ' desligamento' + (lista.length === 1 ? '' : 's');
    body.innerHTML = lista.length ? lista.map(r => {
      const dias = tempoCasaDias(r);
      const corCasa = dias !== null && dias < 90 ? ` style="color:${C.bad} !important;"` : '';
      return `<tr>
        <td class="mono" style="font-size:12px;">${dataBr(r.dataDemissao)}</td>
        <td class="mono" style="font-size:12px;">${dataBr(r.dataAdmis)}</td>
        <td class="num"${corCasa}>${formatarDias(dias)}</td>
        <td>${esc(r.nome) || '—'}</td>
        <td>${esc(r.supervisor) || '—'}</td>
        <td>${esc(r.tipoDesligamento) || '—'}</td>
        <td>${esc(r.causaPareto) || '—'}</td>
      </tr>`;
    }).join('') : `<tr><td colspan="7" style="text-align:center; color:${C.muted};">Nenhum desligamento encontrado</td></tr>`;
  }

  function exportarCsv(){
    const itens = filtrados();
    if(!itens.length){ alert('Nenhum desligamento para exportar com os filtros atuais.'); return; }
    const csv = s => { const t = String(s ?? ''); return /[",;\n\r]/.test(t) ? '"' + t.replace(/"/g,'""') + '"' : t; };
    const linhas = [['DtDemissao','DataAdmis','TempoCasaDias','TempoCasa','Nome','Matricula','Supervisor','TipoDesligamento','CausaPareto','OQueEntra'].join(';')];
    itens.forEach(r => {
      const dias = tempoCasaDias(r);
      linhas.push([dataBr(r.dataDemissao), dataBr(r.dataAdmis), dias === null ? '' : dias, formatarDias(dias), r.nome, r.matricula, r.supervisor, r.tipoDesligamento, r.causaPareto, r.oQueEntra].map(csv).join(';'));
    });
    const blob = new Blob(['\uFEFF' + linhas.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'Turnover_Desligamentos.csv'; a.style.display = 'none';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function carregar(){
    try{
      const resp = await fetch('dados.json?_=' + Date.now());
      if(!resp.ok) throw new Error('HTTP ' + resp.status);
      const data = await resp.json();
      TURNOVER = Array.isArray(data.turnover) ? data.turnover : [];
      render();
    }catch(err){
      console.error('[Turnover] erro ao carregar dados.json:', err);
    }
  }

  montarEstrutura();
  carregar();
  setInterval(carregar, 5 * 60 * 1000);
})();
