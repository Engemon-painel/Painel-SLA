// ======================================================================
// GESTÃO DE ESCALA — módulo do Painel
// ----------------------------------------------------------------------
// Cria o item "📅 Gestão de Escala" no menu e a página: quantos técnicos
// cada supervisão tem por dia e por horário, quantos escalados (X), de
// folga (FO), faltas e atestados. Dados: dados_escala.json (aba
// "ESCALA PAULO" da planilha de escala — sem telefone/endereço).
// No index.html: <script src="gestao_escala.js"></script> antes de </body>
// ======================================================================
(function(){

  // Categorias dos códigos da escala
  const CATS = [
    { k:'X',      label:'Escalado',     cor:'#059669', curto:'X'  },
    { k:'TREIN',  label:'Treinamento',  cor:'#2563eb', curto:'TR' },
    { k:'FO',     label:'Folga',        cor:'#94a3b8', curto:'FO' },
    { k:'FBH',    label:'Folga BH',     cor:'#0e7490', curto:'BH' },
    { k:'FALTA',  label:'Falta',        cor:'#dc2626', curto:'F'  },
    { k:'ATEST',  label:'Atestado',     cor:'#d97706', curto:'AT' },
    { k:'OUTRO',  label:'Outro',        cor:'#5e34b5', curto:'?'  }
  ];
  const CAT = {}; CATS.forEach(c => CAT[c.k] = c);
  function categoria(cod){
    const c = String(cod || '').toUpperCase().trim();
    if(!c) return '';
    if(c === 'X') return 'X';
    if(c.startsWith('TREIN')) return 'TREIN';
    if(c === 'FO' || c === 'FOLGA') return 'FO';
    if(c === 'FBH' || c.startsWith('FOLGA BH')) return 'FBH';
    if(c === 'F' || c.startsWith('FALTA')) return 'FALTA';
    if(c === 'ATM' || c.startsWith('ATEST')) return 'ATEST';
    return 'OUTRO';
  }

  let ESC = null, escErro = null;
  const filtro = { mes:'', dia:'', sup:'', area:'', turno:'', q:'' };
  let chartDias = null, chartSup = null;

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const fmt = n => Number(n).toLocaleString('pt-BR');
  const pad2 = n => String(n).padStart(2, '0');
  const SEMANA = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  const MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const MESES_LONGOS = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  const diaSemana = iso => { const p = iso.split('-'); return new Date(+p[0], +p[1]-1, +p[2]).getDay(); };
  const dm = iso => iso.slice(8,10) + '/' + iso.slice(5,7);
  const mesLabel = ym => MESES[+ym.slice(5,7) - 1] + '/' + ym.slice(2,4);
  function hojeIso(){ const d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth()+1) + '-' + pad2(d.getDate()); }

  // ---------------------------------------------------------------- estrutura
  const CSS = `
    #pagina-escala .es-filtros{ margin:10px 0 14px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
    #pagina-escala .es-lbl{ font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em; font-family:'JetBrains Mono', monospace; }
    #pagina-escala select{ font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); font-family:'JetBrains Mono', monospace; }
    #esDias{ display:flex; gap:4px; flex-wrap:wrap; margin:4px 0 16px; }
    #esDias button{ font-family:'JetBrains Mono', monospace; font-size:11px; min-width:40px; padding:5px 4px; border-radius:8px; border:1px solid var(--line);
      background:var(--panel); color:var(--text); cursor:pointer; line-height:1.15; }
    #esDias button small{ display:block; font-size:9px; color:var(--muted); }
    #esDias button.fds{ background:#f1f4f9; }
    #esDias button.hoje{ border-color:var(--accent); }
    #esDias button.ativo{ background:var(--accent); color:#fff; border-color:var(--accent); }
    #esDias button.ativo small{ color:#d7f0f2; }
    #esDias button .es-pt{ display:block; width:6px; height:6px; border-radius:50%; background:#dc2626; margin:2px auto 0; }
    #esKpis{ grid-template-columns: repeat(auto-fit, minmax(125px,1fr)); }
    #esTabela td, #esTabela th{ white-space:nowrap; }
    #esTabela tr.es-sup td{ background:#e3e9f1; font-weight:700; }
    #esTabela td.nomes{ white-space:normal; font-size:11.5px; color:#b91c1c; min-width:160px; }
    #esMatriz th, #esMatriz td{ padding:4px 3px; text-align:center; font-size:11px; white-space:nowrap; }
    #esMatriz td.nm, #esMatriz th.nm{ text-align:left; position:sticky; left:0; background:#eef1f5; z-index:1; font-size:12px; }
    #esMatriz thead th{ position:sticky; top:0; background:#eef1f5; z-index:2; }
    #esMatriz thead th.nm{ z-index:3; }
    #esMatriz th.sel, #esMatriz td.sel{ outline:2px solid var(--accent); outline-offset:-2px; }
    #esMatriz .cd{ display:inline-block; min-width:24px; padding:2px 3px; border-radius:5px; font-family:'JetBrains Mono', monospace; font-weight:700; font-size:10px; }
    #esMatriz tr.es-grupo td{ background:#e3e9f1; font-weight:700; text-align:left; font-size:12px; }
    #pagina-escala .es-legenda{ display:flex; gap:10px; flex-wrap:wrap; margin:4px 0 10px; }
    #pagina-escala .es-legenda span{ font-size:11px; display:flex; align-items:center; gap:5px; color:var(--muted); }
    #pagina-escala .es-legenda i{ width:10px; height:10px; border-radius:3px; display:inline-block; }
    .es-dupla{ font-size:9.5px; font-weight:700; color:#5e34b5; background:rgba(94,52,181,.12); padding:1px 5px; border-radius:999px; margin-left:4px; }
  `;

  const TEMPLATE = `
    <div class="panel">
      <h2>Gestão de Escala</h2>
      <div class="hint mono" id="esInfo">Carregando dados da escala...</div>

      <div class="es-filtros">
        <label class="es-lbl" for="esMes">Mês</label>
        <select id="esMes" style="min-width:110px;"></select>
        <label class="es-lbl" for="esSup" style="margin-left:8px;">Supervisão</label>
        <select id="esSup" style="min-width:170px;"></select>
        <label class="es-lbl" for="esArea" style="margin-left:8px;">Área</label>
        <select id="esArea" style="min-width:100px;"></select>
        <label class="es-lbl" for="esTurno" style="margin-left:8px;">Turno</label>
        <select id="esTurno" style="min-width:120px;"></select>
      </div>

      <div class="es-lbl" style="margin-bottom:4px;">Dia <span style="text-transform:none; letter-spacing:0;">(ponto vermelho = teve falta ou atestado)</span></div>
      <div id="esDias"></div>

      <div class="kpi-row" id="esKpis"></div>

      <div class="panel" style="margin-bottom:16px;">
        <h2 id="esTituloTabela">Por supervisão e horário</h2>
        <div class="hint">Quantos técnicos cada supervisão tem em cada horário no dia escolhido. Passe o mouse em "Escalados" para ver os nomes.</div>
        <div style="overflow:auto;">
          <table id="esTabela">
            <thead><tr>
              <th>Supervisão / Horário</th><th>Turno</th><th class="num">Técnicos</th><th class="num">Escalados</th>
              <th class="num">Folga</th><th class="num">Folga BH</th><th class="num">Faltas</th><th class="num">Atestado</th><th class="num">Treinam.</th>
              <th>Faltas / atestados</th>
            </tr></thead>
            <tbody id="esTbody"></tbody>
          </table>
        </div>
      </div>

      <div class="grid" style="grid-template-columns: 1.5fr 1fr;">
        <div class="panel">
          <h2>Dia a dia do mês</h2>
          <div class="hint">Clique numa barra para escolher o dia.</div>
          <div style="position:relative; height:300px;"><canvas id="esDiasChart"></canvas></div>
        </div>
        <div class="panel">
          <h2>Faltas e atestados no mês</h2>
          <div class="hint">Por supervisão (dias de falta/atestado somados).</div>
          <div id="esSupWrap" style="position:relative; height:300px;"><canvas id="esSupChart"></canvas></div>
        </div>
      </div>

      <div class="panel">
        <h2>Escala do mês</h2>
        <div class="es-filtros">
          <div class="search" title="Buscar técnico">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
            <input id="esBusca" type="search" placeholder="Nome do técnico..." aria-label="Buscar técnico" style="min-width:220px;" />
          </div>
          <span id="esContagem" class="mono" style="font-size:11px; color:var(--muted);"></span>
          <button type="button" class="link-base" id="esExportar" style="margin-left:auto; cursor:pointer; border:1px solid var(--accent); font-family:'JetBrains Mono', monospace;">⬇ Exportar mês (CSV)</button>
        </div>
        <div class="es-legenda" id="esLegenda"></div>
        <div style="overflow:auto; max-height:70vh; border-radius:8px;">
          <table id="esMatriz"><thead id="esMThead"></thead><tbody id="esMTbody"></tbody></table>
        </div>
      </div>
    </div>
  `;

  function montarEstrutura(){
    if(document.getElementById('pagina-escala')) return;
    const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);

    const nav = document.createElement('div');
    nav.className = 'nav-item'; nav.id = 'nav-escala';
    nav.innerHTML = '<span class="nav-icon">📅</span> Gestão de Escala';
    nav.onclick = () => mostrarPagina('escala');
    const ancora = document.getElementById('nav-aso') || document.getElementById('nav-nr') || document.getElementById('nav-frota');
    if(ancora) ancora.after(nav); else document.querySelector('.sidebar')?.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina'; pg.id = 'pagina-escala'; pg.style.display = 'none';
    pg.innerHTML = TEMPLATE;
    document.querySelector('main.main-content')?.appendChild(pg);

    if(typeof PAGINAS !== 'undefined' && PAGINAS.indexOf('escala') === -1) PAGINAS.push('escala');
    const mostrarOriginal = window.mostrarPagina;
    window.mostrarPagina = function(nome){
      mostrarOriginal(nome);
      if(nome === 'escala') renderEscala();
    };

    document.getElementById('esMes').addEventListener('change', e => { filtro.mes = e.target.value; filtro.dia = ''; renderEscala(); });
    [['esSup','sup'],['esArea','area'],['esTurno','turno']].forEach(([id, campo]) =>
      document.getElementById(id).addEventListener('change', e => { filtro[campo] = e.target.value; renderEscala(); }));
    let t;
    document.getElementById('esBusca').addEventListener('input', e => {
      clearTimeout(t); t = setTimeout(() => { filtro.q = e.target.value.trim().toLowerCase(); renderMatriz(); }, 200);
    });
    document.getElementById('esExportar').addEventListener('click', exportarCsv);
    document.getElementById('esLegenda').innerHTML = CATS.filter(c => c.k !== 'OUTRO').map(c =>
      `<span><i style="background:${c.cor};"></i>${c.curto} = ${c.label}</span>`).join('') + '<span><i style="background:#e5e7eb;"></i>— = sem escala</span>';
  }

  // ---------------------------------------------------------------- dados
  async function carregarEscala(){
    try{
      const r = await fetch('dados_escala.json?_=' + Date.now());
      if(!r.ok) throw new Error('HTTP ' + r.status);
      ESC = await r.json(); escErro = null;
      window.EscalaInfo = { atualizado_em: ESC.atualizado_em };
      window.dispatchEvent(new Event('painel-dados'));
      popularFiltros();
    }catch(e){ escErro = e.message; console.error('[Escala]', e); }
    const pg = document.getElementById('pagina-escala');
    if(pg && pg.style.display !== 'none') renderEscala();
  }

  function mesesDisponiveis(){
    // meses que têm pelo menos um código preenchido
    const set = new Set();
    ESC.datas.forEach((d, i) => { if(ESC.tecnicos.some(t => t.codigos[i])) set.add(d.slice(0,7)); });
    return Array.from(set).sort();
  }

  function popularFiltros(){
    const meses = mesesDisponiveis();
    const hojeMes = hojeIso().slice(0,7);
    if(!filtro.mes || meses.indexOf(filtro.mes) === -1) filtro.mes = meses.indexOf(hojeMes) !== -1 ? hojeMes : meses[meses.length - 1];
    const selMes = document.getElementById('esMes');
    selMes.innerHTML = meses.map(m => `<option value="${m}">${mesLabel(m)}</option>`).join('');
    selMes.value = filtro.mes;
    const preencher = (id, campo, valores) => {
      const sel = document.getElementById(id);
      if(filtro[campo] && valores.indexOf(filtro[campo]) === -1) filtro[campo] = '';
      sel.innerHTML = '<option value="">Todos</option>' + valores.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
      sel.value = filtro[campo];
    };
    const unicos = campo => Array.from(new Set(ESC.tecnicos.map(t => t[campo]).filter(Boolean))).sort((a,b) => a.localeCompare(b,'pt-BR'));
    preencher('esSup', 'sup', unicos('supervisao'));
    preencher('esArea', 'area', unicos('area'));
    preencher('esTurno', 'turno', unicos('turno'));
  }

  const tecnicosFiltrados = () => ESC.tecnicos.filter(t =>
    (!filtro.sup || t.supervisao === filtro.sup) && (!filtro.area || t.area === filtro.area) && (!filtro.turno || t.turno === filtro.turno));
  const idxDatasDoMes = () => ESC.datas.map((d, i) => d.slice(0,7) === filtro.mes ? i : -1).filter(i => i >= 0);

  function contarDia(tecs, i){
    const c = { total:0, X:0, TREIN:0, FO:0, FBH:0, FALTA:0, ATEST:0, OUTRO:0 };
    tecs.forEach(t => { const k = categoria(t.codigos[i]); if(k){ c.total++; c[k]++; } });
    return c;
  }

  // ---------------------------------------------------------------- render
  function renderEscala(){
    const info = document.getElementById('esInfo');
    if(!info) return;
    if(escErro){ info.textContent = 'Não foi possível carregar dados_escala.json (' + escErro + ').'; info.style.color = '#dc2626'; return; }
    if(!ESC){ info.textContent = 'Carregando dados da escala...'; return; }
    info.style.color = '';
    const p = (ESC.atualizado_em || '').split('-');
    info.textContent = 'Escala atualizada em ' + (p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : '—') + ' · ' +
      fmt(ESC.tecnicos.length) + ' técnicos na planilha · ' + MESES_LONGOS[+filtro.mes.slice(5,7) - 1] + ' de ' + filtro.mes.slice(0,4);

    const idx = idxDatasDoMes();
    const tecs = tecnicosFiltrados();
    if(!filtro.dia || idx.every(i => ESC.datas[i] !== filtro.dia)){
      const hoje = hojeIso();
      filtro.dia = idx.map(i => ESC.datas[i]).includes(hoje) ? hoje : ESC.datas[idx[0]];
    }
    renderDias(idx, tecs);
    renderKpis(tecs);
    renderTabela(tecs);
    renderChartDias(idx, tecs);
    renderChartSup(idx, tecs);
    renderMatriz();
  }

  function renderDias(idx, tecs){
    const hoje = hojeIso();
    const box = document.getElementById('esDias');
    box.innerHTML = idx.map(i => {
      const d = ESC.datas[i], ws = diaSemana(d);
      const c = contarDia(tecs, i);
      const cls = [ws === 0 || ws === 6 ? 'fds' : '', d === hoje ? 'hoje' : '', d === filtro.dia ? 'ativo' : ''].join(' ');
      return `<button type="button" class="${cls}" data-dia="${d}" title="${dm(d)} · ${c.X + c.TREIN} escalados · ${c.FO + c.FBH} folga">${d.slice(8,10)}<small>${SEMANA[ws]}</small>${c.FALTA + c.ATEST ? '<span class="es-pt"></span>' : ''}</button>`;
    }).join('');
    box.querySelectorAll('button').forEach(b => b.onclick = () => { filtro.dia = b.dataset.dia; renderEscala(); });
  }

  function iDia(){ return ESC.datas.indexOf(filtro.dia); }

  function renderKpis(tecs){
    const c = contarDia(tecs, iDia());
    const cartao = (icone, titulo, valor, cor, sub) => `<div class="kpi"><div class="label">${icone} ${titulo}</div>
      <div class="value" style="color:${cor}">${fmt(valor)}</div><div class="delta" style="color:var(--muted)">${sub || ''}</div></div>`;
    const pct = v => c.total ? (v / c.total * 100).toFixed(0) + '% da escala' : '';
    const ws = diaSemana(filtro.dia);
    document.getElementById('esKpis').innerHTML =
      cartao('📅', SEMANA[ws] + ' ' + dm(filtro.dia), c.total, 'var(--accent)', 'técnicos na escala') +
      cartao('✅', 'Escalados (X)', c.X, CAT.X.cor, pct(c.X)) +
      cartao('🛌', 'Folga (FO)', c.FO, '#6b7590', pct(c.FO)) +
      cartao('🏦', 'Folga BH', c.FBH, CAT.FBH.cor, pct(c.FBH)) +
      cartao('❌', 'Faltas', c.FALTA, CAT.FALTA.cor, pct(c.FALTA)) +
      cartao('🩺', 'Atestado', c.ATEST, CAT.ATEST.cor, pct(c.ATEST)) +
      cartao('📚', 'Treinamento', c.TREIN, CAT.TREIN.cor, pct(c.TREIN));
  }

  function renderTabela(tecs){
    const i = iDia();
    document.getElementById('esTituloTabela').textContent = 'Por supervisão e horário — ' + SEMANA[diaSemana(filtro.dia)] + ', ' + dm(filtro.dia);
    const grupos = {};
    tecs.forEach(t => {
      const k = categoria(t.codigos[i]); if(!k) return;
      const s = t.supervisao || 'Sem supervisão', h = t.horario || 'Sem horário';
      grupos[s] = grupos[s] || {};
      const g = grupos[s][h] = grupos[s][h] || { turno: t.turno, tecs: [] };
      g.tecs.push({ t, k });
    });
    const linha = (rot, turno, lista, cls) => {
      const qtd = k => lista.filter(x => x.k === k).length;
      const n = k => { const v = qtd(k); return v ? v : '<span style="color:#cbd5e1; font-weight:400;">0</span>'; };
      const escalados = lista.filter(x => x.k === 'X' || x.k === 'TREIN').map(x => x.t.nome).join('\n');
      const ausentes = lista.filter(x => x.k === 'FALTA' || x.k === 'ATEST').map(x => x.t.nome + ' (' + CAT[x.k].label.toLowerCase() + ')');
      const vermelho = k => qtd(k) ? ` style="color:#dc2626 !important; font-weight:700;"` : '';
      return `<tr class="${cls || ''}">
        <td>${esc(rot)}</td><td>${esc(turno || '')}</td><td class="num">${lista.length}</td>
        <td class="num" title="${esc(escalados)}" style="color:#059669; font-weight:700; cursor:help;">${n('X')}</td>
        <td class="num">${n('FO')}</td><td class="num">${n('FBH')}</td>
        <td class="num"${vermelho('FALTA')}>${n('FALTA')}</td><td class="num"${vermelho('ATEST')}>${n('ATEST')}</td>
        <td class="num">${n('TREIN')}</td>
        <td class="nomes">${esc(cls ? '' : ausentes.join(', '))}</td>
      </tr>`;
    };
    const sups = Object.keys(grupos).sort((a,b) => a.localeCompare(b,'pt-BR'));
    let html = '';
    let geral = [];
    sups.forEach(s => {
      const hs = Object.keys(grupos[s]).sort();
      const todos = hs.flatMap(h => grupos[s][h].tecs);
      geral = geral.concat(todos);
      html += linha('👤 ' + s, '', todos, 'es-sup');
      hs.forEach(h => { html += linha('   🕒 ' + h, grupos[s][h].turno, grupos[s][h].tecs); });
    });
    if(sups.length > 1) html += linha('Total', '', geral, 'es-sup');
    document.getElementById('esTbody').innerHTML = html ||
      '<tr><td colspan="10" style="text-align:center; color:var(--muted);">Ninguém na escala neste dia com os filtros escolhidos.</td></tr>';
  }

  function renderChartDias(idx, tecs){
    const canvas = document.getElementById('esDiasChart');
    if(!canvas || typeof Chart === 'undefined') return;
    const ordem = ['X','TREIN','FO','FBH','FALTA','ATEST'];
    const contagens = idx.map(i => contarDia(tecs, i));
    const datasets = ordem.map(k => ({
      label: CAT[k].label, data: contagens.map(c => c[k]), backgroundColor: CAT[k].cor, borderWidth:0,
      datalabels:{ display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) > 0 && (k === 'FALTA' || k === 'ATEST' || k === 'X'),
        color:'#ffffff', font:{ size:9, family:'JetBrains Mono', weight:'700' } }
    }));
    if(chartDias) chartDias.destroy();
    chartDias = new Chart(canvas, {
      type:'bar',
      data:{ labels: idx.map(i => ESC.datas[i].slice(8,10)), datasets },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options:{
        responsive:true, maintainAspectRatio:false,
        onClick:(evt, els) => { if(els.length){ filtro.dia = ESC.datas[idx[els[0].index]]; renderEscala(); } },
        onHover:(evt, els) => { evt.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        plugins:{
          legend:{ position:'bottom', labels:{ color:'#6b7590', font:{ family:'Inter', size:10 }, boxWidth:10, filter:(it, data) => data.datasets[it.datasetIndex].data.some(v => v > 0) } },
          datalabels:{ display:true },
          tooltip:{ callbacks:{ title: items => { const d = ESC.datas[idx[items[0].dataIndex]]; return SEMANA[diaSemana(d)] + ', ' + dm(d); } } }
        },
        scales:{
          x:{ stacked:true, ticks:{ color:'#1b2440', font:{ size:10, weight:'bold' } }, grid:{ display:false } },
          y:{ stacked:true, beginAtZero:true, ticks:{ color:'#1b2440', precision:0 }, grid:{ display:false } }
        }
      }
    });
  }

  function renderChartSup(idx, tecs){
    const canvas = document.getElementById('esSupChart');
    if(!canvas || typeof Chart === 'undefined') return;
    const porSup = {};
    tecs.forEach(t => {
      const s = t.supervisao || 'Sem supervisão';
      porSup[s] = porSup[s] || { FALTA:0, ATEST:0 };
      idx.forEach(i => { const k = categoria(t.codigos[i]); if(k === 'FALTA' || k === 'ATEST') porSup[s][k]++; });
    });
    const sups = Object.keys(porSup).sort((a,b) => (porSup[b].FALTA + porSup[b].ATEST) - (porSup[a].FALTA + porSup[a].ATEST));
    const wrap = document.getElementById('esSupWrap');
    if(wrap) wrap.style.height = Math.max(240, sups.length * 40 + 80) + 'px';
    if(chartSup) chartSup.destroy();
    chartSup = new Chart(canvas, {
      type:'bar',
      data:{ labels: sups, datasets: ['FALTA','ATEST'].map(k => ({
        label: CAT[k].label, data: sups.map(s => porSup[s][k]), backgroundColor: CAT[k].cor, borderWidth:0,
        datalabels:{ display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) > 0, color:'#ffffff', font:{ size:10, family:'JetBrains Mono', weight:'700' } }
      })) },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options:{
        indexAxis:'y', responsive:true, maintainAspectRatio:false, maxBarThickness:26,
        plugins:{ legend:{ position:'bottom', labels:{ color:'#6b7590', font:{ family:'Inter', size:10 }, boxWidth:10 } }, datalabels:{ display:true } },
        scales:{
          x:{ stacked:true, beginAtZero:true, ticks:{ color:'#1b2440', precision:0 }, grid:{ display:false } },
          y:{ stacked:true, ticks:{ color:'#1b2440', font:{ size:11, weight:'bold' } }, grid:{ display:false } }
        }
      }
    });
  }

  function linhasMatriz(){
    return tecnicosFiltrados().filter(t => !filtro.q || [t.nome, t.supervisao, t.horario].join(' ').toLowerCase().includes(filtro.q))
      .sort((a,b) => (a.supervisao || '').localeCompare(b.supervisao || '','pt-BR') || (a.horario || '').localeCompare(b.horario || '') || a.nome.localeCompare(b.nome,'pt-BR'));
  }

  function renderMatriz(){
    if(!ESC) return;
    const idx = idxDatasDoMes();
    const tecs = linhasMatriz();
    document.getElementById('esContagem').textContent = fmt(tecs.length) + ' técnico' + (tecs.length === 1 ? '' : 's');
    document.getElementById('esMThead').innerHTML = '<tr><th class="nm">Técnico</th><th>Horário</th>' +
      idx.map(i => { const d = ESC.datas[i]; return `<th class="${d === filtro.dia ? 'sel' : ''}">${d.slice(8,10)}<br><span style="font-weight:400;">${SEMANA[diaSemana(d)].slice(0,1)}</span></th>`; }).join('') + '</tr>';
    let supAtual = null, html = '';
    tecs.forEach(t => {
      if(t.supervisao !== supAtual){
        supAtual = t.supervisao;
        html += `<tr class="es-grupo"><td class="nm" colspan="${idx.length + 2}">👤 ${esc(supAtual || 'Sem supervisão')}</td></tr>`;
      }
      html += `<tr><td class="nm">${esc(t.nome)}${t.duplado ? '<span class="es-dupla">dupla</span>' : ''}</td><td class="mono" style="font-size:10.5px;">${esc(t.horario)}</td>` +
        idx.map(i => {
          const cod = t.codigos[i], k = categoria(cod), sel = ESC.datas[i] === filtro.dia ? ' class="sel"' : '';
          if(!k) return `<td${sel}><span style="color:#cbd5e1;">—</span></td>`;
          const c = CAT[k];
          const fundo = k === 'FO' ? 'rgba(148,163,184,.25)' : c.cor + '22';
          const cor = k === 'FO' ? '#64748b' : c.cor;
          return `<td${sel}><span class="cd" title="${esc(dm(ESC.datas[i]) + ' · ' + cod)}" style="background:${fundo}; color:${cor};">${esc(k === 'OUTRO' ? cod.slice(0,3) : c.curto)}</span></td>`;
        }).join('') + '</tr>';
    });
    document.getElementById('esMTbody').innerHTML = html ||
      `<tr><td colspan="${idx.length + 2}" style="text-align:center; color:var(--muted);">Nenhum técnico encontrado.</td></tr>`;
  }

  function exportarCsv(){
    if(!ESC) return;
    const idx = idxDatasDoMes(), tecs = linhasMatriz();
    const csvEsc = v => { const s = String(v ?? ''); return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s; };
    const out = [['Supervisao','Area','Turno','Horario','Tecnico'].concat(idx.map(i => dm(ESC.datas[i]))).join(';')];
    tecs.forEach(t => out.push([t.supervisao, t.area, t.turno, t.horario, t.nome].concat(idx.map(i => t.codigos[i] || '')).map(csvEsc).join(';')));
    const blob = new Blob(['\uFEFF' + out.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'Escala_' + filtro.mes + '.csv'; a.style.display = 'none';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  // ---------------------------------------------------------------- início
  function iniciar(){
    montarEstrutura();
    carregarEscala();
    setInterval(carregarEscala, 5 * 60 * 1000);
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
