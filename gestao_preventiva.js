/* =========================================================================
   Painel do ESER — Seção "Preventiva" (Atualização_Preventiva)
   Fonte: dados_preventiva.json (gerado da aba Atualização_Preventiva)
   Requer: Chart.js (datalabels é opcional — usado se já estiver carregado)
   Uso:    <div id="sec-preventiva"></div>
           <script src="gestao_preventiva.js"> (depois do Chart.js)
           (monta sozinho; ou chame PainelPreventiva.montar('id-do-container'))
   ========================================================================= */
(function () {
  const ARQUIVO_JSON = 'dados_preventiva.json';
  const CONTAINER_PADRAO = 'sec-preventiva';

  // Feriados que não contam como dia útil (edite conforme o calendário da operação)
  const FERIADOS = [
    '2026-01-01','2026-02-16','2026-02-17','2026-04-03','2026-04-21','2026-05-01',
    '2026-06-04','2026-09-07','2026-10-12','2026-11-02','2026-11-15','2026-11-20','2026-12-25',
    '2027-01-01','2027-02-08','2027-02-09','2027-03-26','2027-04-21','2027-05-01',
    '2027-05-27','2027-09-07','2027-10-12','2027-11-02','2027-11-15','2027-11-20','2027-12-25'
  ];

  const FILTROS = [
    { k: 'st', rotulo: 'Status' },
    { k: 'ex', rotulo: 'Executor' },
    { k: 't',  rotulo: 'Tipo Preventiva' },
    { k: 'x',  rotulo: 'Equipe Responsável' }
  ];

  // Executor definido pelo Tipo de Preventiva
  const EXECUTOR_POR_TIPO = {
    'energia': 'MOP', 'climatizacao': 'MOP',
    'zeladoria': 'EPS', 'sdai': 'EPS', 'gerador': 'EPS',
    'inspecao termografica': 'EPS', 'spda': 'EPS'
  };
  const semAcento = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const executorDoTipo = t => EXECUTOR_POR_TIPO[semAcento(t).replace(/^preventiva infra - /, '')] || 'Outros';

  // ---- HISTÓRICO MENSAL (meses fechados) — edite/acrescente aqui ----
  // O mês atual entra sozinho, calculado da base (dados_preventiva.json).
  const HISTORICO_MENSAL = [
    { mes: '2026-01', programadas: 920,  executadas: 482 },
    { mes: '2026-02', programadas: 961,  executadas: 896 },
    { mes: '2026-03', programadas: 978,  executadas: 848 },
    { mes: '2026-04', programadas: 1092, executadas: 810 },
    { mes: '2026-05', programadas: 1407, executadas: 1196 },
    { mes: '2026-06', programadas: 1446, executadas: 1284 },
    { mes: '2026-07', programadas: 1327, executadas: 1264 },
    { mes: '2026-08', programadas: 1267, executadas: 1216 },
    { mes: '2026-09', programadas: 1484, executadas: 1420 }
  ];

  const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho',
                 'Agosto','Setembro','Outubro','Novembro','Dezembro'];

  let dados = null, grafico = null, raiz = null;
  const sel = { st: '', ex: '', t: '', x: '' };
  let busca = '';

  const fmt = n => n.toLocaleString('pt-BR');
  const pct = n => (isFinite(n) ? n : 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
  const iso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // "Executada" só conta se a Data Fim - Consolidado (r.d) estiver preenchida
  // e dentro do mês corrente (dados.mes) — vazia ou de outro mês não conta.
  const executadaNoMes = r => r.st === 'Executada' && !!r.d && r.d.slice(0, 7) === dados.mes;

  function diasDoMes(mes) {
    const [a, m] = mes.split('-').map(Number);
    const out = [];
    for (let d = new Date(a, m - 1, 1); d.getMonth() === m - 1; d.setDate(d.getDate() + 1)) {
      const dia = new Date(d), s = iso(dia), sem = dia.getDay();
      out.push({ data: s, dia: dia.getDate(), util: sem !== 0 && sem !== 6 && !FERIADOS.includes(s) });
    }
    return out;
  }

  // D-1: ontem, se o mês é o atual; último dia, se o mês já passou
  function dataReferencia(mes) {
    const ontem = new Date(); ontem.setDate(ontem.getDate() - 1);
    const refMes = iso(ontem).slice(0, 7);
    if (refMes > mes) { const [a, m] = mes.split('-').map(Number); return iso(new Date(a, m, 0)); }
    if (refMes < mes) return mes + '-00';
    return iso(ontem);
  }

  function estilos() {
    if (document.getElementById('pv-estilos')) return;
    // Usa as mesmas classes do painel (.signal-strip, .kpi-row, .kpi, .panel, .pill);
    // aqui ficam só os ajustes específicos da Preventiva.
    const css = `
    .pv-filtros{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 16px}
    .pv-filtros label{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted,#6b7590);text-transform:uppercase;letter-spacing:.05em}
    .pv-filtros label:not(:first-child){margin-left:8px}
    .pv-filtros select{font-family:'JetBrains Mono',monospace;font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid var(--line,#e1e5f0);background:var(--panel,#fff);color:var(--text,#1b2440);min-width:150px}
    .pv-sub{font-family:'Inter',sans-serif;font-size:10px;font-weight:600;color:var(--muted,#6b7590);text-transform:uppercase;margin-left:4px}
    .pv-sep{color:var(--line,#e1e5f0)}
    #pv-kpis{grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}
    .pv-cv{position:relative;height:360px}
    .pv-rolagem{max-height:420px;overflow:auto;border-radius:8px}
    .pv-tab-topo{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}
    .pv-tab-topo .search input{min-width:240px}
    .pv-layout{display:flex;gap:20px;align-items:flex-start}
    .pv-menu{width:200px;flex:0 0 200px;background:var(--panel,#fff);border:1px solid var(--line,#e1e5f0);border-radius:14px;
      padding:10px;position:sticky;top:20px;box-shadow:0 2px 8px rgba(27,36,64,.06)}
    .pv-menu .nav-group{margin-bottom:0}
    .pv-menu .nav-group-label{cursor:default}
    .pv-menu .nav-group-label:hover{color:var(--muted,#6b7590)}
    .pv-conteudo{flex:1;min-width:0}
    @media (max-width:860px){.pv-layout{flex-direction:column}.pv-menu{width:100%;position:static}}
    .pv-legenda{display:flex;flex-direction:column;gap:9px;font-size:12.5px;color:var(--text,#1b2440)}
    .pv-legenda i{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:6px;vertical-align:-1px}
    .pv-legenda b.n{font-family:'JetBrains Mono',monospace;margin-left:4px}
    .pv-op-tab td{white-space:nowrap}
    .pv-sf{display:inline-block;min-width:92px;text-align:center;padding:3px 8px;border-radius:4px;font-weight:600;font-size:12px}
    .pv-sf.verde{background:#16a34a;color:#fff}.pv-sf.amarelo{background:#facc15;color:#1b2440}.pv-sf.vermelho{background:#dc2626;color:#fff}
    .pv-tec-tab tbody tr:nth-child(even){background:rgba(107,117,144,.05)}
    .pv-tec-tab tbody tr:hover{background:rgba(14,124,134,.08)}
    .pv-tec-tab th,.pv-tec-tab td{padding:10px 12px}
    .pv-tec-tab td:first-child{font-weight:600}`;
    const st = document.createElement('style'); st.id = 'pv-estilos'; st.textContent = css;
    document.head.appendChild(st);
  }

  function esqueleto() {
    opMontada = false;
    const [a, m] = dados.mes.split('-').map(Number);
    raiz.innerHTML = `
    <div class="pv">
      <div class="pv-layout">
      <nav class="pv-menu">
        <div class="nav-group">
          <div class="nav-group-label"><span>Preventiva</span></div>
          <div class="nav-group-items">
            <div class="nav-item pv-aba active" data-aba="geral"><span class="nav-icon">📊</span> Visão Geral</div>
            <div class="nav-item pv-aba" data-aba="op"><span class="nav-icon">🛠️</span> Operação</div>
            <div class="nav-item pv-aba" data-aba="tec"><span class="nav-icon">👷</span> Produtividade</div>
          </div>
        </div>
      </nav>
      <div class="pv-conteudo">
      <div id="pv-aba-geral">
      <div class="pv-filtros">
        ${FILTROS.map(f => `<label>${f.rotulo}</label><select data-k="${f.k}"></select>`).join('')}
        <span class="mono" style="font-size:11px;color:var(--muted,#6b7590);margin-left:auto;">
          ${MESES[m - 1]}/${a} · atualizado em ${esc(dados.gerado_em || '')}</span>
      </div>

      <div class="signal-strip" id="pv-signal" title="Aderência em D-1"></div>

      <div class="kpi-row" id="pv-kpis"></div>

      <div class="panel">
        <h2>Realizado x Meta — Acumulado</h2>
        <div class="pv-cv"><canvas id="pv-canvas"></canvas></div>
      </div>
      <div class="panel">
        <h2>Visão mensal</h2>
        <div class="pv-cv"><canvas id="pv-mensal"></canvas></div>
      </div>

      <div class="panel">
        <div class="pv-tab-topo">
          <h2>Preventivas não executadas (<span id="pv-qtd-tab">0</span>)</h2>
          <div class="search">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2"/></svg>
            <input id="pv-busca" type="search" placeholder="Buscar site, WO, município...">
          </div>
        </div>
        <div class="pv-rolagem"><table>
          <thead><tr><th>Status</th><th>Site</th><th>Site Infratel</th><th>Município</th><th>Tipo</th>
          <th>Tipologia</th><th>Executor</th><th>Equipe</th><th>Supervisão</th><th>WO</th></tr></thead>
          <tbody id="pv-tbody"></tbody></table></div>
      </div>
      </div>
      <div id="pv-aba-op" style="display:none;"></div>
      <div id="pv-aba-tec" style="display:none;"></div>
      </div>
      </div>
    </div>`;

    raiz.querySelectorAll('.pv-aba').forEach(bt => bt.addEventListener('click', () => trocarAba(bt.dataset.aba)));

    raiz.querySelectorAll('.pv-filtros select').forEach(s =>
      s.addEventListener('change', () => { sel[s.dataset.k] = s.value; atualizar(); }));
    raiz.querySelector('#pv-busca').addEventListener('input', e => { busca = e.target.value.toLowerCase(); tabela(filtrar()); });
  }

  function filtrar(ignorar) {
    return dados.registros.filter(r => FILTROS.every(f => f.k === ignorar || !sel[f.k] || r[f.k] === sel[f.k]));
  }

  // Opções de cada filtro respeitam os demais filtros (como slicers do Power BI)
  function opcoes() {
    raiz.querySelectorAll('.pv-filtros select').forEach(s => {
      const k = s.dataset.k;
      const vals = [...new Set(filtrar(k).map(r => r[k]))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
      if (sel[k] && !vals.includes(sel[k])) vals.push(sel[k]);
      s.innerHTML = '<option value="">Todos</option>' +
        vals.map(v => `<option value="${esc(v)}"${v === sel[k] ? ' selected' : ''}>${esc(v)}</option>`).join('');
    });
  }

  const COR = { good:'#059669', warn:'#d97706', bad:'#dc2626', accent:'#0e7c86', muted:'#6b7590' };
  // Meta 100% (verde) · a partir de 95% é o 2º patamar (laranja) · abaixo de 95% vermelho
  const META_ADER = 100, PATAMAR2_ADER = 95;
  const corAder = p => p >= META_ADER ? COR.good : p >= PATAMAR2_ADER ? COR.warn : COR.bad;

  function atualizar() {
    opcoes();
    const regs = filtrar();
    const dias = diasDoMes(dados.mes);
    const ref = dataReferencia(dados.mes);
    const duMes = dias.filter(d => d.util).length || 1;
    const duRef = dias.filter(d => d.util && d.data <= ref).length;

    const planejado = regs.length;
    const planD1 = Math.round(planejado * duRef / duMes);
    const realizado = regs.filter(executadaNoMes).length;
    const pendentes = planejado - realizado;
    const aberto = regs.filter(r => r.st === 'Aberta').length;
    // Aderência comparada com a meta proporcional ATÉ HOJE (não até amanhã):
    // conta o que já foi executado hoje, sem exigir ainda o que só vence nos próximos dias.
    const hoje = iso(new Date());
    const duHoje = dias.filter(d => d.util && d.data <= hoje).length;
    const planHoje = Math.round(planejado * duHoje / duMes);
    const saldo = realizado - planHoje;
    const ader = planHoje ? realizado / planHoje * 100 : 0;
    const entrega = planejado ? realizado / planejado * 100 : 0;

    // faixa de sinal (30 blocos) = aderência em D-1
    const strip = raiz.querySelector('#pv-signal');
    const on = Math.round(Math.min(1, ader / META_ADER) * 30);
    strip.innerHTML = Array.from({ length: 30 }, (_, i) =>
      `<div style="background:${i < on ? 'rgba(5,150,105,.7)' : 'rgba(220,38,38,.7)'}"></div>`).join('');

    raiz.querySelector('#pv-kpis').innerHTML = `
      <div class="kpi"><div class="label">📋 Planejado no Mês</div>
        <div class="value">${fmt(planejado)}</div>
        <div class="delta" style="color:${COR.muted}">${duMes} dias úteis no mês</div></div>
      <div class="kpi"><div class="label">✅ Realizado / Meta D-1</div>
        <div class="value" style="display:flex;gap:10px;align-items:baseline;white-space:nowrap;">
          <span style="color:${COR.good}">${fmt(realizado)}</span><span class="pv-sep">|</span><span>${fmt(planD1)}</span></div>
        <div class="delta" style="color:${COR.muted}">${pct(entrega)} de entrega no mês</div></div>
      <div class="kpi"><div class="label">📅 Dias Úteis (D-1)</div>
        <div class="value" style="white-space:nowrap;">${duRef}<span class="pv-sub">decorridos</span>
          <span style="color:${COR.muted}"> / </span>${duMes - duRef}<span class="pv-sub">${duMes - duRef === 1 ? 'restante' : 'restantes'}</span></div>
        <div class="delta" style="color:${COR.muted}">${pct(duRef / duMes * 100)} dos dias úteis do mês</div></div>
      <div class="kpi"><div class="label">📈 Aderência</div>
        <div class="value"><span style="color:${corAder(ader)}">${pct(ader)}</span></div>
        <div class="delta" style="color:${corAder(ader)}">${ader >= META_ADER ? '▲ meta batida' : ader >= PATAMAR2_ADER ? '● 2º patamar (≥95%)' : '▼ abaixo de 95%'} · saldo ${(saldo > 0 ? '+' : '') + fmt(saldo)}</div>
        <div class="delta" style="color:${COR.muted}">Meta 100% · 2º patamar 95%</div></div>
      <div class="kpi"><div class="label">⚠️ Pendentes</div>
        <div class="value" style="display:flex;gap:10px;align-items:baseline;white-space:nowrap;">
          <span style="color:${COR.bad}">${fmt(pendentes)}</span><span class="pv-sep">|</span>
          <span style="font-size:15px;color:${COR.muted};font-weight:600;">Em aberto <span style="color:#1b2440">${fmt(aberto)}</span></span></div></div>`;

    grafico_(regs, dias, ref, duMes, planejado);
    graficoMensal_();
    tabela(regs);
  }

  function grafico_(regs, dias, ref, duMes, planejado) {
    const porDia = {};
    // Só conta quem tem Data Fim - Consolidado preenchida e dentro do mês
    // corrente — vazia ou de outro mês não entra na contagem.
    regs.forEach(r => { if (executadaNoMes(r)) porDia[r.d] = (porDia[r.d] || 0) + 1; });
    let du = 0, acum = 0;
    const rot = [], plan = [], real = [];
    dias.forEach(d => {
      if (d.util) du++;
      acum += porDia[d.data] || 0;
      rot.push(String(d.dia).padStart(2, '0') + '/' + dados.mes.slice(5));
      plan.push(Math.round(planejado * du / duMes));
      real.push(d.data <= ref || porDia[d.data] ? acum : null);
    });

    const temLabels = typeof ChartDataLabels !== 'undefined';
    const cfg = {
      type: 'line',
      data: { labels: rot, datasets: [
        { label: 'Planejado Acumulado', data: plan, borderColor: '#6b7590', backgroundColor: '#6b7590',
          borderDash: [3, 4], borderWidth: 2, pointStyle: 'rect', pointRadius: 4, tension: 0,
          datalabels: { align: 'top', color: '#6b7590' } },
        { label: 'Realizado Acumulado', data: real, borderColor: '#0e7c86', backgroundColor: '#0e7c86',
          borderWidth: 3, pointRadius: 4, tension: 0.25, spanGaps: false,
          datalabels: { align: 'bottom', color: '#000000' } }
      ]},
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 18, bottom: 6, right: 26 } },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top', align: 'start', labels: { usePointStyle: true, boxWidth: 8, boxHeight: 8 } },
          tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmt(c.parsed.y)}` } },
          datalabels: temLabels ? {
            display: c => { const i = c.dataIndex, v = c.dataset.data;
              if (v[i] == null) return false;
              let ult = v.length - 1; while (ult > 0 && v[ult] == null) ult--;
              if (i === ult) return true;                       // sempre mostra o último valor
              if (i >= ult - 2) return false;                   // evita encavalar com o último
              return i % 3 === 0 && (i === 0 || v[i] !== v[i - 1]); },
            font: { size: 10, family: 'JetBrains Mono', weight: '700' }, formatter: v => fmt(v)
          } : undefined
        },
        scales: { y: { beginAtZero: true, grid: { display: false }, ticks: { callback: v => fmt(v) } }, x: { grid: { display: false } } }
      },
      plugins: temLabels ? [ChartDataLabels] : []
    };
    if (grafico) grafico.destroy();
    grafico = new Chart(raiz.querySelector('#pv-canvas'), cfg);
  }

  let graficoMensal = null;
  function graficoMensal_() {
    const ABREV = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    const serie = HISTORICO_MENSAL.filter(h => h.mes !== dados.mes).map(h => ({ ...h }));
    serie.push({ mes: dados.mes, programadas: dados.registros.length,
      executadas: dados.registros.filter(executadaNoMes).length, atual: true });
    serie.sort((a, b) => a.mes.localeCompare(b.mes));
    const rot = serie.map(h => ABREV[+h.mes.slice(5) - 1] + '/' + h.mes.slice(2, 4) + (h.atual ? '*' : ''));
    const pcts = serie.map(h => h.programadas ? h.executadas / h.programadas * 100 : 0);
    const corPct = p => p >= 100 ? '#059669' : p >= 95 ? '#d97706' : '#dc2626';
    const temLabels = typeof ChartDataLabels !== 'undefined';

    if (graficoMensal) graficoMensal.destroy();
    graficoMensal = new Chart(raiz.querySelector('#pv-mensal'), {
      data: {
        labels: rot,
        datasets: [
          { type: 'bar', label: 'Programadas', data: serie.map(h => h.programadas),
            backgroundColor: 'rgba(107,117,144,.6)', borderRadius: 5, order: 2,
            datalabels: { color: '#000', anchor: 'end', align: 'top', font: { size: 9, family: 'JetBrains Mono', weight: '700' } } },
          { type: 'bar', label: 'Executadas', data: serie.map(h => h.executadas),
            backgroundColor: 'rgba(14,124,134,.8)', borderRadius: 5, order: 2,
            datalabels: { color: '#000', anchor: 'end', align: 'top', font: { size: 9, family: 'JetBrains Mono', weight: '700' } } },
          { type: 'line', label: '% Execução', data: pcts, yAxisID: 'y1', order: 1,
            borderColor: '#d97706', backgroundColor: '#d97706', tension: .3, pointRadius: 4,
            pointBackgroundColor: pcts.map(corPct), pointBorderColor: pcts.map(corPct),
            datalabels: { align: 'top', offset: 6, color: c => corPct(pcts[c.dataIndex]),
              font: { size: 10, family: 'JetBrains Mono', weight: '700' }, formatter: v => Math.round(v) + '%',
              backgroundColor: 'rgba(255,255,255,.85)', borderRadius: 4, padding: { top: 1, bottom: 1, left: 4, right: 4 } } },
          { type: 'line', label: 'Meta 100%', data: serie.map(() => 100), yAxisID: 'y1', order: 1,
            borderColor: '#000', borderDash: [5, 5], pointRadius: 0, borderWidth: 1, datalabels: { display: false } }
        ]
      },
      plugins: temLabels ? [ChartDataLabels] : [],
      options: {
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { top: 10 } },
        plugins: {
          legend: { labels: { color: '#6b7590', font: { family: 'Inter', size: 11 }, boxWidth: 12 } },
          tooltip: { callbacks: {
            label: c => c.dataset.yAxisID === 'y1' ? `${c.dataset.label}: ${c.parsed.y.toFixed(1).replace('.', ',')}%` : `${c.dataset.label}: ${fmt(c.parsed.y)}`,
            footer: it => serie[it[0].dataIndex].atual ? 'Mês atual (parcial)' : '' } }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: '#1b2440', font: { weight: 'bold' } } },
          y: { beginAtZero: true, max: Math.ceil(Math.max(...serie.map(h => h.programadas)) * 1.75 / 200) * 200, grid: { display: false }, ticks: { color: '#1b2440' } },
          y1: { position: 'right', min: 0, max: 110, grid: { display: false },
            ticks: { color: '#1b2440', callback: v => v <= 100 ? v + '%' : '' } }
        }
      }
    });
  }

  function tabela(regs) {
    const lista = regs.filter(r => r.st !== 'Executada').filter(r => !busca ||
      [r.s, r.si, r.w, r.m, r.t, r.e, r.x, r.ex, r.tp].some(v => String(v ?? '').toLowerCase().includes(busca)));
    raiz.querySelector('#pv-qtd-tab').textContent = fmt(lista.length);
    raiz.querySelector('#pv-tbody').innerHTML = lista.map(r => `<tr>
      <td><span class="pill ${r.st === 'Aberta' ? 'warn' : 'bad'}">${esc(r.st)}</span></td><td>${esc(r.s)}</td><td>${esc(r.si)}</td>
      <td>${esc(r.m)}</td><td>${esc(r.t)}</td><td>${esc(r.tp)}</td><td>${esc(r.ex)}</td><td>${esc(r.x)}</td><td>${esc(r.e)}</td><td>${esc(r.w)}</td>
    </tr>`).join('') || '<tr><td colspan="10" style="text-align:center;color:#6b7590">Nenhuma preventiva pendente</td></tr>';
  }

  // ======================= ABA OPERAÇÃO =======================
  const FILTROS_OP = [
    { k: 'xpF', rotulo: 'Expurgo' },
    { k: 'st',  rotulo: 'Status' },
    { k: 'ex',  rotulo: 'Executor' },
    { k: 't',   rotulo: 'Tipo Preventiva' },
    { k: 'io',  rotulo: 'Infratel OK' },
    { k: 'x',   rotulo: 'Equipe Responsável' }
  ];
  const selOp = { xpF: '', st: '', ex: '', t: '', io: '', x: '' };
  let buscaOp = '', opMontada = false;
  const FILTROS_TEC = [
    { k: 't', rotulo: 'Tipo de Preventiva' },
    { k: 'x', rotulo: 'Equipe' }
  ];
  const selTec = { tipo: '', x: '' };
  let buscaTec = '', tecMontada = false;

  const temRelatorio = r => String(r.er || '').trim() !== '';
  const fmtPct = v => v === null || v === undefined ? '' : (v * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';

  // Cor do "Status Final" (legenda do Power BI):
  //  verde    = Executada: Infratel 100% + relatório recebido (MOP não exige relatório)
  //  amarelo  = Divergência: relatório recebido, porém Infratel 0%
  //  vermelho = Verificar: Infratel 100%, porém sem relatório entregue (EPS)
  function corStatusFinal(r) {
    const rel = temRelatorio(r), infra100 = r.pc === 1, infra0 = r.pc === 0;
    if (r.ex === 'MOP') return infra100 ? 'verde' : '';
    if (infra100 && rel) return 'verde';
    if (rel && infra0) return 'amarelo';
    if (infra100 && !rel) return 'vermelho';
    return '';
  }

  function filtrarOp(ignorar) {
    return dados.registros.filter(r => FILTROS_OP.every(f => f.k === ignorar || !selOp[f.k] || r[f.k] === selOp[f.k]));
  }

  function montarOp() {
    const alvo = raiz.querySelector('#pv-aba-op');
    alvo.innerHTML = `
      <div class="pv-filtros pv-op-filtros">
        ${FILTROS_OP.map(f => `<label>${f.rotulo}</label><select data-k="${f.k}"></select>`).join('')}
      </div>
      <div class="kpi-row" id="pv-op-ritmo" style="grid-template-columns:repeat(auto-fit,minmax(190px,1fr));margin-bottom:14px;"></div>
      <div class="kpi-row" id="pv-op-kpis" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));"></div>
      <div class="grid" style="grid-template-columns:1.4fr 1fr;">
        <div class="panel">
          <h2>Realizado x Meta Diária</h2>
          <div style="position:relative;height:300px;"><canvas id="pv-op-diario"></canvas></div>
        </div>
        <div class="panel">
          <h2>Entrega por Equipe Responsável</h2>
          <div id="pv-op-equipes" style="display:flex;flex-direction:column;gap:22px;margin-top:18px;"></div>
        </div>
      </div>
      <div class="panel">
        <div class="pv-tab-topo">
          <h2>Preventivas (<span id="pv-op-qtd">0</span>)</h2>
          <div class="search">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2"/></svg>
            <input id="pv-op-busca" type="search" placeholder="Buscar WO, site, chamado...">
          </div>
        </div>
        <div class="pv-rolagem" style="max-height:560px;"><table class="pv-op-tab">
          <thead><tr><th>WO</th><th>Site</th><th>Tipo de Infra</th><th>Equipe Responsável - Engemon</th>
          <th class="num">% Inventário</th><th>Nº do chamado / Acesso</th><th>Infratel OK</th><th>Tipo de Preventiva</th>
          <th class="num">% Cronograma Infratel</th><th>Relatórios Entregues</th><th>Expurgo</th></tr></thead>
          <tbody id="pv-op-tbody"></tbody></table></div>
      </div>`;
    alvo.querySelectorAll('.pv-op-filtros select').forEach(s =>
      s.addEventListener('change', () => { selOp[s.dataset.k] = s.value; atualizarOp(); }));
    alvo.querySelector('#pv-op-busca').addEventListener('input', e => { buscaOp = e.target.value.toLowerCase(); tabelaOp(filtrarOp()); });
    opMontada = true;
  }

  function opcoesOp() {
    raiz.querySelectorAll('.pv-op-filtros select').forEach(s => {
      const k = s.dataset.k;
      const vals = [...new Set(filtrarOp(k).map(r => r[k]))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
      if (selOp[k] && !vals.includes(selOp[k])) vals.push(selOp[k]);
      s.innerHTML = '<option value="">Todos</option>' +
        vals.map(v => `<option value="${esc(v)}"${v === selOp[k] ? ' selected' : ''}>${esc(v)}</option>`).join('');
    });
  }

  function atualizarOp() {
    opcoesOp();
    const regs = filtrarOp();
    const planejado = regs.length;
    const realizado = regs.filter(executadaNoMes).length;
    const entrega = planejado ? realizado / planejado * 100 : 0;
    const totalRel = regs.filter(r => r.ex === 'EPS').length;
    const relEntregues = regs.filter(temRelatorio).length;
    const expurgo = regs.filter(r => r.xp).length;
    const traco = v => v ? v : '--';

    raiz.querySelector('#pv-op-kpis').innerHTML = `
      <div class="kpi"><div class="label">📄 Total Relatórios EPS</div>
        <div class="value">${fmt(totalRel)}</div>
        <div class="delta" style="color:${COR.muted}">Zeladoria, SDAI, Gerador, Termografia e SPDA</div></div>
      <div class="kpi"><div class="label">🚫 Qtd Expurgo</div>
        <div class="value"><span style="color:${COR.warn}">${fmt(expurgo)}</span></div>
        <div class="delta" style="color:${COR.muted}">% Expurgo <b style="color:#1b2440">${pct(planejado ? expurgo / planejado * 100 : 0)}</b></div></div>`;

    ritmoOp(regs);
    tabelaOp(regs);
  }

  let graficoDiario = null;
  function ritmoOp(regs) {
    const dias = diasDoMes(dados.mes);
    const ref = dataReferencia(dados.mes);
    const duMes = dias.filter(d => d.util).length || 1;
    const duRef = dias.filter(d => d.util && d.data <= ref).length;
    const restantes = duMes - duRef;

    const total = regs.length;
    const exec = regs.filter(executadaNoMes);
    const realizadas = exec.length;
    const realD1 = exec.filter(r => r.d && r.d <= ref).length;
    const media = duRef ? realD1 / duRef : 0;
    const necessario = restantes > 0 ? Math.max(0, total - realD1) / restantes : 0;
    const capacidade = Math.round(realD1 + media * restantes);
    const projecao = total ? capacidade / total * 100 : 0;
    const dec = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    raiz.querySelector('#pv-op-ritmo').innerHTML = `
      <div class="kpi"><div class="label">✅ Realizadas | Total WO</div>
        <div class="value" style="display:flex;gap:10px;align-items:baseline;white-space:nowrap;">
          <span style="color:${COR.good}">${fmt(realizadas)}</span><span class="pv-sep">|</span><span>${fmt(total)}</span></div>
        <div class="delta" style="color:${COR.muted}">% Entrega <b style="color:#1b2440">${pct(total ? realizadas / total * 100 : 0)}</b></div></div>
      <div class="kpi"><div class="label">📊 Média Realizada Dia</div>
        <div class="value">${dec(media)}</div>
        <div class="delta" style="color:${COR.muted}">${fmt(realD1)} em ${duRef} dias úteis</div></div>
      <div class="kpi"><div class="label">🎯 Necessário por Dia</div>
        <div class="value"><span style="color:${necessario > media ? COR.bad : COR.good}">${dec(necessario)}</span></div>
        <div class="delta" style="color:${COR.muted}">para fechar ${fmt(total)} no mês</div></div>
      <div class="kpi"><div class="label">📅 Dias Úteis Restantes</div>
        <div class="value">${restantes}</div>
        <div class="delta" style="color:${COR.muted}">${duRef} / ${duMes} dias úteis</div></div>
      <div class="kpi"><div class="label">🔮 Capacidade | % Projeção D-1</div>
        <div class="value" style="display:flex;gap:10px;align-items:baseline;white-space:nowrap;">
          <span>${fmt(capacidade)}</span><span class="pv-sep">|</span><span style="color:${corAder(projecao)}">${pct(projecao)}</span></div>
        <div class="delta" style="color:${COR.muted}">no ritmo atual até o fim do mês</div></div>`;

    // gráfico Realizado x Meta Diária
    // Meta do dia = meta original do dia + o que ficou faltando nos dias anteriores
    // (meta acumulada até hoje − realizado acumulado até ontem). Se sobrou, abate do dia seguinte.
    // Dias depois de D-1 usam a meta de recuperação (o que falta ÷ dias úteis restantes).
    const porDia = {};
    exec.forEach(r => { if (r.d) porDia[r.d] = (porDia[r.d] || 0) + 1; });
    const metaOriginal = total / duMes;
    const realDia = [], metaDia = [], deficitAnt = [], corBarra = [];
    let du = 0, realAcumAnt = 0;
    dias.forEach(d => {
      const feito = porDia[d.data] || 0;
      if (d.util) du++;
      let meta = null, deficit = null;
      if (d.util) {
        if (d.data <= ref) {
          const planAcum = metaOriginal * du;
          meta = Math.max(0, planAcum - realAcumAnt);
          deficit = meta - metaOriginal;
        } else {
          meta = necessario;
          deficit = necessario - metaOriginal;
        }
      }
      realDia.push(d.data <= ref || feito ? feito : null);
      metaDia.push(meta === null ? null : Math.round(meta * 100) / 100);
      deficitAnt.push(deficit);
      corBarra.push(meta === null ? 'rgba(107,117,144,.55)' : feito >= meta - 0.005 ? 'rgba(5,150,105,.8)' : 'rgba(220,38,38,.75)');
      realAcumAnt += feito;
    });

    const cv = raiz.querySelector('#pv-op-diario');
    if (graficoDiario) graficoDiario.destroy();
    graficoDiario = new Chart(cv, {
      data: {
        labels: dias.map(d => String(d.dia).padStart(2, '0') + '/' + dados.mes.slice(5)),
        datasets: [
          { type: 'bar', label: 'Realizado Dia', data: realDia, backgroundColor: corBarra, borderRadius: 3, order: 3,
            datalabels: { display: c => c.dataset.data[c.dataIndex] > 0, anchor: 'end', align: 'top', color: '#000',
              font: { size: 9, family: 'JetBrains Mono', weight: '700' } } },
          { type: 'line', label: 'Meta do Dia (com acúmulo)', data: metaDia, spanGaps: false,
            borderColor: COR.warn, backgroundColor: COR.warn, borderWidth: 2, stepped: 'middle',
            pointRadius: 3, pointHoverRadius: 5, order: 1,
            datalabels: { display: c => c.dataset.data[c.dataIndex] !== null, align: 'top', offset: 4, color: COR.warn,
              font: { size: 9, family: 'JetBrains Mono', weight: '700' }, formatter: v => Math.round(v) } },
          { type: 'line', label: 'Meta Original', data: dias.map(d => d.util ? Math.round(metaOriginal * 100) / 100 : null),
            borderColor: '#6b7590', borderWidth: 1.5, borderDash: [5, 4], pointRadius: 0, order: 2, spanGaps: true,
            datalabels: { display: false } }
        ]
      },
      plugins: typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [],
      options: {
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { top: 14 } },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { labels: { color: COR.muted, font: { family: 'Inter', size: 11 }, boxWidth: 12 } },
          tooltip: { callbacks: {
            label: c => c.parsed.y === null ? null : `${c.dataset.label}: ${dec(c.parsed.y)}`,
            afterBody: it => {
              const i = it[0].dataIndex, dAnt = deficitAnt[i];
              if (dAnt === null) return 'Dia não útil';
              return dAnt >= 0 ? `Acúmulo dos dias anteriores: +${dec(dAnt)}` : `Sobra dos dias anteriores: ${dec(dAnt)}`;
            } } }
        },
        scales: { x: { grid: { display: false }, ticks: { font: { size: 10 } } },
          y: { beginAtZero: true, grid: { color: '#e1e5f0' } } }
      }
    });

    // barras por Equipe Responsável
    const porEquipe = {};
    regs.forEach(r => {
      const k = r.x || 'Sem equipe';
      porEquipe[k] = porEquipe[k] || { feitas: 0, total: 0 };
      porEquipe[k].total++;
      if (executadaNoMes(r)) porEquipe[k].feitas++;
    });
    raiz.querySelector('#pv-op-equipes').innerHTML = Object.keys(porEquipe).sort().map(k => {
      const { feitas, total: tot } = porEquipe[k];
      const p = tot ? feitas / tot * 100 : 0;
      const cor = p >= 100 ? '#16a34a' : p >= 80 ? '#eab308' : '#dc2626';
      return `<div style="display:grid;grid-template-columns:110px 1fr auto;gap:12px;align-items:center;">
        <b style="font-size:13px;text-align:right;">${esc(k)}</b>
        <div><div style="background:#e5e7eb;border-radius:999px;height:14px;overflow:hidden;">
          <div style="width:${Math.min(100, p)}%;background:${cor};height:100%;border-radius:999px;"></div></div>
          <div class="mono" style="font-size:11px;color:${COR.muted};margin-top:4px;">${p.toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%</div></div>
        <b class="mono" style="font-size:13px;">${fmt(feitas)} / ${fmt(tot)}</b></div>`;
    }).join('') || `<div style="color:${COR.muted}">Sem dados</div>`;
  }

  // ======================= ABA PRODUTIVIDADE (POR TÉCNICO) =======================
  // Depende do campo r.tec em cada registro — nome do técnico, obtido via JOIN
  // entre a aba Atualização_Preventiva (WO) e a aba Validação (coluna Executor),
  // feito na geração do dados_preventiva.json. Sem esse campo, tudo cai em
  // "Sem técnico identificado".
  function montarTec() {
    const alvo = raiz.querySelector('#pv-aba-tec');
    alvo.innerHTML = `
      <div class="pv-filtros">
        ${FILTROS_TEC.map(f => `<label>${f.rotulo}</label><select data-k="${f.k}"></select>`).join('')}
      </div>
      <div class="kpi-row" id="pv-tec-kpis" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr));margin-bottom:14px;"></div>
      <div class="panel">
        <div class="pv-tab-topo">
          <h2>Produtividade por Técnico (<span id="pv-tec-qtd">0</span>)</h2>
          <div class="search">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2"/></svg>
            <input id="pv-tec-busca" type="search" placeholder="Buscar técnico...">
          </div>
        </div>
        <div class="pv-rolagem" style="max-height:560px;">
          <table class="pv-tec-tab">
            <thead><tr><th>#</th><th>Técnico</th><th>Equipe</th><th class="num">Qtd WO</th><th class="num">% do Total</th><th class="num">Dias Trabalhados</th><th class="num">Média WO/Dia</th></tr></thead>
            <tbody id="pv-tec-tbody"></tbody>
          </table>
        </div>
      </div>`;
    alvo.querySelectorAll('.pv-filtros select').forEach(s =>
      s.addEventListener('change', () => { selTec[s.dataset.k] = s.value; renderTec(); }));
    alvo.querySelector('#pv-tec-busca').addEventListener('input', e => { buscaTec = e.target.value.toLowerCase(); renderTec(); });
    tecMontada = true;
  }

  function atualizarTec() { renderTec(); }

  // Base: dados.registros (Atualização_Preventiva), uma linha por preventiva
  // de verdade — só preventivas Executadas e com técnico identificado (via
  // Validação) entram na produtividade.
  function filtrarTec(ignorar) {
    return dados.registros.filter(r =>
      executadaNoMes(r) && r.tec &&
      FILTROS_TEC.every(f => f.k === ignorar || !selTec[f.k] || r[f.k] === selTec[f.k]));
  }

  function opcoesTec() {
    raiz.querySelectorAll('.pv-filtros select').forEach(s => {
      const k = s.dataset.k;
      const vals = [...new Set(filtrarTec(k).map(r => r[k]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
      if (selTec[k] && !vals.includes(selTec[k])) vals.push(selTec[k]);
      s.innerHTML = '<option value="">Todos</option>' +
        vals.map(v => `<option value="${esc(v)}"${v === selTec[k] ? ' selected' : ''}>${esc(v)}</option>`).join('');
    });
  }

  function renderTec() {
    opcoesTec();
    const base = filtrarTec();
    const porTec = {};
    base.forEach(r => {
      const nome = r.tec;
      if (!porTec[nome]) porTec[nome] = { qtd: 0, dias: new Set(), x: r.x || '' };
      porTec[nome].qtd++;
      const dataRef = r.ini || r.d;
      if (dataRef) porTec[nome].dias.add(dataRef);
      if (!porTec[nome].x && r.x) porTec[nome].x = r.x;
    });

    const totalWo = base.length; // já é contagem de preventivas reais (1 linha = 1 WO)
    let linhasTodas = Object.entries(porTec).map(([nome, info]) => {
      const dias = info.dias.size;
      return { nome, qtd: info.qtd, dias, media: dias ? info.qtd / dias : 0, x: info.x,
        pctTotal: totalWo ? info.qtd / totalWo * 100 : 0 };
    }).sort((a, b) => b.qtd - a.qtd);

    const mediaGeral = linhasTodas.length ? linhasTodas.reduce((s, l) => s + l.media, 0) / linhasTodas.length : 0;

    raiz.querySelector('#pv-tec-kpis').innerHTML = `
      <div class="kpi"><div class="label">👷 Técnicos</div><div class="value">${fmt(linhasTodas.length)}</div></div>
      <div class="kpi"><div class="label">🧾 Total WO</div><div class="value">${fmt(totalWo)}</div></div>
      <div class="kpi"><div class="label">📊 Média Geral WO/Dia</div><div class="value">${mediaGeral.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div></div>`;

    // Tabela (aplica busca)
    let linhas = linhasTodas;
    if (buscaTec) linhas = linhas.filter(l => l.nome.toLowerCase().includes(buscaTec));

    raiz.querySelector('#pv-tec-qtd').textContent = fmt(linhas.length);
    raiz.querySelector('#pv-tec-tbody').innerHTML = linhas.map((l, i) => {
      const corMedia = l.media >= mediaGeral ? COR.good : COR.bad;
      return `<tr>
        <td class="mono" style="color:${COR.muted}">${i + 1}</td>
        <td>${esc(l.nome)}</td>
        <td>${l.x ? `<span class="pill">${esc(l.x)}</span>` : '--'}</td>
        <td class="num mono">${fmt(l.qtd)}</td>
        <td class="num mono">${pct(l.pctTotal)}</td>
        <td class="num mono">${fmt(l.dias)}</td>
        <td class="num mono" style="font-weight:700;color:${corMedia}">${l.media.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="7" style="text-align:center;color:#6b7590">Nenhum técnico encontrado</td></tr>';
  }

  function tabelaOp(regs) {
    const lista = regs.filter(r => !buscaOp ||
      [r.w, r.si, r.s, r.ch, r.t, r.ti, r.x, r.m].some(v => String(v ?? '').toLowerCase().includes(buscaOp)))
      .sort((a, b) => String(a.w).localeCompare(String(b.w)));
    raiz.querySelector('#pv-op-qtd').textContent = fmt(lista.length);
    raiz.querySelector('#pv-op-tbody').innerHTML = lista.map(r => {
      return `<tr>
        <td class="mono" style="font-size:12px;">${esc(r.w)}</td><td>${esc(r.si)}</td><td>${esc(r.ti) || '-'}</td><td>${esc(r.x)}</td>
        <td class="num">${r.pi === null || r.pi === undefined ? '' : fmt(r.pi)}</td><td class="mono" style="font-size:12px;">${esc(r.ch)}</td>
        <td>${r.io === 'OK' ? '<span class="pill good">OK</span>' : `<span class="pill bad">${esc(r.io) || '—'}</span>`}</td>
        <td>Preventiva infra - ${esc(r.t)}</td>
        <td class="num">${fmtPct(r.pc)}</td><td>${esc(r.er)}</td><td>${esc(r.xp)}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="11" style="text-align:center;color:#6b7590">Nenhuma preventiva encontrada</td></tr>';
  }

  function trocarAba(aba) {
    raiz.querySelectorAll('.pv-aba').forEach(b => b.classList.toggle('active', b.dataset.aba === aba));
    raiz.querySelector('#pv-aba-geral').style.display = aba === 'geral' ? '' : 'none';
    raiz.querySelector('#pv-aba-op').style.display = aba === 'op' ? '' : 'none';
    raiz.querySelector('#pv-aba-tec').style.display = aba === 'tec' ? '' : 'none';
    if (aba === 'op') { if (!opMontada) montarOp(); atualizarOp(); if (graficoDiario) graficoDiario.resize(); }
    else if (aba === 'tec') { if (!tecMontada) montarTec(); atualizarTec(); }
    else { if (grafico) grafico.resize(); if (graficoMensal) graficoMensal.resize(); }
  }

  async function montar(idContainer) {
    raiz = document.getElementById(idContainer || CONTAINER_PADRAO);
    if (!raiz) return;
    estilos();
    try {
      dados = window.DADOS_PREVENTIVA ||
        await fetch(ARQUIVO_JSON + '?v=' + Date.now(), { cache: 'no-store' }).then(r => { if (!r.ok) throw r.status; return r.json(); });
      dados.registros.forEach(r => { r.ex = executorDoTipo(r.t); r.xpF = r.xp || 'Sem expurgo'; });
      esqueleto();
      atualizar();
    } catch (e) {
      raiz.innerHTML = `<div style="padding:20px;color:#b00020">Não foi possível carregar ${ARQUIVO_JSON} (${esc(e)}).</div>`;
    }
  }

  window.PainelPreventiva = { montar, redesenhar: () => { if (grafico) grafico.resize(); if (graficoMensal) graficoMensal.resize(); } };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => montar());
  else montar();
})();
