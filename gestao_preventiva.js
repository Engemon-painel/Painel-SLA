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

  const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho',
                 'Agosto','Setembro','Outubro','Novembro','Dezembro'];

  let dados = null, grafico = null, raiz = null;
  const sel = { st: '', ex: '', t: '', x: '' };
  let busca = '';

  const fmt = n => n.toLocaleString('pt-BR');
  const pct = n => (isFinite(n) ? n : 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
  const iso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

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
    .pv-abas{display:flex;gap:6px;margin:0 0 16px;flex-wrap:wrap}
    .pv-aba{font-family:'JetBrains Mono',monospace;font-size:12px;padding:8px 16px;border-radius:999px;border:1px solid var(--line,#e1e5f0);
      color:var(--muted,#6b7590);cursor:pointer;background:var(--panel,#fff);transition:.15s}
    .pv-aba:hover{border-color:var(--accent,#0e7c86);color:var(--text,#1b2440)}
    .pv-aba.ativa{background:var(--accent,#0e7c86);color:#fff;border-color:var(--accent,#0e7c86);font-weight:600}
    .pv-legenda{display:flex;flex-direction:column;gap:9px;font-size:12.5px;color:var(--text,#1b2440)}
    .pv-legenda i{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:6px;vertical-align:-1px}
    .pv-legenda b.n{font-family:'JetBrains Mono',monospace;margin-left:4px}
    .pv-op-tab td{white-space:nowrap}
    .pv-sf{display:inline-block;min-width:92px;text-align:center;padding:3px 8px;border-radius:4px;font-weight:600;font-size:12px}
    .pv-sf.verde{background:#16a34a;color:#fff}.pv-sf.amarelo{background:#facc15;color:#1b2440}.pv-sf.vermelho{background:#dc2626;color:#fff}`;
    const st = document.createElement('style'); st.id = 'pv-estilos'; st.textContent = css;
    document.head.appendChild(st);
  }

  function esqueleto() {
    opMontada = false;
    const [a, m] = dados.mes.split('-').map(Number);
    raiz.innerHTML = `
    <div class="pv">
      <div class="pv-abas">
        <button type="button" class="pv-aba ativa" data-aba="geral">📊 Visão Geral</button>
        <button type="button" class="pv-aba" data-aba="op">🛠️ Operação</button>
      </div>
      <div id="pv-aba-geral">
      <div class="pv-filtros">
        ${FILTROS.map(f => `<label>${f.rotulo}</label><select data-k="${f.k}"></select>`).join('')}
        <span class="mono" style="font-size:11px;color:var(--muted,#6b7590);margin-left:auto;">
          ${MESES[m - 1]}/${a} · atualizado em ${esc(dados.gerado_em || '')}</span>
      </div>

      <div class="signal-strip" id="pv-signal" title="Aderência em D-1"></div>

      <div class="kpi-row" id="pv-kpis"></div>

      <div class="panel" style="margin-bottom:16px;">
        <h2>Realizado x Meta — Acumulado</h2>
        <div class="hint">Meta distribuída pelos dias úteis do mês · realizado pela data de conclusão</div>
        <div class="pv-cv"><canvas id="pv-canvas"></canvas></div>
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
    const realizado = regs.filter(r => r.st === 'Executada').length;   // igual ao Power BI
    const pendentes = planejado - realizado;
    const aberto = regs.filter(r => r.st === 'Aberta').length;
    const saldo = realizado - planD1;
    const ader = planD1 ? realizado / planD1 * 100 : 0;
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
    tabela(regs);
  }

  function grafico_(regs, dias, ref, duMes, planejado) {
    const porDia = {};
    regs.forEach(r => { if (r.st === 'Executada' && r.d) porDia[r.d] = (porDia[r.d] || 0) + 1; });
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
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 16, bottom: 6 } },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top', align: 'start', labels: { usePointStyle: true, boxWidth: 8, boxHeight: 8 } },
          tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmt(c.parsed.y)}` } },
          datalabels: temLabels ? {
            display: c => { const i = c.dataIndex, v = c.dataset.data;
              return v[i] != null && (i === 0 || v[i] !== v[i - 1]) && i % 2 === 0 || i === v.length - 1 && v[i] != null; },
            font: { size: 10, family: 'JetBrains Mono', weight: '700' }, formatter: v => fmt(v)
          } : undefined
        },
        scales: { y: { beginAtZero: true, ticks: { callback: v => fmt(v) } }, x: { grid: { display: false } } }
      },
      plugins: temLabels ? [ChartDataLabels] : []
    };
    if (grafico) grafico.destroy();
    grafico = new Chart(raiz.querySelector('#pv-canvas'), cfg);
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
      <div class="kpi-row" id="pv-op-kpis" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));"></div>
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
          <th style="text-align:center;">Status Final</th><th class="num">% Cronograma Infratel</th><th>Relatórios Entregues</th><th>Expurgo</th></tr></thead>
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
    const realizado = regs.filter(r => r.st === 'Executada').length;
    const entrega = planejado ? realizado / planejado * 100 : 0;
    const totalRel = regs.filter(r => r.ex === 'EPS').length;
    const relEntregues = regs.filter(temRelatorio).length;
    const expurgo = regs.filter(r => r.xp).length;
    const cores = { verde: 0, amarelo: 0, vermelho: 0 };
    regs.forEach(r => { if (r.st === 'Executada') { const c = corStatusFinal(r); if (c) cores[c]++; } });
    const traco = v => v ? v : '--';

    raiz.querySelector('#pv-op-kpis').innerHTML = `
      <div class="kpi"><div class="label">📋 Planejado no Mês</div>
        <div class="value">${fmt(planejado)}</div>
        <div class="delta" style="color:${COR.muted}">% Entrega <b style="color:#1b2440">${pct(entrega)}</b></div></div>
      <div class="kpi"><div class="label">📄 Relatórios Entregues | Total</div>
        <div class="value" style="display:flex;gap:10px;align-items:baseline;white-space:nowrap;">
          <span>${traco(relEntregues && fmt(relEntregues))}</span><span class="pv-sep">|</span><span>${fmt(totalRel)}</span></div>
        <div class="delta" style="color:${COR.muted}">% Volume total <b style="color:#1b2440">${relEntregues ? pct(relEntregues / planejado * 100) : '--'}</b>
          · % Relatórios <b style="color:#1b2440">${relEntregues && totalRel ? pct(relEntregues / totalRel * 100) : '--'}</b></div></div>
      <div class="kpi"><div class="label">🚫 Qtd Expurgo</div>
        <div class="value"><span style="color:${COR.warn}">${fmt(expurgo)}</span></div>
        <div class="delta" style="color:${COR.muted}">% Expurgo <b style="color:#1b2440">${pct(planejado ? expurgo / planejado * 100 : 0)}</b></div></div>
      <div class="kpi"><div class="label">Legenda "Executada"</div>
        <div class="pv-legenda">
          <span><i style="background:#16a34a"></i><b>Executada:</b> Infratel 100% + relatório recebido <b class="n">${fmt(cores.verde)}</b></span>
          <span><i style="background:#facc15"></i><b>Divergência:</b> relatório recebido, Infratel 0% <b class="n">${fmt(cores.amarelo)}</b></span>
          <span><i style="background:#dc2626"></i><b>Verificar:</b> Infratel 100%, sem relatório <b class="n">${fmt(cores.vermelho)}</b></span>
        </div></div>`;

    tabelaOp(regs);
  }

  function tabelaOp(regs) {
    const lista = regs.filter(r => !buscaOp ||
      [r.w, r.si, r.s, r.ch, r.t, r.ti, r.x, r.m].some(v => String(v ?? '').toLowerCase().includes(buscaOp)))
      .sort((a, b) => String(a.w).localeCompare(String(b.w)));
    raiz.querySelector('#pv-op-qtd').textContent = fmt(lista.length);
    raiz.querySelector('#pv-op-tbody').innerHTML = lista.map(r => {
      const cor = r.st === 'Executada' ? corStatusFinal(r) : '';
      const status = cor ? `<span class="pv-sf ${cor}">${esc(r.st)}</span>` : `<span class="pv-sf">${esc(r.st)}</span>`;
      return `<tr>
        <td class="mono" style="font-size:12px;">${esc(r.w)}</td><td>${esc(r.si)}</td><td>${esc(r.ti) || '-'}</td><td>${esc(r.x)}</td>
        <td class="num">${r.pi === null || r.pi === undefined ? '' : fmt(r.pi)}</td><td class="mono" style="font-size:12px;">${esc(r.ch)}</td>
        <td>${r.io === 'OK' ? '<span class="pill good">OK</span>' : `<span class="pill bad">${esc(r.io) || '—'}</span>`}</td>
        <td>Preventiva infra - ${esc(r.t)}</td><td style="text-align:center;">${status}</td>
        <td class="num">${fmtPct(r.pc)}</td><td>${esc(r.er)}</td><td>${esc(r.xp)}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="12" style="text-align:center;color:#6b7590">Nenhuma preventiva encontrada</td></tr>';
  }

  function trocarAba(aba) {
    raiz.querySelectorAll('.pv-aba').forEach(b => b.classList.toggle('ativa', b.dataset.aba === aba));
    raiz.querySelector('#pv-aba-geral').style.display = aba === 'geral' ? '' : 'none';
    raiz.querySelector('#pv-aba-op').style.display = aba === 'op' ? '' : 'none';
    if (aba === 'op') { if (!opMontada) montarOp(); atualizarOp(); }
    else if (grafico) grafico.resize();
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

  window.PainelPreventiva = { montar, redesenhar: () => grafico && grafico.resize() };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => montar());
  else montar();
})();
