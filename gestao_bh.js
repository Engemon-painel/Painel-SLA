/* =========================================================================
   Painel-SLA · Banco de Horas — Visão Gerencial  (gestao_bh.js)
   Fonte: dados_bh.json, gerado pelo Office Script gerar_dados_bh.ts a partir
   da aba "Banco de Horas _ ANALITICO".
   Monta tudo dentro de <div id="sec-bh">. Requer Chart.js 4
   (chartjs-plugin-datalabels é opcional).
   Mostra saldo POSITIVO (idade das horas) e NEGATIVO (horas a repor) lado a lado.
   ========================================================================= */
(function () {
  'use strict';

  const ARQUIVO = 'dados_bh.json';
  const CONTAINER = 'sec-bh';

  // ----- faixas -----
  // Positivo: idade das horas no banco (dias). Os índices batem com o array bh[] do JSON.
  const FX_POS = ['0-30', '31-60', '61-90', '91-120', '121-150', '151-180', '>180', 'PRÉ-BASE / IDADE INDETERMINADA'];
  const ROT_POS = ['0–30 d', '31–60 d', '61–90 d', '91–120 d', '121–150 d', '151–180 d', '> 180 d', 'Pré-base'];
  const ACAO_POS = ['Saldo saudável', 'Acompanhar', 'Acompanhar', 'Planejar folgas', 'Programar folgas',
    'Compensação imediata', 'Compensação imediata', 'Validar origem das horas'];
  // Negativo: tamanho do débito
  const FX_NEG = ['0 A -8H', '-8H01 A -16H', '-16H01 A -24H', 'ABAIXO DE -24H'];
  const ROT_NEG = ['0 a -8 h', '-8 a -16 h', '-16 a -24 h', 'Abaixo de -24 h'];
  const ACAO_NEG = ['Monitorar', 'Acompanhamento próximo', 'Plano de reposição', 'Reposição imediata'];

  const CRIT = {
    'SAUDÁVEL': '#16a34a', 'LEVE': '#4ade80', 'ATENÇÃO': '#facc15', 'CRÍTICO': '#f97316',
    'MUITO CRÍTICO': '#dc2626', 'IDADE INDETERMINADA': '#94a3b8'
  };
  const ORDEM_CRIT = ['MUITO CRÍTICO', 'CRÍTICO', 'ATENÇÃO', 'LEVE', 'SAUDÁVEL', 'IDADE INDETERMINADA'];
  const COR_FX_POS = ['#16a34a', '#84cc16', '#facc15', '#f59e0b', '#f97316', '#ef4444', '#b91c1c', '#94a3b8'];
  const COR_FX_NEG = ['#fca5a5', '#f87171', '#ef4444', '#b91c1c'];
  const VERDE = '#0f9d58', VERMELHO = '#dc2626', AZUL = '#0b3a75';

  let dados = null;
  const graficos = {};
  const filtro = { co: '', sup: '', tp: '', cr: '', st: '', busca: '' };

  // ----- util -----
  const $ = (s, r) => (r || document).querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = (n, d) => Number(n || 0).toLocaleString('pt-BR', { maximumFractionDigits: d || 0, minimumFractionDigits: d || 0 });
  const sinal = n => (n > 0 ? '+' : n < 0 ? '−' : '');
  const hTot = n => sinal(Math.round(n)) + nf(Math.abs(Math.round(n))) + ' h';
  const hMin = n => { // 90.5 -> "+90h30"
    const a = Math.abs(n), h = Math.floor(a), m = Math.round((a - h) * 60);
    const [hh, mm] = m === 60 ? [h + 1, 0] : [h, m];
    return sinal(n) + hh + 'h' + String(mm).padStart(2, '0');
  };
  const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
  const unicos = (arr, k) => [...new Set(arr.map(i => i[k]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const tag = (txt, cor) => {
    const escuro = ['#facc15', '#4ade80', '#84cc16', '#fca5a5', '#94a3b8'].includes(cor);
    return `<span class="bh-tag" style="background:${cor};color:${escuro ? '#1e293b' : '#fff'}">${esc(txt)}</span>`;
  };
  const corCrit = c => CRIT[c] || '#94a3b8';
  const nomeCrit = c => c ? c.charAt(0) + c.slice(1).toLowerCase() : '—';
  const rotFx = it => it.tp === 'NEGATIVO'
    ? (ROT_NEG[FX_NEG.indexOf(it.fx)] || it.fx)
    : (ROT_POS[FX_POS.indexOf(it.fx)] || it.fx);
  const temDL = () => typeof window.ChartDataLabels !== 'undefined';
  const ehCritico = it => it.cr === 'CRÍTICO' || it.cr === 'MUITO CRÍTICO';

  // ----- CSS -----
  function css() {
    if ($('#bh-css')) return;
    const C = '#' + CONTAINER;
    const s = document.createElement('style');
    s.id = 'bh-css';
    s.textContent = `
${C}{--bh-borda:#e2e8f0;--bh-txt:#0f172a;--bh-sub:#64748b;color:var(--bh-txt)}
${C} .bh-topo{display:flex;flex-wrap:wrap;gap:12px 24px;align-items:center;justify-content:space-between;background:linear-gradient(100deg,#0b2545,#0b3a75);color:#fff;border-radius:14px;padding:16px 22px;margin-bottom:12px}
${C} .bh-topo h2{margin:0;font-size:24px;line-height:1.1}
${C} .bh-topo h2 small{display:block;font-size:15px;font-weight:500;opacity:.85}
${C} .bh-topo p{margin:0;font-size:13px;opacity:.8}
${C} .bh-filtros{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
${C} .bh-filtros select,${C} .bh-filtros input{padding:7px 10px;border:1px solid #cbd5e1;border-radius:8px;font:inherit;font-size:13px;background:#fff;color:#0f172a}
${C} .bh-filtros button{padding:7px 12px;border:0;border-radius:8px;background:#e2e8f0;font:inherit;font-size:13px;cursor:pointer}
${C} .bh-g{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:12px}
${C} .bh-c{background:#fff;border:1px solid var(--bh-borda);border-radius:14px;padding:14px 16px;min-width:0}
${C} .bh-c h3{margin:0 0 2px;font-size:15px;color:#0b2545}
${C} .bh-c .sub{margin:0 0 10px;font-size:12px;color:var(--bh-sub)}
${C} .s3{grid-column:span 3}${C} .s4{grid-column:span 4}${C} .s5{grid-column:span 5}${C} .s6{grid-column:span 6}${C} .s7{grid-column:span 7}${C} .s8{grid-column:span 8}${C} .s12{grid-column:span 12}
${C} .bh-kpis{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px}
${C} .bh-k{border-radius:14px;padding:14px 16px;color:#fff;min-width:0}
${C} .bh-k small{display:block;font-size:13px;opacity:.9;font-weight:600}
${C} .bh-k b{display:block;font-size:30px;line-height:1.15;margin:4px 0 2px;white-space:nowrap}
${C} .bh-k span{font-size:12px;opacity:.85}
${C} .bh-box{position:relative;height:260px}
${C} .bh-box.b{height:320px}
${C} .bh-box.p{height:190px}
${C} .bh-tipo{display:flex;gap:14px;align-items:center}
${C} .bh-tipo .bh-box{flex:0 0 150px;height:150px}
${C} .bh-tipo ul{list-style:none;margin:0;padding:0;font-size:13px}
${C} .bh-tipo li{margin:6px 0}
${C} .bh-dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:middle}
${C} .bh-crit{display:grid;grid-template-columns:1fr;gap:12px}
${C} .bh-crit h4{margin:0 0 6px;font-size:12px;text-transform:uppercase;letter-spacing:.04em}
${C} .bh-bar{display:grid;grid-template-columns:130px 1fr 72px;gap:8px;align-items:center;font-size:12px;margin:5px 0}
${C} .bh-bar i{display:block;height:12px;border-radius:6px;background:#f1f5f9;overflow:hidden}
${C} .bh-bar i em{display:block;height:100%;border-radius:6px}
${C} .bh-bar b{text-align:right}
${C} .bh-tab{overflow:auto;max-height:360px}
${C} table{width:100%;border-collapse:collapse;font-size:12px}
${C} th,${C} td{padding:7px 8px;border-bottom:1px solid #eef2f7;text-align:left;white-space:nowrap}
${C} th{position:sticky;top:0;background:#f1f5f9;color:#334155;z-index:1}
${C} td.num,${C} th.num{text-align:right}
${C} td.nm{max-width:240px;overflow:hidden;text-overflow:ellipsis}${C} td.nm small{display:block;color:#64748b;font-size:11px;overflow:hidden;text-overflow:ellipsis}
${C} .bh-tag{display:inline-block;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:600}
${C} .pos{color:${VERDE};font-weight:700}${C} .neg{color:${VERMELHO};font-weight:700}
${C} .bh-ins{margin:10px 0 0;padding:10px 12px;background:#f8fafc;border-radius:10px;font-size:12px}
${C} .bh-ins ul{margin:4px 0 0;padding-left:18px}
${C} .bh-ins li{margin:3px 0}
${C} .bh-rod{font-size:11px;color:var(--bh-sub);margin-top:10px;text-align:right}
@media (max-width:1100px){${C} .bh-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}${C} .bh-g>*{grid-column:1/-1!important}}
@media (max-width:600px){${C} .bh-kpis{grid-template-columns:1fr}}`;
    document.head.appendChild(s);
  }

  // ----- layout -----
  function layout(el) {
    el.innerHTML = `
    <div class="bh-topo">
      <h2>Banco de Horas<small>Visão Gerencial</small></h2>
      <p>Saldo positivo (idade das horas) e saldo negativo (horas a repor)</p>
    </div>
    <div class="bh-filtros">
      <select data-f="co"><option value="">Coordenador: todos</option></select>
      <select data-f="sup"><option value="">Supervisor: todos</option></select>
      <select data-f="tp"><option value="">Tipo de saldo: todos</option><option value="POSITIVO">Positivo</option><option value="NEGATIVO">Negativo</option></select>
      <select data-f="cr"><option value="">Criticidade: todas</option></select>
      <select data-f="st"><option value="">Status: todos</option></select>
      <input data-f="busca" type="search" placeholder="Buscar nome ou matrícula">
      <button type="button" class="bh-limpar">Limpar filtros</button>
    </div>
    <div class="bh-g">
      <div class="s8"><div class="bh-kpis" id="bh-kpis"></div></div>
      <div class="bh-c s4">
        <h3>Tipo de saldo</h3>
        <div class="bh-tipo"><div class="bh-box"><canvas id="bh-g-tipo"></canvas></div><ul id="bh-leg-tipo"></ul></div>
      </div>

      <div class="bh-c s5">
        <h3>Saldo positivo · horas por idade</h3>
        <p class="sub">Há quanto tempo as horas estão no banco (dias)</p>
        <div class="bh-box"><canvas id="bh-g-pos"></canvas></div>
      </div>
      <div class="bh-c s3">
        <h3>Saldo negativo · por faixa</h3>
        <p class="sub">Colaboradores e horas a repor</p>
        <div class="bh-box"><canvas id="bh-g-neg"></canvas></div>
      </div>
      <div class="bh-c s4">
        <h3>Criticidade</h3>
        <p class="sub">% de colaboradores em cada nível, dentro do tipo de saldo</p>
        <div class="bh-crit"><div id="bh-crit-pos"></div><div id="bh-crit-neg"></div></div>
      </div>

      <div class="bh-c s12">
        <h3>Saldo por supervisor</h3>
        <p class="sub">Horas positivas e negativas da equipe · o número no fim é o saldo líquido</p>
        <div class="bh-box b" id="bh-box-sup"><canvas id="bh-g-sup"></canvas></div>
      </div>

      <div class="bh-c s6">
        <h3 style="color:${VERDE}">▲ Colaboradores com saldo positivo (prioridade)</h3>
        <p class="sub" id="bh-sub-pos"></p>
        <div class="bh-tab"><table><thead><tr><th>Matrícula</th><th>Nome / supervisor</th><th class="num">Saldo</th><th>Faixa mais crítica</th><th class="num">&gt; 150 d</th><th>Criticidade</th></tr></thead><tbody id="bh-t-pos"></tbody></table></div>
      </div>
      <div class="bh-c s6">
        <h3 style="color:${VERMELHO}">▼ Colaboradores com saldo negativo (prioridade)</h3>
        <p class="sub" id="bh-sub-neg"></p>
        <div class="bh-tab"><table><thead><tr><th>Matrícula</th><th>Nome / supervisor</th><th class="num">Saldo</th><th>Faixa</th><th>Criticidade</th></tr></thead><tbody id="bh-t-neg"></tbody></table></div>
      </div>

      <div class="bh-c s4">
        <h3>Ação por faixa · positivo</h3>
        <div class="bh-tab" style="max-height:none"><table><thead><tr><th>Faixa</th><th class="num">Colab.</th><th class="num">Horas</th><th>Ação recomendada</th></tr></thead><tbody id="bh-a-pos"></tbody></table></div>
      </div>
      <div class="bh-c s4">
        <h3>Ação por faixa · negativo</h3>
        <div class="bh-tab" style="max-height:none"><table><thead><tr><th>Faixa</th><th class="num">Colab.</th><th class="num">Horas</th><th>Ação recomendada</th></tr></thead><tbody id="bh-a-neg"></tbody></table></div>
      </div>
      <div class="bh-c s4">
        <h3>Resumo por tipo de saldo</h3>
        <table><thead><tr><th>Tipo</th><th class="num">Colab.</th><th class="num">%</th><th class="num">Horas</th></tr></thead><tbody id="bh-resumo"></tbody></table>
        <div class="bh-ins"><b>💡 Insights</b><ul id="bh-ins"></ul></div>
      </div>
    </div>
    <div class="bh-rod" id="bh-rod"></div>`;

    const it = dados.itens;
    const opc = (f, vals, rot) => {
      const s = $(`select[data-f="${f}"]`, el);
      vals.forEach(v => s.insertAdjacentHTML('beforeend', `<option value="${esc(v)}">${esc(rot ? rot(v) : v)}</option>`));
    };
    opc('co', unicos(it, 'co'));
    opc('sup', unicos(it, 'sup'));
    opc('cr', ORDEM_CRIT.filter(c => it.some(i => i.cr === c)), nomeCrit);
    opc('st', unicos(it, 'st'));
    if (!it.some(i => i.st)) $('select[data-f="st"]', el).style.display = 'none';

    el.querySelectorAll('[data-f]').forEach(c => c.addEventListener(c.tagName === 'INPUT' ? 'input' : 'change', () => {
      filtro[c.dataset.f] = c.value.trim();
      atualizar();
    }));
    $('.bh-limpar', el).addEventListener('click', () => {
      Object.keys(filtro).forEach(k => filtro[k] = '');
      el.querySelectorAll('[data-f]').forEach(c => c.value = '');
      atualizar();
    });
  }

  function filtrados() {
    const b = filtro.busca.toLowerCase();
    return dados.itens.filter(i =>
      (!filtro.co || i.co === filtro.co) && (!filtro.sup || i.sup === filtro.sup) &&
      (!filtro.tp || i.tp === filtro.tp) && (!filtro.cr || i.cr === filtro.cr) &&
      (!filtro.st || i.st === filtro.st) &&
      (!b || i.n.toLowerCase().includes(b) || i.m.includes(b)));
  }

  // ----- gráficos -----
  function desenhar(id, cfg) {
    if (graficos[id]) graficos[id].destroy();
    const cv = document.getElementById(id);
    if (!cv) return;
    if (temDL()) cfg.plugins = (cfg.plugins || []).concat([window.ChartDataLabels]);
    cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, locale: 'pt-BR' }, cfg.options || {});
    graficos[id] = new Chart(cv, cfg);
  }

  function gTipo(pos, neg) {
    const tot = pos.length + neg.length;
    desenhar('bh-g-tipo', {
      type: 'doughnut',
      data: { labels: ['Positivo', 'Negativo'], datasets: [{ data: [pos.length, neg.length], backgroundColor: [VERDE, VERMELHO], borderWidth: 0 }] },
      plugins: [{
        id: 'centro', afterDraw(ch) {
          const { ctx, chartArea: a } = ch, x = (a.left + a.right) / 2, y = (a.top + a.bottom) / 2;
          ctx.save(); ctx.textAlign = 'center'; ctx.fillStyle = '#0b2545';
          ctx.font = '700 22px sans-serif'; ctx.fillText(pct(pos.length, tot) + '%', x, y + 2);
          ctx.font = '12px sans-serif'; ctx.fillStyle = '#64748b'; ctx.fillText('positivo', x, y + 18);
          ctx.restore();
        }
      }],
      options: { cutout: '72%', plugins: { legend: { display: false }, datalabels: { display: false } } }
    });
    const hp = pos.reduce((s, i) => s + i.h, 0), hn = neg.reduce((s, i) => s + i.h, 0);
    $('#bh-leg-tipo').innerHTML = `
      <li><span class="bh-dot" style="background:${VERDE}"></span><b>Positivo</b> · ${nf(pos.length)} (${pct(pos.length, tot)}%)<br><span class="pos">${hTot(hp)}</span></li>
      <li><span class="bh-dot" style="background:${VERMELHO}"></span><b>Negativo</b> · ${nf(neg.length)} (${pct(neg.length, tot)}%)<br><span class="neg">${hTot(hn)}</span></li>`;
  }

  function gPos(pos) {
    const horas = FX_POS.map((_, k) => pos.reduce((s, i) => s + (i.bh[k] || 0), 0));
    const colab = FX_POS.map((_, k) => pos.filter(i => (i.bh[k] || 0) > 0).length);
    desenhar('bh-g-pos', {
      type: 'bar',
      data: { labels: ROT_POS, datasets: [{ label: 'Horas', data: horas.map(h => Math.round(h)), backgroundColor: COR_FX_POS, borderRadius: 4 }] },
      options: {
        layout: { padding: { top: 22 } },
        scales: { x: { grid: { display: false }, ticks: { font: { size: 11 } } }, y: { beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { callback: v => nf(v) + ' h' } } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: c => `${nf(horas[c.dataIndex])} h · ${colab[c.dataIndex]} colaboradores com horas nessa faixa` } },
          datalabels: { anchor: 'end', align: 'end', font: { weight: '700', size: 11 }, color: '#0f172a', formatter: v => v ? nf(v) + ' h' : '' }
        }
      }
    });
  }

  function gNeg(neg) {
    const qt = FX_NEG.map(f => neg.filter(i => i.fx === f).length);
    const hs = FX_NEG.map(f => neg.filter(i => i.fx === f).reduce((s, i) => s + i.h, 0));
    desenhar('bh-g-neg', {
      type: 'bar',
      data: { labels: ROT_NEG, datasets: [{ label: 'Colaboradores', data: qt, backgroundColor: COR_FX_NEG, borderRadius: 4 }] },
      options: {
        layout: { padding: { top: 22 } },
        scales: { x: { grid: { display: false }, ticks: { font: { size: 10 }, maxRotation: 0, autoSkip: false } }, y: { beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { precision: 0 } } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: c => `${qt[c.dataIndex]} colaboradores · ${hTot(hs[c.dataIndex])}` } },
          datalabels: { anchor: 'end', align: 'end', font: { weight: '700', size: 11 }, color: '#0f172a', formatter: v => v || '' }
        }
      }
    });
  }

  function barrasCrit(alvo, titulo, cor, lista) {
    const tot = lista.length;
    const niveis = ORDEM_CRIT.filter(c => lista.some(i => i.cr === c));
    $(alvo).innerHTML = `<h4 style="color:${cor}">${titulo} · ${nf(tot)}</h4>` + (niveis.length ? niveis.map(c => {
      const n = lista.filter(i => i.cr === c).length, p = pct(n, tot);
      return `<div class="bh-bar" title="${n} colaboradores"><span>${nomeCrit(c)}</span><i><em style="width:${p}%;background:${corCrit(c)}"></em></i><b>${p}% <small style="font-weight:400;color:#64748b">(${n})</small></b></div>`;
    }).join('') : '<p class="sub">Sem colaboradores</p>');
  }

  function gSup(itens) {
    const g = {};
    itens.forEach(i => {
      const k = i.sup || '(sem supervisor)';
      g[k] = g[k] || { p: 0, n: 0 };
      if (i.h >= 0) g[k].p += i.h; else g[k].n += i.h;
    });
    const sups = Object.keys(g).sort((a, b) => (g[b].p - g[b].n) - (g[a].p - g[a].n));
    $('#bh-box-sup').style.height = Math.max(220, sups.length * 28 + 70) + 'px';
    const liq = sups.map(s => g[s].p + g[s].n);
    desenhar('bh-g-sup', {
      type: 'bar',
      data: {
        labels: sups,
        datasets: [
          { label: 'Positivo', data: sups.map(s => Math.round(g[s].p)), backgroundColor: VERDE, stack: 's', borderRadius: 3 },
          { label: 'Negativo', data: sups.map(s => Math.round(g[s].n)), backgroundColor: VERMELHO, stack: 's', borderRadius: 3 }
        ]
      },
      plugins: [{
        id: 'liquido', afterDatasetsDraw(ch) {
          const { ctx } = ch, mp = ch.getDatasetMeta(0);
          if (mp.hidden) return;
          ctx.save(); ctx.font = '700 11px sans-serif'; ctx.textBaseline = 'middle';
          mp.data.forEach((el, k) => { ctx.fillStyle = liq[k] >= 0 ? VERDE : VERMELHO; ctx.fillText(hTot(liq[k]), Math.max(el.x, ch.scales.x.getPixelForValue(0)) + 6, el.y); });
          ctx.restore();
        }
      }],
      options: {
        indexAxis: 'y', layout: { padding: { right: 70 } },
        scales: { x: { stacked: true, grid: { color: '#f1f5f9' }, ticks: { callback: v => nf(v) + ' h' } }, y: { stacked: true, grid: { display: false } } },
        plugins: {
          legend: { position: 'top', align: 'end' },
          tooltip: { callbacks: { label: c => `${c.dataset.label}: ${hTot(c.raw)}`, footer: its => 'Saldo líquido: ' + hTot(liq[its[0].dataIndex]) } },
          datalabels: { display: false }
        }
      }
    });
  }

  // ----- KPIs / tabelas -----
  function kpis(itens, pos, neg) {
    const soma = l => l.reduce((s, i) => s + i.h, 0);
    const crit = itens.filter(ehCritico);
    const cards = [
      [AZUL, 'Saldo total', hTot(soma(itens)), 'Saldo consolidado do banco'],
      [VERDE, '▲ Saldo positivo', hTot(soma(pos)), `${nf(pos.length)} colaboradores · horas acumuladas`],
      [VERMELHO, '▼ Saldo negativo', hTot(soma(neg)), `${nf(neg.length)} colaboradores · horas a repor`],
      ['#ea8a00', '⚠ Casos críticos', nf(crit.length), `${crit.filter(i => i.tp === 'POSITIVO').length} positivos · ${crit.filter(i => i.tp === 'NEGATIVO').length} negativos`],
      ['#334155', 'Colaboradores', nf(itens.length), 'no filtro atual']
    ];
    $('#bh-kpis').innerHTML = cards.map(c => `<div class="bh-k" style="background:${c[0]}"><small>${c[1]}</small><b>${c[2]}</b><span>${c[3]}</span></div>`).join('');
  }

  function tabelas(pos, neg) {
    const velhas = i => (i.bh[5] || 0) + (i.bh[6] || 0); // horas com mais de 150 dias
    const peso = c => ORDEM_CRIT.indexOf(c) < 0 ? 99 : ORDEM_CRIT.indexOf(c);
    const p = pos.slice().sort((a, b) => peso(a.cr) - peso(b.cr) || velhas(b) - velhas(a) || b.h - a.h);
    const n = neg.slice().sort((a, b) => a.h - b.h);
    $('#bh-t-pos').innerHTML = p.map(i => `<tr><td>${esc(i.m)}</td><td class="nm" title="${esc(i.n)} · ${esc(i.sup)}">${esc(i.n)}<small>${esc(i.sup)}</small></td>
      <td class="num pos">${hMin(i.h)}</td><td>${esc(rotFx(i))}</td><td class="num">${velhas(i) ? hMin(velhas(i)).replace('+', '') : '—'}</td>
      <td>${tag(nomeCrit(i.cr), corCrit(i.cr))}</td></tr>`).join('') || '<tr><td colspan="6">Nenhum colaborador</td></tr>';
    $('#bh-t-neg').innerHTML = n.map(i => `<tr><td>${esc(i.m)}</td><td class="nm" title="${esc(i.n)} · ${esc(i.sup)}">${esc(i.n)}<small>${esc(i.sup)}</small></td>
      <td class="num neg">${hMin(i.h)}</td><td>${esc(rotFx(i))}</td><td>${tag(nomeCrit(i.cr), corCrit(i.cr))}</td></tr>`).join('') || '<tr><td colspan="5">Nenhum colaborador</td></tr>';
    $('#bh-sub-pos').textContent = `${nf(p.length)} colaboradores · ordenado por criticidade e horas mais antigas`;
    $('#bh-sub-neg').textContent = `${nf(n.length)} colaboradores · do maior débito para o menor`;
  }

  function acoes(pos, neg) {
    // positivo: colaborador entra na faixa mais crítica dele; horas = soma das horas daquela idade
    let linhasPos = FX_POS.map((f, k) => ({
      rot: ROT_POS[k], cor: COR_FX_POS[k], acao: ACAO_POS[k],
      q: pos.filter(i => i.fx === f).length, h: pos.reduce((s, i) => s + (i.bh[k] || 0), 0)
    }));
    linhasPos.unshift(...linhasPos.splice(0, 7).reverse()); // >180 d primeiro, pré-base no fim
    linhasPos = linhasPos.filter(l => l.q || l.h);
    $('#bh-a-pos').innerHTML = linhasPos.map(l => `<tr><td>${esc(l.rot)}</td><td class="num">${l.q}</td><td class="num pos">${hTot(l.h)}</td>
      <td><span class="bh-dot" style="background:${l.cor}"></span>${l.acao}</td></tr>`).join('') || '<tr><td colspan="4">—</td></tr>';
    const linhasNeg = FX_NEG.map((f, k) => ({
      rot: ROT_NEG[k], cor: COR_FX_NEG[k], acao: ACAO_NEG[k],
      q: neg.filter(i => i.fx === f).length, h: neg.filter(i => i.fx === f).reduce((s, i) => s + i.h, 0)
    })).reverse().filter(l => l.q);
    $('#bh-a-neg').innerHTML = linhasNeg.map(l => `<tr><td>${esc(l.rot)}</td><td class="num">${l.q}</td><td class="num neg">${hTot(l.h)}</td>
      <td><span class="bh-dot" style="background:${l.cor}"></span>${l.acao}</td></tr>`).join('') || '<tr><td colspan="4">—</td></tr>';
  }

  function resumo(itens, pos, neg) {
    const tot = itens.length, soma = l => l.reduce((s, i) => s + i.h, 0);
    $('#bh-resumo').innerHTML = `
      <tr><td><span class="bh-dot" style="background:${VERDE}"></span>Positivo</td><td class="num">${pos.length}</td><td class="num">${pct(pos.length, tot)}%</td><td class="num pos">${hTot(soma(pos))}</td></tr>
      <tr><td><span class="bh-dot" style="background:${VERMELHO}"></span>Negativo</td><td class="num">${neg.length}</td><td class="num">${pct(neg.length, tot)}%</td><td class="num neg">${hTot(soma(neg))}</td></tr>`;
    const hp = soma(pos);
    const h150 = pos.reduce((s, i) => s + (i.bh[5] || 0) + (i.bh[6] || 0), 0);
    const ins = [];
    if (tot) ins.push(`${pct(pos.length, tot)}% dos colaboradores estão com saldo positivo e ${pct(neg.length, tot)}% com saldo negativo (${hTot(soma(neg))}).`);
    if (hp) ins.push(`${pct(h150, hp)}% das horas positivas (${hTot(h150)}) têm mais de 150 dias no banco.`);
    const mc = itens.filter(i => i.cr === 'MUITO CRÍTICO');
    if (mc.length) ins.push(`${mc.length} casos muito críticos: ${mc.filter(i => i.tp === 'POSITIVO').length} positivos e ${mc.filter(i => i.tp === 'NEGATIVO').length} negativos.`);
    const pb = pos.reduce((s, i) => s + (i.bh[7] || 0), 0);
    if (pb) ins.push(`${hTot(pb)} estão em pré-base (idade indeterminada) e precisam de validação.`);
    $('#bh-ins').innerHTML = ins.map(t => `<li>${t}</li>`).join('') || '<li>Sem dados no filtro.</li>';
  }

  function atualizar() {
    const itens = filtrados();
    const pos = itens.filter(i => i.tp !== 'NEGATIVO');
    const neg = itens.filter(i => i.tp === 'NEGATIVO');
    kpis(itens, pos, neg);
    gTipo(pos, neg);
    gPos(pos);
    gNeg(neg);
    barrasCrit('#bh-crit-pos', '▲ Positivo', VERDE, pos);
    barrasCrit('#bh-crit-neg', '▼ Negativo', VERMELHO, neg);
    gSup(itens);
    tabelas(pos, neg);
    acoes(pos, neg);
    resumo(itens, pos, neg);
  }

  async function iniciar() {
    const el = document.getElementById(CONTAINER);
    if (!el) return;
    if (typeof Chart === 'undefined') { el.innerHTML = '<p>Chart.js não carregado.</p>'; return; }
    try {
      const r = await fetch(ARQUIVO + '?v=' + Date.now(), { cache: 'no-store' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      dados = await r.json();
    } catch (e) {
      el.innerHTML = `<p style="color:#dc2626">Não foi possível carregar ${ARQUIVO}: ${esc(e.message)}</p>`;
      return;
    }
    dados.itens = dados.itens || [];
    css();
    layout(el);
    atualizar();
    if (dados.geradoEm) $('#bh-rod').textContent = 'Atualizado em ' + new Date(dados.geradoEm).toLocaleString('pt-BR');
  }

  // se a seção estiver escondida ao carregar, chame bhResize() quando ela aparecer
  window.bhResize = () => Object.values(graficos).forEach(g => g.resize());

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
