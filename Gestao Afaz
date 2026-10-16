/* =========================================================================
   Painel-SLA · Seção AFAZ  (gestao_afaz.js)
   Lê dados_afaz.json (gerado pelo Office Script gerar_dados_afaz.ts) e monta
   a seção dentro de <div id="sec-afaz">.
   Visuais:
     1. Top 10 sites com mais pendências (por Tipo)
     2. Tipo: Backlog x Entrante
     3. Dias no Status por Status (faixas de dias)
     4. Pendências por Prioridade (por Tipo)
     5. Programado x Executado mês a mês
   Requer Chart.js 4 (e, opcional, chartjs-plugin-datalabels) já carregados.
   ========================================================================= */
(function () {
  'use strict';

  const ARQUIVO = 'dados_afaz.json';
  const CONTAINER = 'sec-afaz';

  const COR_TIPO = { Backlog: '#60a5fa', Entrante: '#1d4ed8', 'Sem tipo': '#94a3b8' };
  const ORDEM_TIPO = ['Backlog', 'Entrante', 'Sem tipo'];
  const ORDEM_PR = ['P0', 'P1', 'P2', 'P3'];
  const FAIXAS = [
    { nome: 'Até 15 dias', max: 15, cor: '#22c55e' },
    { nome: '16 a 30 dias', max: 30, cor: '#a3e635' },
    { nome: '31 a 60 dias', max: 60, cor: '#facc15' },
    { nome: '61 a 90 dias', max: 90, cor: '#f97316' },
    { nome: 'Mais de 90 dias', max: Infinity, cor: '#dc2626' }
  ];
  const COR_PROG = '#94a3b8', COR_EXEC = '#1d4ed8';
  const MES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

  let dados = null, dadosBrutos = null;
  const graficos = {};
  const filtro = { tp: '', pr: '', st: '', ar: '', site: '', busca: '' };

  // ---------------------------------------------------------------- util
  const $ = (sel, raiz) => (raiz || document).querySelector(sel);
  const nf = n => Number(n || 0).toLocaleString('pt-BR');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const unicos = (arr, campo) => [...new Set(arr.map(i => i[campo]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const rotMes = m => { const [a, mm] = m.split('-'); return MES_ABREV[+mm - 1] + '/' + a.slice(2); };
  const dataBR = s => s ? s.split('-').reverse().join('/') : '—';
  const temDL = () => typeof window.ChartDataLabels !== 'undefined';

  function faixaDe(d) { return FAIXAS.findIndex(f => d <= f.max); }

  // ---------------------------------------------------------------- estilos
  function injetarCSS() {
    if ($('#afz-css')) return;
    const css = `
    #${CONTAINER} .afz-filtros{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 14px}
    #${CONTAINER} .afz-filtros select,#${CONTAINER} .afz-filtros input{padding:6px 10px;border:1px solid #cbd5e1;border-radius:8px;font:inherit;font-size:13px;background:#fff;color:#0f172a}
    #${CONTAINER} .afz-chip{display:inline-flex;gap:6px;align-items:center;padding:4px 10px;border-radius:999px;background:#1d4ed8;color:#fff;font-size:12px;cursor:pointer}
    #${CONTAINER} .afz-limpar{padding:6px 12px;border:0;border-radius:8px;background:#e2e8f0;color:#0f172a;font:inherit;font-size:13px;cursor:pointer}
    #${CONTAINER} .afz-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:14px}
    #${CONTAINER} .afz-kpi{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:12px 14px}
    #${CONTAINER} .afz-kpi small{display:block;font-size:12px;color:#64748b}
    #${CONTAINER} .afz-kpi b{display:block;font-size:24px;line-height:1.2;color:#0f172a}
    #${CONTAINER} .afz-kpi span{font-size:11px;color:#64748b}
    #${CONTAINER} .afz-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:12px}
    #${CONTAINER} .afz-card{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:12px 14px;min-width:0;grid-column:span 3}
    #${CONTAINER} .afz-card.c2{grid-column:span 2}
    #${CONTAINER} .afz-card.c4{grid-column:span 4}
    #${CONTAINER} .afz-card.full{grid-column:1/-1}
    #${CONTAINER} .afz-card h3{margin:0 0 2px;font-size:14px;color:#0f172a}
    #${CONTAINER} .afz-card p.sub{margin:0 0 8px;font-size:12px;color:#64748b}
    #${CONTAINER} .afz-box{position:relative;height:300px}
    #${CONTAINER} .afz-box.alto{height:360px}
    #${CONTAINER} .afz-tipo{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:10px;align-items:center}
    #${CONTAINER} .afz-leg{font-size:12px;color:#334155}
    #${CONTAINER} .afz-leg div{border-left:4px solid;padding:6px 10px;margin-bottom:8px;background:#f8fafc;border-radius:0 8px 8px 0}
    #${CONTAINER} .afz-leg b{font-size:13px}
    #${CONTAINER} .afz-tab{overflow:auto;max-height:420px}
    #${CONTAINER} table{width:100%;border-collapse:collapse;font-size:12px}
    #${CONTAINER} th,#${CONTAINER} td{padding:6px 8px;border-bottom:1px solid #e2e8f0;text-align:left;white-space:nowrap}
    #${CONTAINER} td.ti{white-space:normal;min-width:240px}
    #${CONTAINER} th{position:sticky;top:0;background:#f1f5f9;z-index:1}
    #${CONTAINER} .afz-tag{display:inline-block;padding:1px 8px;border-radius:999px;color:#fff;font-size:11px}
    #${CONTAINER} .afz-rodape{font-size:11px;color:#64748b;margin-top:8px}
    @media (max-width:900px){#${CONTAINER} .afz-grid,#${CONTAINER} .afz-tipo{grid-template-columns:1fr}#${CONTAINER} .afz-card{grid-column:1/-1!important}}`;
    const st = document.createElement('style');
    st.id = 'afz-css';
    st.textContent = css;
    document.head.appendChild(st);
  }

  // ---------------------------------------------------------------- layout
  function montarLayout(el) {
    el.innerHTML = `
      <div class="afz-filtros">
        <select data-f="tp"><option value="">Tipo: todos</option></select>
        <select data-f="pr"><option value="">Prioridade: todas</option></select>
        <select data-f="st"><option value="">Status: todos</option></select>
        <select data-f="ar"><option value="">Área/Sistema: todas</option></select>
        <input data-f="busca" type="search" placeholder="Buscar site, título, responsável...">
        <span id="afz-chip-site"></span>
        <button class="afz-limpar" type="button">Limpar filtros</button>
      </div>
      <div class="afz-kpis" id="afz-kpis"></div>
      <div class="afz-grid">
        <div class="afz-card">
          <h3>Top 10 sites com mais pendências</h3>
          <p class="sub">Separado por tipo · clique na barra para filtrar o site</p>
          <div class="afz-box alto"><canvas id="afz-g-sites"></canvas></div>
        </div>
        <div class="afz-card">
          <h3>Tipo: Backlog x Entrante</h3>
          <p class="sub">Pendências atuais por tipo de carteira</p>
          <div class="afz-tipo">
            <div class="afz-box"><canvas id="afz-g-tipo"></canvas></div>
            <div class="afz-leg" id="afz-leg-tipo"></div>
          </div>
        </div>
        <div class="afz-card full">
          <h3>Dias no status atual, por status</h3>
          <p class="sub">Quantos AFAZ estão parados em cada status e há quanto tempo · média de dias ao lado da barra</p>
          <div class="afz-box alto"><canvas id="afz-g-dias"></canvas></div>
        </div>
        <div class="afz-card c2">
          <h3>Pendências por prioridade</h3>
          <p class="sub">P0 a P3, separado por tipo</p>
          <div class="afz-box"><canvas id="afz-g-prio"></canvas></div>
        </div>
        <div class="afz-card c4">
          <h3>Programado x Executado · mês a mês</h3>
          <p class="sub">Executado = mudanças para Finalizado, Resolvido ou Pendência Claro no mês (Logs) · % executado abaixo do mês</p>
          <div class="afz-box"><canvas id="afz-g-prog"></canvas></div>
        </div>
        <div class="afz-card full">
          <h3>AFAZ pendentes</h3>
          <p class="sub" id="afz-tab-sub"></p>
          <div class="afz-tab"><table>
            <thead><tr><th>AFAZ</th><th>Site</th><th>Prioridade</th><th>Tipo</th><th>Status</th><th>Dias no status</th><th>Título</th><th>Área/Sistema</th><th>Responsável</th><th>Previsão</th></tr></thead>
            <tbody id="afz-tbody"></tbody>
          </table></div>
        </div>
      </div>
      <div class="afz-rodape" id="afz-rodape"></div>`;

    // opções dos filtros
    const it = dados.itens;
    const preencher = (campo, valores) => {
      const sel = $(`select[data-f="${campo}"]`, el);
      valores.forEach(v => sel.insertAdjacentHTML('beforeend', `<option value="${esc(v)}">${esc(v)}</option>`));
    };
    preencher('tp', ORDEM_TIPO.filter(t => it.some(i => i.tp === t)));
    preencher('pr', ORDEM_PR.filter(p => it.some(i => i.pr === p)));
    preencher('st', unicos(it, 'st'));
    preencher('ar', unicos(it, 'ar'));

    el.querySelectorAll('[data-f]').forEach(c => c.addEventListener(c.tagName === 'INPUT' ? 'input' : 'change', () => {
      filtro[c.dataset.f] = c.value.trim();
      atualizar();
    }));
    $('.afz-limpar', el).addEventListener('click', () => {
      Object.keys(filtro).forEach(k => filtro[k] = '');
      el.querySelectorAll('[data-f]').forEach(c => c.value = '');
      atualizar();
    });
  }

  // ---------------------------------------------------------------- filtro
  function filtrados() {
    const b = filtro.busca.toLowerCase();
    return dados.itens.filter(i =>
      (!filtro.tp || i.tp === filtro.tp) &&
      (!filtro.pr || i.pr === filtro.pr) &&
      (!filtro.st || i.st === filtro.st) &&
      (!filtro.ar || i.ar === filtro.ar) &&
      (!filtro.site || i.s === filtro.site) &&
      (!b || [i.s, i.ti, i.resp, i.ar, i.id].some(v => String(v || '').toLowerCase().includes(b))));
  }

  // ---------------------------------------------------------------- charts
  function desenhar(id, cfg) {
    if (graficos[id]) graficos[id].destroy();
    const cv = document.getElementById(id);
    if (!cv) return;
    if (temDL()) cfg.plugins = (cfg.plugins || []).concat([window.ChartDataLabels]);
    cfg.options = cfg.options || {};
    cfg.options.responsive = true;
    cfg.options.maintainAspectRatio = false;
    cfg.options.locale = 'pt-BR';
    graficos[id] = new Chart(cv, cfg);
  }

  // rótulo dentro da barra só se couber
  const dlDentro = {
    color: '#fff', font: { size: 10, weight: '600' }, formatter: v => v ? nf(v) : '',
    display: c => {
      const v = c.dataset.data[c.dataIndex];
      if (!v) return false;
      const esc = c.chart.scales[c.chart.options.indexAxis === 'y' ? 'x' : 'y'];
      return Math.abs(esc.getPixelForValue(v) - esc.getPixelForValue(0)) > 24; // só se couber
    }
  };

  function porTipo(itens, chave, ordem) {
    const rotulos = ordem;
    const tipos = ORDEM_TIPO.filter(t => itens.some(i => i.tp === t));
    return tipos.map(t => ({
      label: t, backgroundColor: COR_TIPO[t], stack: 's', datalabels: dlDentro,
      data: rotulos.map(r => itens.filter(i => i[chave] === r && i.tp === t).length)
    }));
  }

  function totalNoFim(datasets) {
    // plugin simples: escreve o total no fim da barra empilhada horizontal/vertical
    return {
      id: 'totais',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const horiz = chart.options.indexAxis === 'y';
        const metas = chart.data.datasets.map((d, k) => chart.getDatasetMeta(k)).filter(m => !m.hidden);
        if (!metas.length) return;
        ctx.save();
        ctx.font = '600 11px sans-serif';
        ctx.fillStyle = '#0f172a';
        chart.data.labels.forEach((_, i) => {
          const tot = metas.reduce((s, m) => s + (chart.data.datasets[m.index].data[i] || 0), 0);
          if (!tot) return;
          const el = metas[metas.length - 1].data[i];
          const txt = typeof datasets === 'function' ? datasets(i, tot) : nf(tot);
          if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, el.x + 6, el.y); }
          else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, el.x, el.y - 4); }
        });
        ctx.restore();
      }
    };
  }

  // 1. Top 10 sites
  function gSites(itens) {
    const cont = {};
    itens.forEach(i => { if (i.s) cont[i.s] = (cont[i.s] || 0) + 1; });
    const top = Object.entries(cont).sort((a, b) => b[1] - a[1]).slice(0, 10).map(e => e[0]);
    desenhar('afz-g-sites', {
      type: 'bar',
      data: { labels: top, datasets: porTipo(itens.filter(i => top.includes(i.s)), 's', top) },
      plugins: [totalNoFim()],
      options: {
        indexAxis: 'y', layout: { padding: { right: 36 } },
        scales: { x: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' } }, y: { stacked: true, grid: { display: false } } },
        plugins: { legend: { position: 'bottom' } },
        onClick: (e, els, ch) => {
          if (!els.length) return;
          filtro.site = ch.data.labels[els[0].index];
          setTimeout(atualizar, 0); // não destruir o gráfico dentro do próprio evento
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });
  }

  // 2. Tipo
  function gTipo(itens) {
    const tipos = ORDEM_TIPO.filter(t => itens.some(i => i.tp === t));
    const qt = tipos.map(t => itens.filter(i => i.tp === t).length);
    const tot = qt.reduce((a, b) => a + b, 0);
    desenhar('afz-g-tipo', {
      type: 'doughnut',
      data: { labels: tipos, datasets: [{ data: qt, backgroundColor: tipos.map(t => COR_TIPO[t]), borderWidth: 2, borderColor: '#fff' }] },
      options: {
        cutout: '58%',
        plugins: {
          legend: { position: 'bottom' },
          datalabels: { color: '#fff', font: { weight: '700', size: 12 }, formatter: v => tot ? `${nf(v)}\n${Math.round(v / tot * 100)}%` : '', textAlign: 'center' }
        }
      }
    });

    // legenda explicativa: período de abertura de cada tipo + tratados
    const leg = tipos.map((t, k) => {
      const datas = itens.filter(i => i.tp === t && i.ab).map(i => i.ab).sort();
      const periodo = datas.length ? `Abertos de ${dataBR(datas[0])} a ${dataBR(datas[datas.length - 1])}` : '';
      const trat = (dados.tratados || {})[t] || 0;
      return `<div style="border-color:${COR_TIPO[t]}"><b>${esc(t)}</b> · ${nf(qt[k])} pendentes<br>${periodo}<br>${nf(trat)} já tratados</div>`;
    }).join('');
    $('#afz-leg-tipo').innerHTML = leg +
      '<div style="border-color:#cbd5e1">Backlog = carteira que já estava em aberto · Entrante = demanda nova que entrou na carteira</div>';
  }

  // 3. Dias no status por status
  function gDias(itens) {
    const grupos = {};
    itens.forEach(i => { (grupos[i.st] = grupos[i.st] || []).push(i.dst || 0); });
    const status = Object.keys(grupos).sort((a, b) => grupos[b].length - grupos[a].length);
    const medias = status.map(s => grupos[s].reduce((a, b) => a + b, 0) / grupos[s].length);
    const datasets = FAIXAS.map((f, k) => ({
      label: f.nome, backgroundColor: f.cor, stack: 'd',
      datalabels: Object.assign({}, dlDentro, { color: k === 1 || k === 2 ? '#0f172a' : '#fff' }),
      data: status.map(s => grupos[s].filter(d => faixaDe(d) === k).length)
    }));
    desenhar('afz-g-dias', {
      type: 'bar',
      data: { labels: status, datasets },
      plugins: [totalNoFim((i, tot) => `${nf(tot)} · média ${Math.round(medias[i])} d`)],
      options: {
        indexAxis: 'y', layout: { padding: { right: 120 } },
        scales: { x: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' } }, y: { stacked: true, grid: { display: false } } },
        plugins: {
          legend: { position: 'bottom' },
          tooltip: { callbacks: { footer: its => `Média no status: ${Math.round(medias[its[0].dataIndex])} dias · Máx: ${Math.max(...grupos[status[its[0].dataIndex]])} dias` } }
        }
      }
    });
  }

  // 4. Prioridade
  function gPrio(itens) {
    const prs = ORDEM_PR.filter(p => itens.some(i => i.pr === p));
    desenhar('afz-g-prio', {
      type: 'bar',
      data: { labels: prs, datasets: porTipo(itens, 'pr', prs) },
      plugins: [totalNoFim()],
      options: {
        layout: { padding: { top: 18 } },
        scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' } } },
        plugins: { legend: { position: 'bottom' } }
      }
    });
  }

  // 5. Programado x Executado
  function gProg() {
    const prog = dados.programado || {}, exe = dados.executado || {};
    const mesAtual = new Date().toISOString().slice(0, 7);
    const meses = [...new Set([...Object.keys(prog), ...Object.keys(exe)])].sort();
    const vProg = meses.map(m => prog[m] != null ? prog[m] : null);
    const vExe = meses.map(m => m <= mesAtual ? (exe[m] || 0) : null);
    const pct = meses.map((m, k) => vProg[k] && vExe[k] != null ? Math.round(vExe[k] / vProg[k] * 100) : null);
    const corPct = v => v == null ? '#64748b' : v >= 100 ? '#16a34a' : v >= 80 ? '#ca8a04' : '#dc2626';
    const rotulo = { anchor: 'end', align: 'end', offset: 2, font: { size: 10, weight: '600' }, formatter: v => v == null ? '' : nf(v) };
    desenhar('afz-g-prog', {
      type: 'bar',
      data: {
        // cada rótulo do eixo X tem duas linhas: mês e % executado
        labels: meses.map((m, k) => [rotMes(m), pct[k] == null ? '' : pct[k] + '%']),
        datasets: [
          { label: 'Programado', data: vProg, backgroundColor: COR_PROG, datalabels: Object.assign({ color: '#475569' }, rotulo) },
          { label: 'Executado', data: vExe, backgroundColor: COR_EXEC, datalabels: Object.assign({ color: COR_EXEC }, rotulo) }
        ]
      },
      options: {
        layout: { padding: { top: 20 } },
        scales: {
          x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkip: false,
               color: c => c.tick && c.tick.label && c.tick.label[1] ? corPct(pct[c.index]) : '#475569' } },
          y: { beginAtZero: true, grid: { color: '#f1f5f9' }, title: { display: true, text: 'Qtde AFAZ' } }
        },
        plugins: {
          legend: { position: 'bottom' },
          tooltip: { callbacks: {
            title: its => rotMes(meses[its[0].dataIndex]),
            footer: its => pct[its[0].dataIndex] == null ? '' : `Executado: ${pct[its[0].dataIndex]}% do programado` } }
        }
      }
    });
    return { mesAtual, prog: prog[mesAtual], exe: exe[mesAtual] || 0 };
  }

  // ---------------------------------------------------------------- KPIs / tabela
  function kpis(itens, pm) {
    const n = t => itens.filter(i => i.tp === t).length;
    const trat = Object.values(dados.tratados || {}).reduce((a, b) => a + b, 0);
    const medDias = itens.length ? Math.round(itens.reduce((s, i) => s + (i.dst || 0), 0) / itens.length) : 0;
    const ader = pm.prog ? Math.round(pm.exe / pm.prog * 100) + '%' : '—';
    const cards = [
      ['AFAZ pendentes', nf(itens.length), `média de ${medDias} dias no status`],
      ['Backlog pendente', nf(n('Backlog')), ''],
      ['Entrante pendente', nf(n('Entrante')), ''],
      ['P0 pendentes', nf(itens.filter(i => i.pr === 'P0').length), ''],
      ['Total tratado', nf(trat), 'Finalizado, Resolvido, Pendência Claro'],
      [`Executado ${rotMes(pm.mesAtual)}`, nf(pm.exe), pm.prog ? `de ${nf(pm.prog)} programados · ${ader}` : 'sem programação no mês']
    ];
    $('#afz-kpis').innerHTML = cards.map(c => `<div class="afz-kpi"><small>${c[0]}</small><b>${c[1]}</b><span>${c[2]}</span></div>`).join('');
  }

  function tabela(itens) {
    const lista = itens.slice().sort((a, b) => (b.dst || 0) - (a.dst || 0));
    const LIM = 500;
    $('#afz-tbody').innerHTML = lista.slice(0, LIM).map(i => {
      const f = FAIXAS[faixaDe(i.dst || 0)];
      return `<tr><td>${esc(i.id)}</td><td>${esc(i.s)}</td><td>${esc(i.pr)}</td>
        <td><span class="afz-tag" style="background:${COR_TIPO[i.tp] || '#94a3b8'}">${esc(i.tp)}</span></td>
        <td>${esc(i.st)}</td><td><span class="afz-tag" style="background:${f.cor};color:${f.max === 30 || f.max === 60 ? '#0f172a' : '#fff'}">${nf(i.dst)} d</span></td>
        <td class="ti">${esc(i.ti)}</td><td>${esc(i.ar)}</td><td>${esc(i.resp)}</td><td>${dataBR(i.prev)}</td></tr>`;
    }).join('');
    $('#afz-tab-sub').textContent = `${nf(lista.length)} itens, do mais parado para o mais recente` + (lista.length > LIM ? ` · mostrando os ${LIM} primeiros (use os filtros)` : '');
  }

  // ---------------------------------------------------------------- ciclo
  function atualizar() {
    const itens = filtrados();
    $('#afz-chip-site').innerHTML = filtro.site ? `<span class="afz-chip" title="Remover">Site: ${esc(filtro.site)} ✕</span>` : '';
    const chip = $('#afz-chip-site .afz-chip');
    if (chip) chip.onclick = () => { filtro.site = ''; atualizar(); };
    gSites(itens);
    gTipo(itens);
    gDias(itens);
    gPrio(itens);
    const pm = gProg();
    kpis(itens, pm);
    tabela(itens);
  }

  // Aceita o JSON novo ({itens, tratados, programado, executado}) e também o formato antigo
  // (lista com todos os AFAZ e campos oc/ci/tst). Retorna true se for o formato antigo.
  function normalizar(d) {
    if (Array.isArray(d.itens) && d.itens.length) return false;
    const lista = Array.isArray(d) ? d
      : (d.registros || d.dados || d.afaz || d.itens ||
         Object.values(d).find(v => Array.isArray(v) && v.length && v[0] && typeof v[0] === 'object' && 'st' in v[0]) || []);
    const TRATADOS = ['finalizado', 'resolvido', 'pendência claro', 'pendencia claro'];
    const tipoDe = r => {
      const t = String(r.tp || r.oc || r.ci || '').trim();
      return /backlog/i.test(t) ? 'Backlog' : /entrante/i.test(t) ? 'Entrante' : 'Sem tipo';
    };
    const isoData = v => {
      if (!v) return null;
      const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
      return String(v).slice(0, 10);
    };
    const tratados = {};
    const itens = [];
    lista.forEach(r => {
      const st = String(r.st || '').trim();
      const tp = tipoDe(r);
      if (TRATADOS.includes(st.toLowerCase())) { tratados[tp] = (tratados[tp] || 0) + 1; return; }
      const dst = r.dst != null ? Number(r.dst) : Math.round((Number(r.tst) || 0) / 24);
      itens.push({ id: r.id, pr: r.pr, s: r.s, st, ti: r.ti, ab: isoData(r.ab), ar: r.ar, resp: r.resp, prev: isoData(r.prev), tp, dst });
    });
    d.itens = itens;
    d.tratados = d.tratados || tratados;
    d.programado = d.programado || {};
    d.executado = d.executado || {};
    d.geradoEm = d.geradoEm || d.gerado_em || d.atualizado_em || null;
    return lista.length > 0;
  }

  // Lê o JSON mesmo quando o Power Automate grava embrulhado: {"result":"{...}"}, {"body":...} ou texto
  async function lerJson(url) {
    const r = await fetch(url + '?v=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) throw new Error('HTTP ' + r.status + ' ao abrir ' + url);
    let d = JSON.parse((await r.text()).replace(/^\uFEFF/, ''));
    for (let i = 0; i < 4; i++) {
      if (typeof d === 'string') { d = JSON.parse(d); continue; }
      if (d && typeof d === 'object' && !Array.isArray(d)) {
        const emb = ['result', 'body', 'Result', 'Body', 'content', 'value'].find(k => k in d && Object.keys(d).length <= 3);
        if (emb) { d = d[emb]; continue; }
      }
      break;
    }
    return d;
  }
  function diagnostico(el, nome, d) {
    const tipo = Array.isArray(d) ? 'lista com ' + d.length + ' itens' : typeof d;
    const chaves = d && typeof d === 'object' && !Array.isArray(d) ? Object.keys(d).slice(0, 15).join(', ') : '—';
    el.insertAdjacentHTML('afterbegin', `<div style="padding:12px 14px;margin-bottom:12px;border-radius:10px;background:#fee2e2;color:#991b1b;font-size:13px">
      ⚠ ${nome} foi lido, mas não tem registros no formato esperado.<br>Tipo recebido: <b>${tipo}</b> · campos no topo: <b>${chaves}</b><br>
      Mande um print desta mensagem para ajustar.</div>`);
  }

  async function iniciar() {
    const el = document.getElementById(CONTAINER);
    if (!el) return;
    if (typeof Chart === 'undefined') { el.innerHTML = '<p>Chart.js não carregado.</p>'; return; }
    try {
      dados = await lerJson(ARQUIVO);
      dadosBrutos = Array.isArray(dados) ? dados : Object.assign({}, dados);
    } catch (e) {
      el.innerHTML = `<p style="color:#dc2626">Não foi possível carregar ${ARQUIVO}: ${esc(e.message)}</p>`;
      return;
    }
    if (Array.isArray(dados)) dados = { registros: dados };
    const formatoAntigo = normalizar(dados);
    injetarCSS();
    montarLayout(el);
    atualizar();
    if (dados.geradoEm) {
      $('#afz-rodape').textContent = 'Dados gerados em ' + new Date(dados.geradoEm).toLocaleString('pt-BR');
    }
    if (!dados.itens.length && !Object.keys(dados.tratados || {}).length) diagnostico(el, ARQUIVO, dadosBrutos);
    if (formatoAntigo) {
      el.insertAdjacentHTML('afterbegin', '<div style="padding:10px 14px;margin-bottom:12px;border-radius:10px;background:#fef3c7;color:#92400e;font-size:13px">' +
        '⚠ O dados_afaz.json está no formato antigo (gerado pelo Office Script anterior). O painel funciona, mas o gráfico ' +
        'Programado x Executado fica vazio. Troque o script do fluxo pelo gerar_dados_afaz.ts novo.</div>');
    }
  }

  // gráficos criados com a seção escondida ficam com tamanho 0: redimensiona ao aparecer
  window.afazResize = () => Object.values(graficos).forEach(g => g.resize());

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
