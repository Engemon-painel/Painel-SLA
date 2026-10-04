/* =========================================================================
   Painel do ESER — Seção "AFAZ"
   Fonte: dados_afaz.json (gerado da aba qAFAZAtual)
   Requer: Chart.js (datalabels opcional)
   Uso:    <div id="sec-afaz"></div>
           <script src="gestao_afaz.js"> (depois do Chart.js)
   ========================================================================= */
(function () {
  const ARQUIVO_JSON = 'dados_afaz.json';
  const CONTAINER_PADRAO = 'sec-afaz';

  // Statuses considerados "tratados/fechados" (igual à medida do Power BI).
  const TRATADOS = new Set(['Finalizado', 'Resolvido', 'Pendência Claro']);

  const FILTROS = [
    { k: 'st', rotulo: 'Status' },
    { k: 'oc', rotulo: 'Origem Carteira' },
    { k: 'eq', rotulo: 'Equipe Engemon' },
    { k: 'sp', rotulo: 'Site Prioritário' }
  ];

  const COR = { good: '#059669', warn: '#d97706', bad: '#dc2626', accent: '#0e7c86', muted: '#6b7590' };
  const COR_CI = { 'Backlog': '#60a5fa', 'Entrante': '#1d4ed8', 'Prioritário': '#7c3aed' };

  // =======================================================================
  // CALENDÁRIO DO PLANO — "Acompanhamento Semanal - Plan"
  // Semanas ISO (segunda a domingo) de SEMANA_INICIO até SEMANA_FIM do
  // ANO_PLANO. 2026 tem 53 semanas ISO, por isso termina em W53.
  // =======================================================================
  const ANO_PLANO = 2026;
  const SEMANA_INICIO = 26;
  const SEMANA_FIM = 53;

  // Backlog Planejado Restante — meta de cada semana (a linha azul-clara
  // do Power BI). Valores lidos do gráfico do Power BI: CONFIRA, principalmente
  // W27 a W30, onde os números ficavam sobrepostos na imagem.
  const PLANO_BACKLOG = {
    26: 2076, 27: 2007, 28: 1883, 29: 1818, 30: 1716, 31: 1619, 32: 1528,
    33: 1437, 34: 1345, 35: 1254, 36: 1167, 37: 1081, 38: 998, 39: 910,
    40: 826, 41: 744, 42: 662, 43: 580, 44: 498, 45: 436, 46: 374,
    47: 313, 48: 251, 49: 192, 50: 135, 51: 77, 52: 20, 53: 0
  };

  // Backlog Real Restante — fechamento de cada semana JÁ PASSADA (linha
  // azul-escura do Power BI). O dados_afaz.json é só a foto de hoje, então
  // o histórico fica aqui. A semana atual é calculada sozinha a partir da
  // foto (itens Backlog ainda pendentes). Quando a semana fechar, acrescente
  // o número dela aqui. CONFIRA W27 a W30 (números sobrepostos na imagem).
  const HISTORICO_BACKLOG_REAL = {
    27: 2019, 28: 1977, 29: 1918, 30: 1855, 31: 1812, 32: 1774, 33: 1715,
    34: 1618, 35: 1616, 36: 1546, 37: 1518, 38: 1477, 39: 1468, 40: 1450
  };

  // Quais carteiras contam como "Backlog Real Restante" no cálculo da semana atual.
  const CARTEIRAS_BACKLOG_REAL = new Set(['Backlog']);

  const COR_PLANO = { planejado: '#38a5f5', real: '#2a2a8f', entrantes: '#7b1fa2' };

  let dados = null, raiz = null, busca = '';
  const sel = { st: '', oc: '', eq: '', sp: '' };
  let gStatus = null, gEquipe = null, gPrioridade = null, gAging = null, gSemanal = null, gPlano = null;

  const fmt = n => n.toLocaleString('pt-BR');
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const pad2 = n => (n < 10 ? '0' : '') + n;

  // Segunda-feira da semana ISO "semana" do "ano".
  function segundaDaSemanaIso(ano, semana) {
    const jan4 = new Date(ano, 0, 4);
    const diaSemana = (jan4.getDay() + 6) % 7; // segunda = 0
    const segunda = new Date(ano, 0, 4 - diaSemana);
    segunda.setDate(segunda.getDate() + (semana - 1) * 7);
    return segunda;
  }

  // Lista de semanas do plano: [{ num: 26, label: 'W26', periodo: '22/06 – 28/06' }, ...]
  function calendarioPlano() {
    const semanas = [];
    for (let w = SEMANA_INICIO; w <= SEMANA_FIM; w++) {
      const ini = segundaDaSemanaIso(ANO_PLANO, w);
      const fim = new Date(ini); fim.setDate(fim.getDate() + 6);
      semanas.push({
        num: w,
        label: 'W' + pad2(w),
        periodo: pad2(ini.getDate()) + '/' + pad2(ini.getMonth() + 1) + ' – ' + pad2(fim.getDate()) + '/' + pad2(fim.getMonth() + 1)
      });
    }
    return semanas;
  }

  // "W07" -> 7
  const numSemana = s => { const m = String(s || '').match(/(\d+)/); return m ? Number(m[1]) : null; };

  function estilos() {
    if (document.getElementById('afz-estilos')) return;
    const css = `
    .afz-filtros{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 16px}
    .afz-filtros label{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted,#6b7590);text-transform:uppercase;letter-spacing:.05em}
    .afz-filtros label:not(:first-child){margin-left:8px}
    .afz-filtros select{font-family:'JetBrains Mono',monospace;font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid var(--line,#e1e5f0);background:var(--panel,#fff);color:var(--text,#1b2440);min-width:150px}
    #afz-kpis{grid-template-columns:repeat(auto-fit,minmax(180px,1fr))}
    .afz-cv{position:relative;height:320px}
    .afz-cv-plano{position:relative;height:380px}
    .afz-rolagem{max-height:420px;overflow:auto;border-radius:8px}
    .afz-tab-topo{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}
    .afz-tab-topo .search input{min-width:240px}`;
    const st = document.createElement('style'); st.id = 'afz-estilos'; st.textContent = css;
    document.head.appendChild(st);
  }

  function esqueleto() {
    raiz.innerHTML = `
    <div class="afz">
      <div class="afz-filtros">
        ${FILTROS.map(f => `<label>${f.rotulo}</label><select data-k="${f.k}"></select>`).join('')}
        <span class="mono" style="font-size:11px;color:var(--muted,#6b7590);margin-left:auto;">
          atualizado em ${esc(dados.gerado_em || '')}</span>
      </div>

      <div class="kpi-row" id="afz-kpis"></div>

      <div class="grid">
        <div class="panel">
          <h2>Qtde AFAZ por Status e Carteira</h2>
          <div class="afz-cv"><canvas id="afz-status"></canvas></div>
        </div>
        <div class="panel">
          <h2>Qtde AFAZ por Equipe Engemon e Carteira</h2>
          <div class="afz-cv"><canvas id="afz-equipe"></canvas></div>
        </div>
      </div>
      <div class="grid">
        <div class="panel">
          <h2>Qtde AFAZ por Prioridade e Carteira</h2>
          <div class="afz-cv"><canvas id="afz-prioridade"></canvas></div>
        </div>
        <div class="panel">
          <h2>Aging Médio (Dias) por Status</h2>
          <div class="afz-cv"><canvas id="afz-aging"></canvas></div>
        </div>
      </div>

      <div class="panel" style="margin-bottom:16px;">
        <h2>Acompanhamento Semanal — Plan</h2>
        <div class="afz-cv-plano"><canvas id="afz-plano"></canvas></div>
      </div>

      <div class="panel" style="margin-bottom:16px;">
        <h2>Prioritários em Esteira (por semana)</h2>
        <div class="afz-cv"><canvas id="afz-semanal"></canvas></div>
      </div>

      <div class="panel">
        <div class="afz-tab-topo">
          <h2>Itens em Aberto (<span id="afz-qtd-tab">0</span>)</h2>
          <div class="search">
            <input id="afz-busca" type="search" placeholder="Buscar site, título, responsável...">
          </div>
        </div>
        <div class="afz-rolagem"><table>
          <thead><tr><th>Status</th><th>Site</th><th>Prioridade</th><th>Título</th><th>Área/Sistema</th>
          <th>Responsável</th><th>Carteira</th><th>Equipe</th><th>Aging (dias)</th></tr></thead>
          <tbody id="afz-tbody"></tbody></table></div>
      </div>
    </div>`;

    raiz.querySelectorAll('.afz-filtros select').forEach(s =>
      s.addEventListener('change', () => { sel[s.dataset.k] = s.value; atualizar(); }));
    raiz.querySelector('#afz-busca').addEventListener('input', e => { busca = e.target.value.toLowerCase(); tabela(filtrar()); });
  }

  function filtrar(ignorar) {
    return dados.registros.filter(r => FILTROS.every(f => f.k === ignorar || !sel[f.k] || r[f.k] === sel[f.k]));
  }

  function opcoes() {
    raiz.querySelectorAll('.afz-filtros select').forEach(s => {
      const k = s.dataset.k;
      const vals = [...new Set(filtrar(k).map(r => r[k]))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
      if (sel[k] && !vals.includes(sel[k])) vals.push(sel[k]);
      s.innerHTML = '<option value="">Todos</option>' +
        vals.map(v => `<option value="${esc(v)}"${v === sel[k] ? ' selected' : ''}>${esc(v)}</option>`).join('');
    });
  }

  function atualizar() {
    opcoes();
    const regs = filtrar();
    const pendentes = regs.filter(r => !TRATADOS.has(r.st));
    const tratados = regs.filter(r => TRATADOS.has(r.st));
    const backlogPendente = pendentes.filter(r => r.ci === 'Backlog').length;
    const entrantesPendentes = pendentes.filter(r => r.ci === 'Entrante').length;
    const prioritariosPendentes = pendentes.filter(r => r.ci === 'Prioritário').length;

    raiz.querySelector('#afz-kpis').innerHTML = `
      <div class="kpi"><div class="label">📥 Pendentes AFAZ</div><div class="value">${fmt(pendentes.length)}</div></div>
      <div class="kpi"><div class="label">📦 Backlog Pendente</div><div class="value">${fmt(backlogPendente)}</div></div>
      <div class="kpi"><div class="label">🆕 Entrantes Pendentes</div><div class="value">${fmt(entrantesPendentes)}</div></div>
      <div class="kpi"><div class="label">⭐ Prioritários Pendentes</div><div class="value" style="color:${COR.bad}">${fmt(prioritariosPendentes)}</div></div>
      <div class="kpi"><div class="label">✅ Total Tratado</div><div class="value" style="color:${COR.good}">${fmt(tratados.length)}</div></div>`;

    graficoEmpilhado('afz-status', gStatus, v => (gStatus = v), pendentes, 'st');
    graficoEmpilhado('afz-equipe', gEquipe, v => (gEquipe = v), pendentes, 'eq');
    graficoEmpilhado('afz-prioridade', gPrioridade, v => (gPrioridade = v), pendentes, 'pr', ['P0', 'P1', 'P2', 'P3']);
    graficoAging(pendentes);
    graficoPlano();
    graficoSemanal();
    tabela(pendentes);
  }

  // "Acompanhamento Semanal - Plan": as três linhas do Power BI no calendário
  // W26–W53. É série histórica, então NÃO respeita os filtros da página.
  function graficoPlano() {
    const semanas = calendarioPlano();

    // Entrantes acumulados vêm do dados_afaz.json. A série começa no ano
    // anterior (W35 de 2025...), por isso os rótulos se repetem: como ela está
    // em ordem cronológica, a última ocorrência de cada "Wnn" é a de 2026.
    const serieEntr = Array.isArray(dados.entrantesAcumulados) ? dados.entrantesAcumulados : [];
    const entrPorSemana = {};
    serieEntr.forEach(s => { const n = numSemana(s.semana); if (n !== null) entrPorSemana[n] = s.entrantesAcumulados; });
    const semanaAtual = serieEntr.length ? numSemana(serieEntr[serieEntr.length - 1].semana) : null;

    // Backlog real: histórico fixo + semana atual calculada da foto de hoje.
    const realPorSemana = Object.assign({}, HISTORICO_BACKLOG_REAL);
    if (semanaAtual !== null && realPorSemana[semanaAtual] === undefined) {
      realPorSemana[semanaAtual] = dados.registros.filter(r => !TRATADOS.has(r.st) && CARTEIRAS_BACKLOG_REAL.has(r.ci)).length;
    }

    // Depois da semana atual não há dado real: fica em branco (só o plano segue até W53).
    const valorAte = (mapa, w) => (semanaAtual !== null && w > semanaAtual) ? null : (mapa[w] ?? null);

    const labels = semanas.map(s => s.label);
    const planejado = semanas.map(s => PLANO_BACKLOG[s.num] ?? null);
    const real = semanas.map(s => valorAte(realPorSemana, s.num));
    const entrantes = semanas.map(s => valorAte(entrPorSemana, s.num));

    const temLabels = typeof ChartDataLabels !== 'undefined';
    const rotulo = (cor, align) => temLabels
      ? { display: ctx => ctx.dataset.data[ctx.dataIndex] !== null, align: align, anchor: 'center', offset: 4,
          color: cor, font: { size: 9, weight: '700' }, formatter: v => fmt(v) }
      : undefined;

    if (gPlano) gPlano.destroy();
    gPlano = new Chart(raiz.querySelector('#afz-plano'), {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          { label: 'Backlog Planejado Restante', data: planejado, borderColor: COR_PLANO.planejado, backgroundColor: COR_PLANO.planejado,
            borderWidth: 3, tension: 0.15, pointRadius: 2, datalabels: rotulo('#1b2440', 'bottom') },
          { label: 'Backlog Real Restante', data: real, borderColor: COR_PLANO.real, backgroundColor: COR_PLANO.real,
            borderWidth: 3, tension: 0.15, pointRadius: 2, spanGaps: false, datalabels: rotulo('#1b2440', 'top') },
          { label: 'Entrantes em Esteira Acumulados', data: entrantes, borderColor: COR_PLANO.entrantes, backgroundColor: COR_PLANO.entrantes,
            borderWidth: 3, tension: 0.15, pointRadius: 2, spanGaps: false, datalabels: rotulo('#1b2440', 'bottom') }
        ]
      },
      plugins: temLabels ? [ChartDataLabels] : [],
      options: {
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { top: 18 } },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top', align: 'start', labels: { boxWidth: 10, font: { size: 11 } } },
          tooltip: {
            callbacks: {
              title: items => {
                const s = semanas[items[0].dataIndex];
                return s.label + ' · ' + s.periodo + (s.num === semanaAtual ? ' (semana atual)' : '');
              }
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { size: 10 }, autoSkip: false, maxRotation: 0 } },
          y: { display: false, beginAtZero: true }
        }
      }
    });
  }

  // Prioritários em Esteira por semana — série histórica pronta do
  // dados_afaz.json; não respeita os filtros da página.
  function graficoSemanal() {
    const serie = Array.isArray(dados.semanal) ? dados.semanal : [];
    const temLabels = typeof ChartDataLabels !== 'undefined';
    if (gSemanal) gSemanal.destroy();
    gSemanal = new Chart(raiz.querySelector('#afz-semanal'), {
      type: 'line',
      data: {
        labels: serie.map(s => s.semana),
        datasets: [{ label: 'Prioritários em Esteira', data: serie.map(s => s.prioritariosEmEsteira),
          borderColor: '#7c3aed', backgroundColor: '#7c3aed', tension: 0.3, pointRadius: 4 }]
      },
      plugins: temLabels ? [ChartDataLabels] : [],
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false },
          datalabels: temLabels ? { align: 'top', color: '#7c3aed', font: { size: 10, weight: '700' } } : undefined },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: '#e1e5f0' } } }
      }
    });
  }

  // Barras horizontais empilhadas por Carteira Indicador (Backlog / Entrante / Prioritário).
  function graficoEmpilhado(canvasId, atual, setAtual, regs, campo, ordemFixa) {
    const CIS = ['Backlog', 'Entrante', 'Prioritário'];
    const porCat = {};
    regs.forEach(r => {
      const cat = r[campo] || 'Sem valor';
      if (!porCat[cat]) porCat[cat] = { Backlog: 0, Entrante: 0, 'Prioritário': 0 };
      if (porCat[cat][r.ci] !== undefined) porCat[cat][r.ci]++;
    });
    let cats = Object.keys(porCat);
    cats = ordemFixa ? ordemFixa.filter(c => porCat[c]) :
      cats.sort((a, b) => (porCat[b].Backlog + porCat[b].Entrante + porCat[b]['Prioritário']) - (porCat[a].Backlog + porCat[a].Entrante + porCat[a]['Prioritário']));

    const temLabels = typeof ChartDataLabels !== 'undefined';
    if (atual) atual.destroy();
    const novo = new Chart(raiz.querySelector('#' + canvasId), {
      type: 'bar',
      data: {
        labels: cats,
        datasets: CIS.map(ci => ({ label: ci, data: cats.map(c => porCat[c][ci]), backgroundColor: COR_CI[ci], stack: 's' }))
      },
      plugins: temLabels ? [ChartDataLabels] : [],
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { boxWidth: 10, font: { size: 11 } } },
          datalabels: temLabels ? { color: '#fff', font: { size: 9, weight: '700' }, formatter: v => v || '' } : undefined
        },
        scales: { x: { stacked: true, beginAtZero: true, grid: { color: '#e1e5f0' } }, y: { stacked: true, grid: { display: false } } }
      }
    });
    setAtual(novo);
  }

  function graficoAging(regs) {
    const soma = {}, qtd = {};
    regs.forEach(r => {
      if (r.tst == null) return;
      const st = r.st;
      soma[st] = (soma[st] || 0) + r.tst / 24;
      qtd[st] = (qtd[st] || 0) + 1;
    });
    const linhas = Object.keys(soma).map(st => ({ st, media: soma[st] / qtd[st] })).sort((a, b) => b.media - a.media);

    const temLabels = typeof ChartDataLabels !== 'undefined';
    if (gAging) gAging.destroy();
    gAging = new Chart(raiz.querySelector('#afz-aging'), {
      type: 'bar',
      data: { labels: linhas.map(l => l.st), datasets: [{ label: 'Aging Médio (dias)', data: linhas.map(l => Math.round(l.media * 10) / 10), backgroundColor: 'rgba(14,124,134,.75)', borderRadius: 4 }] },
      plugins: temLabels ? [ChartDataLabels] : [],
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, datalabels: temLabels ? { anchor: 'end', align: 'end', color: '#000', font: { size: 10, weight: '700' } } : undefined },
        scales: { x: { beginAtZero: true, grid: { color: '#e1e5f0' } }, y: { grid: { display: false }, ticks: { font: { size: 11 } } } }
      }
    });
  }

  function tabela(regs) {
    const lista = regs.filter(r => !busca ||
      [r.s, r.ti, r.resp, r.ar].some(v => String(v ?? '').toLowerCase().includes(busca)));
    raiz.querySelector('#afz-qtd-tab').textContent = fmt(lista.length);
    raiz.querySelector('#afz-tbody').innerHTML = lista.slice(0, 500).map(r => `<tr>
      <td><span class="pill ${r.st === 'Aberto' ? 'warn' : 'bad'}">${esc(r.st)}</span></td>
      <td>${esc(r.s)}</td><td>${esc(r.pr)}</td><td>${esc(r.ti)}</td><td>${esc(r.ar)}</td>
      <td>${esc(r.resp)}</td><td>${esc(r.oc)}</td><td>${esc(r.eq)}</td>
      <td class="num mono">${r.tst != null ? fmt(Math.round(r.tst / 24)) : '--'}</td>
    </tr>`).join('') || '<tr><td colspan="9" style="text-align:center;color:#6b7590">Nenhum item encontrado</td></tr>';
  }

  async function montar(idContainer) {
    raiz = document.getElementById(idContainer || CONTAINER_PADRAO);
    if (!raiz) return;
    estilos();
    try {
      dados = window.DADOS_AFAZ ||
        await fetch(ARQUIVO_JSON + '?v=' + Date.now(), { cache: 'no-store' }).then(r => { if (!r.ok) throw r.status; return r.json(); });
      esqueleto();
      atualizar();
    } catch (e) {
      raiz.innerHTML = `<div style="padding:20px;color:#b00020">Não foi possível carregar ${ARQUIVO_JSON} (${esc(e)}).</div>`;
    }
  }

  function redesenhar() {
    [gStatus, gEquipe, gPrioridade, gAging, gSemanal, gPlano].forEach(g => { if (g) g.resize(); });
  }

  window.PainelAFAZ = { montar, redesenhar };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => montar());
  else montar();
})();
