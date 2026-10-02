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

  let dados = null, raiz = null, busca = '';
  const sel = { st: '', oc: '', eq: '', sp: '' };
  let gStatus = null, gEquipe = null, gPrioridade = null, gAging = null, gSemanal = null, gEntrantesAcum = null;

  const fmt = n => n.toLocaleString('pt-BR');
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  function estilos() {
    if (document.getElementById('afz-estilos')) return;
    const css = `
    .afz-filtros{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 16px}
    .afz-filtros label{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted,#6b7590);text-transform:uppercase;letter-spacing:.05em}
    .afz-filtros label:not(:first-child){margin-left:8px}
    .afz-filtros select{font-family:'JetBrains Mono',monospace;font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid var(--line,#e1e5f0);background:var(--panel,#fff);color:var(--text,#1b2440);min-width:150px}
    #afz-kpis{grid-template-columns:repeat(auto-fit,minmax(180px,1fr))}
    .afz-cv{position:relative;height:320px}
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

      <div class="grid">
        <div class="panel">
          <h2>Prioritários em Esteira (por semana)</h2>
          <div class="afz-cv"><canvas id="afz-semanal"></canvas></div>
        </div>
        <div class="panel">
          <h2>Acompanhamento Semanal — Entrantes Acumulados</h2>
          <div class="afz-cv"><canvas id="afz-entrantes-acum"></canvas></div>
        </div>
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
    graficoSemanal();
    graficoEntrantesAcumulados();
    tabela(pendentes);
  }

  // Entrantes em Esteira Acumulados: entre os itens Entrante (não
  // prioritário, não tratado) AINDA PENDENTES HOJE, acumula pela semana
  // de Abertura — não depende de histórico diário, calculado direto da
  // foto atual, e já vem pronto do dados_afaz.json.
  function graficoEntrantesAcumulados() {
    const serie = Array.isArray(dados.entrantesAcumulados) ? dados.entrantesAcumulados : [];
    const temLabels = typeof ChartDataLabels !== 'undefined';
    if (gEntrantesAcum) gEntrantesAcum.destroy();
    gEntrantesAcum = new Chart(raiz.querySelector('#afz-entrantes-acum'), {
      type: 'line',
      data: {
        labels: serie.map(s => s.semana),
        datasets: [{ label: 'Entrantes em Esteira Acumulados', data: serie.map(s => s.entrantesAcumulados),
          borderColor: COR_CI['Entrante'], backgroundColor: COR_CI['Entrante'], tension: 0.2, pointRadius: 3, fill: false }]
      },
      plugins: temLabels ? [ChartDataLabels] : [],
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false },
          datalabels: temLabels ? { align: 'top', color: COR_CI['Entrante'], font: { size: 9, weight: '700' } } : undefined },
        scales: { x: { grid: { display: false }, ticks: { font: { size: 10 } } }, y: { beginAtZero: true, grid: { color: '#e1e5f0' } } }
      }
    });
  }

  // Prioritários em Esteira por semana — não respeita os filtros da página,
  // pois é uma série histórica (já vem pronta do dados_afaz.json), e não
  // existe filtro histórico por Status/Equipe/Carteira em cada semana passada.
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

  // Gráfico de barras horizontais empilhadas por Carteira Indicador
  // (Backlog / Entrante / Prioritário — sites prioritários já saem das
  // outras duas categorias, igual à medida "Carteira Indicador" do Power BI)
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
    [gStatus, gEquipe, gPrioridade, gAging, gSemanal, gEntrantesAcum].forEach(g => { if (g) g.resize(); });
  }

  window.PainelAFAZ = { montar, redesenhar };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => montar());
  else montar();
})();
