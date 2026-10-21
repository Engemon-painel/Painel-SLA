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
   Aba "Prioridades em aberto": cards por Tipo Site (coluna PRIORIDADE INFRAESTRUTURA
   da aba Afaz, campo "pi" no JSON). Se o JSON ainda não tiver "pi", usa o quadro antigo
   (Concentrador / Estratégico / Isola Móvel / Isola GPON / GPON = Sim).
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
  // prioridade do site (aba Prioridades da planilha): 0 = mais crítico
  const NIVEIS_PS = ['0', '1', '2', '3', '4', 'sem'];
  const COR_PS = { '0': '#b91c1c', '1': '#ea580c', '2': '#ca8a04', '3': '#2563eb', '4': '#64748b', sem: '#cbd5e1' };
  // modoPi = o JSON traz a coluna PRIORIDADE INFRAESTRUTURA (campo "pi": P0, P1 ... P4).
  // Nesse caso a PRIORIDADE 0 a 4 de todo o painel vem dela (e não mais da aba Prioridades).
  let modoPi = false;
  const rotPS = k => k === 'sem' ? (modoPi ? 'Sem prioridade' : 'Fora da lista') : 'PRIORIDADE ' + k;
  const psDe = i => {
    if (modoPi) {
      const m = String(i.pi == null ? '' : i.pi).match(/[0-4]/);
      return m ? m[0] : 'sem';
    }
    return (i.ps === null || i.ps === undefined || i.ps === '') ? 'sem' : String(i.ps);
  };
  // demais colunas da aba Prioridades (quadro antigo, usado só se o JSON não tiver "pi")
  const ATRIB = [
    { k: 'cc', nome: 'Concentrador' }, { k: 'es', nome: 'Estratégico' },
    { k: 'im', nome: 'Isola Móvel' }, { k: 'ig', nome: 'Isola GPON' }, { k: 'gp', nome: 'GPON' }
  ];
  // vazio: "Fora da lista" se o site não está na aba Prioridades; "Em branco" se está, mas a coluna não foi preenchida
  const valAt = (i, k) => String(i[k] == null ? '' : i[k]).trim() || (psDe(i) === 'sem' ? 'Fora da lista' : 'Em branco');
  // Tipo Site = coluna PRIORIDADE INFRAESTRUTURA da aba Afaz (campo "pi")
  const SEM_PI = 'Sem tipo site';
  const temPi = () => !!(dados && dados.itens.some(i => 'pi' in i));
  const piDe = i => String(i.pi == null ? '' : i.pi).trim() || SEM_PI;
  const PALETA_PI = ['#16a34a', '#0891b2', '#7c3aed', '#db2777', '#ea580c', '#0d9488', '#4f46e5', '#65a30d', '#b45309', '#475569'];
  let corPi = {};   // cor fixa por Tipo Site (definida na carga, não muda com o filtro)
  const COR_PR_AFAZ = { P0: '#1e3a8a', P1: '#2563eb', P2: '#60a5fa', P3: '#bfdbfe' };
  const COR_PROG = '#94a3b8', COR_EXEC = '#1d4ed8', COR_ENTRANTE = '#f97316';
  // status fechados (= executados): não aparecem nas pendências, só no gráfico Meta x Executado
  const FECHADOS = ['finalizado', 'resolvido', 'pendencia claro'];
  const normSt = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
  const ehFechado = s => FECHADOS.includes(normSt(s));
  const MES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

  let dados = null, dadosBrutos = null;
  const graficos = {};
  const filtro = { tp: '', pr: '', ps: '', st: '', ar: '', site: '', cls: '', busca: '' }; // cls = "campo|valor"
  let abaAtual = 'geral';

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
    #${CONTAINER} .afz-abas{display:flex;gap:4px;border-bottom:2px solid #e2e8f0;margin:0 0 14px}
    #${CONTAINER} .afz-abas button{padding:8px 16px;border:0;background:none;font:inherit;font-size:14px;font-weight:600;color:#64748b;cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-2px}
    #${CONTAINER} .afz-abas button.on{color:#1d4ed8;border-bottom-color:#1d4ed8}
    #${CONTAINER} .afz-abas button span{display:inline-block;margin-left:6px;padding:0 7px;border-radius:999px;background:#fee2e2;color:#b91c1c;font-size:11px}
    #${CONTAINER} .afz-kpi.ps{border-left:5px solid}
    #${CONTAINER} .afz-mx-leg{display:flex;flex-wrap:wrap;gap:10px;align-items:center;font-size:11px;color:#64748b;margin:0 0 8px}
    #${CONTAINER} .afz-mx-leg span{display:inline-flex;align-items:center;gap:4px;color:#0f172a;font-weight:600}
    #${CONTAINER} .afz-mx-leg i{width:10px;height:10px;border-radius:2px;display:inline-block}
    #${CONTAINER} .afz-mx .lin{display:grid;grid-template-columns:96px 52px minmax(0,1fr) 62px;gap:8px;align-items:center;padding:7px 6px;border-radius:8px;cursor:pointer}
    #${CONTAINER} .afz-mx .lin:hover{background:#f8fafc}
    #${CONTAINER} .afz-mx .lin.on{background:#eff6ff;outline:1px solid #93c5fd}
    #${CONTAINER} .afz-mx .rot b{display:block;font-size:12px;color:#0f172a}
    #${CONTAINER} .afz-mx .rot small{font-size:10px;color:#b91c1c}
    #${CONTAINER} .afz-mx .tot{font-size:16px;font-weight:700;color:#0f172a;text-align:right}
    #${CONTAINER} .afz-mx .trilha{background:#f1f5f9;border-radius:4px;height:22px;overflow:hidden}
    #${CONTAINER} .afz-mx .barra{display:flex;height:100%}
    #${CONTAINER} .afz-mx .seg{height:100%;font-size:10px;font-weight:600;display:flex;align-items:center;justify-content:center;white-space:nowrap;overflow:hidden;border-right:1px solid #fff}
    #${CONTAINER} .afz-mx .seg:hover{filter:brightness(1.15)}
    #${CONTAINER} .afz-mx .txt{display:flex;flex-wrap:wrap;gap:10px;font-size:11px;color:#475569;margin-top:3px}
    #${CONTAINER} .afz-mx .txt span{display:inline-flex;align-items:center;gap:3px}
    #${CONTAINER} .afz-mx .txt span.zero{opacity:.4}
    #${CONTAINER} .afz-mx .txt i{width:8px;height:8px;border-radius:2px;display:inline-block}
    #${CONTAINER} .afz-mx .txt b{color:#0f172a}
    #${CONTAINER} .afz-mx .atr{font-size:11px;color:#dc2626;text-align:right}
    #${CONTAINER} .afz-cls{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px}
    #${CONTAINER} .afz-cls-b{border:1px solid #e2e8f0;border-left:5px solid #16a34a;border-radius:10px;padding:10px 12px;min-width:0;cursor:pointer;background:#fff}
    #${CONTAINER} .afz-cls-b:hover{background:#f0fdf4}
    #${CONTAINER} .afz-cls-b.on{background:#dcfce7;outline:2px solid #16a34a}
    #${CONTAINER} .afz-cls-b.tot{border-left-color:#1d4ed8}
    #${CONTAINER} .afz-cls-b.tot.on{background:#dbeafe;outline-color:#1d4ed8}
    #${CONTAINER} .afz-cls-b h4{margin:0 0 4px;font-size:12px;color:#64748b;font-weight:600}
    #${CONTAINER} .afz-cls-b b.n{font-size:26px;line-height:1.1;color:#0f172a;margin-right:6px}
    #${CONTAINER} .afz-cls-b small{font-size:11px;color:#64748b}
    #${CONTAINER} .afz-cls-b .det{font-size:11px;color:#64748b;margin:2px 0 6px}
    #${CONTAINER} .afz-cls-b .ps{display:flex;flex-wrap:wrap;gap:4px}
    #${CONTAINER} .afz-aviso{padding:10px 14px;margin-bottom:12px;border-radius:10px;background:#fef3c7;color:#92400e;font-size:13px}
    #${CONTAINER} .afz-sem-res{display:flex;flex-wrap:wrap;gap:8px 22px;margin:4px 0 8px;font-size:12px;color:#64748b}
    #${CONTAINER} .afz-sem-res b{font-size:16px;color:#0f172a;margin-right:4px}
    #${CONTAINER} .afz-sem-res i{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:baseline}
    #${CONTAINER} .afz-sem-res .ruim{color:#dc2626;font-weight:600}
    #${CONTAINER} .afz-sem-res .bom{color:#16a34a;font-weight:600}
    #${CONTAINER} .afz-st-chips{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0 4px}
    #${CONTAINER} .afz-st-chips button{display:inline-flex;align-items:baseline;gap:6px;padding:4px 10px;border:1px solid #e2e8f0;border-radius:999px;background:#fff;font:inherit;font-size:12px;color:#475569;cursor:pointer}
    #${CONTAINER} .afz-st-chips button b{font-size:13px;color:#0f172a}
    #${CONTAINER} .afz-st-chips button small{font-size:11px;color:#94a3b8}
    #${CONTAINER} .afz-st-chips button:hover{border-color:#93c5fd}
    #${CONTAINER} .afz-st-chips button.on{background:#1d4ed8;border-color:#1d4ed8;color:#fff}
    #${CONTAINER} .afz-st-chips button.on b,#${CONTAINER} .afz-st-chips button.on small{color:#fff}
    @media (max-width:900px){#${CONTAINER} .afz-grid,#${CONTAINER} .afz-tipo{grid-template-columns:1fr}#${CONTAINER} .afz-card{grid-column:1/-1!important}}`;
    const st = document.createElement('style');
    st.id = 'afz-css';
    st.textContent = css;
    document.head.appendChild(st);
  }

  // ---------------------------------------------------------------- layout
  function montarLayout(el) {
    el.innerHTML = `
      <div class="afz-abas">
        <button type="button" data-aba="geral" class="on">Visão geral</button>
        <button type="button" data-aba="prio">Prioridades em aberto<span id="afz-aba-cont"></span></button>
      </div>
      <div class="afz-filtros">
        <select data-f="tp"><option value="">Tipo: todos</option></select>
        <select data-f="pr"><option value="">Prioridade AFAZ: todas</option></select>
        <select data-f="ps"><option value="">PRIORIDADE (site): todas</option></select>
        <select data-f="st"><option value="">Status: todos</option></select>
        <select data-f="ar"><option value="">Área/Sistema: todas</option></select>
        <input data-f="busca" type="search" placeholder="Buscar site, título, responsável...">
        <span id="afz-chip-site"></span>
        <span id="afz-chip-cls"></span>
        <button class="afz-limpar" type="button">Limpar filtros</button>
      </div>
      <div id="afz-v-geral">
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
        <div class="afz-card c4">
          <h3>Dias no status atual, por status</h3>
          <p class="sub">Quantos AFAZ estão parados em cada status e há quanto tempo · média de dias ao lado da barra</p>
          <div class="afz-box alto"><canvas id="afz-g-dias"></canvas></div>
        </div>
        <div class="afz-card c2">
          <h3>Entraram por mês</h3>
          <p class="sub" id="afz-ent-sub">Entrantes pelo mês da Abertura · todos os status</p>
          <div class="afz-box alto"><canvas id="afz-g-ent"></canvas></div>
        </div>
        <div class="afz-card c2">
          <h3>Pendências por prioridade</h3>
          <p class="sub">P0 a P3, separado por tipo</p>
          <div class="afz-box"><canvas id="afz-g-prio"></canvas></div>
        </div>
        <div class="afz-card c4">
          <h3 id="afz-prog-tit">Meta Backlog x Executado Backlog x Executado Entrante</h3>
          <div class="afz-st-chips" id="afz-prog-st"></div>
          <p class="sub" id="afz-prog-sub">Meta x executado no mês · executado empilhado (Backlog + Entrante) com o total em cima · % da meta de Backlog abaixo de cada mês</p>
          <div class="afz-box"><canvas id="afz-g-prog"></canvas></div>
        </div>
        <div class="afz-card c2">
          <h3>Por status e carteira</h3>
          <p class="sub" id="afz-stc-sub">Tratados (status atual na aba Afaz) em Backlog e Entrante · clique na barra para ver o status no Meta x Executado</p>
          <div class="afz-box alto"><canvas id="afz-g-stc"></canvas></div>
        </div>
        <div class="afz-card c4">
          <h3>Visão semanal · Backlog x Entrantes</h3>
          <p class="sub" id="afz-sem-sub">Semana de segunda a domingo · executados contados a partir da W26</p>
          <div class="afz-sem-res" id="afz-sem-res"></div>
          <div class="afz-box alto"><canvas id="afz-g-sem"></canvas></div>
        </div>
        <div class="afz-card c2">
          <h3>Alarmes em aberto</h3>
          <p class="sub">Área/Sistema com "alarme" · só AFAZ em aberto (não conta os executados)</p>
          <div class="afz-box" style="height:230px"><canvas id="afz-g-alarme"></canvas></div>
          <div class="afz-leg" id="afz-leg-alarme" style="margin-top:8px"></div>
        </div>
        <div class="afz-card c4">
          <h3>Alarmes em aberto por mês de abertura</h3>
          <p class="sub" id="afz-alm-sub">Mês da coluna Abertura · Backlog e Entrante empilhados com o total em cima</p>
          <div class="afz-box alto"><canvas id="afz-g-alarme-mes"></canvas></div>
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
      </div>
      <div id="afz-v-prio" hidden>
        <div id="afz-ps-aviso"></div>
        <div class="afz-kpis" id="afz-kpis-ps"></div>
        <div class="afz-grid">
          <div class="afz-card full">
            <h3 id="afz-cls-tit">AFAZ abertos por Tipo Site</h3>
            <p class="sub" id="afz-cls-sub">Coluna PRIORIDADE INFRAESTRUTURA da aba Afaz · quebra pela PRIORIDADE do site · clique num card para filtrar a tabela</p>
            <div class="afz-cls" id="afz-cls"></div>
          </div>
          ${modoPi ? '' : `<div class="afz-card">
            <h3>AFAZ abertos por PRIORIDADE</h3>
            <p class="sub">Coluna PRIORIDADE da aba Prioridades (0 = mais crítico) · clique na linha para filtrar</p>
            <div class="afz-mx-leg" id="afz-mx-leg"></div>
            <div class="afz-mx" id="afz-mx"></div>
          </div>`}
          <div class="afz-card${modoPi ? ' full' : ''}">
            <h3>Tempo parado, por PRIORIDADE</h3>
            <p class="sub">Pendências em aberto por faixa de dias no status atual · média de dias ao lado da barra</p>
            <div class="afz-box"><canvas id="afz-g-ps"></canvas></div>
          </div>
          <div class="afz-card full">
            <h3>Prioridades em aberto</h3>
            <p class="sub" id="afz-ps-sub"></p>
            <div class="afz-tab"><table>
              <thead><tr id="afz-ps-head"></tr></thead>
              <tbody id="afz-ps-tbody"></tbody>
            </table></div>
          </div>
        </div>
      </div>
      <div class="afz-rodape" id="afz-rodape"></div>`;

    el.querySelectorAll('.afz-abas button').forEach(b => b.addEventListener('click', () => trocarAba(b.dataset.aba)));

    // opções dos filtros
    const it = dados.itens;
    const preencher = (campo, valores, prefixo) => {
      const sel = $(`select[data-f="${campo}"]`, el);
      valores.forEach(v => sel.insertAdjacentHTML('beforeend', `<option value="${esc(v)}">${esc((prefixo || '') + v)}</option>`));
    };
    preencher('tp', ORDEM_TIPO.filter(t => it.some(i => i.tp === t)));
    preencher('pr', ORDEM_PR.filter(p => it.some(i => i.pr === p)), 'AFAZ: ');
    NIVEIS_PS.filter(k => it.some(i => psDe(i) === k)).forEach(k =>
      $('select[data-f="ps"]', el).insertAdjacentHTML('beforeend', `<option value="${k}">${esc(rotPS(k))}</option>`));
    // Status: os em aberto (pendências) e os fechados (executados, vêm do execPorStatus do JSON)
    const selSt = $('select[data-f="st"]', el);
    const abertos = unicos(it, 'st').filter(v => !ehFechado(v));
    const fechados = Object.keys(dados.execPorStatus || {}).filter(ehFechado).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    if (!fechados.length) fechados.push('Finalizado', 'Pendência Claro', 'Resolvido');
    selSt.insertAdjacentHTML('beforeend',
      `<optgroup label="Em aberto (pendências)">${abertos.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('')}</optgroup>` +
      `<optgroup label="Executados (fechados)">${fechados.map(v => `<option value="${esc(v)}">${esc(v)} · executado</option>`).join('')}</optgroup>`);
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
  function passaCls(i) {
    if (!filtro.cls) return true;
    const [ck, ...resto] = filtro.cls.split('|');
    const v = resto.join('|');
    if (ck === 'pi') return piDe(i) === v;
    if (ck === 'piAll') return piDe(i) !== SEM_PI;
    if (ck === 'any') return algumSim(i);
    return valAt(i, ck) === v;
  }

  function filtrados(semPrAfaz) {
    const b = filtro.busca.toLowerCase();
    return dados.itens.filter(i =>
      (!filtro.tp || i.tp === filtro.tp) &&
      (semPrAfaz || !filtro.pr || i.pr === filtro.pr) &&
      (!filtro.ps || psDe(i) === filtro.ps) &&
      (!filtro.st || ehFechado(filtro.st) || i.st === filtro.st) &&   // status fechado só muda o gráfico Meta x Executado
      (!filtro.ar || i.ar === filtro.ar) &&
      (!filtro.site || i.s === filtro.site) &&
      passaCls(i) &&
      (!b || [i.s, i.ti, i.resp, i.ar, i.id, i.pi].some(v => String(v || '').toLowerCase().includes(b))));
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

  // 2b. Alarmes em aberto: Área/Sistema que contém "alarme" (Alarme, ALARMES...), Backlog x Entrante
  const ehAlarme = i => /alarm/i.test(String(i.ar || '').normalize('NFD').replace(/[\u0300-\u036f]/g, ''));
  function gAlarme(itens) {
    const al = itens.filter(ehAlarme);
    const tipos = ORDEM_TIPO.filter(t => al.some(i => i.tp === t));
    const qt = tipos.map(t => al.filter(i => i.tp === t).length);
    const tot = al.length;
    const pctAbertos = itens.length ? Math.round(tot / itens.length * 100) : 0;
    desenhar('afz-g-alarme', {
      type: 'doughnut',
      data: { labels: tipos, datasets: [{ data: qt, backgroundColor: tipos.map(t => COR_TIPO[t]), borderWidth: 2, borderColor: '#fff' }] },
      plugins: [{
        id: 'centroAlarme',
        afterDraw(ch) {
          const m = ch.getDatasetMeta(0).data[0];
          if (!m) return;
          const { ctx } = ch; ctx.save();
          ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
          ctx.fillStyle = '#0f172a'; ctx.font = '700 22px sans-serif'; ctx.fillText(nf(tot), m.x, m.y - 8);
          ctx.fillStyle = '#64748b'; ctx.font = '11px sans-serif'; ctx.fillText('alarmes em aberto', m.x, m.y + 12);
          ctx.restore();
        }
      }],
      options: {
        cutout: '62%',
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8 } },
          datalabels: { color: '#fff', font: { weight: '700', size: 12 }, textAlign: 'center',
            formatter: v => tot && v ? `${nf(v)}\n${Math.round(v / tot * 100)}%` : '' },
          tooltip: { callbacks: { label: c => `${c.label}: ${nf(c.raw)} (${tot ? Math.round(c.raw / tot * 100) : 0}% dos alarmes)` } }
        }
      }
    });
    const linhas = tipos.map((t, k) => {
      const doTipo = itens.filter(i => i.tp === t).length;
      return `<div style="border-color:${COR_TIPO[t]}"><b>${esc(t)}</b> · ${nf(qt[k])} alarmes · ${tot ? Math.round(qt[k] / tot * 100) : 0}% dos alarmes<br>` +
        `${doTipo ? Math.round(qt[k] / doTipo * 100) : 0}% de todo o ${esc(t)} em aberto (${nf(doTipo)})</div>`;
    }).join('');
    $('#afz-leg-alarme').innerHTML = tot
      ? linhas + `<div style="border-color:#cbd5e1">Alarmes = <b>${pctAbertos}%</b> dos ${nf(itens.length)} AFAZ em aberto</div>`
      : '<div style="border-color:#cbd5e1">Nenhum AFAZ de alarme em aberto com os filtros atuais</div>';
  }

  // 2c. Alarmes em aberto por mês de Abertura (últimos 12 meses; o que é mais antigo vira uma coluna "Até ...")
  function gAlarmeMes(itens) {
    const al = itens.filter(ehAlarme);
    const MESES_VISIVEIS = 12;
    const mesAb = i => String(i.ab || '').slice(0, 7);
    const hojeMes = new Date().toLocaleDateString('sv-SE').slice(0, 7);
    const ultimo = [hojeMes, ...al.map(mesAb).filter(Boolean)].sort().slice(-1)[0];
    const [ua, um] = ultimo.split('-').map(Number);
    const meses = [];
    for (let k = MESES_VISIVEIS - 1; k >= 0; k--) {
      const d = new Date(Date.UTC(ua, um - 1 - k, 1));
      meses.push(d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'));
    }
    const corte = meses[0];
    const antigos = al.filter(i => mesAb(i) && mesAb(i) < corte);
    const semData = al.filter(i => !mesAb(i));
    const ant = new Date(Date.UTC(+corte.slice(0, 4), +corte.slice(5, 7) - 2, 1));
    const rotAnt = 'Até ' + rotMes(ant.getUTCFullYear() + '-' + String(ant.getUTCMonth() + 1).padStart(2, '0'));
    const cols = (antigos.length ? [{ rot: rotAnt, f: i => mesAb(i) && mesAb(i) < corte }] : [])
      .concat(meses.map(m => ({ rot: rotMes(m), f: i => mesAb(i) === m })));
    const tipos = ORDEM_TIPO.filter(t => al.some(i => i.tp === t));
    desenhar('afz-g-alarme-mes', {
      type: 'bar',
      data: {
        labels: cols.map(c => c.rot),
        datasets: tipos.map(t => ({
          label: t, data: cols.map(c => al.filter(i => i.tp === t && c.f(i)).length), backgroundColor: COR_TIPO[t], stack: 'a',
          datalabels: Object.assign({}, dlDentro, { color: '#fff' })
        }))
      },
      plugins: [totalNoFim()],
      options: {
        layout: { padding: { top: 18 } },
        scales: {
          x: { stacked: true, grid: { display: false }, border: { display: false }, ticks: { maxRotation: 0, autoSkip: false, color: '#475569', font: { size: 10 } } },
          y: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' }, border: { display: false }, ticks: { color: '#94a3b8' } }
        },
        plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8 } } }
      }
    });
    $('#afz-alm-sub').textContent = `${nf(al.length)} alarmes em aberto pelo mês da coluna Abertura · Backlog e Entrante empilhados com o total em cima` +
      (antigos.length ? ` · "${rotAnt}" junta os ${nf(antigos.length)} abertos antes de ${rotMes(corte)}` : '') +
      (semData.length ? ` · ${nf(semData.length)} sem data de abertura (fora do gráfico)` : '');
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

  // 3b. Quantos AFAZ entraram por mês: todos os AFAZ pelo mês da Abertura, de qualquer tipo e status
  //     (janeiro do ano atual até o mês atual) · fonte: logsMes[mês].Abertura (script novo)
  //     ou semanal.ea + semanal.ba (Entrante + Backlog abertos por dia)
  function gEntMes() {
    const ano = new Date().getFullYear(), hoje = new Date().toLocaleDateString('sv-SE').slice(0, 7);
    const meses = [];
    for (let m = 1; ano + '-' + String(m).padStart(2, '0') <= hoje; m++) meses.push(ano + '-' + String(m).padStart(2, '0'));
    const tot = {};
    const L = dados.logsMes, S = dados.semanal || {};
    if (L && Object.keys(L).length) {
      Object.keys(L).forEach(m => {
        const ab = (L[m] || {}).Abertura || {};
        tot[m] = Object.values(ab).reduce((a, b) => a + Number(b || 0), 0);
      });
    } else if (S.ea || S.ba) {
      [S.ea || {}, S.ba || {}].forEach(o => Object.keys(o).forEach(d => { const m = d.slice(0, 7); tot[m] = (tot[m] || 0) + Number(o[d]); }));
    } else {
      $('#afz-ent-sub').textContent = '⚠ Para ver este gráfico, atualize o Office Script do fluxo e rode o fluxo de novo.';
      return;
    }
    const vTot = meses.map(m => tot[m] || 0);
    const soma = vTot.reduce((a, b) => a + b, 0);
    $('#afz-ent-sub').textContent = `${nf(soma)} AFAZ abertos em ${ano} · mês da Abertura · todos os tipos e status`;
    desenhar('afz-g-ent', {
      type: 'bar',
      data: {
        labels: meses.map(m => MES_ABREV[+m.slice(5, 7) - 1]),
        datasets: [
          { label: 'Entraram', data: vTot, backgroundColor: COR_TIPO.Entrante,
            borderRadius: 4, borderSkipped: false,
            datalabels: { anchor: 'end', align: 'end', offset: 2, color: '#0f172a', font: { size: 11, weight: '700' },
              formatter: v => v ? nf(v) : '' } }
        ]
      },
      options: {
        layout: { padding: { top: 20 } },
        scales: {
          x: { grid: { display: false }, border: { display: false },
            ticks: { maxRotation: 0, autoSkip: false, color: '#475569', font: { size: 11 } },
            title: { display: true, text: String(ano), color: '#475569' } },
          y: { beginAtZero: true, display: false, grid: { display: false } }
        },
        plugins: {
          legend: { display: false, position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8 } },
          tooltip: { callbacks: { label: c => `${nf(c.raw)} AFAZ abertos` } }
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

  // quantos executados de cada status fechado (total da base) — clique filtra só este gráfico
  function chipsStatus(eps, stF) {
    const box = $('#afz-prog-st');
    if (!box) return;
    if (!eps) { box.innerHTML = ''; return; }
    const soma = o => Object.values(o || {}).reduce((a, b) => a + Number(b || 0), 0);
    const nomes = Object.keys(eps).filter(ehFechado).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const lin = nomes.map(n => ({ n, b: soma(eps[n].b), e: soma(eps[n].e) }));
    const tb = lin.reduce((s, l) => s + l.b, 0), te = lin.reduce((s, l) => s + l.e, 0);
    const chip = (rot, b, e, val, on) => `<button type="button" class="${on ? 'on' : ''}" data-st="${esc(val)}"
      title="Backlog ${nf(b)} + Entrante ${nf(e)}">${esc(rot)} <b>${nf(b + e)}</b><small>${nf(b)} + ${nf(e)}</small></button>`;
    box.innerHTML = chip('Todos executados', tb, te, '', !stF) + lin.map(l => chip(l.n, l.b, l.e, l.n, normSt(l.n) === normSt(stF))).join('');
    box.querySelectorAll('button').forEach(bt => bt.addEventListener('click', () => {
      filtro.st = bt.dataset.st;
      $('select[data-f="st"]').value = filtro.st;
      atualizar();
    }));
  }

  // 5. Meta Backlog x Executado Backlog x Executado Entrante
  function gProg() {
    // O JSON deve conter executadoBacklog e executadoEntrante por mês.
    // "programado" é mantido como alternativa para a meta do Backlog.
    const prog = dados.metaBacklog || dados.programado || {};
    let back = dados.execBacklog || dados.executadoBacklog || {};
    let ent = dados.execEntrante || dados.executadoEntrante || {};
    // JSON com logsMes: este gráfico vira "Abertura x Devolvido x Finalizado x Resolvido" por mês (aba Logs)
    if (dados.logsMes && Object.keys(dados.logsMes).length) {
      gLogsMes();
      const mesAt = new Date().toLocaleDateString('sv-SE').slice(0, 7);
      return { mesAtual: mesAt, prog: prog[mesAt], exe: back[mesAt] || 0, entrante: ent[mesAt] || 0 };
    }
    // filtro de Status num status fechado: mostra só o executado com aquele status
    const stF = ehFechado(filtro.st) ? filtro.st : '';
    const eps = dados.execPorStatus || null;
    let semDetalhe = false;
    if (stF) {
      const k = eps ? Object.keys(eps).find(x => normSt(x) === normSt(stF)) : null;
      if (k) { back = eps[k].b || {}; ent = eps[k].e || {}; }
      else if (eps) { back = {}; ent = {}; }
      else semDetalhe = true;   // JSON antigo, sem o detalhe por status
    }
    $('#afz-prog-sub').textContent = (stF && !semDetalhe ? `Somente status ${stF} · ` : '') +
      'Meta x executado no mês · executado empilhado (Backlog + Entrante) com o total em cima · % da meta de Backlog abaixo de cada mês' +
      (/logs/i.test(dados.origemExec || '') ? ' · executado pela aba Logs: conta cada status de fechamento por onde o AFAZ passou, uma vez por mês' : '') +
      (semDetalhe ? ' · ⚠ filtro de status não aplicado: atualize o Office Script (gerar_dados_afaz.ts) para ter o executado por status' : '');
    chipsStatus(eps, stF);
    const mesAtual = new Date().toLocaleDateString('sv-SE').slice(0, 7);
    const meses = [...new Set([
      ...Object.keys(prog), ...Object.keys(back), ...Object.keys(ent)
    ])].sort();
    const vProg = meses.map(m => prog[m] != null ? Number(prog[m]) : null);
    const vBack = meses.map(m => m <= mesAtual ? Number(back[m] || 0) : null);
    const vEnt = meses.map(m => m <= mesAtual ? Number(ent[m] || 0) : null);
    const vTot = meses.map((m, k) => vBack[k] == null ? null : vBack[k] + vEnt[k]);
    const pct = meses.map((m, k) => vProg[k] > 0 && vBack[k] != null
      ? Math.round(vBack[k] / vProg[k] * 100) : null);
    // rótulo dentro da parte da coluna empilhada, só se couber
    const dentro = cor => ({
      color: cor, font: { size: 10, weight: '600' }, anchor: 'center', align: 'center',
      formatter: v => v ? nf(v) : '',
      display: c => {
        const v = c.dataset.data[c.dataIndex];
        if (!v) return false;
        const y = c.chart.scales.y;
        return Math.abs(y.getPixelForValue(0) - y.getPixelForValue(v)) >= 15;
      }
    });
    // total (Backlog + Entrante) escrito em cima da coluna do executado
    const totalExec = {
      id: 'totalExec',
      afterDatasetsDraw(chart) {
        const iB = chart.data.datasets.findIndex(d => d.stack === 'exec' && d.label.includes('Backlog'));
        const iE = chart.data.datasets.findIndex(d => d.stack === 'exec' && d.label.includes('Entrante'));
        const mB = chart.getDatasetMeta(iB), mE = chart.getDatasetMeta(iE);
        const { ctx } = chart;
        ctx.save();
        ctx.font = '700 11px sans-serif'; ctx.fillStyle = '#0f172a';
        ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        vTot.forEach((t, k) => {
          if (t == null || !t) return;
          const els = [!mB.hidden && mB.data[k], !mE.hidden && mE.data[k]].filter(Boolean);
          if (!els.length) return;
          const topo = Math.min(...els.map(e => e.y));
          ctx.fillText(nf(t), els[0].x, topo - 4);
        });
        ctx.restore();
      }
    };
    desenhar('afz-g-prog', {
      type: 'bar',
      data: {
        labels: meses.map((m, k) => [rotMes(m), pct[k] == null ? '—' : pct[k] + '%']),
        datasets: [
          { label: 'Meta Backlog', data: vProg, stack: 'meta', backgroundColor: COR_PROG, borderRadius: 4,
            datalabels: { anchor: 'end', align: 'end', offset: 2, color: '#64748b', font: { size: 10, weight: '600' },
              formatter: v => v == null ? '' : nf(v) } },
          { label: 'Executado Backlog', data: vBack, stack: 'exec', backgroundColor: COR_EXEC,
            datalabels: dentro('#fff') },
          { label: 'Executado Entrante', data: vEnt, stack: 'exec', backgroundColor: COR_ENTRANTE,
            borderRadius: { topLeft: 4, topRight: 4 }, borderSkipped: false,
            datalabels: dentro('#fff') }
        ]
      },
      plugins: [totalExec],
      options: {
        layout: { padding: { top: 22 } },
        datasets: { bar: { categoryPercentage: 0.62, barPercentage: 0.92 } },
        scales: {
          x: { stacked: true, grid: { display: false }, border: { display: false },
            ticks: { maxRotation: 0, autoSkip: false, color: '#475569' } },
          y: { stacked: true, beginAtZero: true, grid: { display: false }, border: { display: false },
            ticks: { color: '#94a3b8' }, title: { display: true, text: 'Qtde AFAZ', color: '#94a3b8' } }
        },
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'rectRounded', boxWidth: 10, boxHeight: 10, padding: 16 } },
          tooltip: { callbacks: {
            title: its => rotMes(meses[its[0].dataIndex]),
            footer: its => {
              const k = its[0].dataIndex;
              const tot = vTot[k] == null ? '' : `Total executado: ${nf(vTot[k])} (Backlog ${nf(vBack[k])} + Entrante ${nf(vEnt[k])})\n`;
              return tot + (pct[k] == null ? 'Meta não informada' : `Atingimento da meta Backlog: ${pct[k]}%`);
            }
          } }
        }
      }
    });
    // KPIs do topo usam sempre o executado total (o filtro de status fechado muda só o gráfico)
    const backT = dados.execBacklog || dados.executadoBacklog || {}, entT = dados.execEntrante || dados.executadoEntrante || {};
    return { mesAtual, prog: prog[mesAtual], exe: backT[mesAtual] || 0,
      entrante: entT[mesAtual] || 0 };
  }

  // 5b. Por status e carteira: tratados (status atual) por status de fechamento, empilhado Backlog + Entrante
  function gStatusCarteira() {
    const T = dados.tratadosPorStatus;
    if (!T) {
      $('#afz-stc-sub').textContent = '⚠ Para ver este gráfico, atualize o Office Script do fluxo (gerar_dados_afaz.ts novo) e rode o fluxo de novo.';
      return;
    }
    // junta variações de escrita do mesmo status (ex.: "Pendencia Claro" e "Pendência Claro")
    const porSt = {};
    Object.keys(T).forEach(st => {
      const k = Object.keys(porSt).find(x => normSt(x) === normSt(st)) || st;
      porSt[k] = porSt[k] || {};
      Object.keys(T[st]).forEach(tp => { porSt[k][tp] = (porSt[k][tp] || 0) + Number(T[st][tp] || 0); });
    });
    const tipos = ['Backlog', 'Entrante'].filter(t => !filtro.tp || filtro.tp === t);
    const total = st => tipos.reduce((s, t) => s + (porSt[st][t] || 0), 0);
    const status = Object.keys(porSt).sort((a, b) => total(b) - total(a));
    desenhar('afz-g-stc', {
      type: 'bar',
      data: {
        labels: status,
        datasets: tipos.map(t => ({
          label: t, data: status.map(st => porSt[st][t] || 0), backgroundColor: COR_TIPO[t], stack: 's',
          borderRadius: 3, borderSkipped: false, barPercentage: 0.75, categoryPercentage: 0.85,
          datalabels: Object.assign({}, dlDentro, { color: '#fff', font: { size: 11, weight: '700' } })
        }))
      },
      plugins: [totalNoFim()],
      options: {
        indexAxis: 'y', layout: { padding: { right: 44 } },
        scales: {
          x: { stacked: true, beginAtZero: true, display: false, grid: { display: false } },
          y: { stacked: true, grid: { display: false }, border: { display: false }, ticks: { color: '#334155', font: { size: 12 } } }
        },
        plugins: {
          legend: { position: 'top', align: 'start', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8, padding: 14 } },
          tooltip: { callbacks: { footer: its => { const st = status[its[0].dataIndex]; return `Total ${st}: ${nf(total(st))}`; } } }
        },
        onClick: (e, els, ch) => {
          if (!els.length) return;
          filtro.st = ch.data.labels[els[0].index];
          const sel = $('select[data-f="st"]');
          const op = [...sel.options].find(o => normSt(o.value) === normSt(filtro.st));
          if (op) { filtro.st = op.value; sel.value = op.value; }
          setTimeout(atualizar, 0);
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });
  }

  // 5c. Status por mês (aba Logs): Devolvido x Finalizado x Resolvido — desenhado no card do Meta x Executado, barras lado a lado
  //     filtro próprio: Todos / Backlog / Entrante
  const ST_LOG = [
    { nome: 'Devolvido', cor: '#1e3a8a', txt: '#475569' },
    { nome: 'Finalizado', cor: '#16a34a', txt: '#475569' },
    { nome: 'Resolvido', cor: '#facc15', txt: '#475569' }
  ];
  let tpLogs = '';   // '' = Backlog + Entrante
  function gLogsMes() {
    const L = dados.logsMes;
    const box = $('#afz-prog-st');
    $('#afz-prog-tit').textContent = 'Devolvido x Finalizado x Resolvido por mês';
    if (!L || !Object.keys(L).length) {
      box.innerHTML = '';
      $('#afz-prog-sub').textContent = '⚠ Para ver este gráfico, atualize o Office Script do fluxo (script_afaz.ts novo, que lê a aba Logs) e rode o fluxo de novo.';
      return;
    }
    // só de janeiro do ano atual até o mês atual
    const anoAtual = new Date().getFullYear(), mesHoje = new Date().toLocaleDateString('sv-SE').slice(0, 7);
    const meses = [];
    for (let m = 1; anoAtual + '-' + String(m).padStart(2, '0') <= mesHoje; m++) meses.push(anoAtual + '-' + String(m).padStart(2, '0'));
    const tiposDe = t => t ? [t] : ['Backlog', 'Entrante', 'Sem tipo'];
    const val = (m, st, t) => tiposDe(t).reduce((s, x) => s + Number(((L[m] || {})[st] || {})[x] || 0), 0);
    const totalTipo = t => meses.reduce((s, m) => s + ST_LOG.reduce((a, st) => a + val(m, st.nome, t), 0), 0);

    // chips Todos / Backlog / Entrante
    const chip = (rot, t) => `<button type="button" class="${tpLogs === t ? 'on' : ''}" data-tp="${t}">${rot} <b>${nf(totalTipo(t))}</b></button>`;
    box.innerHTML = chip('Todos', '') + chip('Backlog', 'Backlog') + chip('Entrante', 'Entrante');
    box.querySelectorAll('button').forEach(bt => bt.addEventListener('click', () => { tpLogs = bt.dataset.tp; gLogsMes(); }));

    const anos = [...new Set(meses.map(m => m.slice(0, 4)))];
    const total = totalTipo(tpLogs);
    $('#afz-prog-sub').textContent = `${tpLogs || 'Backlog + Entrante'} · total no período: ${nf(total)} · ` +
      'cada mudança de status registrada na aba Logs, no mês da Data';

    desenhar('afz-g-prog', {
      type: 'bar',
      data: {
        labels: meses.map(m => MES_ABREV[+m.slice(5, 7) - 1] + (anos.length > 1 ? '/' + m.slice(2, 4) : '')),
        datasets: ST_LOG.map(st => ({
          label: st.nome, backgroundColor: st.cor, data: meses.map(m => val(m, st.nome, tpLogs)),
          datalabels: { anchor: 'end', align: 'end', offset: 2, color: st.txt, font: { size: 11, weight: '600' },
            formatter: v => v ? nf(v) : '' }
        }))
      },
      options: {
        layout: { padding: { top: 22 } },
        datasets: { bar: { categoryPercentage: 0.78, barPercentage: 0.92 } },
        scales: {
          x: { grid: { display: false }, border: { display: false },
            ticks: { maxRotation: 0, autoSkip: false, color: '#0f172a', font: { size: 12, weight: '600' } },
            title: { display: anos.length === 1, text: anos[0], color: '#0f172a', font: { size: 12, weight: '600' } } },
          y: { beginAtZero: true, display: false, grid: { display: false } }
        },
        plugins: {
          legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8, padding: 14 } },
          tooltip: { callbacks: {
            footer: its => {
              const m = meses[its[0].dataIndex];
              const t = tpLogs ? '' : `\nBacklog ${nf(val(m, its[0].dataset.label, 'Backlog'))} · Entrante ${nf(val(m, its[0].dataset.label, 'Entrante'))}`;
              return t;
            }
          } }
        }
      }
    });
  }

  // 6. Visão semanal (semana ISO: segunda a domingo), a partir da W26
  //   Backlog planejado restante = Backlog em aberto no início da W26 − meta mensal (aba Programação) distribuída por dia
  //   Backlog real restante      = Backlog em aberto no início da W26 − Backlog executado desde a W26
  //   Entrantes em esteira       = Entrantes abertos desde a W26 − Entrantes executados desde a W26 (acumulado)
  const SEM_INICIAL = 26;
  const COR_PLAN = '#38bdf8', COR_REAL = '#1e3a8a', COR_ESTEIRA = '#7e22ce';
  const isoD = d => d.toISOString().slice(0, 10);
  function inicioSemanaIso(ano, sem) {        // segunda-feira da semana ISO
    const j4 = new Date(Date.UTC(ano, 0, 4));
    const seg = new Date(j4); seg.setUTCDate(j4.getUTCDate() - ((j4.getUTCDay() + 6) % 7) + (sem - 1) * 7);
    return seg;
  }
  const semanasNoAno = ano => { const d = new Date(Date.UTC(ano, 11, 28)); const t = new Date(d); t.setUTCDate(d.getUTCDate() + 3 - ((d.getUTCDay() + 6) % 7)); const j1 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return Math.ceil(((t - j1) / 86400000 + 1) / 7); };

  function gSemanal() {
    const S = dados.semanal;
    if (!S) {
      $('#afz-sem-res').innerHTML = '<div class="afz-aviso">Para ver a visão semanal, atualize o Office Script do fluxo (gerar_dados_afaz.ts novo) e rode o fluxo de novo.</div>';
      return;
    }
    const ano = Number((Object.keys(dados.programado || {}).sort()[0] || String(new Date().getFullYear())).slice(0, 4));
    const ini = inicioSemanaIso(ano, SEM_INICIAL), iniIso = isoD(ini);
    const ultSem = semanasNoAno(ano);
    const hojeIso = new Date().toLocaleDateString('sv-SE');
    const somaAte = (o, ate, de) => Object.keys(o || {}).reduce((t, d) => t + ((!de || d >= de) && d <= ate ? Number(o[d]) : 0), 0);
    // Backlog em aberto no começo da W26
    const inicio = S.bTotal - (S.bSemData || 0) - Object.keys(S.bx || {}).reduce((t, d) => t + (d < iniIso ? Number(S.bx[d]) : 0), 0);
    const prog = dados.programado || {};
    const diasMes = (a, m) => new Date(Date.UTC(a, m, 0)).getUTCDate();
    const metaDia = d => { const k = isoD(d).slice(0, 7); const v = Number(prog[k] || 0); return v ? v / diasMes(d.getUTCFullYear(), d.getUTCMonth() + 1) : 0; };

    const sem = [];
    let plan = inicio;
    for (let w = SEM_INICIAL; w <= ultSem; w++) {
      const s = inicioSemanaIso(ano, w), f = new Date(s); f.setUTCDate(s.getUTCDate() + 6);
      for (let i = 0; i < 7; i++) { const d = new Date(s); d.setUTCDate(s.getUTCDate() + i); plan -= metaDia(d); }
      const sIso = isoD(s), fIso = isoD(f), passou = sIso <= hojeIso;
      const bxW = somaAte(S.bx, fIso, sIso), eaW = somaAte(S.ea, fIso, sIso), exW = somaAte(S.ex, fIso, sIso);
      sem.push({
        w, sIso, fIso, plan: Math.max(0, Math.round(plan)),
        real: passou ? inicio - somaAte(S.bx, fIso, iniIso) : null,
        esteira: passou ? somaAte(S.ea, fIso, iniIso) - somaAte(S.ex, fIso, iniIso) : null,
        bxW, eaW, exW, atual: sIso <= hojeIso && hojeIso <= fIso
      });
    }
    const cur = sem.find(x => x.atual) || sem.filter(x => x.real != null).slice(-1)[0];
    const ddmm = iso => iso.slice(8, 10) + '/' + iso.slice(5, 7);
    if (cur) {
      const dif = cur.real - cur.plan;
      $('#afz-sem-res').innerHTML = `
        <span><i style="background:${COR_REAL}"></i><b>${nf(cur.real)}</b>Backlog real restante (W${cur.w})</span>
        <span><i style="background:${COR_PLAN}"></i><b>${nf(cur.plan)}</b>planejado</span>
        <span class="${dif > 0 ? 'ruim' : 'bom'}">${dif > 0 ? nf(dif) + ' acima do planejado' : nf(-dif) + ' abaixo do planejado'}</span>
        <span><i style="background:${COR_ESTEIRA}"></i><b>${nf(cur.esteira)}</b>entrantes em esteira</span>
        <span>Na W${cur.w}: ${nf(cur.bxW)} Backlog executados · ${nf(cur.eaW)} entrantes entraram · ${nf(cur.exW)} executados</span>`;
    }
    $('#afz-sem-sub').textContent = `Semana de segunda a domingo · W${SEM_INICIAL} começa em ${ddmm(iniIso)}/${ano} · ` +
      `Backlog em aberto no início da W${SEM_INICIAL}: ${nf(inicio)} · planejado = meta mensal da aba Programação distribuída pelos dias`;

    const rot = (cor, pos) => ({
      display: 'auto', color: cor, align: pos, anchor: 'center', offset: 4, clamp: true,
      font: { size: 10, weight: '600' }, formatter: v => v == null ? '' : nf(v)
    });
    desenhar('afz-g-sem', {
      type: 'line',
      data: {
        labels: sem.map(x => ['W' + x.w, ddmm(x.sIso)]),
        datasets: [
          { label: 'Backlog planejado restante', data: sem.map(x => x.plan), borderColor: COR_PLAN, backgroundColor: COR_PLAN,
            borderWidth: 2.5, pointRadius: 0, pointHoverRadius: 4, tension: 0, datalabels: rot('#0284c7', 'bottom') },
          { label: 'Backlog real restante', data: sem.map(x => x.real), borderColor: COR_REAL, backgroundColor: COR_REAL,
            borderWidth: 3, pointRadius: 3, pointHoverRadius: 5, tension: 0.25, datalabels: rot(COR_REAL, 'top') },
          { label: 'Entrantes em esteira (acumulado)', data: sem.map(x => x.esteira), borderColor: COR_ESTEIRA, backgroundColor: COR_ESTEIRA,
            borderWidth: 3, pointRadius: 3, pointHoverRadius: 5, tension: 0.25, datalabels: rot(COR_ESTEIRA, 'bottom') }
        ]
      },
      options: {
        layout: { padding: { top: 18, right: 12 } },
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { display: false }, border: { display: false }, ticks: { maxRotation: 0, autoSkip: false, color: '#475569', font: { size: 10 } } },
          y: { beginAtZero: true, grid: { color: '#f1f5f9' }, border: { display: false }, ticks: { color: '#94a3b8' } }
        },
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8, padding: 16 } },
          tooltip: { callbacks: {
            title: its => { const x = sem[its[0].dataIndex]; return `W${x.w} · ${ddmm(x.sIso)} a ${ddmm(x.fIso)}`; },
            footer: its => { const x = sem[its[0].dataIndex]; return x.real == null ? 'Semana futura: só o planejado' :
              `Na semana: ${nf(x.bxW)} Backlog executados\nEntrantes: ${nf(x.eaW)} entraram · ${nf(x.exW)} executados`; }
          } }
        }
      }
    });
  }

  // ---------------------------------------------------------------- KPIs / tabela
  function kpis(itens, pm) {
    const n = t => itens.filter(i => i.tp === t).length;
    const tr = dados.tratados || {};
    const trat = Object.values(tr).reduce((a, b) => a + b, 0);
    const medDias = itens.length ? Math.round(itens.reduce((s, i) => s + (i.dst || 0), 0) / itens.length) : 0;
    const ader = pm.prog ? Math.round(pm.exe / pm.prog * 100) + '%' : '—';
    const mes = rotMes(pm.mesAtual);
    const tp = filtro.tp;   // filtro de Tipo: os cartões de tratado/executado mostram só esse tipo
    const cardTratado = tp
      ? [`${tp} tratado`, nf(tr[tp] || 0), `de ${nf(trat)} no total · Finalizado, Resolvido, Pendência Claro`]
      : ['Total tratado', nf(trat), `Backlog ${nf(tr.Backlog || 0)} · Entrante ${nf(tr.Entrante || 0)}${tr['Sem tipo'] ? ' · Sem tipo ' + nf(tr['Sem tipo']) : ''}`];
    const cardsExec =
      tp === 'Backlog' ? [[`Executado Backlog ${mes}`, nf(pm.exe), pm.prog ? `de ${nf(pm.prog)} de meta · ${ader}` : 'sem meta no mês']] :
      tp === 'Entrante' ? [[`Executado Entrante ${mes}`, nf(pm.entrante), `de ${nf(pm.exe + pm.entrante)} executados no mês`]] :
      [[`Executado Backlog ${mes}`, nf(pm.exe), pm.prog ? `de ${nf(pm.prog)} de meta · ${ader}` : 'sem meta no mês'],
       [`Executado total ${mes}`, nf(pm.exe + pm.entrante), `Backlog ${nf(pm.exe)} + Entrante ${nf(pm.entrante)}`]];
    const cards = [
      ['AFAZ pendentes', nf(itens.length), `média de ${medDias} dias no status`],
      ['Backlog pendente', nf(n('Backlog')), ''],
      ['Entrante pendente', nf(n('Entrante')), ''],
      ['P0 pendentes', nf(itens.filter(i => i.pr === 'P0').length), ''],
      cardTratado,
      ...cardsExec
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

  // ---------------------------------------------------------------- aba Prioridades em aberto
  const hojeISO = () => new Date().toLocaleDateString('sv-SE');
  const situacao = i => !i.prev ? 'Sem previsão' : (i.prev < hojeISO() ? 'Atrasado' : 'No prazo');
  const COR_SIT = { 'Atrasado': '#dc2626', 'No prazo': '#16a34a', 'Sem previsão': '#94a3b8' };
  const rankPr = p => { const k = ORDEM_PR.indexOf(p); return k < 0 ? 99 : k; };
  const rankPs = k => k === 'sem' ? 99 : Number(k);

  function abaPrio(itens) {
    const temPs = dados.itens.some(i => 'ps' in i);
    const avisos = [];
    if (!temPs && !modoPi) avisos.push('⚠ O dados_afaz.json ainda não traz a prioridade do site. Atualize o Office Script do fluxo ' +
      '(gerar_dados_afaz.ts novo, que lê a aba Prioridades) e rode o fluxo de novo.');
    if (!temPi()) avisos.push('⚠ O dados_afaz.json ainda não traz a coluna PRIORIDADE INFRAESTRUTURA (Tipo Site). ' +
      'Troque o Office Script do fluxo pelo novo e rode o fluxo de novo. Enquanto isso, o quadro mostra as colunas antigas da aba Prioridades.');
    $('#afz-ps-aviso').innerHTML = avisos.map(a => `<div class="afz-aviso">${a}</div>`).join('');

    const niveis = NIVEIS_PS.filter(k => itens.some(i => psDe(i) === k));
    quadroClassificacao(itens);

    // KPIs por prioridade do site (no modo PRIORIDADE INFRAESTRUTURA os cards do quadro já fazem esse papel)
    $('#afz-kpis-ps').hidden = modoPi;
    if (!modoPi) $('#afz-kpis-ps').innerHTML = niveis.map(k => {
      const g = itens.filter(i => psDe(i) === k);
      const sites = new Set(g.map(i => i.s)).size;
      const atr = g.filter(i => situacao(i) === 'Atrasado').length;
      return `<div class="afz-kpi ps" style="border-left-color:${COR_PS[k]}"><small>${esc(rotPS(k))}</small>
        <b>${nf(g.length)}</b><span>AFAZ abertos · ${nf(sites)} sites · ${nf(atr)} atrasados</span></div>`;
    }).join('');

    // por PRIORIDADE (aba Prioridades): total de AFAZ abertos, sites e atrasados (só no modo antigo)
    if (!modoPi) {
    const maxLinha = Math.max(1, ...niveis.map(k => itens.filter(i => psDe(i) === k).length));
    $('#afz-mx-leg').innerHTML = '';
    $('#afz-mx').innerHTML = niveis.map(k => {
      const g = itens.filter(i => psDe(i) === k);
      const atr = g.filter(i => situacao(i) === 'Atrasado').length;
      const sites = new Set(g.map(i => i.s)).size;
      const pct = itens.length ? Math.round(g.length / itens.length * 100) : 0;
      return `<div class="lin${filtro.ps === k ? ' on' : ''}" data-ps="${k}">
        <div class="rot"><b>${k === 'sem' ? 'Fora da lista' : 'PRIORIDADE ' + k}</b>${k === '0' ? '<small>mais crítico</small>' : ''}</div>
        <div class="tot">${nf(g.length)}</div>
        <div><div class="trilha"><div class="barra" style="width:${Math.max(g.length / maxLinha * 100, 2)}%;background:${COR_PS[k]}"></div></div>
          <div class="txt"><span>${nf(sites)} sites</span><span>${pct}% dos AFAZ abertos</span></div></div>
        <div class="atr">${nf(atr)} atras.</div></div>`;
    }).join('');
    $('#afz-mx').querySelectorAll('.lin').forEach(l => l.addEventListener('click', () => {
      filtro.ps = filtro.ps === l.dataset.ps ? '' : l.dataset.ps;
      $(`select[data-f="ps"]`).value = filtro.ps;
      atualizar();
    }));
    }

    // faixas de dias por prioridade do site
    const medias = niveis.map(k => {
      const g = itens.filter(i => psDe(i) === k);
      return g.length ? g.reduce((s, i) => s + (i.dst || 0), 0) / g.length : 0;
    });
    desenhar('afz-g-ps', {
      type: 'bar',
      data: {
        labels: niveis.map(rotPS),
        datasets: FAIXAS.map((f, j) => ({
          label: f.nome, backgroundColor: f.cor, stack: 'p',
          datalabels: Object.assign({}, dlDentro, { color: j === 1 || j === 2 ? '#0f172a' : '#fff' }),
          data: niveis.map(k => itens.filter(i => psDe(i) === k && faixaDe(i.dst || 0) === j).length)
        }))
      },
      plugins: [totalNoFim((i, tot) => `${nf(tot)} · média ${Math.round(medias[i])} d`)],
      options: {
        indexAxis: 'y', layout: { padding: { right: 120 } },
        scales: { x: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' } }, y: { stacked: true, grid: { display: false } } },
        plugins: { legend: { position: 'bottom' } },
        onClick: (e, els, ch) => {
          if (!els.length) return;
          filtro.ps = niveis[els[0].index];
          $(`select[data-f="ps"]`).value = filtro.ps;
          setTimeout(atualizar, 0);
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });

    // tabela: site mais crítico primeiro, depois o mais parado
    const comPi = false;   // a PRIORIDADE já vem da coluna PRIORIDADE INFRAESTRUTURA: não repete em outra coluna
    const colsClasse = modoPi ? '' : ATRIB.map(a => `<th>${esc(a.nome)}</th>`).join('');
    $('#afz-ps-head').innerHTML = `<th>AFAZ</th><th>Site</th><th>PRIORIDADE</th>${colsClasse}<th>Tipo</th><th>Status</th>` +
      '<th>Dias no status</th><th>Previsão</th><th>Situação</th><th>Título</th><th>Responsável</th>';
    const lista = itens.slice().sort((a, b) =>
      (rankPs(psDe(a)) - rankPs(psDe(b))) || ((b.dst || 0) - (a.dst || 0)));
    const LIM = 500;
    $('#afz-ps-tbody').innerHTML = lista.slice(0, LIM).map(i => {
      const k = psDe(i), f = FAIXAS[faixaDe(i.dst || 0)], sit = situacao(i);
      return `<tr><td>${esc(i.id)}</td><td>${esc(i.s)}</td>
        <td><span class="afz-tag" style="background:${COR_PS[k]};color:${k === 'sem' ? '#0f172a' : '#fff'}">${k === 'sem' ? '—' : k}</span></td>
        ${modoPi ? '' : comPi ? celPi(i) : ATRIB.map(a => celAt(i[a.k])).join('')}
        <td><span class="afz-tag" style="background:${COR_TIPO[i.tp] || '#94a3b8'}">${esc(i.tp)}</span></td>
        <td>${esc(i.st)}</td><td><span class="afz-tag" style="background:${f.cor};color:${f.max === 30 || f.max === 60 ? '#0f172a' : '#fff'}">${nf(i.dst)} d</span></td>
        <td>${dataBR(i.prev)}</td><td><span class="afz-tag" style="background:${COR_SIT[sit]}">${sit}</span></td>
        <td class="ti">${esc(i.ti)}</td><td>${esc(i.resp)}</td></tr>`;
    }).join('');
    const atr = lista.filter(i => situacao(i) === 'Atrasado').length;
    $('#afz-ps-sub').textContent = `${nf(lista.length)} AFAZ em aberto · ${nf(atr)} com previsão vencida · ordenado pela PRIORIDADE do site e pelos dias parado` +
      (lista.length > LIM ? ` · mostrando os ${LIM} primeiros (use os filtros)` : '');

    // contador na aba: pendências em sites prioridade 0 e 1
    const crit = itens.filter(i => k01(psDe(i))).length;
    $('#afz-aba-cont').textContent = crit ? nf(crit) : '';
    $('#afz-aba-cont').title = 'Pendências em sites prioridade 0 e 1';
  }
  const k01 = k => k === '0' || k === '1';

  // célula da tabela: "Sim" em destaque, "Não" apagado, outros valores (PONTA, HUBS E CENTRAIS...) em texto normal
  function celAt(v) {
    const t = String(v == null ? '' : v).trim();
    if (!t) return '<td style="color:#94a3b8">—</td>';
    if (/^sim$/i.test(t)) return '<td><span class="afz-tag" style="background:#16a34a">Sim</span></td>';
    if (/^n[aã]o$/i.test(t)) return `<td style="color:#94a3b8">${esc(t)}</td>`;
    return `<td>${esc(t)}</td>`;
  }
  // célula Tipo Site: etiqueta com a mesma cor do card
  function celPi(i) {
    const v = piDe(i);
    if (v === SEM_PI) return '<td style="color:#94a3b8">—</td>';
    return `<td><span class="afz-tag" style="background:${corPi[v] || '#475569'}">${esc(v)}</span></td>`;
  }

  // quadro antigo: colunas da aba Prioridades com "Sim"
  const ehSim = v => /^sim$/i.test(String(v == null ? '' : v).trim());
  const algumSim = i => ATRIB.some(a => ehSim(i[a.k]));

  // quebra pela PRIORIDADE do site (selos 0: 7 · 1: 498 ...)
  const selosPs = g => NIVEIS_PS.filter(k => g.some(i => psDe(i) === k)).map(k =>
    `<span class="afz-tag" style="background:${COR_PS[k]};color:${k === 'sem' ? '#0f172a' : '#fff'}" title="${esc(rotPS(k))}">${k === 'sem' ? '—' : k}: ${nf(g.filter(i => psDe(i) === k).length)}</span>`).join(' ');

  // selos Backlog x Entrante (ex.: Backlog 180 · Entrante 64)
  const selosTipo = g => ORDEM_TIPO.filter(t => g.some(i => i.tp === t)).map(t =>
    `<span class="afz-tag" style="background:${COR_TIPO[t]}" title="${esc(t)}">${esc(t)}: ${nf(g.filter(i => i.tp === t).length)}</span>`).join(' ');

  // modo PRIORIDADE INFRAESTRUTURA: um card por PRIORIDADE, sempre de 0 a 4, com Backlog x Entrante
  function quadroPrioridades(itens) {
    $('#afz-cls-tit').textContent = 'AFAZ abertos por PRIORIDADE';
    $('#afz-cls-sub').textContent = 'Coluna PRIORIDADE INFRAESTRUTURA da aba Afaz (0 = mais crítico) · Backlog x Entrante em cada prioridade · clique num card para filtrar · painel v28';
    const niveis = ['0', '1', '2', '3', '4'];
    if (itens.some(i => psDe(i) === 'sem')) niveis.push('sem');
    const cards = niveis.map(k => ({ k, titulo: rotPS(k), cor: COR_PS[k], g: itens.filter(i => psDe(i) === k) }));
    cards.push({ k: '', titulo: 'Total em aberto', g: itens, total: true });

    $('#afz-cls').innerHTML = cards.map(c => {
      const sites = new Set(c.g.map(i => i.s)).size;
      const atr = c.g.filter(i => situacao(i) === 'Atrasado').length;
      const on = c.total ? !filtro.ps : filtro.ps === c.k;
      const borda = c.total ? '' : ` style="border-left-color:${c.cor}"`;
      return `<div class="afz-cls-b${c.total ? ' tot' : ''}${on ? ' on' : ''}" data-ps="${c.k}"${borda}>
        <h4>${esc(c.titulo)}${c.k === '0' ? ' · mais crítico' : ''}</h4>
        <b class="n">${nf(c.g.length)}</b><small>AFAZ abertos</small>
        <div class="det">${nf(sites)} sites · ${nf(atr)} atrasados</div>
        <div class="ps">${selosTipo(c.g) || '<span class="det">nenhum</span>'}</div></div>`;
    }).join('');

    $('#afz-cls').querySelectorAll('.afz-cls-b').forEach(cd => cd.addEventListener('click', () => {
      const k = cd.dataset.ps;
      filtro.ps = (!k || filtro.ps === k) ? '' : k;
      $('select[data-f="ps"]').value = filtro.ps;
      atualizar();
    }));
  }

  function quadroClassificacao(itens) {
    if (modoPi) return quadroPrioridades(itens);
    let cards;
    if (temPi()) {
      // um card por Tipo Site (coluna PRIORIDADE INFRAESTRUTURA), do maior para o menor
      $('#afz-cls-tit').textContent = 'AFAZ abertos por Tipo Site';
      $('#afz-cls-sub').textContent = 'Coluna PRIORIDADE INFRAESTRUTURA da aba Afaz · quebra pela PRIORIDADE do site · clique num card para filtrar a tabela · painel v28 (modo Tipo Site: valores sem 0 a 4)';
      const cont = {};
      itens.forEach(i => { const v = piDe(i); cont[v] = (cont[v] || 0) + 1; });
      // mantém o card filtrado visível mesmo se a combinação de filtros zerar
      if (filtro.cls.startsWith('pi|')) { const v = filtro.cls.slice(3); if (!(v in cont)) cont[v] = 0; }
      const tipos = Object.keys(cont).sort((a, b) =>
        (a === SEM_PI) - (b === SEM_PI) || cont[b] - cont[a] || a.localeCompare(b, 'pt-BR'));
      cards = tipos.map(v => ({
        chave: 'pi|' + v, titulo: v, cor: v === SEM_PI ? '#94a3b8' : (corPi[v] || '#475569'),
        g: itens.filter(i => piDe(i) === v)
      }));
      const comTipo = itens.filter(i => piDe(i) !== SEM_PI);
      if (tipos.filter(v => v !== SEM_PI).length > 1)
        cards.push({ chave: 'piAll|', titulo: 'Todos com Tipo Site', g: comTipo, total: true });
    } else {
      $('#afz-cls-tit').textContent = 'AFAZ abertos em sites marcados como "Sim"';
      $('#afz-cls-sub').textContent = 'Colunas da aba Prioridades · quebra pela PRIORIDADE do site · clique num card para filtrar a tabela';
      cards = ATRIB.map(a => ({ chave: a.k + '|Sim', titulo: a.nome + ' = Sim', g: itens.filter(i => ehSim(i[a.k])) }));
      cards.push({ chave: 'any|Sim', titulo: 'Pelo menos uma coluna = Sim', g: itens.filter(algumSim), total: true });
    }

    $('#afz-cls').innerHTML = cards.map(c => {
      const sites = new Set(c.g.map(i => i.s)).size;
      const atr = c.g.filter(i => situacao(i) === 'Atrasado').length;
      const porPs = selosPs(c.g);
      const borda = c.cor && !c.total ? ` style="border-left-color:${c.cor}"` : '';
      return `<div class="afz-cls-b${c.total ? ' tot' : ''}${filtro.cls === c.chave ? ' on' : ''}" data-cls="${esc(c.chave)}"${borda}>
        <h4>${esc(c.titulo)}</h4>
        <b class="n">${nf(c.g.length)}</b><small>AFAZ abertos</small>
        <div class="det">${nf(sites)} sites · ${nf(atr)} atrasados</div>
        <div class="ps">${porPs || '<span class="det">nenhum</span>'}</div></div>`;
    }).join('');

    $('#afz-cls').querySelectorAll('.afz-cls-b').forEach(cd => cd.addEventListener('click', () => {
      filtro.cls = filtro.cls === cd.dataset.cls ? '' : cd.dataset.cls;
      atualizar();
    }));
  }

  function trocarAba(aba) {
    abaAtual = aba;
    document.querySelectorAll(`#${CONTAINER} .afz-abas button`).forEach(b => b.classList.toggle('on', b.dataset.aba === aba));
    $('#afz-v-geral').hidden = aba !== 'geral';
    $('#afz-v-prio').hidden = aba !== 'prio';
    $('select[data-f="pr"]').hidden = aba === 'prio'; // P0..P3 não se aplica nesta aba
    window.afazResize();
  }

  // ---------------------------------------------------------------- ciclo
  function rotuloCls() {
    if (!filtro.cls) return '';
    const [ck, ...cv] = filtro.cls.split('|');
    if (ck === 'pi') return 'Tipo Site: ' + cv.join('|');
    if (ck === 'piAll') return 'Tipo Site: todos preenchidos';
    if (ck === 'any') return 'Pelo menos um: Sim';
    return ((ATRIB.find(a => a.k === ck) || {}).nome || ck) + ': ' + cv.join('|');
  }

  function atualizar() {
    const itens = filtrados();
    $('#afz-chip-site').innerHTML = filtro.site ? `<span class="afz-chip" title="Remover">Site: ${esc(filtro.site)} ✕</span>` : '';
    const chip = $('#afz-chip-site .afz-chip');
    if (chip) chip.onclick = () => { filtro.site = ''; atualizar(); };
    $('#afz-chip-cls').innerHTML = filtro.cls ? `<span class="afz-chip" title="Remover">${esc(rotuloCls())} ✕</span>` : '';
    const chip2 = $('#afz-chip-cls .afz-chip');
    if (chip2) chip2.onclick = () => { filtro.cls = ''; atualizar(); };
    gSites(itens);
    gTipo(itens);
    gAlarme(itens);
    gAlarmeMes(itens);
    gDias(itens);
    gEntMes();
    gPrio(itens);
    const pm = gProg();
    gStatusCarteira();
    gSemanal();
    kpis(itens, pm);
    tabela(itens);
    abaPrio(filtrados(true)); // a aba Prioridades usa só a PRIORIDADE do site
  }

  // cor fixa para cada Tipo Site (o mais frequente pega a primeira cor da paleta)
  function definirCoresPi() {
    corPi = {};
    const cont = {};
    dados.itens.forEach(i => { const v = piDe(i); if (v !== SEM_PI) cont[v] = (cont[v] || 0) + 1; });
    Object.keys(cont).sort((a, b) => cont[b] - cont[a] || a.localeCompare(b, 'pt-BR'))
      .forEach((v, k) => { corPi[v] = PALETA_PI[k % PALETA_PI.length]; });
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
    d.metaBacklog = d.metaBacklog || d.programado;
    d.executadoBacklog = d.executadoBacklog || {};
    d.executadoEntrante = d.executadoEntrante || {};
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
    definirCoresPi();
    // a coluna PRIORIDADE INFRAESTRUTURA vem como P0 ... P4: se tiver, ela vira a PRIORIDADE do painel
    modoPi = dados.itens.some(i => /[0-4]/.test(String(i.pi == null ? '' : i.pi)));
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
