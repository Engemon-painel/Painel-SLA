/* =========================================================================
   Painel-SLA · Banco de Horas — Visão Gerencial  (gestao_bh.js)
   Fonte: dados_bh.json, gerado pelo Office Script gerar_dados_bh.ts a partir
   da aba "Banco de Horas _ ANALITICO".
   Monta tudo dentro de <div id="sec-bh">. Requer Chart.js 4
   (chartjs-plugin-datalabels é opcional).
   Mostra saldo POSITIVO (idade das horas) e NEGATIVO (horas a repor) lado a lado.

   ALTERADO (18/10/2026): três abas.
     📊 Visão geral            — o painel de antes
     ⏳ Por faixa de tempo      — colaboradores por idade das horas no banco
     📈 Evolução por supervisor — saldo mês a mês (ou semana a semana) e quem
        aumentou/diminuiu. O histórico vem das versões antigas do dados_bh.json
        guardadas no GitHub (cada atualização do fluxo é uma versão), então não
        precisa mudar o Office Script.
   ========================================================================= */
(function () {
  'use strict';

  const ARQUIVO = 'dados_bh.json';
  const CONTAINER = 'sec-bh';

  // Repositório do GitHub (dono/nome). Vazio = descobre pelo endereço do painel.
  const REPO_GITHUB = '';
  const MAX_PERIODOS = 12;   // quantos meses/semanas mostrar na evolução

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
  const VERDE = '#0f9d58', VERMELHO = '#dc2626', AZUL = '#0b3a75', LARANJA = '#ea580c';
  const PALETA = ['#0b3a75', '#0f9d58', '#dc2626', '#ea8a00', '#7c3aed', '#0891b2', '#db2777', '#65a30d', '#475569', '#b45309', '#2563eb', '#be123c'];

  let dados = null, dadosBrutos = null;
  const graficos = {};
  const filtro = { co: '', sup: '', tp: '', cr: '', st: 'Ativo', busca: '' };

  // estado das abas novas
  let aba = 'geral';
  let tipoSel = '';             // '' | POSITIVO | NEGATIVO (aba por faixa)
  let faixaSel = '';            // '' = todas · 'P:0-30' / 'N:0 A -8H'
  let ultimosItens = [];
  let metrica = 'liq';          // liq | pos | neg
  let gran = 'mes';             // mes | sem
  let hist = null, histErro = '', carregandoHist = false;

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
  const pad2 = n => String(n).padStart(2, '0');
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

  // padroniza os textos de status
  const normalizarStatus = v => {
    const s = String(v || '').trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
    if (!s) return 'Não informado';
    if (['ATIVO', 'ATIVA'].includes(s)) return 'Ativo';
    if (['DEMITIDO', 'DEMITIDA', 'DESLIGADO', 'DESLIGADA'].includes(s)) return 'Demitido';
    if (s === 'FERIAS') return 'Férias';
    if (['AFASTADO', 'AFASTADA'].includes(s)) return 'Afastado';
    return String(v).trim(); // outros status (ex.: PLACA) ficam como vieram
  };

  // ----- CSS -----
  function css() {
    if ($('#bh-css')) return;
    const C = '#' + CONTAINER;
    const s = document.createElement('style');
    s.id = 'bh-css';
    s.textContent = `
${C}{--bh-borda:#e2e8f0;--bh-txt:#0f172a;--bh-sub:#64748b;color:var(--bh-txt)}
${C} [hidden]{display:none!important}
${C} .bh-topo{display:flex;flex-wrap:wrap;gap:12px 24px;align-items:center;justify-content:space-between;background:linear-gradient(100deg,#0b2545,#0b3a75);color:#fff;border-radius:14px;padding:16px 22px;margin-bottom:12px}
${C} .bh-topo h2{margin:0;font-size:24px;line-height:1.1}
${C} .bh-topo h2 small{display:block;font-size:15px;font-weight:500;opacity:.85}
${C} .bh-topo p{margin:0;font-size:13px;opacity:.8}
${C} .bh-layout{display:grid;grid-template-columns:210px minmax(0,1fr);gap:16px;align-items:start}
${C} .bh-nav{position:sticky;top:12px;background:#fff;border:1px solid var(--bh-borda);border-radius:14px;padding:12px 10px;display:flex;flex-direction:column;gap:6px}
${C} .bh-nav .rot{font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#64748b;padding:2px 8px 6px}
${C} .bh-nav button{display:flex;align-items:center;gap:10px;width:100%;text-align:left;padding:11px 12px;border:0;border-radius:10px;background:#f8fafc;font:inherit;font-size:14px;font-weight:600;color:#1b2440;cursor:pointer;line-height:1.2}
${C} .bh-nav button:hover{background:#eef3f8}
${C} .bh-nav button.on{background:#e6f2f3;color:#0e7c86}
${C} .bh-nav .ic{font-size:18px;width:22px;text-align:center;flex:0 0 22px}
${C} .bh-conteudo{min-width:0}
@media (max-width:900px){${C} .bh-layout{grid-template-columns:1fr}${C} .bh-nav{position:static;flex-direction:row;overflow-x:auto}${C} .bh-nav .rot{display:none}${C} .bh-nav button{white-space:nowrap;width:auto}}
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
${C} .bh-box.p{height:220px}
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
${C} .pos,${C} td.num.pos{color:${VERDE}!important;font-weight:700}${C} .neg,${C} td.num.neg{color:${VERMELHO}!important;font-weight:700}
${C} th{text-transform:none;letter-spacing:0;font-size:12px;font-weight:600}
${C} th.sel{background:#dbeafe;color:#0b3a75}
${C} .bh-ins{margin:10px 0 0;padding:10px 12px;background:#f8fafc;border-radius:10px;font-size:12px}
${C} .bh-ins ul{margin:4px 0 0;padding-left:18px}
${C} .bh-ins li{margin:3px 0}
${C} .bh-rod{font-size:11px;color:var(--bh-sub);margin-top:10px;text-align:right}
${C} .bh-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
${C} .bh-chip{border:1px solid #cbd5e1;background:#fff;border-radius:999px;padding:5px 12px;font:inherit;font-size:12px;cursor:pointer;color:#0f172a}
${C} .bh-chip:hover{border-color:#0b3a75}
${C} .bh-chip.on{background:#0b3a75;color:#fff;border-color:#0b3a75}
${C} .bh-chip .bh-dot{width:8px;height:8px}
${C} .bh-seg{display:inline-flex;background:#f1f5f9;border-radius:999px;padding:3px;gap:2px}
${C} .bh-seg button{border:0;background:transparent;padding:6px 12px;border-radius:999px;font:inherit;font-size:12px;cursor:pointer;color:#475569}
${C} .bh-seg button.on{background:#0b3a75;color:#fff}
${C} .bh-evtopo{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:10px}
${C} .bh-evtopo .ctl{display:flex;flex-wrap:wrap;gap:8px}
${C} .bh-aviso{padding:10px 12px;border-radius:10px;background:#fef3c7;color:#92400e;font-size:12px;margin-bottom:12px}
${C} .bh-vazio{padding:30px 10px;text-align:center;color:#64748b;font-size:13px}
${C} .bh-d{display:block;font-size:10px;font-weight:700}
${C} .bh-d.up{color:${LARANJA}}${C} .bh-d.dn{color:${VERDE}}
${C} tr.tot td{font-weight:700;background:#f8fafc}
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
    <div class="bh-layout">
    <nav class="bh-nav">
      <div class="rot">Banco de Horas</div>
      <button type="button" data-aba="geral" class="on"><span class="ic">📊</span>Visão Geral</button>
      <button type="button" data-aba="faixa"><span class="ic">👥</span>Funcionários por Faixa</button>
      <button type="button" data-aba="evol"><span class="ic">📈</span>Evolução por Supervisor</button>
    </nav>
    <div class="bh-conteudo">
    <div class="bh-filtros">
      <select data-f="co"><option value="">Coordenador: todos</option></select>
      <select data-f="sup"><option value="">Supervisor: todos</option></select>
      <select data-f="tp"><option value="">Tipo de saldo: todos</option><option value="POSITIVO">Positivo</option><option value="NEGATIVO">Negativo</option></select>
      <select data-f="cr"><option value="">Criticidade: todas</option></select>
      <select data-f="st"><option value="">Status: todos</option></select>
      <input data-f="busca" type="search" placeholder="Buscar nome ou matrícula">
      <button type="button" class="bh-limpar">Limpar filtros</button>
    </div>

    <!-- ================= ABA: VISÃO GERAL ================= -->
    <div class="bh-aba" data-p="geral">
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
    </div>

    <!-- ================= ABA: FUNCIONÁRIOS POR FAIXA ================= -->
    <div class="bh-aba" data-p="faixa" hidden>
      <div class="bh-c">
        <div class="bh-evtopo">
          <div>
            <h3>Funcionários por faixa</h3>
            <p class="sub" id="bh-fx-sub" style="margin:0"></p>
          </div>
          <div class="bh-seg" id="bh-fx-tipo"></div>
        </div>
        <div class="bh-chips" id="bh-chips"></div>
        <div class="bh-tab" style="max-height:620px;margin-top:12px"><table><thead id="bh-fx-head"></thead><tbody id="bh-fx-body"></tbody></table></div>
      </div>
    </div>

    <!-- ================= ABA: EVOLUÇÃO POR SUPERVISOR ================= -->
    <div class="bh-aba" data-p="evol" hidden>
    <div id="bh-ev-msg"></div>
    <div class="bh-g">
      <div class="bh-c s12 bh-evtopo">
        <div>
          <h3>Evolução do banco por supervisor</h3>
          <p class="sub" id="bh-ev-sub" style="margin:0">Carregando histórico…</p>
        </div>
        <div class="ctl">
          <div class="bh-seg" id="bh-met">
            <button type="button" data-m="liq" class="on">Saldo líquido</button>
            <button type="button" data-m="pos">Só positivo</button>
            <button type="button" data-m="neg">Só débito</button>
          </div>
          <div class="bh-seg" id="bh-gran">
            <button type="button" data-g="mes" class="on">Mês a mês</button>
            <button type="button" data-g="sem">Semana a semana</button>
          </div>
        </div>
      </div>
      <div class="bh-c s5">
        <h3 id="bh-var-tit">Quem aumentou ou diminuiu</h3>
        <p class="sub" id="bh-var-sub"></p>
        <div class="bh-box b" id="bh-box-var"><canvas id="bh-g-var"></canvas></div>
        <div class="bh-vazio" id="bh-var-vazio" hidden></div>
      </div>
      <div class="bh-c s7">
        <h3>Saldo por período</h3>
        <p class="sub">Clique no nome na legenda para esconder ou mostrar · use o filtro de supervisor para ver um só</p>
        <div class="bh-box b"><canvas id="bh-g-ev"></canvas></div>
      </div>
      <div class="bh-c s12">
        <h3>Tabela por período</h3>
        <p class="sub">Saldo no fim de cada período e variação em relação ao anterior · <span style="color:${LARANJA};font-weight:700">▲ laranja = banco cresceu</span> · <span style="color:${VERDE};font-weight:700">▼ verde = banco diminuiu</span></p>
        <div class="bh-tab" style="max-height:540px"><table><thead id="bh-ev-head"></thead><tbody id="bh-ev-body"></tbody></table></div>
      </div>
    </div>
    </div>

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
    // ----- filtro de status: padroniza os textos e começa mostrando só os ativos -----
    it.forEach(i => { i.st = normalizarStatus(i.st); });
    const ORDEM_ST = ['Ativo', 'Férias', 'Afastado', 'Demitido'];
    const statusDisponiveis = unicos(it, 'st').sort((a, b) =>
      (ORDEM_ST.indexOf(a) < 0 ? 99 : ORDEM_ST.indexOf(a)) - (ORDEM_ST.indexOf(b) < 0 ? 99 : ORDEM_ST.indexOf(b)) || a.localeCompare(b, 'pt-BR'));
    opc('st', statusDisponiveis);
    const statusPadrao = () => statusDisponiveis.includes('Ativo') ? 'Ativo' : '';
    filtro.st = statusPadrao();
    $('select[data-f="st"]', el).value = filtro.st;

    el.querySelectorAll('[data-f]').forEach(c => c.addEventListener(c.tagName === 'INPUT' ? 'input' : 'change', () => {
      filtro[c.dataset.f] = c.value.trim();
      atualizar();
    }));
    $('.bh-limpar', el).addEventListener('click', () => {
      Object.keys(filtro).forEach(k => filtro[k] = '');
      filtro.st = statusPadrao();
      faixaSel = ''; tipoSel = '';
      el.querySelectorAll('[data-f]').forEach(c => c.value = c.dataset.f === 'st' ? filtro.st : '');
      atualizar();
    });

    // ----- troca de aba -----
    el.querySelectorAll('.bh-nav button').forEach(b => b.addEventListener('click', () => {
      aba = b.dataset.aba;
      el.querySelectorAll('.bh-nav button').forEach(x => x.classList.toggle('on', x === b));
      el.querySelectorAll('.bh-aba').forEach(p => { p.hidden = p.dataset.p !== aba; });
      if (aba === 'evol' && !hist && !carregandoHist) abrirHistorico();
      setTimeout(() => window.bhResize(), 30);
    }));
    // ----- botões da evolução -----
    el.querySelectorAll('#bh-met button').forEach(b => b.addEventListener('click', () => {
      metrica = b.dataset.m;
      el.querySelectorAll('#bh-met button').forEach(x => x.classList.toggle('on', x === b));
      renderEvolucao();
    }));
    el.querySelectorAll('#bh-gran button').forEach(b => b.addEventListener('click', () => {
      if (gran === b.dataset.g) return;
      gran = b.dataset.g;
      el.querySelectorAll('#bh-gran button').forEach(x => x.classList.toggle('on', x === b));
      abrirHistorico();
    }));
  } // fim do layout()

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

  // =====================================================================
  // ABA: POR FAIXA DE TEMPO
  // =====================================================================
  function renderFaixas(itens) {
    ultimosItens = itens;
    const ehNeg = i => i.tp === 'NEGATIVO';
    const chaveFx = i => (ehNeg(i) ? 'N:' : 'P:') + i.fx;
    const corFx = i => ehNeg(i) ? (COR_FX_NEG[FX_NEG.indexOf(i.fx)] || '#94a3b8') : (COR_FX_POS[FX_POS.indexOf(i.fx)] || '#94a3b8');

    // botões Todos / Positivo / Negativo
    const nNeg = itens.filter(ehNeg).length, nPos = itens.length - nNeg;
    const seg = $('#bh-fx-tipo');
    seg.innerHTML = [['', 'Todos', itens.length], ['POSITIVO', '▲ Positivo', nPos], ['NEGATIVO', '▼ Negativo', nNeg]]
      .map(([v, r, q]) => `<button type="button" data-t="${v}" class="${tipoSel === v ? 'on' : ''}">${r} · ${nf(q)}</button>`).join('');
    seg.querySelectorAll('button').forEach(b => b.onclick = () => { tipoSel = b.dataset.t; faixaSel = ''; renderFaixas(ultimosItens); });

    const doTipo = itens.filter(i => !tipoSel || (tipoSel === 'NEGATIVO') === ehNeg(i));

    // botões de faixa (só as que têm gente)
    const faixas = [];
    if (tipoSel !== 'NEGATIVO') FX_POS.forEach((f, k) => faixas.push({ ch: 'P:' + f, rot: '▲ ' + ROT_POS[k], cor: COR_FX_POS[k] }));
    if (tipoSel !== 'POSITIVO') FX_NEG.forEach((f, k) => faixas.push({ ch: 'N:' + f, rot: '▼ ' + ROT_NEG[k], cor: COR_FX_NEG[k] }));
    faixas.forEach(x => { x.q = doTipo.filter(i => chaveFx(i) === x.ch).length; });
    const chips = $('#bh-chips');
    chips.innerHTML = `<button type="button" class="bh-chip${!faixaSel ? ' on' : ''}" data-k="">Todas as faixas · ${nf(doTipo.length)}</button>` +
      faixas.filter(x => x.q).map(x => `<button type="button" class="bh-chip${faixaSel === x.ch ? ' on' : ''}" data-k="${esc(x.ch)}"><span class="bh-dot" style="background:${x.cor}"></span>${esc(x.rot)} · ${nf(x.q)}</button>`).join('');
    chips.querySelectorAll('.bh-chip').forEach(b => b.onclick = () => { faixaSel = b.dataset.k; renderFaixas(ultimosItens); });

    // lista: positivos (horas mais antigas primeiro) e depois negativos (maior débito primeiro)
    const lista = doTipo.filter(i => !faixaSel || chaveFx(i) === faixaSel).sort((a, b) =>
      (ehNeg(a) - ehNeg(b)) ||
      (ehNeg(a) ? a.h - b.h : (FX_POS.indexOf(b.fx) - FX_POS.indexOf(a.fx)) || b.h - a.h));
    const somaH = lista.reduce((s, i) => s + i.h, 0);
    $('#bh-fx-sub').textContent = `${nf(lista.length)} funcionários · saldo ${hTot(somaH)} · positivo: faixa = idade das horas mais antigas · negativo: faixa = tamanho do débito`;

    const mostrarIdade = tipoSel !== 'NEGATIVO';
    $('#bh-fx-head').innerHTML = `<tr><th>Nome</th><th>Supervisor</th><th>Tipo</th><th class="num">Saldo</th><th>Faixa</th>` +
      (mostrarIdade ? ROT_POS.map(r => `<th class="num" title="Horas no banco com essa idade (só saldo positivo)">${esc(r)}</th>`).join('') : '') + '</tr>';

    $('#bh-fx-body').innerHTML = lista.map(i => {
      const neg = ehNeg(i);
      const idade = mostrarIdade ? FX_POS.map((_, k) => {
        const v = neg ? 0 : (i.bh[k] || 0);
        return v > 0 ? `<td class="num" style="background:${COR_FX_POS[k]}22">${hMin(v).replace('+', '')}</td>` : '<td class="num" style="color:#cbd5e1">—</td>';
      }).join('') : '';
      return `<tr>
        <td class="nm" title="${esc(i.n)} · ${esc(i.m)}">${esc(i.n)}<small>${esc(i.m)}</small></td>
        <td>${esc(i.sup || '—')}</td>
        <td>${neg ? tag('▼ Negativo', VERMELHO) : tag('▲ Positivo', VERDE)}</td>
        <td class="num ${neg ? 'neg' : 'pos'}">${hMin(i.h)}</td>
        <td><span class="bh-dot" style="background:${corFx(i)}"></span>${esc(rotFx(i))}</td>${idade}</tr>`;
    }).join('') || `<tr><td colspan="${mostrarIdade ? FX_POS.length + 5 : 5}">Nenhum funcionário nesse filtro</td></tr>`;
  }

  // =====================================================================
  // ABA: EVOLUÇÃO POR SUPERVISOR
  // O histórico vem das versões antigas do dados_bh.json no GitHub: para cada
  // mês (ou semana) pega a última versão gravada naquele período.
  // =====================================================================
  function repoAtual() {
    if (REPO_GITHUB) return REPO_GITHUB;
    const h = location.hostname, partes = location.pathname.split('/').filter(Boolean);
    if (!/\.github\.io$/i.test(h) || !partes.length) return '';
    return h.split('.')[0] + '/' + partes[0];
  }
  const MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  function chavePer(dataIso) {
    const d = new Date(dataIso);
    if (gran === 'mes') return d.getFullYear() + '-' + pad2(d.getMonth() + 1);
    const seg = new Date(d.getFullYear(), d.getMonth(), d.getDate() - ((d.getDay() + 6) % 7)); // segunda-feira
    return seg.getFullYear() + '-' + pad2(seg.getMonth() + 1) + '-' + pad2(seg.getDate());
  }
  function rotPer(p) {
    if (p.length === 7) return MESES[+p.slice(5, 7) - 1] + '/' + p.slice(2, 4);
    return 'sem ' + p.slice(8, 10) + '/' + p.slice(5, 7);
  }

  let commitsCache = null;
  const snapCache = {};
  async function listarCommits(repo) {
    if (commitsCache) return commitsCache;
    const CH = 'bhCommits:' + repo;
    try { const c = JSON.parse(sessionStorage.getItem(CH) || 'null'); if (c && Date.now() - c.em < 30 * 60 * 1000) return (commitsCache = c.l); } catch (e) { }
    const l = [];
    for (let p = 1; p <= 3; p++) {
      const r = await fetch(`https://api.github.com/repos/${repo}/commits?path=${encodeURIComponent(ARQUIVO)}&per_page=100&page=${p}`);
      if (!r.ok) {
        if (!l.length) throw new Error(r.status === 403 ? 'limite de consultas do GitHub atingido, tente de novo em alguns minutos' : 'o GitHub respondeu ' + r.status);
        break;
      }
      const j = await r.json();
      j.forEach(c => l.push({ sha: c.sha, d: c.commit.committer.date }));
      if (j.length < 100) break;
    }
    try { sessionStorage.setItem(CH, JSON.stringify({ em: Date.now(), l })); } catch (e) { }
    return (commitsCache = l);
  }
  async function versaoAntiga(repo, sha) {
    if (snapCache[sha]) return snapCache[sha];
    const CH = 'bhSnap:' + sha;
    try { const c = sessionStorage.getItem(CH); if (c) return (snapCache[sha] = JSON.parse(c)); } catch (e) { }
    const d = await lerJson(`https://raw.githubusercontent.com/${repo}/${sha}/${ARQUIVO}`);
    const it = itensDe(d).map(i => ({ m: String(i.m || ''), n: String(i.n || ''), sup: i.sup || '', co: i.co || '', st: normalizarStatus(i.st), h: +i.h || 0 }));
    try { sessionStorage.setItem(CH, JSON.stringify(it)); } catch (e) { }
    return (snapCache[sha] = it);
  }
  async function carregarHistorico() {
    histErro = '';
    const repo = repoAtual();
    if (!repo) {
      histErro = 'O histórico só funciona no painel publicado (endereço github.io). Aqui aparece só a foto atual.';
      hist = [{ per: chavePer(new Date().toISOString()), itens: dados.itens }];
      return;
    }
    const commits = await listarCommits(repo);              // do mais novo para o mais antigo
    const porPer = {};
    commits.forEach(c => { const k = chavePer(c.d); if (!porPer[k]) porPer[k] = c.sha; });
    const pers = Object.keys(porPer).sort().slice(-MAX_PERIODOS);
    const fotos = await Promise.all(pers.map(p => versaoAntiga(repo, porPer[p]).then(itens => ({ per: p, itens })).catch(() => null)));
    hist = fotos.filter(f => f && f.itens.length);
    if (!hist.length) {
      histErro = 'Não encontrei versões antigas do ' + ARQUIVO + ' no GitHub. Mostrando só a foto atual.';
      hist = [{ per: chavePer(new Date().toISOString()), itens: dados.itens }];
    }
  }
  function abrirHistorico() {
    carregandoHist = true;
    hist = null;
    renderEvolucao();
    carregarHistorico()
      .catch(e => {
        histErro = 'Não consegui ler o histórico: ' + e.message + '. Mostrando só a foto atual.';
        hist = [{ per: chavePer(new Date().toISOString()), itens: dados.itens }];
      })
      .finally(() => { carregandoHist = false; renderEvolucao(); });
  }

  function renderEvolucao() {
    const msg = $('#bh-ev-msg');
    if (!msg) return;
    if (!hist) {
      $('#bh-ev-sub').textContent = 'Carregando histórico do GitHub…';
      msg.innerHTML = '';
      return;
    }
    const pers = hist.map(h => h.per);
    const n = pers.length;
    const nomeMet = { liq: 'saldo líquido', pos: 'saldo positivo', neg: 'débito (horas a repor)' }[metrica];
    const val = i => metrica === 'pos' ? Math.max(i.h, 0) : metrica === 'neg' ? Math.max(-i.h, 0) : i.h;
    const fmt = v => metrica === 'neg' ? nf(Math.round(v)) + ' h' : hTot(v);
    const b = filtro.busca.toLowerCase();
    const passa = i => (!filtro.co || i.co === filtro.co) && (!filtro.sup || i.sup === filtro.sup) &&
      (!filtro.st || i.st === filtro.st) && (!b || String(i.n).toLowerCase().includes(b) || String(i.m).includes(b));

    // tabela supervisor x período
    const tab = {}, total = Array(n).fill(0);
    hist.forEach((f, k) => f.itens.filter(passa).forEach(i => {
      const s = i.sup || '(sem supervisor)';
      if (!tab[s]) tab[s] = Array(n).fill(null);
      tab[s][k] = (tab[s][k] || 0) + val(i);
      total[k] += val(i);
    }));
    const sups = Object.keys(tab);
    const delta = s => (n >= 2 && tab[s][n - 1] != null && tab[s][n - 2] != null) ? tab[s][n - 1] - tab[s][n - 2] : null;

    $('#bh-ev-sub').textContent = `${nomeMet} · ${n} ${gran === 'mes' ? (n === 1 ? 'mês' : 'meses') : (n === 1 ? 'semana' : 'semanas')} de histórico (${rotPer(pers[0])}${n > 1 ? ' a ' + rotPer(pers[n - 1]) : ''}) · filtros de tipo e criticidade não se aplicam aqui`;
    const avisos = [];
    if (histErro) avisos.push(histErro);
    if (n < 2 && !histErro) avisos.push(`O histórico está começando: só existe ${rotPer(pers[0])}. A comparação aparece quando houver ${gran === 'mes' ? 'o próximo mês' : 'a próxima semana'} — enquanto isso, experimente "Semana a semana".`);
    msg.innerHTML = avisos.map(a => `<div class="bh-aviso">⚠ ${esc(a)}</div>`).join('');

    // ---- gráfico: variação no último período ----
    const comDelta = sups.filter(s => delta(s) != null && Math.round(delta(s)) !== 0).sort((a, c) => delta(c) - delta(a));
    const box = $('#bh-box-var'), vazio = $('#bh-var-vazio');
    $('#bh-var-tit').textContent = n >= 2 ? `Quem aumentou ou diminuiu · ${rotPer(pers[n - 2])} → ${rotPer(pers[n - 1])}` : 'Quem aumentou ou diminuiu';
    $('#bh-var-sub').innerHTML = `Variação do ${nomeMet} no último período · <span style="color:${LARANJA};font-weight:700">laranja = cresceu</span> · <span style="color:${VERDE};font-weight:700">verde = diminuiu</span>`;
    if (!comDelta.length) {
      if (graficos['bh-g-var']) { graficos['bh-g-var'].destroy(); delete graficos['bh-g-var']; }
      box.hidden = true; vazio.hidden = false;
      vazio.textContent = n < 2 ? 'Precisa de pelo menos dois períodos para comparar.' : 'Nenhum supervisor mudou de saldo no último período.';
    } else {
      box.hidden = false; vazio.hidden = true;
      box.style.height = Math.max(220, comDelta.length * 28 + 50) + 'px';
      const ds = comDelta.map(delta);
      desenhar('bh-g-var', {
        type: 'bar',
        data: { labels: comDelta, datasets: [{ label: 'Variação', data: ds.map(Math.round), backgroundColor: ds.map(d => d > 0 ? LARANJA : VERDE), borderRadius: 3 }] },
        options: {
          indexAxis: 'y', layout: { padding: { left: 6, right: 60 } },
          scales: { x: { grid: { color: '#f1f5f9' }, ticks: { callback: v => nf(v) + ' h' } }, y: { grid: { display: false } } },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: c => `${ds[c.dataIndex] > 0 ? '▲ aumentou' : '▼ diminuiu'} ${nf(Math.abs(Math.round(ds[c.dataIndex])))} h`,
                footer: its => { const s = comDelta[its[0].dataIndex]; return `${rotPer(pers[n - 2])}: ${fmt(tab[s][n - 2])} → ${rotPer(pers[n - 1])}: ${fmt(tab[s][n - 1])}`; }
              }
            },
            datalabels: { anchor: 'end', align: c => ds[c.dataIndex] > 0 ? 'end' : 'start', font: { weight: '700', size: 11 }, color: c => ds[c.dataIndex] > 0 ? LARANJA : VERDE, formatter: v => (v > 0 ? '▲ +' : '▼ −') + nf(Math.abs(v)) + ' h' }
          }
        }
      });
    }

    // ---- gráfico: linha por supervisor ----
    const ultimo = s => { for (let k = n - 1; k >= 0; k--) if (tab[s][k] != null) return tab[s][k]; return 0; };
    const ordem = sups.slice().sort((a, c) => Math.abs(ultimo(c)) - Math.abs(ultimo(a)));
    desenhar('bh-g-ev', {
      type: 'line',
      data: {
        labels: pers.map(rotPer),
        datasets: ordem.map((s, k) => ({
          label: s, data: tab[s].map(v => v == null ? null : Math.round(v)),
          borderColor: PALETA[k % PALETA.length], backgroundColor: PALETA[k % PALETA.length],
          tension: .25, pointRadius: n === 1 ? 5 : 3, borderWidth: 2, spanGaps: true,
          hidden: !filtro.sup && k >= 8   // com muitos supervisores, começa mostrando os 8 maiores
        }))
      },
      options: {
        interaction: { mode: 'nearest', intersect: false },
        scales: { x: { grid: { display: false } }, y: { grid: { color: '#f1f5f9' }, ticks: { callback: v => nf(v) + ' h' } } },
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } },
          tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmt(c.raw)}` } },
          datalabels: { display: false }
        }
      }
    });

    // ---- tabela ----
    const celula = (arr, k) => {
      const v = arr[k];
      if (v == null) return '<td class="num" style="color:#cbd5e1">—</td>';
      const a = k ? arr[k - 1] : null, d = a == null ? null : v - a;
      const seta = d != null && Math.round(d) !== 0
        ? `<small class="bh-d ${d > 0 ? 'up' : 'dn'}">${d > 0 ? '▲' : '▼'} ${nf(Math.abs(Math.round(d)))} h</small>` : '';
      return `<td class="num">${fmt(v)}${seta}</td>`;
    };
    const tendencia = arr => {
      const vals = arr.map((v, k) => [v, k]).filter(x => x[0] != null);
      if (vals.length < 2) return '<td class="num" style="color:#cbd5e1">—</td>';
      const d = vals[vals.length - 1][0] - vals[0][0];
      if (Math.round(d) === 0) return '<td class="num">= estável</td>';
      return `<td class="num" style="color:${d > 0 ? LARANJA : VERDE};font-weight:700">${d > 0 ? '▲ +' : '▼ −'}${nf(Math.abs(Math.round(d)))} h</td>`;
    };
    $('#bh-ev-head').innerHTML = `<tr><th>Supervisor</th>${pers.map(p => `<th class="num">${rotPer(p)}</th>`).join('')}<th class="num">No período</th></tr>`;
    const linhas = sups.slice().sort((a, c) => (delta(c) ?? -1e9) - (delta(a) ?? -1e9) || Math.abs(ultimo(c)) - Math.abs(ultimo(a)));
    $('#bh-ev-body').innerHTML = (sups.length > 1 ? `<tr class="tot"><td>Total</td>${pers.map((_, k) => celula(total, k)).join('')}${tendencia(total)}</tr>` : '') +
      linhas.map(s => `<tr><td class="nm" title="${esc(s)}">${esc(s)}</td>${pers.map((_, k) => celula(tab[s], k)).join('')}${tendencia(tab[s])}</tr>`).join('') ||
      `<tr><td colspan="${n + 2}">Nenhum dado no filtro</td></tr>`;
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
    renderFaixas(itens);
    if (hist) renderEvolucao();
  }

  // Lê o JSON mesmo quando o Power Automate grava embrulhado: {"result":"{...}"}, {"body":...} ou texto
  async function lerJson(url) {
    const r = await fetch(url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now(), { cache: 'no-store' });
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
  // acha a lista de colaboradores dentro do JSON
  function itensDe(d) {
    if (Array.isArray(d)) return d;
    if (d && Array.isArray(d.itens)) return d.itens;
    return (d && Object.values(d).find(v => Array.isArray(v) && v.length && v[0] && typeof v[0] === 'object' && 'h' in v[0])) || [];
  }
  function diagnostico(el, nome, d) {
    const tipo = Array.isArray(d) ? 'lista com ' + d.length + ' itens' : typeof d;
    const chaves = d && typeof d === 'object' && !Array.isArray(d) ? Object.keys(d).slice(0, 15).join(', ') : '—';
    el.insertAdjacentHTML('afterbegin', `<div style="padding:12px 14px;margin-bottom:12px;border-radius:10px;background:#fee2e2;color:#991b1b;font-size:13px">
      ⚠ ${nome} foi lido, mas não tem registros no formato esperado.<br>Tipo recebido: <b>${tipo}</b> · campos no topo: <b>${chaves}</b><br>
      Mande um print desta mensagem para ajustar.</div>`);
  }

  async function iniciar() {
    let el = document.getElementById(CONTAINER);
    if (!el) {
      // index sem o <div id="sec-bh">: usa a página antiga de BH e troca o conteúdo dela
      const pag = document.getElementById('pagina-bancohoras');
      if (!pag) return;
      pag.innerHTML = `<div id="${CONTAINER}"></div>`;
      el = document.getElementById(CONTAINER);
    }
    if (typeof Chart === 'undefined') { el.innerHTML = '<p>Chart.js não carregado.</p>'; return; }
    try {
      dados = await lerJson(ARQUIVO);
      dadosBrutos = Array.isArray(dados) ? dados : Object.assign({}, dados);
    } catch (e) {
      el.innerHTML = `<p style="color:#dc2626">Não foi possível carregar ${ARQUIVO}: ${esc(e.message)}</p>`;
      return;
    }
    if (Array.isArray(dados)) dados = { itens: dados };
    if (!Array.isArray(dados.itens)) dados.itens = itensDe(dados);
    css();
    layout(el);
    atualizar();
    if (!dados.itens.length) diagnostico(el, ARQUIVO, dadosBrutos);
    if (dados.geradoEm) $('#bh-rod').textContent = 'Atualizado em ' + new Date(dados.geradoEm).toLocaleString('pt-BR');
  }

  // se a seção estiver escondida ao carregar, chame bhResize() quando ela aparecer
  window.bhResize = () => Object.values(graficos).forEach(g => g.resize());

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
