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
     📝 Pendências             — o gestor justifica as pendências de ponto
        (dados_pend_bh.json, gerado pelo script_pendencias_bh); a justificativa vai para o fluxo do Power
        Automate, que grava na tabela tbJustificativasPainel (aba JUSTIFICATIVAS_PAINEL).
     📋 Andamento por Supervisor — quanto cada supervisor já tratou das pendências.
     📈 Evolução por supervisor — saldo mês a mês (ou semana a semana) e quem
        aumentou/diminuiu.
        · Mês a mês: historico_bh.json (gerado da planilha de histórico, coluna
          "Saldo atual" de cada folha de 15 a 14) + a folha em andamento, que é
          a foto atual do dados_bh.json. O supervisor vem do dados_bh.json atual
          (cruzando pela matrícula; se não achar, pelo nome).
        · Semana a semana: versões antigas do dados_bh.json guardadas no GitHub.
   ========================================================================= */
(function () {
  'use strict';

  const ARQUIVO = 'dados_bh.json';
  const ARQ_HIST = 'historico_bh.json';   // folhas fechadas (15 a 14)
  // ----- Pendências (justificativa do gestor) -----
  const ARQ_PEND = 'dados_pend_bh.json';
  // histórico fixo do que já foi tratado na TRATATIVA_GESTORES antes do painel (não muda)
  const ARQ_PEND_HIST = 'historico_pend_bh.json';   // gerado pelo script_pendencias_bh
  // Cole aqui a URL do gatilho "Quando uma solicitação HTTP for recebida" do fluxo "BH - Gravar justificativa"
  const FLUXO_JUSTIFICATIVA = 'https://defaultc54e138a79f54294aa9bb004cf8d7f.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/05/workflows/c543ddf154d648c79b86fb6191c44377/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dtmz5AkW_myFfa3v7OG_VGqXggeIukyWjsjiKe3571U';
  // Tratativas que o gestor pode escolher em cada tipo de pendência
  // Tratativas que o gestor pode escolher (as mesmas da fórmula do Status_Automatico)
  const TRATATIVAS = {
    'HORA EXTRA': ['PAGAMENTO HE', 'BANCO DE HORAS', 'AJUSTAR PONTO', 'APENAS JUSTIFICAR'],
    'PONTO ABERTO': ['AJUSTAR PONTO', 'AUSENTE', 'APENAS JUSTIFICAR'],
    'SEM MARCACAO': ['AJUSTAR PONTO', 'AUSENTE', 'APENAS JUSTIFICAR'],
    'INTERJORNADA': ['APENAS JUSTIFICAR', 'AJUSTAR PONTO'],
    'FOLGA NAO CUMPRIDA': ['APENAS JUSTIFICAR', 'AJUSTAR PONTO', 'AUSENTE'],
    'ATRASO': ['BANCO DE HORAS', 'APENAS JUSTIFICAR'],
    'FALTA': ['AUSENTE', 'BANCO DE HORAS', 'APENAS JUSTIFICAR']   // AUSENTE = falta confirmada (RH desconta)   // BANCO DE HORAS sempre: desconta do saldo, mesmo que fique negativo
  };
  const TRAT_PADRAO = ['AJUSTAR PONTO', 'AUSENTE', 'APENAS JUSTIFICAR'];
  const TAREFAS_HE = [
    'HE01 — Hora Extra Operacional oriunda de atividade Programada',
    'HE02 — Hora Extra Operacional oriunda de chamado Emergencial',
    'HE1 — HORA EXTRA COBERTURA',
    'HE03 — ACOMPANHAMENTO TERCEIROS, visitas'
  ];
  const SEM_SUP = 'Sem supervisor (saiu do BH)';   // está no histórico mas não no dados_bh.json de hoje, então não dá para saber o supervisor
  const CONTAINER = 'sec-bh';

  // Repositório do GitHub (dono/nome). Vazio = descobre pelo endereço do painel.
  const REPO_GITHUB = '';
  const MAX_PERIODOS = 13;   // quantos meses/semanas mostrar na evolução

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
  let cmpDe = '', cmpAte = '';   // períodos comparados no gráfico de variação
  let excluidos = new Set();     // matrículas fora da evolução (demitidos e Cirion do histórico)

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
    if (['ATIVO', 'ATIVA', 'NORMAL'].includes(s)) return 'Ativo';
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
${C} .bh-cmp{display:flex;flex-wrap:wrap;align-items:center;gap:6px;font-size:12px;color:#475569;margin:0 0 10px}
${C} .bh-cmp select{padding:5px 8px;border:1px solid #cbd5e1;border-radius:8px;font:inherit;font-size:12px;background:#fff}
${C} tr.tot td{font-weight:700;background:#e2e8f0;position:sticky;bottom:0;border-top:2px solid #94a3b8}
${C} .bh-cont{margin-left:auto;background:#dc2626;color:#fff;font-size:11px;font-weight:700;border-radius:999px;padding:2px 8px}
@media (max-width:900px){${C} .bh-pd-graf{grid-template-columns:1fr!important}}
${C} .bh-pd-kpis{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px;margin-bottom:12px}
${C} .bh-pd-kpis .bh-k{border:0;text-align:left;font:inherit;cursor:pointer;opacity:.75;transition:opacity .15s,transform .15s}
${C} .bh-pd-kpis .bh-k b{font-size:26px}
${C} .bh-pd-kpis .bh-k:hover,${C} .bh-pd-kpis .bh-k.on{opacity:1}
${C} .bh-pd-kpis .bh-k.on{box-shadow:0 0 0 3px #fff,0 0 0 5px #0b3a75}
${C} .bh-sel{padding:6px 10px;border:1px solid #cbd5e1;border-radius:8px;font:inherit;font-size:12px;background:#fff}
${C} .bh-pd-det{max-width:300px;white-space:normal;min-width:200px;font-size:11.5px;color:#334155}
${C} .bh-pd-just{max-width:260px;white-space:normal;min-width:160px}
${C} .bh-pd-just small,${C} .bh-pd-env,${C} .bh-pd-err{display:block;font-size:11px;color:#64748b;white-space:normal}
${C} .bh-pd-env{color:#2563eb}${C} .bh-pd-err{color:#dc2626;font-weight:600;max-width:200px}
${C} .bh-pd-btn{border:1px solid #cbd5e1;background:#fff;border-radius:8px;padding:6px 12px;font:inherit;font-size:12px;font-weight:600;cursor:pointer;color:#0b3a75;white-space:nowrap}
${C} .bh-pd-btn:hover{border-color:#0b3a75}
${C} .bh-pd-btn.on{background:#0b3a75;color:#fff;border-color:#0b3a75}
${C} .bh-pd-btn:disabled{opacity:.6;cursor:wait}
${C} .bh-pd-lote{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:10px;padding:8px 12px;background:#eff6ff;border-radius:10px;font-size:13px}
${C} .bh-modal{position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9999;display:flex;align-items:flex-start;justify-content:center;padding:40px 14px;overflow:auto}
${C} .bh-modal-box{background:#fff;border-radius:16px;width:min(720px,100%);padding:18px 20px;box-shadow:0 20px 50px rgba(0,0,0,.35)}
${C} .bh-modal-topo{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
${C} .bh-modal-topo h3{margin:0;font-size:17px;color:#0b2545}
${C} .bh-x{border:0;background:transparent;font-size:18px;cursor:pointer;color:#64748b}
${C} .bh-modal-lista{max-height:150px;overflow:auto;background:#f8fafc;border-radius:10px;padding:8px 12px;font-size:12px;margin-bottom:12px}
${C} .bh-modal-lista div{margin:3px 0}${C} .bh-modal-lista span{color:#64748b}
${C} .bh-form{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px}
${C} .bh-form label{display:flex;flex-direction:column;gap:4px;font-size:12px;font-weight:600;color:#334155}
${C} .bh-form label.largo{grid-column:1/-1}
${C} .bh-form select,${C} .bh-form input,${C} .bh-form textarea{padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;font:inherit;font-size:13px;font-weight:400;color:#0f172a;background:#fff}
${C} .bh-form textarea{resize:vertical}
${C} .bh-prev{margin:12px 0 4px;font-size:12px;font-weight:700}
${C} .bh-modal-pe{display:flex;gap:8px;align-items:center;margin-top:10px;flex-wrap:wrap}
@media (max-width:700px){${C} .bh-form{grid-template-columns:1fr}${C} .bh-pd-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
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
      <button type="button" data-aba="pend"><span class="ic">📝</span>Pendências<span class="bh-cont" hidden></span></button>
      <button type="button" data-aba="atraso"><span class="ic">⏰</span>Atrasos<span class="bh-cont" hidden></span></button>
      <button type="button" data-aba="falta"><span class="ic">🚫</span>Faltas<span class="bh-cont" hidden></span></button>
      <button type="button" data-aba="rh"><span class="ic">🧾</span>Resposta do RH<span class="bh-cont" hidden></span></button>
      <button type="button" data-aba="andamento"><span class="ic">📋</span>Andamento por Supervisor</button>
    </nav>
    <div class="bh-conteudo">
    <div class="bh-filtros">
      <select data-f="co"><option value="">Coordenador: todos</option></select>
      <select data-f="sup"><option value="">Supervisor: todos</option></select>
      <select data-f="tp" data-nao="pend andamento atraso falta rh"><option value="">Tipo de saldo: todos</option><option value="POSITIVO">Positivo</option><option value="NEGATIVO">Negativo</option></select>
      <select data-f="cr" data-nao="pend andamento atraso falta rh"><option value="">Criticidade: todas</option></select>
      <select data-f="st" data-nao="pend andamento atraso falta rh"><option value="">Status: todos</option></select>
      <select data-pf="tipo" data-so="pend andamento" hidden><option value="">Pendência: todas</option></select>
      <select data-pf="sit" data-so="pend" hidden></select>
      <select data-pf="folha" data-so="pend andamento atraso falta rh" hidden title="Folha (dia 15 a 14)"></select>
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

    <!-- ================= ABA: PENDÊNCIAS ================= -->
    <div class="bh-aba" data-p="pend" hidden>
      <div id="bh-pd-msg"></div>
      <div class="bh-pd-kpis" id="bh-pd-kpis"></div>
      <div class="bh-pd-graf" style="display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.2fr);gap:12px;margin-bottom:12px">
        <div class="bh-c">
          <h3>Pendências por tipo</h3>
          <p class="sub" id="bh-pd-g-sub" style="margin:0 0 6px">Quantidade de cada pendência por situação · clique na barra para filtrar a tabela</p>
          <div style="position:relative;height:200px"><canvas id="bh-g-pdtipo"></canvas></div>
        </div>
        <div class="bh-c">
          <h3>Hora extra por supervisor · folha atual</h3>
          <p class="sub" id="bh-pd-he-sub" style="margin:0 0 6px"></p>
          <div id="bh-pd-he-box" style="position:relative;height:200px;overflow-y:auto"><div id="bh-pd-he-alt" style="position:relative;height:200px"><canvas id="bh-g-pdhe"></canvas></div></div>
        </div>
      </div>
      <div class="bh-c">
        <div class="bh-evtopo">
          <div>
            <h3>Pendências de ponto para o gestor justificar</h3>
            <p class="sub" id="bh-pd-sub" style="margin:0">Carregando pendências…</p>
          </div>
          <div class="ctl">
          </div>
        </div>
        <div class="bh-pd-lote" id="bh-pd-lote" hidden></div>
        <div class="bh-tab" style="max-height:640px;margin-top:10px"><table><thead id="bh-pd-head"></thead><tbody id="bh-pd-body"></tbody></table></div>
      </div>
    </div>

    <!-- ================= ABA: ANDAMENTO POR SUPERVISOR ================= -->
    <!-- ================= ABA: ATRASOS ================= -->
    <div class="bh-aba" data-p="atraso" hidden>
      <div class="bh-pd-kpis" id="bh-at-kpis"></div>
      <div class="bh-c" style="margin-bottom:12px">
        <h3>Atrasos por supervisor</h3>
        <p class="sub" id="bh-at-g-sub" style="margin:0 0 6px"></p>
        <div style="position:relative;height:260px"><canvas id="bh-g-at"></canvas></div>
      </div>
      <div class="bh-c">
        <h3>Atrasos para o gestor justificar</h3>
        <p class="sub" id="bh-at-sub" style="margin:0">Carregando…</p>
        <div class="bh-tab" style="max-height:640px;margin-top:10px"><table><thead id="bh-at-head"></thead><tbody id="bh-at-body"></tbody></table></div>
      </div>
    </div>

    <!-- ================= ABA: FALTAS ================= -->
    <div class="bh-aba" data-p="falta" hidden>
      <div class="bh-pd-kpis" id="bh-ft-kpis"></div>
      <div class="bh-pd-graf" style="display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:12px;margin-bottom:12px">
        <div class="bh-c">
          <h3>Absenteísmo por supervisor</h3>
          <p class="sub" id="bh-abs-sub" style="margin:0 0 6px"></p>
          <div style="position:relative;height:260px" id="bh-abs-alt"><canvas id="bh-g-abs"></canvas></div>
        </div>
        <div class="bh-c">
          <h3>Funcionários que mais faltaram</h3>
          <p class="sub" id="bh-absf-sub" style="margin:0 0 6px"></p>
          <div style="position:relative;height:260px" id="bh-absf-alt"><canvas id="bh-g-absf"></canvas></div>
        </div>
      </div>
      <div class="bh-c" style="margin-bottom:12px">
        <h3>Faltas por supervisor</h3>
        <p class="sub" id="bh-ft-g-sub" style="margin:0 0 6px"></p>
        <div style="position:relative;height:260px"><canvas id="bh-g-ft"></canvas></div>
      </div>
      <div class="bh-c">
        <h3>Faltas para o gestor justificar</h3>
        <p class="sub" id="bh-ft-sub" style="margin:0">Carregando…</p>
        <div class="bh-tab" style="max-height:640px;margin-top:10px"><table><thead id="bh-ft-head"></thead><tbody id="bh-ft-body"></tbody></table></div>
      </div>
    </div>

    <!-- ================= ABA: RESPOSTA DO RH ================= -->
    <div class="bh-aba" data-p="rh" hidden>
      <div class="bh-pd-kpis" id="bh-rh-kpis"></div>
      <div class="bh-c">
        <h3>Resposta do RH</h3>
        <p class="sub" id="bh-rh-sub" style="margin:0">Carregando…</p>
        <div class="bh-tab" style="max-height:640px;margin-top:10px"><table><thead id="bh-rh-head"></thead><tbody id="bh-rh-body"></tbody></table></div>
      </div>
    </div>

    <div class="bh-aba" data-p="andamento" hidden>
      <div class="bh-pd-kpis" id="bh-an-kpis"></div>
      <div class="bh-g">
        <div class="bh-c s12">
          <div class="bh-evtopo">
            <div>
              <h3>Andamento das pendências por supervisor</h3>
              <p class="sub" id="bh-an-sub" style="margin:0">Carregando…</p>
            </div>
          </div>
          <div class="bh-box b" id="bh-box-an" style="margin-top:10px"><canvas id="bh-g-an"></canvas></div>
        </div>
        <div class="bh-c s12">
          <h3>Detalhe por supervisor</h3>
          <p class="sub">% tratado = justificadas ÷ total · "Ver pendências" abre a aba Pendências já filtrada no supervisor</p>
          <div class="bh-tab" style="max-height:620px"><table><thead id="bh-an-head"></thead><tbody id="bh-an-body"></tbody></table></div>
        </div>
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
        <div class="bh-cmp">Comparar <select id="bh-cmp-de"></select> com <select id="bh-cmp-ate"></select></div>
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
    el.querySelectorAll('[data-pf]').forEach(c => c.addEventListener('change', () => {
      pendFiltro[c.dataset.pf] = c.value; pendSel.clear();
      if (pend) { renderPend(); renderAndamento(); }
    }));
    $('.bh-limpar', el).addEventListener('click', () => {
      Object.keys(filtro).forEach(k => filtro[k] = '');
      filtro.st = statusPadrao();
      faixaSel = ''; tipoSel = '';
      pendFiltro.tipo = ''; pendFiltro.sit = 'tratar'; pendFiltro.folha = null; pendFiltro.semAlmoco = false; pendSel.clear();
      el.querySelectorAll('[data-f]').forEach(c => c.value = c.dataset.f === 'st' ? filtro.st : '');
      atualizar();
    });

    // ----- troca de aba -----
    el.querySelectorAll('.bh-nav button').forEach(b => b.addEventListener('click', () => irAba(b.dataset.aba)));
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

  function irAba(nome) {
    const el = document.getElementById(CONTAINER);
    aba = nome;
    el.querySelectorAll('.bh-nav button').forEach(x => x.classList.toggle('on', x.dataset.aba === nome));
    el.querySelectorAll('.bh-aba').forEach(p => { p.hidden = p.dataset.p !== nome; });
    el.querySelectorAll('.bh-filtros [data-so]').forEach(c => { c.hidden = !c.dataset.so.split(' ').includes(nome); });
    el.querySelectorAll('.bh-filtros [data-nao]').forEach(c => { c.hidden = c.dataset.nao.split(' ').includes(nome); });
    if (nome === 'evol' && !hist && !carregandoHist) abrirHistorico();
    setTimeout(() => window.bhResize(), 30);
    window.scrollTo(0, el.getBoundingClientRect().top + window.scrollY - 10);
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
  const PARCIAIS = new Set();   // períodos ainda em andamento
  function rotPer(p) {
    if (p.length === 7) return MESES[+p.slice(5, 7) - 1] + '/' + p.slice(2, 4) + (PARCIAIS.has(p) ? ' (parcial)' : '');
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
  // ---- mês a mês: folhas fechadas do historico_bh.json + folha em andamento (dados_bh.json)
  // matrícula pode vir "02 - 005225", "005225", "5225", "02005225"... gera várias chaves para casar
  const chavesMat = m => {
    const txt = String(m == null ? '' : m), tudo = txt.replace(/\D/g, ''), fim = txt.split('-').pop().replace(/\D/g, '');
    const semZero = v => v.replace(/^0+/, '');
    return [...new Set([tudo, semZero(tudo), fim, semZero(fim), tudo.slice(-6), semZero(tudo.slice(-6)), semZero(tudo.slice(-5))].filter(v => v && v.length >= 3))];
  };
  const chaveNome = n => String(n || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
  const chaveNomeCurta = n => { const p = chaveNome(n).split(' '); return p.length > 1 ? p[0] + ' ' + p[p.length - 1] : p[0]; };
  const isoBr = iso => String(iso || '').split('-').reverse().join('/');
  let histArquivo = null, cruzamento = null;
  // só centro de custo da Claro (Descricao C.C); se o item não tiver C.C., mantém
  const ehClaro = cc => !cc || /CLARO/i.test(String(cc));
  const ccDe = i => i.cc || i.projeto || i.ccd || i.descCC || '';
  const foraDaEvolucao = i => chavesMat(i.m).some(k => excluidos.has(k));
  async function carregarFolhas() {
    if (!histArquivo) histArquivo = await lerJson(ARQ_HIST);
    const porMat = {}, porNome = {}, porCurto = {};
    dados.itens.forEach(i => {
      chavesMat(i.m).forEach(k => { if (!porMat[k]) porMat[k] = i; });
      porNome[chaveNome(i.n)] = i;
      const c = chaveNomeCurta(i.n); porCurto[c] = porCurto[c] === undefined ? i : null; // nome curto só se for único
    });
    const achar = x => {
      for (const k of chavesMat(x.m)) if (porMat[k]) return porMat[k];
      return porNome[chaveNome(x.n)] || porCurto[chaveNomeCurta(x.n)] || null;
    };
    const naoAchados = new Map();
    // quem teve DEMITIDO em qualquer folha ou C.C. Cirion no histórico não entra (nem na folha atual)
    excluidos = new Set();
    const ex = histArquivo.excluir || {};
    [].concat(ex.demitidos || [], ex.cirion || []).forEach(m => chavesMat(m).forEach(k => excluidos.add(k)));
    const folhas = (histArquivo.periodos || []).map(p => ({
      per: p.comp,
      tit: 'Folha ' + isoBr(p.ini) + ' a ' + isoBr(p.fim) + ' · saldo atual no fechamento',
      itens: (p.itens || []).filter(x => ehClaro(x.cc) && x.st !== 'DEMITIDO' && !foraDaEvolucao(x)).map(x => {
        const a = achar(x);
        if (!a) { naoAchados.set(x.m, x.n); return null; }   // não está no BH de hoje: sem supervisor, não entra
        return { m: x.m, n: x.n, sup: a.sup || '', co: a.co || '', st: normalizarStatus(x.st), h: +x.h || 0 };
      }).filter(Boolean)
    }));
    // folha em andamento (15 deste mês a 14 do próximo) = saldo de hoje
    const hoje = new Date();
    const fim = new Date(hoje.getFullYear(), hoje.getMonth() + (hoje.getDate() >= 15 ? 1 : 0), 14);
    const comp = fim.getFullYear() + '-' + pad2(fim.getMonth() + 1);
    PARCIAIS.clear();
    if (!folhas.some(f => f.per === comp)) {
      const ini = new Date(fim.getFullYear(), fim.getMonth() - 1, 15);
      PARCIAIS.add(comp);
      folhas.push({ per: comp, tit: 'Folha ' + pad2(ini.getDate()) + '/' + pad2(ini.getMonth() + 1) + ' a 14/' + pad2(fim.getMonth() + 1) + ' em andamento · saldo de hoje', itens: dados.itens });
    }
    hist = folhas.sort((a, b) => a.per.localeCompare(b.per)).slice(-MAX_PERIODOS);
    cruzamento = naoAchados.size
      ? `${naoAchados.size} funcionários do histórico não estão no BH de hoje e não entram na evolução (ex.: ${[...naoAchados].slice(0, 3).map(([m, n]) => m + ' ' + n).join(' · ')}).`
      : '';
    cmpDe = ''; cmpAte = '';
  }

  async function carregarHistorico() {
    histErro = '';
    if (gran === 'mes') return carregarFolhas();
    PARCIAIS.clear();
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
        histErro = 'Não consegui ler o histórico' + (gran === 'mes' ? ' (' + ARQ_HIST + ')' : '') + ': ' + e.message + '. Mostrando só a foto atual.';
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
    // demitidos (coluna Situação) ficam sempre fora da evolução
    const passa = i => i.st !== 'Demitido' && ehClaro(ccDe(i)) && !foraDaEvolucao(i) && (!filtro.co || i.co === filtro.co) && (!filtro.sup || i.sup === filtro.sup) &&
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
    // períodos comparados: padrão = folha em andamento x a folha fechada antes da última (ex.: set/26 → nov/26 parcial)
    let iB = pers.indexOf(cmpAte), iA = pers.indexOf(cmpDe);
    if (iB < 0) iB = n - 1;
    if (iA < 0 || iA >= iB) iA = Math.max(0, PARCIAIS.has(pers[n - 1]) && n >= 3 ? n - 3 : n - 2);
    cmpDe = pers[iA]; cmpAte = pers[iB];
    const opcoes = sel => pers.map((p, k) => `<option value="${p}"${k === sel ? ' selected' : ''}>${rotPer(p)}</option>`).join('');
    const sDe = $('#bh-cmp-de'), sAte = $('#bh-cmp-ate');
    sDe.innerHTML = opcoes(iA); sAte.innerHTML = opcoes(iB);
    sDe.onchange = () => { cmpDe = sDe.value; renderEvolucao(); };
    sAte.onchange = () => { cmpAte = sAte.value; if (pers.indexOf(cmpDe) >= pers.indexOf(cmpAte)) cmpDe = ''; renderEvolucao(); };
    const delta = s => (n >= 2 && iA !== iB && tab[s][iB] != null && tab[s][iA] != null) ? tab[s][iB] - tab[s][iA] : null;

    $('#bh-ev-sub').textContent = `${nomeMet} · ${n} ${gran === 'mes' ? (n === 1 ? 'mês' : 'meses') : (n === 1 ? 'semana' : 'semanas')} de histórico (${rotPer(pers[0])}${n > 1 ? ' a ' + rotPer(pers[n - 1]) : ''}) ${gran === 'mes' ? ' · folhas de 15 a 14, saldo no fechamento' : ''} · só C.C. Claro · sem demitidos e sem Cirion · filtros de tipo e criticidade não se aplicam aqui${tab[SEM_SUP] ? ' · "' + SEM_SUP + '" = funcionários do histórico que não estão mais no BH de hoje (geralmente desligados), sem supervisor conhecido' : ''}`;
    const avisos = [];
    if (histErro) avisos.push(histErro);
    if (gran === 'mes' && cruzamento) avisos.push(cruzamento);
    if (n < 2 && !histErro) avisos.push(`O histórico está começando: só existe ${rotPer(pers[0])}. A comparação aparece quando houver ${gran === 'mes' ? 'o próximo mês' : 'a próxima semana'} — enquanto isso, experimente "Semana a semana".`);
    msg.innerHTML = avisos.map(a => `<div class="bh-aviso">⚠ ${esc(a)}</div>`).join('');

    // ---- gráfico: variação no último período ----
    const comDelta = sups.filter(s => delta(s) != null && Math.round(delta(s)) !== 0).sort((a, c) => delta(c) - delta(a));
    const box = $('#bh-box-var'), vazio = $('#bh-var-vazio');
    $('#bh-var-tit').textContent = n >= 2 ? `Quem aumentou ou diminuiu · ${rotPer(pers[iA])} → ${rotPer(pers[iB])}` : 'Quem aumentou ou diminuiu';
    $('#bh-var-sub').innerHTML = `Variação do ${nomeMet} entre os dois períodos escolhidos · <span style="color:${LARANJA};font-weight:700">laranja = cresceu</span> · <span style="color:${VERDE};font-weight:700">verde = diminuiu</span>`;
    if (!comDelta.length) {
      if (graficos['bh-g-var']) { graficos['bh-g-var'].destroy(); delete graficos['bh-g-var']; }
      box.hidden = true; vazio.hidden = false;
      vazio.textContent = n < 2 ? 'Precisa de pelo menos dois períodos para comparar.' : 'Nenhum supervisor mudou de saldo entre esses dois períodos — escolha outra folha em "De".';
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
                footer: its => { const s = comDelta[its[0].dataIndex]; return `${rotPer(pers[iA])}: ${fmt(tab[s][iA])} → ${rotPer(pers[iB])}: ${fmt(tab[s][iB])}`; }
              }
            },
            datalabels: { anchor: 'end', align: c => ds[c.dataIndex] > 0 ? 'end' : 'start', font: { weight: '700', size: 11 }, color: c => ds[c.dataIndex] > 0 ? LARANJA : VERDE, formatter: v => (v > 0 ? '▲ +' : '▼ −') + nf(Math.abs(v)) + ' h' }
          }
        }
      });
      setTimeout(() => graficos['bh-g-var'] && graficos['bh-g-var'].resize(), 30);
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
    $('#bh-ev-head').innerHTML = `<tr><th>Supervisor</th>${hist.map(h => `<th class="num" title="${esc(h.tit || '')}">${rotPer(h.per)}</th>`).join('')}<th class="num">No período</th></tr>`;
    const linhas = sups.slice().sort((a, c) => (delta(c) ?? -1e9) - (delta(a) ?? -1e9) || Math.abs(ultimo(c)) - Math.abs(ultimo(a)));
    // "sem supervisor" vai para o fim da lista e o Total fica na última linha
    linhas.sort((a, c) => (a === SEM_SUP) - (c === SEM_SUP));
    $('#bh-ev-body').innerHTML = (linhas.map(s => `<tr${s === SEM_SUP ? ' style="color:#64748b;font-style:italic"' : ''}><td class="nm" title="${esc(s === SEM_SUP ? 'Funcionários do histórico que não estão mais no Banco de Horas de hoje (geralmente desligados) — por isso não dá para saber o supervisor' : s)}">${esc(s)}</td>${pers.map((_, k) => celula(tab[s], k)).join('')}${tendencia(tab[s])}</tr>`).join('') ||
      `<tr><td colspan="${n + 2}">Nenhum dado no filtro</td></tr>`) +
      (sups.length > 1 ? `<tr class="tot"><td>Total</td>${pers.map((_, k) => celula(total, k)).join('')}${tendencia(total)}</tr>` : '');
  }

  // =====================================================================
  // ABA: PENDÊNCIAS — o gestor justifica cada pendência de ponto
  // Leitura: dados_pend_bh.json (Office Script exportar_pendencias_bh.ts).
  // Gravação: POST para o fluxo do Power Automate (FLUXO_JUSTIFICATIVA), que roda o
  // script_pendencias_bh e grava na tbJustificativasPainel (aba JUSTIFICATIVAS_PAINEL),
  // o histórico das tratativas. Chave = Matrícula 6 dígitos | Data | Pendência.
  // =====================================================================
  // colunas copiadas pelo botão "Copiar para Excel" = colunas da tbJustificativasPainel (cole no fim da tabela)
  const COLS_J = ['Chave', 'ID_Pendencia', 'Pendencia', 'Nome', 'Matricula_6', 'Centro de Custo', 'Turno', 'Data', 'Dia',
    'Detalhe / O que verificar', 'Supervisor', 'Coordenador', 'Gerente', 'Área Técnica', 'Status_Colaborador', 'BH_ATUAL_H',
    'Horas_Pendencia_h', 'Pagamento_HE_Permitido', 'Tratativa_Gestor', 'Tarefa_HE', 'Chamado', 'Local_Atendimento',
    'Horario_Ajuste_Marcacoes', 'Justificativa_Gestor', 'Data_Resposta_Gestor', 'Status_RH', 'Observacao_RH', 'Data_Ajuste_RH',
    'Status_Automatico', 'Status Pendência', 'Origem'];
  // "Status Pendência" como na planilha (LISTAS, coluna De/para)
  const ORDEM_SIT = ['Erro na Justificativa', 'Pendente Operação', 'Justificado'];
  const COR_SIT = { 'Pendente Operação': '#f59e0b', 'Erro na Justificativa': '#dc2626', 'Justificado': '#0f9d58' };
  const SITS = [['tratar', 'A tratar'], ['Pendente Operação', 'Pendente Operação'], ['Erro na Justificativa', 'Com erro'], ['Justificado', 'Justificado'], ['', 'Todas']];
  const ehATratar = s => s === 'Pendente Operação' || s === 'Erro na Justificativa';
  const fechadaRH = (i, j) => (j.auto || calcularStatus(j, i.hePerm, i.tipo)[0]) === 'CONCLUÍDO';   // RH marcou AJUSTADO

  let pend = null, pendErro = '', pendGerado = '', pendFonte = '', pendHistInfo = '', pendFimBase = '';
  const pendSel = new Set();
  const pendFiltro = { tipo: '', sit: 'tratar', folha: null };   // folha: null = folha atual (padrão) · '' = todas
  // ALTERADO (21/10/2026): período da folha = sempre do dia 15 ao dia 14 do mês seguinte
  const folhaIniDe = iso => {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return '';
    let a = +m[1], mes = +m[2];
    if (+m[3] < 15) { mes--; if (mes < 1) { mes = 12; a--; } }
    return a + '-' + String(mes).padStart(2, '0') + '-15';
  };
  const rotFolha = ini => {
    if (!ini) return '';
    const [a, m] = ini.split('-').map(Number);
    const f = new Date(Date.UTC(a, m, 14));
    return '15/' + String(m).padStart(2, '0') + '/' + a + ' a 14/' + String(f.getUTCMonth() + 1).padStart(2, '0') + '/' + f.getUTCFullYear();
  };
  // ALTERADO (21/10/2026): saldo do dia da pendência (coluna "No dia")
  //   HORA EXTRA → o "HE: hh:mm" do Detalhe (a base calcula: batidas x turno) = horas a mais (+)
  //   demais     → horas trabalhadas (pares de batidas) − jornada do turno do dia; sem batida = deve o dia todo (−)
  //   batidas ímpares (ponto aberto) → não dá para calcular
  const DIAS_SEM = { DOM: 0, SEG: 1, TER: 2, QUA: 3, QUI: 4, SEX: 5, SAB: 6 };
  const semAc = t => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
  const minDe = t => { const m = String(t).match(/(\d{1,2}):(\d{2})/); return m ? +m[1] * 60 + +m[2] : NaN; };
  function jornadaDia(det, turno, dia) {
    const td = String(det || '').match(/Turno \(dia\):\s*([^|]+(?:\|\s*(?:DOM|SEG|TER|QUA|QUI|SEX|SAB)[^|]*)*)/i);
    const txt = semAc(td ? td[1] : turno), w = DIAS_SEM[semAc(dia).slice(0, 3)];
    const re = /(?:(DOM|SEG|TER|QUA|QUI|SEX|SAB)(?:\s*A\s*(DOM|SEG|TER|QUA|QUI|SEX|SAB))?|12X36)[^0-9]*(\d{1,2}:\d{2})\s*AS\s*(\d{1,2}:\d{2})/g;
    let m;
    while ((m = re.exec(txt))) {
      const a = m[1] ? DIAS_SEM[m[1]] : 0, b = m[1] ? (m[2] ? DIAS_SEM[m[2]] : a) : 6;
      if (w == null || (w >= a && w <= b)) {
        let j = minDe(m[4]) - minDe(m[3]); if (j < 0) j += 1440;
        return j > 360 ? j - 60 : j;          // desconta 1h de intervalo em jornada acima de 6h
      }
    }
    return null;
  }
  // ALTERADO (23/10/2026): percentual da hora extra pela escala do dia
  //   dia de trabalho da escala → 50% · sábado → 50% · domingo ou dia de folga da escala ("FOLGA TER"...) → 100%
  //   12x36 → 50% (o turno do dia não diz se era folga) · feriados não são considerados (não há lista no painel)
  function percHE(i) {
    if (i.tipo !== 'HORA EXTRA') return null;
    const det = String(i.det || '');
    const td = det.match(/Turno \(dia\):\s*([^|]+(?:\|\s*(?:FOLGA\s+)?(?:DOM|SEG|TER|QUA|QUI|SEX|SAB)[^|]*)*)/i);
    const turno = semAc(td ? td[1] : i.turno);
    let w = DIAS_SEM[semAc(i.dia).slice(0, 3)];
    if (w == null && i.d) w = new Date(String(i.d).slice(0, 10) + 'T12:00:00Z').getUTCDay();
    if (/\d+\s*X\s*\d+/.test(turno)) return { p: 50, motivo: 'escala 12x36' };
    const nomes = ['DOM', 'SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB'];
    if (w != null && new RegExp('FOLGA\\s+' + nomes[w] + '\\b').test(turno)) return { p: 100, motivo: 'dia de folga da escala (FOLGA ' + nomes[w] + ')' };
    if (w === 0 && !/\bDOM\b/.test(turno)) return { p: 100, motivo: 'domingo (fora da escala)' };
    if (w === 6) return { p: 50, motivo: 'sábado (50%)' };   // sábado é sempre 50%, mesmo fora da escala
    if (jornadaDia(det, i.turno, i.dia) == null) return { p: 100, motivo: 'dia fora da escala (' + (nomes[w] || '?') + ')' };
    return { p: 50, motivo: 'dia de trabalho da escala' };
  }
  const tagPerc = i => { const x = percHE(i); return x ? `<small title="${esc(x.motivo)}" style="display:inline-block;margin-top:2px;padding:0 6px;border-radius:999px;font-weight:700;color:#fff;background:${x.p === 100 ? '#dc2626' : '#d97706'}">HE ${x.p}%</small>` : ''; };
  const ehSemAlmoco = i => i.tipo === 'HORA EXTRA' && heSemAlmoco(i) != null;
  function heSemAlmoco(i) {
    const det = String(i.det || '');
    const mb = det.match(/Batidas:\s*([0-9:\s\/]*)/i);
    const b = mb ? (mb[1].match(/\d{1,2}:\d{2}/g) || []).map(minDe) : [];
    if (b.length !== 2) return null;
    const jor = jornadaDia(det, i.turno, i.dia);          // já sem a 1h de intervalo quando a jornada passa de 6h
    if (jor == null) return null;
    let trab = b[1] - b[0]; if (trab < 0) trab += 1440;
    if (trab <= 360) return null;                           // jornada curta: sem intervalo obrigatório
    return (trab - 60 - jor) / 60;
  }
  function saldoDia(i) {
    const det = String(i.det || '');
    if (i.tipo === 'HORA EXTRA') {
      // ALTERADO (23/10/2026): só entrada e saída (almoço não marcado) em jornada acima de 6h →
      // a base conta o almoço como HE; aqui desconta 1h de intervalo e recalcula
      const semAlmoco = heSemAlmoco(i);
      if (semAlmoco != null) return semAlmoco;
      const he = det.match(/HE:\s*(\d{1,3}):(\d{2})/i);
      if (he) return (+he[1] * 60 + +he[2]) / 60;
      return +i.horas || null;
    }
    const mb = det.match(/Batidas:\s*([0-9:\s\/]*)/i);
    if (!mb && i.tipo !== 'SEM MARCACAO') return null;   // INTERJORNADA, FOLGA NÃO CUMPRIDA...: sem batidas no detalhe
    const b = mb ? (mb[1].match(/\d{1,2}:\d{2}/g) || []).map(minDe) : [];
    if (b.length % 2) return null;
    const jor = jornadaDia(det, i.turno, i.dia);
    if (jor == null) return null;
    let trab = 0;
    for (let k = 0; k < b.length; k += 2) { let x = b[k + 1] - b[k]; if (x < 0) x += 1440; trab += x; }
    return (trab - jor) / 60;
  }
  const folhaSel = () => pendFiltro.folha == null ? folhaIniDe(hojeIso()) : pendFiltro.folha;
  // ALTERADO (21/10/2026): pendência só aparece depois de 2 dias úteis (seg a sex) da data da batida,
  // a não ser que o gestor já tenha respondido (ex.: batida sexta 16/10 → aparece a partir de quarta 21/10)
  const somarDiasUteis = (iso, n) => {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return '';
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    while (n > 0) { d.setUTCDate(d.getUTCDate() + 1); const w = d.getUTCDay(); if (w !== 0 && w !== 6) n--; }
    return d.toISOString().slice(0, 10);
  };
  const liberada = i => (i.just && i.just.trat) || hojeIso() > somarDiasUteis(i.d, 2);

  const up = v => String(v == null ? '' : v).trim().toUpperCase();
  const dataBr = iso => { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + '/' + m[2] + '/' + m[1] : String(iso || ''); };
  const hojeIso = () => { const d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); };

  // MESMA regra da fórmula do Status_Automatico da planilha (e dos Office Scripts).
  // Colunas da fórmula: B=Pendencia R=Pagamento_HE_Permitido T=Tratativa U=Tarefa_HE X=Horário Y=Justificativa AA=Status_RH
  function calcularStatus(j, hePerm, tipo) {
    const trat = up(j.trat), rh = up(j.stRH), pend = up(tipo);
    const auto =
      !trat ? 'AGUARDANDO GESTOR' :
      !String(j.texto || '').trim() ? 'INCOMPLETO - FALTA JUSTIFICATIVA' :
      rh === 'NÃO REALIZADO - DEVOLVIDO' ? 'DEVOLVIDO AO GESTOR' :
      trat === 'PAGAMENTO HE' && pend !== 'HORA EXTRA' ? 'INVÁLIDO - PAGAMENTO HE SOMENTE PARA HORA EXTRA' :
      trat === 'BANCO DE HORAS' && pend !== 'HORA EXTRA' && pend !== 'ATRASO' && pend !== 'FALTA' ? 'INVÁLIDO - BANCO DE HORAS SOMENTE PARA HORA EXTRA' :
      trat === 'PAGAMENTO HE' && up(hePerm) === 'NÃO' ? 'INVÁLIDO - BH NEGATIVO NÃO PERMITE PAGAMENTO' :
      trat === 'PAGAMENTO HE' && !String(j.tarefa || '').trim() ? 'INCOMPLETO - FALTA TAREFA HE' :
      trat === 'AJUSTAR PONTO' && (pend === 'PONTO ABERTO' || pend === 'SEM MARCACAO') && !String(j.horario || '').trim() ? 'INCOMPLETO - FALTA HORÁRIO/MARCAÇÃO' :
      ['PAGAMENTO HE', 'AJUSTAR PONTO', 'AUSENTE'].includes(trat) && rh === 'AJUSTADO' ? 'CONCLUÍDO' :
      trat === 'PAGAMENTO HE' ? 'PRONTO PARA RH - HE' :
      trat === 'AJUSTAR PONTO' ? 'AGUARDANDO RH - AJUSTE DE PONTO' :
      trat === 'AUSENTE' ? 'AGUARDANDO RH - AUSÊNCIA' :
      trat === 'BANCO DE HORAS' ? 'TRATADO - BANCO DE HORAS' :
      trat === 'APENAS JUSTIFICAR' ? 'TRATADO - JUSTIFICADO' : 'CONCLUÍDO';
    return [auto, statusPendencia(auto)];
  }
  // Status Pendência a partir do Status_Automatico
  function statusPendencia(auto) {
    if (auto === 'AGUARDANDO GESTOR') return 'Pendente Operação';
    if (/^(INCOMPLETO|INVÁLIDO|DEVOLVIDO)/.test(auto)) return 'Erro na Justificativa';
    return 'Justificado';
  }

  // justificativas enviadas por este navegador e ainda não vistas no JSON
  const CH_LOCAL = 'bhJustEnviadas';
  const locais = () => { try { return JSON.parse(localStorage.getItem(CH_LOCAL) || '{}'); } catch (e) { return {}; } };
  const salvarLocais = o => { try { localStorage.setItem(CH_LOCAL, JSON.stringify(o)); } catch (e) { } };
  function justDe(it) {
    const l = locais()[it.id];
    if (l && (!pendGerado || new Date(pendGerado) < new Date(l.em))) return Object.assign({}, l.j, { enviado: true });
    return it.just || {};
  }
  const sitDe = it => { const j = justDe(it); const s = j.status; return COR_SIT[s] ? s : calcularStatus(j, it.hePerm, it.tipo)[1]; };

  async function carregarPend() {
    try {
      const [d, h] = await Promise.all([lerJson(ARQ_PEND), lerJson(ARQ_PEND_HIST).catch(() => null)]);
      pend = Array.isArray(d) ? d : (d.itens || []);
      pendAbs = (d && Array.isArray(d.absenteismo)) ? d.absenteismo : null;
      // junta o histórico da TRATATIVA_GESTORES: mesma pendência = matrícula (6 dígitos) | data | tipo
      const chaveP = i => String(i.m || '').replace(/\D/g, '').padStart(6, '0').slice(-6) + '|' + String(i.d || '').slice(0, 10) + '|' + String(i.tipo || '').trim().toUpperCase();
      const hItens = h ? (Array.isArray(h) ? h : (h.itens || [])) : [];
      pendHistInfo = h ? `histórico: ${nf(hItens.length)} pendências` : 'histórico: arquivo ' + ARQ_PEND_HIST + ' não encontrado';
      if (hItens.length) {
        let aplicadas = 0, soHist = 0, foraPeriodo = 0, ignoradas = 0;
        const porChave = {};
        pend.forEach(i => { porChave[chaveP(i)] = i; });
        // ALTERADO (19/10/2026): última data coberta pela base de pendências atual. Pendência do histórico
        // com data DEPOIS disso não "sumiu da base" — a base é que ainda não chegou nela — então entra
        // mesmo sem justificativa (ex.: 13 e 14/10 do fechamento, que estão só na TRATATIVA_GESTORES).
        const fimBase = pend.filter(i => !i.hist).reduce((m, i) => String(i.d || '') > m ? String(i.d || '') : m, '');
        pendFimBase = fimBase;
        // força da resposta: Justificado > Erro na Justificativa > Pendente Operação
        const peso = (j, i) => { const s = j && COR_SIT[j.status] ? j.status : calcularStatus(j || {}, i.hePerm, i.tipo)[1]; return s === 'Justificado' ? 2 : s === 'Erro na Justificativa' ? 1 : 0; };
        hItens.forEach(x => {
          const atual = porChave[chaveP(x)];
          if (atual) {
            // ALTERADO (19/10/2026): usa a resposta do histórico quando o painel não tem resposta, ou quando a do
            // histórico é melhor (ex.: erro corrigido) — a não ser que o gestor tenha respondido pelo painel depois.
            const ja = atual.just || {}, jh = x.just || {};
            if (!jh.trat) return;
            const painelDepois = ja.dataResp && (!jh.dataResp || ja.dataResp > jh.dataResp);
            if (!ja.trat || (!painelDepois && (peso(jh, x) > peso(ja, atual) || (peso(jh, x) === peso(ja, atual) && jh.dataResp && jh.dataResp > (ja.dataResp || ''))))) {
              atual.just = jh; aplicadas++;
            }
          } else {
            const st = x.just && COR_SIT[x.just.status] ? x.just.status : calcularStatus(x.just || {}, x.hePerm, x.tipo)[1];
            if (st === 'Justificado') { x.hist = true; pend.push(x); porChave[chaveP(x)] = x; soHist++; }
            else if (!fimBase || String(x.d || '') > fimBase) { x.hist = true; pend.push(x); porChave[chaveP(x)] = x; foraPeriodo++; }
            // dentro do período da base e sem justificativa: sumiu da base = resolvido no ponto
            else ignoradas++;
          }
        });
        pendHistInfo += ` (${nf(aplicadas)} respostas aplicadas na base atual, ${nf(soHist)} justificadas que saíram da base` +
          (foraPeriodo ? `, ${nf(foraPeriodo)} a tratar com data depois da base atual${fimBase ? ' (' + dataBr(fimBase) + ')' : ''}` : '') +
          (ignoradas ? ', ' + nf(ignoradas) + ' sem justificativa que saíram da base (ignoradas)' : '') + ')';
      }
      pendGerado = (d && d.geradoEm) || '';
      pendFonte = d && d.arquivo ? `arquivo ${d.arquivo} · base ${d.base} · ${d.historico} respostas no histórico` : '';
      // limpa o que já voltou no JSON
      const l = locais(); let mudou = false;
      Object.keys(l).forEach(id => { if (pendGerado && new Date(pendGerado) >= new Date(l[id].em)) { delete l[id]; mudou = true; } });
      if (mudou) salvarLocais(l);
    } catch (e) {
      pend = [];
      pendErro = 'Não encontrei o ' + ARQ_PEND + ' no repositório. Ele é gerado pelo script_pendencias_bh no fluxo diário.';
    }
    renderPend();
    renderAndamento();
  }

  let pendAbs = null;   // absenteísmo por colaborador e folha (dados_pend_bh.json → absenteismo)
  const TIPOS_PROPRIOS = ['ATRASO', 'FALTA'];   // têm aba própria: não entram em Pendências/Andamento
  function pendFiltrados(soTipo) {
    const b = filtro.busca.toLowerCase(), tr = v => String(v || '').trim();
    const fo = folhaSel();
    return (pend || []).filter(i => liberada(i) && (soTipo ? i.tipo === soTipo : !TIPOS_PROPRIOS.includes(i.tipo)) &&
      (!fo || folhaIniDe(i.d) === fo) &&
      (!filtro.co || tr(i.co) === tr(filtro.co)) && (!filtro.sup || tr(i.sup) === tr(filtro.sup)) &&
      (!filtro.tp || i.tpSaldo === filtro.tp) &&
      (!b || String(i.n).toLowerCase().includes(b) || String(i.m).includes(b)));
  }

  function renderPend() {
    const msg = $('#bh-pd-msg');
    if (!msg) return;
    if (!pend) { $('#bh-pd-sub').textContent = 'Carregando pendências…'; return; }
    const avisos = [];
    if (pendErro) avisos.push(pendErro);
    msg.innerHTML = avisos.map(a => `<div class="bh-aviso">⚠ ${esc(a)}</div>`).join('');

    const base = pendFiltrados();

    // ALTERADO (23/10/2026): alerta das HE com almoço não marcado (só entrada e saída) — a base conta o almoço como HE
    const semAlm = base.filter(ehSemAlmoco);
    if (semAlm.length) {
      const hBase = semAlm.reduce((t, i) => { const m = String(i.det || '').match(/HE:\s*(\d{1,3}):(\d{2})/i); return t + (m ? +m[1] + m[2] / 60 : 0); }, 0);
      const hReal = semAlm.reduce((t, i) => t + Math.max(0, heSemAlmoco(i) || 0), 0);
      const colab = new Set(semAlm.map(i => i.m)).size;
      msg.insertAdjacentHTML('beforeend', `<div class="bh-aviso" style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;background:#fff7ed;border:1px solid #fdba74;color:#9a3412">
        <span style="font-size:20px">⚙️</span>
        <span style="flex:1;min-width:260px"><b>${nf(semAlm.length)} hora${semAlm.length > 1 ? 's' : ''} extra${semAlm.length > 1 ? 's' : ''} com erro de parametrização do almoço</b> (${nf(colab)} colaborador${colab > 1 ? 'es' : ''}) ·
          a base considerou <b>${hMin(hBase).replace('+', '')}</b> de HE, mas descontando o intervalo de almoço sobram <b>${hMin(hReal).replace('+', '')}</b>.
          Só tem entrada e saída e o sistema contou o almoço como hora extra: ajustar a parametrização do ponto (intervalo pré-assinalado).</span>
        <button type="button" class="bh-pd-btn${pendFiltro.semAlmoco ? ' on' : ''}" id="bh-pd-semalm">${pendFiltro.semAlmoco ? 'Mostrar todas' : 'Ver esses casos'}</button></div>`);
      $('#bh-pd-semalm').onclick = () => { pendFiltro.semAlmoco = !pendFiltro.semAlmoco; if (pendFiltro.semAlmoco) { pendFiltro.sit = ''; pendFiltro.tipo = ''; } pendSel.clear(); renderPend(); };
    } else pendFiltro.semAlmoco = false;
    const cont = {}; ORDEM_SIT.forEach(s => cont[s] = 0);
    base.forEach(i => { cont[sitDe(i)] = (cont[sitDe(i)] || 0) + 1; });
    const aTratar = cont['Pendente Operação'] + cont['Erro na Justificativa'];

    // contador no menu
    const nav = document.querySelector('#' + CONTAINER + ' .bh-nav button[data-aba="pend"] .bh-cont');
    if (nav) { nav.textContent = aTratar; nav.hidden = !aTratar; }

    // cards
    const cards = [['tratar', 'A tratar', aTratar, '#b45309'], ['Pendente Operação', 'Pendente Operação', cont['Pendente Operação'], COR_SIT['Pendente Operação']],
      ['Erro na Justificativa', 'Com erro', cont['Erro na Justificativa'], COR_SIT['Erro na Justificativa']],
      ['Justificado', 'Justificadas', cont['Justificado'], COR_SIT['Justificado']], ['', 'Total', base.length, '#334155']];
    $('#bh-pd-kpis').innerHTML = cards.map(c => `<button type="button" class="bh-k${pendFiltro.sit === c[0] ? ' on' : ''}" data-s="${esc(c[0])}" style="background:${c[3]}"><small>${c[1]}</small><b>${nf(c[2])}</b><span>${c[0] === 'tratar' ? 'pendente operação + com erro' : c[0] === '' ? (folhaSel() ? 'folha ' + rotFolha(folhaSel()) : 'pendências no filtro · todas as folhas') : pct(c[2], base.length) + '% do total'}</span></button>`).join('');
    $('#bh-pd-kpis').querySelectorAll('button').forEach(b => b.onclick = () => { pendFiltro.sit = b.dataset.s; pendSel.clear(); renderPend(); });

    // tipo de pendência
    const tipos = {}; base.forEach(i => tipos[i.tipo] = (tipos[i.tipo] || 0) + 1);
    gPendTipo(base, tipos);
    gHeSup();
    const raiz = document.getElementById(CONTAINER);
    const sTipo = raiz.querySelector('[data-pf="tipo"]');
    sTipo.innerHTML = `<option value="">Pendência: todas (${nf(base.length)})</option>` + Object.keys(tipos).sort().map(t => `<option value="${esc(t)}"${pendFiltro.tipo === t ? ' selected' : ''}>${esc(t)} (${nf(tipos[t])})</option>`).join('');
    const sFolha = raiz.querySelector('[data-pf="folha"]');
    if (sFolha) {
      const atualF = folhaIniDe(hojeIso());
      const folhas = [...new Set((pend || []).filter(liberada).map(i => folhaIniDe(i.d)).filter(Boolean).concat([atualF]))].sort().reverse();
      sFolha.innerHTML = `<option value=""${folhaSel() === '' ? ' selected' : ''}>Folha: todas</option>` +
        folhas.map(f => `<option value="${f}"${folhaSel() === f ? ' selected' : ''}>Folha ${rotFolha(f)}${f === atualF ? ' (atual)' : ''}</option>`).join('');
    }
    const sSit = raiz.querySelector('[data-pf="sit"]');
    sSit.innerHTML = SITS.map(([v, r]) => `<option value="${esc(v)}"${pendFiltro.sit === v ? ' selected' : ''}>Situação: ${r}</option>`).join('');

    const lista = base.filter(i => {
      const s = sitDe(i);
      return (!pendFiltro.tipo || i.tipo === pendFiltro.tipo) &&
        (!pendFiltro.sit || (pendFiltro.sit === 'tratar' ? ehATratar(s) : s === pendFiltro.sit)) &&
        (!pendFiltro.semAlmoco || ehSemAlmoco(i));
    }).sort((a, b) => ORDEM_SIT.indexOf(sitDe(a)) - ORDEM_SIT.indexOf(sitDe(b)) || (b.dias || 0) - (a.dias || 0) || String(a.n).localeCompare(String(b.n), 'pt-BR'));
    [...pendSel].forEach(id => { if (!lista.some(i => i.id === id)) pendSel.delete(id); });

    $('#bh-pd-sub').textContent = `${nf(lista.length)} pendências na lista · marque várias do mesmo tipo para justificar de uma vez` + (pendGerado ? ' · gerado em ' + new Date(pendGerado).toLocaleString('pt-BR') : '') + (pendFonte ? ' · ' + pendFonte : '') + (pendHistInfo ? ' · ' + pendHistInfo : '');

    $('#bh-pd-head').innerHTML = `<tr><th style="width:26px"><input type="checkbox" id="bh-pd-todos" title="Marcar todas"></th><th>Situação</th><th>Pendência</th><th>Colaborador</th><th>Data</th><th>Detalhe / o que verificar</th><th>Supervisor</th><th class="num">Saldo BH</th><th class="num">Horas</th><th class="num" title="Saldo do dia: hora extra feita (+) ou horas que ficou devendo (−), pelas batidas x turno">No dia</th><th>Justificativa</th><th></th></tr>`;
    const MAX = 600;
    $('#bh-pd-body').innerHTML = lista.slice(0, MAX).map(i => {
      const j = justDe(i), s = sitDe(i), auto = j.auto || calcularStatus(j, i.hePerm, i.tipo)[0];
      const fechado = fechadaRH(i, j);
      return `<tr data-id="${esc(i.id)}">
        <td>${fechado ? '' : `<input type="checkbox" class="bh-pd-ck"${pendSel.has(i.id) ? ' checked' : ''}>`}</td>
        <td>${tag(s, COR_SIT[s] || '#94a3b8')}${j.enviado ? '<small class="bh-pd-env">enviado · aguardando atualização</small>' : ''}${s === 'Erro na Justificativa' ? `<small class="bh-pd-err">${esc(auto)}</small>` : ''}</td>
        <td><b>${esc(i.tipo)}</b>${ehSemAlmoco(i) ? '<small style="display:block;margin-top:2px;padding:1px 6px;border-radius:999px;background:#fed7aa;color:#9a3412;font-weight:700;width:max-content" title="Erro de parametrização do almoço: só entrada e saída, e o sistema contou o almoço como hora extra">⚙️ Erro de parametrização do almoço</small>' : ''}${i.hist ? `<small style="display:block;color:#64748b">${pendFimBase && String(i.d || '') > pendFimBase ? 'só no histórico (TRATATIVA_GESTORES)' : 'fora da base atual'}</small>` : ''}</td>
        <td class="nm" title="${esc(i.n)} · ${esc(i.m)}">${esc(i.n)}<small>${esc(i.m)} · ${esc(i.area || '')}</small></td>
        <td>${dataBr(i.d).slice(0, 5)}<small style="display:block;color:#64748b">${esc(i.dia || '')}</small></td>
        <td class="bh-pd-det" title="${esc(i.det)}">${esc(i.det)}</td>
        <td>${esc(i.sup || '—')}</td>
        <td class="num ${i.saldo < 0 ? 'neg' : 'pos'}">${hMin(+i.saldo || 0)}<small style="display:block;font-weight:500;color:#64748b">HE paga: ${esc(i.hePerm)}</small></td>
        <td class="num">${i.horas ? hMin(+i.horas).replace('+', '') : '—'}</td>
        <td class="num ${saldoDia(i) == null ? '' : saldoDia(i) < 0 ? 'neg' : 'pos'}" title="${saldoDia(i) == null ? 'Batidas incompletas: não dá para calcular' : i.tipo === 'HORA EXTRA' && heSemAlmoco(i) != null ? 'Erro de parametrização do almoço: descontada 1h de intervalo (o sistema considerou o almoço como hora extra)' : saldoDia(i) < 0 ? 'Ficou devendo no dia' : 'Hora extra feita no dia'}">${saldoDia(i) == null ? '—' : hMin(saldoDia(i))}${i.tipo === 'HORA EXTRA' ? '<span style="display:block">' + tagPerc(i) + '</span>' : ''}${i.tipo === 'HORA EXTRA' && heSemAlmoco(i) != null ? '<small style="display:block;font-weight:500;color:#b45309">erro de parametrização do almoço</small>' : ''}<small style="display:block;font-weight:500;color:#64748b">folha ${esc(rotFolha(folhaIniDe(i.d)).replace(/\/\d{4}/g, ''))}</small></td>
        <td class="bh-pd-just">${j.trat ? `<b>${esc(j.trat)}</b>${j.tarefa ? ' · ' + esc(String(j.tarefa).split('—')[0].trim()) : ''}${j.chamado ? ' · ch. ' + esc(j.chamado) : ''}<small>${esc(j.texto || '')}</small>${j.dataResp ? `<small>respondido em ${esc(dataBr(j.dataResp))}</small>` : ''}${j.obsRH ? `<small>RH: ${esc(j.obsRH)}</small>` : ''}` : '<span style="color:#cbd5e1">—</span>'}</td>
        <td>${fechado ? '' : `<button type="button" class="bh-pd-btn">${j.trat ? 'Editar' : 'Justificar'}</button>`}</td></tr>`;
    }).join('') || '<tr><td colspan="12">Nenhuma pendência nesse filtro 🎉</td></tr>';
    if (lista.length > MAX) $('#bh-pd-body').insertAdjacentHTML('beforeend', `<tr><td colspan="12" style="color:#64748b">Mostrando ${MAX} de ${nf(lista.length)} · use os filtros para reduzir</td></tr>`);

    $('#bh-pd-body').querySelectorAll('tr[data-id]').forEach(tr => {
      const id = tr.dataset.id, ck = tr.querySelector('.bh-pd-ck'), bt = tr.querySelector('.bh-pd-btn');
      if (ck) ck.onchange = () => { ck.checked ? pendSel.add(id) : pendSel.delete(id); loteBar(lista); };
      if (bt) bt.onclick = () => abrirForm([id]);
    });
    const todos = $('#bh-pd-todos');
    todos.onchange = () => { lista.slice(0, MAX).forEach(i => { if (!fechadaRH(i, justDe(i))) todos.checked ? pendSel.add(i.id) : pendSel.delete(i.id); }); renderPend(); };
    loteBar(lista);
    renderAtraso();
  }

  // =====================================================================
  // ABA: ATRASOS (pendências tipo ATRASO: entrada mais de 5 min depois do início do turno)
  // =====================================================================
  const minDet = (det, rot) => { const m = String(det || '').match(new RegExp(rot + ':\\s*(\\d{1,2}:\\d{2})', 'i')); return m ? m[1] : '—'; };
  function renderTipo(T) {
    const P = '#bh-' + T.pre + '-';
    if (!$(P + 'kpis') || !pend) return;
    const base = pendFiltrados(T.tipo);
    const cont = {}; ORDEM_SIT.forEach(st => cont[st] = 0);
    base.forEach(i => { cont[sitDe(i)] = (cont[sitDe(i)] || 0) + 1; });
    const aTratar = cont['Pendente Operação'] + cont['Erro na Justificativa'];
    const minTot = base.reduce((t, i) => t + Math.round((+i.horas || 0) * 60), 0);
    const nav = document.querySelector('#' + CONTAINER + ' .bh-nav button[data-aba="' + T.aba + '"] .bh-cont');
    if (nav) { nav.textContent = aTratar; nav.hidden = !aTratar; }
    const colab = new Set(base.map(i => i.m)).size;
    const cards = [['A tratar', aTratar, '#b45309', 'pendente operação + com erro'],
      ['Justificados', cont['Justificado'], COR_SIT['Justificado'], pct(cont['Justificado'], base.length) + '% do total'],
      [T.nome, base.length, '#334155', `${nf(colab)} colaboradores${folhaSel() ? ' · folha ' + rotFolha(folhaSel()) : ''}`],
      ['Tempo somado', hMin(minTot / 60).replace('+', ''), '#1e3a8a', base.length ? T.media(minTot, base.length) : '']];
    $(P + 'kpis').innerHTML = cards.map(c => `<div class="bh-k" style="background:${c[2]};opacity:1;cursor:default"><small>${c[0]}</small><b>${typeof c[1] === 'number' ? nf(c[1]) : c[1]}</b><span>${c[3]}</span></div>`).join('');

    // gráfico por supervisor (a tratar x justificado)
    const porSup = {};
    base.forEach(i => { const k = String(i.sup || '—').trim(); const o = porSup[k] = porSup[k] || { t: 0, j: 0 }; sitDe(i) === 'Justificado' ? o.j++ : o.t++; });
    const sups = Object.keys(porSup).sort((a, b) => (porSup[b].t + porSup[b].j) - (porSup[a].t + porSup[a].j)).slice(0, 12);
    $(P + 'g-sub').textContent = `${nf(base.length)} ${T.nome.toLowerCase()} · ${T.regra}`;
    desenhar('bh-g-' + T.pre, {
      type: 'bar',
      data: { labels: sups, datasets: [
        { label: 'A tratar', data: sups.map(k => porSup[k].t), backgroundColor: '#f59e0b', stack: 'a', borderRadius: 3 },
        { label: 'Justificados', data: sups.map(k => porSup[k].j), backgroundColor: COR_SIT['Justificado'], stack: 'a', borderRadius: 3 }] },
      options: {
        indexAxis: 'y',
        scales: { x: { stacked: true, beginAtZero: true, ticks: { precision: 0 }, grid: { color: '#f1f5f9' } }, y: { stacked: true, grid: { display: false } } },
        plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8 } },
          datalabels: { color: '#fff', font: { weight: '700', size: 11 }, formatter: v => v || '' } }
      }
    });

    // tabela
    const lista = base.slice().sort((a, b) => ORDEM_SIT.indexOf(sitDe(a)) - ORDEM_SIT.indexOf(sitDe(b)) || (+b.horas || 0) - (+a.horas || 0));
    $(P + 'sub').textContent = `${nf(lista.length)} ${T.nome.toLowerCase()} · ${T.regraTab}`;
    $(P + 'head').innerHTML = '<tr><th>Situação</th><th>Colaborador</th><th>Data</th><th>Turno</th>' + T.cols.map(c => `<th class="num">${c[0]}</th>`).join('') + '<th>Supervisor</th><th>Justificativa</th><th></th></tr>';
    $(P + 'body').innerHTML = lista.map(i => {
      const j = justDe(i), st = sitDe(i), fechado = fechadaRH(i, j);
      return `<tr data-id="${esc(i.id)}">
        <td>${tag(st, COR_SIT[st] || '#94a3b8')}${j.enviado ? '<small class="bh-pd-env">enviado · aguardando atualização</small>' : ''}</td>
        <td class="nm" title="${esc(i.n)} · ${esc(i.m)}">${esc(i.n)}<small>${esc(i.m)} · ${esc(i.area || '')}</small></td>
        <td>${dataBr(i.d).slice(0, 5)}<small style="display:block;color:#64748b">${esc(i.dia || '')}</small></td>
        <td style="white-space:normal;max-width:220px">${esc(i.turno || '')}</td>
        ${T.cols.map(c => `<td class="num${c[2] ? ' ' + c[2] : ''}">${c[1](i)}</td>`).join('')}
        <td>${esc(i.sup || '—')}</td>
        <td class="bh-pd-just">${j.trat ? `<b>${esc(j.trat)}</b><small>${esc(j.texto || '')}</small>${j.dataResp ? `<small>respondido em ${esc(dataBr(j.dataResp))}</small>` : ''}` : '<span style="color:#cbd5e1">—</span>'}</td>
        <td>${fechado ? '' : `<button type="button" class="bh-pd-btn">${j.trat ? 'Editar' : 'Justificar'}</button>`}</td></tr>`;
    }).join('') || `<tr><td colspan="${7 + T.cols.length}">Nenhum registro nesse filtro 🎉</td></tr>`;
    $(P + 'body').querySelectorAll('tr[data-id] .bh-pd-btn').forEach(bt => bt.onclick = () => abrirForm([bt.closest('tr').dataset.id]));
  }

  const TIPOS_ABA = {
    atraso: { tipo: 'ATRASO', pre: 'at', aba: 'atraso', nome: 'Atrasos',
      media: (m, n) => 'média de ' + Math.round(m / n) + ' min por atraso',
      regra: 'entrada mais de 5 min depois do início do turno · só dias com entrada e saída marcadas',
      regraTab: 'tolerância de 5 minutos · aparecem 2 dias úteis depois da data',
      cols: [['Início', i => minDet(i.det, 'Início do turno')], ['Entrada', i => minDet(i.det, 'Entrada')], ['Atraso', i => hMin(-(+i.horas || 0)), 'neg']] },
    falta: { tipo: 'FALTA', pre: 'ft', aba: 'falta', nome: 'Faltas',
      media: (m, n) => 'média de ' + hMin(m / 60 / n).replace('+', '') + ' por falta',
      regra: 'dias marcados como "Ausente" na coluna observacao do Espelho de ponto (sem abono)',
      regraTab: 'observacao = Ausente no Espelho de ponto · aparecem 2 dias úteis depois da data',
      cols: [['Horas ausentes', i => hMin(-(+i.horas || 0)), 'neg']] }
  };
  const renderAtraso = () => { renderTipo(TIPOS_ABA.atraso); renderTipo(TIPOS_ABA.falta); gAbsenteismo(); renderRH(); };

  // ALTERADO (23/10/2026): aba RESPOSTA DO RH — pendências já respondidas pelo gestor com o que o RH preencheu
  //   na JUSTIFICATIVAS_PAINEL (Pagamento_HE_Permitido, Status_RH, Observacao_RH, Data_Ajuste_RH)
  let rhFiltro = '';   // '' = todas · 'aguardando' · 'AJUSTADO' · 'devolvido' · 'outros'
  const grupoRH = j => { const t = up(j.stRH); return !t ? 'aguardando' : t === 'AJUSTADO' ? 'AJUSTADO' : /DEVOLV/.test(t) ? 'devolvido' : 'outros'; };
  function renderRH() {
    if (!$('#bh-rh-kpis') || !pend) return;
    const b = filtro.busca.toLowerCase(), tr = v => String(v || '').trim(), fo = folhaSel();
    const base = pend.filter(i => { const j = justDe(i); return j.trat && (!fo || folhaIniDe(i.d) === fo) &&
      (!filtro.co || tr(i.co) === tr(filtro.co)) && (!filtro.sup || tr(i.sup) === tr(filtro.sup)) &&
      (!b || String(i.n).toLowerCase().includes(b) || String(i.m).includes(b)); });
    const cont = { aguardando: 0, AJUSTADO: 0, devolvido: 0, outros: 0 };
    base.forEach(i => cont[grupoRH(justDe(i))]++);
    const nav = document.querySelector('#' + CONTAINER + ' .bh-nav button[data-aba="rh"] .bh-cont');
    if (nav) { nav.textContent = cont.aguardando; nav.hidden = !cont.aguardando; }
    const cards = [['aguardando', 'Aguardando RH', cont.aguardando, '#b45309', 'gestor respondeu, RH ainda não'],
      ['AJUSTADO', 'Ajustado pelo RH', cont.AJUSTADO, COR_SIT['Justificado'], pct(cont.AJUSTADO, base.length) + '% do total'],
      ['devolvido', 'Devolvido ao gestor', cont.devolvido, VERMELHO, 'RH pediu correção'],
      ['', 'Respondidas', base.length, '#334155', folhaSel() ? 'folha ' + rotFolha(folhaSel()) : 'todas as folhas']];
    $('#bh-rh-kpis').innerHTML = cards.map(c => `<button type="button" class="bh-k${rhFiltro === c[0] ? ' on' : ''}" data-r="${c[0]}" style="background:${c[3]}"><small>${c[1]}</small><b>${nf(c[2])}</b><span>${c[4]}</span></button>`).join('');
    $('#bh-rh-kpis').querySelectorAll('button').forEach(bt => bt.onclick = () => { rhFiltro = rhFiltro === bt.dataset.r ? '' : bt.dataset.r; renderRH(); });

    const lista = base.filter(i => !rhFiltro || grupoRH(justDe(i)) === rhFiltro)
      .sort((a, c) => (justDe(c).dataResp || '').localeCompare(justDe(a).dataResp || '') || String(c.d).localeCompare(String(a.d)));
    $('#bh-rh-sub').textContent = `${nf(lista.length)} pendências respondidas pelo gestor · o RH preenche Pagamento_HE_Permitido, Status_RH, Observacao_RH e Data_Ajuste_RH na aba JUSTIFICATIVAS_PAINEL · atualiza no fluxo diário`;
    $('#bh-rh-head').innerHTML = '<tr><th>Colaborador</th><th>Pendência</th><th>Data</th><th>Supervisor</th><th>Tratativa do gestor</th><th>Respondido em</th>' +
      '<th>Pagamento HE permitido</th><th>Status RH</th><th>Observação RH</th><th>Data ajuste RH</th></tr>';
    const corRH = j => ({ aguardando: '#f59e0b', AJUSTADO: COR_SIT['Justificado'], devolvido: VERMELHO, outros: '#64748b' })[grupoRH(j)];
    const MAX = 600;
    $('#bh-rh-body').innerHTML = lista.slice(0, MAX).map(i => {
      const j = justDe(i), perm = up(i.hePerm);
      return `<tr>
        <td class="nm" title="${esc(i.n)} · ${esc(i.m)}">${esc(i.n)}<small>${esc(i.m)}</small></td>
        <td><b>${esc(i.tipo)}</b></td>
        <td>${dataBr(i.d).slice(0, 5)}<small style="display:block;color:#64748b">${esc(i.dia || '')}</small></td>
        <td>${esc(i.sup || '—')}</td>
        <td class="bh-pd-just"><b>${esc(j.trat)}</b><small>${esc(j.texto || '')}</small></td>
        <td>${j.dataResp ? esc(dataBr(j.dataResp)) : '—'}</td>
        <td>${i.tipo !== 'HORA EXTRA' ? '<span style="color:#cbd5e1">—</span>' : tag(perm || '—', perm === 'NÃO' || perm === 'NAO' ? VERMELHO : perm ? COR_SIT['Justificado'] : '#94a3b8')}</td>
        <td>${tag(j.stRH || 'Aguardando RH', corRH(j))}</td>
        <td style="white-space:normal;max-width:260px">${esc(j.obsRH || '') || '<span style="color:#cbd5e1">—</span>'}</td>
        <td>${j.dataRH ? esc(dataBr(j.dataRH)) : '—'}</td></tr>`;
    }).join('') || '<tr><td colspan="10">Nenhuma resposta nesse filtro</td></tr>';
    if (lista.length > MAX) $('#bh-rh-body').insertAdjacentHTML('beforeend', `<tr><td colspan="10" style="color:#64748b">Mostrando ${MAX} de ${nf(lista.length)} · use os filtros</td></tr>`);
  }

  // ALTERADO (22/10/2026): absenteísmo por supervisor = dias de falta ÷ dias previstos (trabalhou ou faltou)
  //   vem pronto do script (dados_pend_bh.json → absenteismo); segue Coordenador, Supervisor, busca e Folha
  function gAbsenteismo() {
    if (!$('#bh-g-abs')) return;
    if (!pendAbs) {
      if ($('#bh-absf-sub')) $('#bh-absf-sub').textContent = 'Aguardando o script novo (veja o aviso ao lado).';
      $('#bh-abs-sub').textContent = '⚠ Para ver o absenteísmo, atualize o script Pendencias no Excel (versão que lê o Espelho de ponto) e rode o fluxo diário.';
      return;
    }
    const fo = folhaSel(), b = filtro.busca.toLowerCase(), tr = v => String(v || '').trim();
    const base = pendAbs.filter(a => (!fo || a.f === fo) && (!filtro.co || tr(a.co) === tr(filtro.co)) &&
      (!filtro.sup || tr(a.sup) === tr(filtro.sup)) && (!b || String(a.n).toLowerCase().includes(b) || String(a.m).includes(b)));
    const g = {};
    base.forEach(a => {
      const k = tr(a.sup) || 'Sem supervisor';
      const o = g[k] = g[k] || { p: 0, a: 0, hp: 0, ha: 0, c: new Set(), cf: new Set() };
      o.p += a.dPrev; o.a += a.dAus; o.hp += a.hPrev; o.ha += a.hAus; o.c.add(a.m); if (a.dAus) o.cf.add(a.m);
    });
    const taxa = o => o.p ? o.a / o.p * 100 : 0;
    const sups = Object.keys(g).filter(k => g[k].p).sort((x, y) => taxa(g[y]) - taxa(g[x]));
    const tp = sups.reduce((t, k) => t + g[k].p, 0), ta = sups.reduce((t, k) => t + g[k].a, 0);
    const geral = tp ? ta / tp * 100 : 0;
    const f1 = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const tpass = base.reduce((t, a) => t + (a.dPass || 0), 0);
    $('#bh-abs-sub').textContent = `${fo ? 'Folha ' + rotFolha(fo) : 'Todas as folhas'} · geral ${f1(geral)}% (${nf(ta)} dias de falta sobre ${nf(tp)} dias de trabalho somando todos os colaboradores na folha inteira` +
      (tpass ? `, ${nf(tpass)} já passados` : '') + ') · dias da folha = os que já passaram pelo ponto + os que faltam pela escala (12x36 dia sim, dia não)';
    $('#bh-abs-alt').style.height = Math.max(220, sups.length * 26 + 60) + 'px';
    const cor = v => v >= 10 ? VERMELHO : v >= 5 ? '#f59e0b' : VERDE;
    desenhar('bh-g-abs', {
      type: 'bar',
      data: { labels: sups, datasets: [{ label: 'Absenteísmo (%)', data: sups.map(k => Math.round(taxa(g[k]) * 10) / 10),
        backgroundColor: sups.map(k => cor(taxa(g[k]))), borderRadius: 3,
        datalabels: { anchor: 'end', align: 'right', color: '#0f172a', font: { size: 11, weight: '700' },
          formatter: (v, c) => { const o = g[sups[c.dataIndex]]; return `${f1(v)}% · ${o.a} faltas · ${o.c.size} colab.`; } } }] },
      plugins: [{
        id: 'linhaGeral',
        afterDatasetsDraw(ch) {
          const x = ch.scales.x.getPixelForValue(geral), { top, bottom } = ch.chartArea, ctx = ch.ctx;
          ctx.save(); ctx.strokeStyle = '#334155'; ctx.setLineDash([5, 4]); ctx.beginPath(); ctx.moveTo(x, top); ctx.lineTo(x, bottom); ctx.stroke();
          ctx.setLineDash([]); ctx.fillStyle = '#334155'; ctx.font = '600 10px sans-serif'; ctx.fillText('geral ' + f1(geral) + '%', x + 4, top + 10); ctx.restore();
        }
      }],
      options: {
        indexAxis: 'y', layout: { padding: { right: 120 } },
        scales: { x: { beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { callback: v => v + '%' } }, y: { grid: { display: false }, ticks: { font: { size: 11 } } } },
        plugins: { legend: { display: false },
          tooltip: { callbacks: {
            label: c => `Absenteísmo: ${f1(c.raw)}%`,
            footer: its => { const o = g[sups[its[0].dataIndex]];
              const med = o.c.size ? (o.p / o.c.size).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) : '0';
              return `${o.a} dias de falta sobre ${o.p} dias de trabalho da equipe na folha\n` +
                `(${o.c.size} colaboradores × ~${med} dias cada)\n${o.cf.size} de ${o.c.size} colaboradores faltaram`; } } } },
        onClick: (e, els) => {
          if (!els.length) return;
          const k = sups[els[0].index]; if (k === 'Sem supervisor') return;
          filtro.sup = filtro.sup === k ? '' : k;
          const sel = document.querySelector('#' + CONTAINER + ' [data-f="sup"]'); if (sel) sel.value = filtro.sup;
          setTimeout(atualizar, 0);
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });
    gAbsFuncionario(base, f1);
  }

  // ALTERADO (22/10/2026): ranking dos funcionários que mais faltaram (dias de falta; desempate pelo %)
  function gAbsFuncionario(base, f1) {
    if (!$('#bh-g-absf')) return;
    const p = {};
    base.forEach(a => {   // junta as folhas quando o filtro é "todas"
      const o = p[a.m] = p[a.m] || { n: a.n, m: a.m, sup: a.sup, d: 0, a: 0, ha: 0 };
      o.d += a.dPrev; o.a += a.dAus; o.ha += a.hAus;
    });
    const todos = Object.values(p).filter(o => o.a > 0);
    const top = todos.sort((x, y) => y.a - x.a || (y.a / y.d) - (x.a / x.d) || String(x.n).localeCompare(String(y.n), 'pt-BR')).slice(0, 15);
    $('#bh-absf-sub').textContent = todos.length
      ? `${nf(todos.length)} funcionários com falta · mostrando os ${top.length} que mais faltaram · clique para buscar o funcionário`
      : 'Nenhuma falta nesse filtro 🎉';
    $('#bh-absf-alt').style.height = Math.max(220, top.length * 24 + 50) + 'px';
    const curto = n => { const w = String(n || '').split(' '); return w.length > 2 ? w[0] + ' ' + w[w.length - 1] : n; };
    desenhar('bh-g-absf', {
      type: 'bar',
      data: { labels: top.map(o => curto(o.n)), datasets: [{ label: 'Dias de falta', data: top.map(o => o.a),
        backgroundColor: top.map(o => { const t = o.a / o.d * 100; return t >= 10 ? VERMELHO : t >= 5 ? '#f59e0b' : VERDE; }), borderRadius: 3,
        datalabels: { anchor: 'end', align: 'right', color: '#0f172a', font: { size: 11, weight: '700' },
          formatter: (v, c) => { const o = top[c.dataIndex]; return `${v} falta${v > 1 ? 's' : ''} · ${f1(o.a / o.d * 100)}%`; } } }] },
      options: {
        indexAxis: 'y', layout: { padding: { right: 110 } },
        scales: { x: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: '#f1f5f9' } }, y: { grid: { display: false }, ticks: { font: { size: 11 } } } },
        plugins: { legend: { display: false },
          tooltip: { callbacks: {
            title: its => top[its[0].dataIndex].n,
            label: c => { const o = top[c.dataIndex]; return `${o.a} dias de falta sobre ${o.d} dias de trabalho na folha (${f1(o.a / o.d * 100)}%)`; },
            footer: its => { const o = top[its[0].dataIndex]; return `${hMin(o.ha).replace('+', '')} ausentes · matrícula ${o.m} · ${String(o.sup || '—').trim()}`; } } } },
        onClick: (e, els) => {
          if (!els.length) return;
          const o = top[els[0].index];
          filtro.busca = filtro.busca === o.m ? '' : o.m;
          const inp = document.querySelector('#' + CONTAINER + ' [data-f="busca"]'); if (inp) inp.value = filtro.busca;
          setTimeout(atualizar, 0);
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });
  }

  // ALTERADO (21/10/2026): gráfico por pendência (barras por tipo, empilhadas por situação)
  function gPendTipo(base, tipos) {
    const nomes = Object.keys(tipos).sort((a, b) => tipos[b] - tipos[a] || a.localeCompare(b, 'pt-BR'));
    const ROT_SIT = { 'Erro na Justificativa': 'Com erro', 'Pendente Operação': 'Pendente Operação', 'Justificado': 'Justificadas' };
    const sits = ['Pendente Operação', 'Erro na Justificativa', 'Justificado'];
    const totais = nomes.map(t => tipos[t]);
    $('#bh-pd-g-sub').textContent = `${nf(base.length)} pendências${folhaSel() ? ' na folha ' + rotFolha(folhaSel()) : ''} · por situação · clique na barra para filtrar a tabela` +
      (pendFiltro.tipo ? ` · filtrando: ${pendFiltro.tipo}` : '');
    desenhar('bh-g-pdtipo', {
      type: 'bar',
      data: {
        labels: nomes,
        datasets: sits.map(st => ({
          label: ROT_SIT[st], backgroundColor: COR_SIT[st], stack: 'p', borderRadius: 3,
          data: nomes.map(t => base.filter(i => i.tipo === t && sitDe(i) === st).length),
          datalabels: { color: '#fff', font: { size: 11, weight: '700' }, formatter: v => v ? nf(v) : '',
            display: c => { const v = c.dataset.data[c.dataIndex]; const x = c.chart.scales.x; return v && Math.abs(x.getPixelForValue(v) - x.getPixelForValue(0)) > 22; } }
        }))
      },
      plugins: [{
        id: 'totPend',
        afterDatasetsDraw(ch) {
          const { ctx } = ch, metas = ch.data.datasets.map((d, k) => ch.getDatasetMeta(k)).filter(m => !m.hidden);
          if (!metas.length) return;
          ctx.save(); ctx.font = '700 12px sans-serif'; ctx.fillStyle = '#0f172a'; ctx.textBaseline = 'middle';
          ch.data.labels.forEach((_, k) => {
            const el = metas.reduce((a, m) => (m.data[k] && m.data[k].x > (a ? a.x : -1) ? m.data[k] : a), null);
            if (el) ctx.fillText(nf(totais[k]), el.x + 6, el.y);
          });
          ctx.restore();
        }
      }],
      options: {
        indexAxis: 'y', layout: { padding: { right: 40 } },
        scales: {
          x: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { precision: 0 } },
          y: { stacked: true, grid: { display: false },
            ticks: { color: c => nomes[c.index] === pendFiltro.tipo ? '#1d4ed8' : '#334155', font: c => ({ size: 12, weight: nomes[c.index] === pendFiltro.tipo ? '700' : '500' }) } }
        },
        plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8 } } },
        onClick: (e, els) => {
          if (!els.length) return;
          const t = nomes[els[0].index];
          pendFiltro.tipo = pendFiltro.tipo === t ? '' : t; pendSel.clear();
          setTimeout(() => { renderPend(); renderAndamento(); }, 0);
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });
  }

  // ALTERADO (22/10/2026): total de hora extra feita por supervisor na folha atual (15 a 14)
  //   soma o "HE: hh:mm" de cada pendência HORA EXTRA da folha atual, de qualquer situação
  //   (respeita Coordenador, Supervisor e busca; não depende do filtro de folha nem da regra dos 2 dias úteis)
  function gHeSup() {
    if (!$('#bh-g-pdhe')) return;
    const fo = folhaIniDe(hojeIso()), b = filtro.busca.toLowerCase(), tr = v => String(v || '').trim();
    const he = (pend || []).filter(i => i.tipo === 'HORA EXTRA' && folhaIniDe(i.d) === fo &&
      (!filtro.co || tr(i.co) === tr(filtro.co)) && (!filtro.sup || tr(i.sup) === tr(filtro.sup)) &&
      (!b || String(i.n).toLowerCase().includes(b) || String(i.m).includes(b)));
    const g = {};
    he.forEach(i => {
      const k = tr(i.sup) || 'Sem supervisor', h = saldoDia(i) || 0;
      const o = g[k] = g[k] || { h: 0, q: 0, just: 0, colab: new Set(), h50: 0, h100: 0 };
      o.h += h; o.q++; o.colab.add(i.m); if (sitDe(i) === 'Justificado') o.just += h;
      const pc = percHE(i); if (pc && pc.p === 100) o.h100 += Math.max(0, h); else o.h50 += Math.max(0, h);
    });
    const sups = Object.keys(g).sort((a, c) => g[c].h - g[a].h);
    const tot = sups.reduce((t, k) => t + g[k].h, 0);
    $('#bh-pd-he-sub').textContent = `Folha ${rotFolha(fo)} · ${hMin(tot).replace('+', '')} de hora extra em ${nf(he.length)} lançamentos · verde = já justificada`;
    $('#bh-pd-he-alt').style.height = Math.max(200, sups.length * 26 + 50) + 'px';
    const aTratar = sups.map(k => Math.round((g[k].h - g[k].just) * 100) / 100), justif = sups.map(k => Math.round(g[k].just * 100) / 100);
    desenhar('bh-g-pdhe', {
      type: 'bar',
      data: { labels: sups, datasets: [
        { label: 'A tratar', data: aTratar, backgroundColor: '#f59e0b', stack: 'h', borderRadius: 3, datalabels: { display: false } },
        { label: 'Justificada', data: justif, backgroundColor: COR_SIT['Justificado'], stack: 'h', borderRadius: 3, datalabels: { display: false } }] },
      plugins: [{
        id: 'totHe',
        afterDatasetsDraw(ch) {
          const { ctx } = ch, m0 = ch.getDatasetMeta(0), m1 = ch.getDatasetMeta(1);
          ctx.save(); ctx.font = '700 11px sans-serif'; ctx.fillStyle = '#0f172a'; ctx.textBaseline = 'middle';
          sups.forEach((k, i) => {
            const el = [!m1.hidden && m1.data[i], !m0.hidden && m0.data[i]].filter(Boolean).reduce((a, e) => (!a || e.x > a.x ? e : a), null);
            if (el) ctx.fillText(`${hMin(g[k].h).replace('+', '')} · ${g[k].q} HE`, el.x + 6, el.y);
          });
          ctx.restore();
        }
      }],
      options: {
        indexAxis: 'y', layout: { padding: { right: 100 } },
        scales: { x: { stacked: true, beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { callback: v => v + 'h' } },
          y: { stacked: true, grid: { display: false }, ticks: { font: { size: 11 } } } },
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8 } },
          tooltip: { callbacks: {
            label: c => `${c.dataset.label}: ${hMin(c.raw).replace('+', '')}`,
            footer: its => { const o = g[sups[its[0].dataIndex]]; return `Total ${hMin(o.h).replace('+', '')} · ${o.q} lançamentos · ${o.colab.size} colaboradores\nHE 50%: ${hMin(o.h50).replace('+', '')} · HE 100%: ${hMin(o.h100).replace('+', '')}`; } } }
        },
        onClick: (e, els) => {
          if (!els.length) return;
          const k = sups[els[0].index];
          if (k === 'Sem supervisor') return;
          filtro.sup = filtro.sup === k ? '' : k;
          const sel = document.querySelector('#' + CONTAINER + ' [data-f="sup"]');
          if (sel) sel.value = filtro.sup;
          setTimeout(atualizar, 0);
        },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
      }
    });
  }

  function loteBar(lista) {
    const bar = $('#bh-pd-lote');
    if (!pendSel.size) { bar.hidden = true; return; }
    const sel = lista.filter(i => pendSel.has(i.id));
    const tipos = [...new Set(sel.map(i => i.tipo))];
    bar.hidden = false;
    bar.innerHTML = `<b>${nf(sel.length)} selecionada(s)</b>` +
      (tipos.length > 1 ? `<span class="bh-pd-err">Selecione só um tipo de pendência para justificar junto (marcou: ${esc(tipos.join(', '))})</span>`
        : `<button type="button" class="bh-pd-btn on" id="bh-pd-lote-ok">Justificar selecionadas</button>`) +
      `<button type="button" class="bh-pd-btn" id="bh-pd-lote-limpar">Limpar seleção</button>`;
    const ok = $('#bh-pd-lote-ok'); if (ok) ok.onclick = () => abrirForm([...pendSel]);
    $('#bh-pd-lote-limpar').onclick = () => { pendSel.clear(); renderPend(); };
  }

  function abrirForm(ids) {
    const itens = ids.map(id => pend.find(i => i.id === id)).filter(Boolean);
    if (!itens.length) return;
    const tipo = itens[0].tipo;
    const permitido = itens.every(i => up(i.hePerm) !== 'NÃO');
    const atual = itens.length === 1 ? justDe(itens[0]) : {};
    // ATRASO: desconta do banco só quando todos os selecionados têm saldo positivo; senão, apenas justificar
    const temBanco = itens.every(i => (+i.saldo || 0) > 0);
    const trats = TRATATIVAS[tipo] || TRAT_PADRAO;
    const ehHE = tipo === 'HORA EXTRA';

    const modal = document.createElement('div');
    modal.className = 'bh-modal';
    modal.innerHTML = `<div class="bh-modal-box" role="dialog" aria-modal="true">
      <div class="bh-modal-topo"><h3>Justificar · ${esc(tipo)}</h3><button type="button" class="bh-x" title="Fechar">✕</button></div>
      <div class="bh-modal-lista">${itens.slice(0, 8).map(i => `<div><b>${esc(i.n)}</b> · ${dataBr(i.d)} (${esc(i.dia || '')}) · <span>${esc(i.det)}</span></div>`).join('')}${itens.length > 8 ? `<div>… e mais ${itens.length - 8}</div>` : ''}</div>
      ${ehHE && !permitido ? '<div class="bh-aviso">⚠ BH negativo: a HE não pode ser paga. Use "BANCO DE HORAS" para compensar o débito.</div>' : ''}
      ${tipo === 'FALTA' ? '<div class="bh-aviso">AUSENTE = falta confirmada (o RH desconta) · BANCO DE HORAS = as horas da falta saem do banco (pode ficar negativo) · APENAS JUSTIFICAR = só registra o motivo.</div>' : ''}
      ${tipo === 'ATRASO' ? (temBanco ? '<div class="bh-aviso" style="background:#ecfdf5;color:#065f46">Colaborador com saldo positivo: o atraso pode ser descontado do banco de horas.</div>' : '<div class="bh-aviso">Colaborador sem saldo positivo: escolhendo "BANCO DE HORAS", o atraso é descontado e o banco fica (mais) negativo.</div>') : ''}
      <div class="bh-form">

        <label>Tratativa<select data-c="trat"><option value="">— escolha —</option>${trats.map(t => `<option value="${esc(t)}"${atual.trat === t ? ' selected' : ''}${t === 'PAGAMENTO HE' && !permitido ? ' disabled' : ''}>${esc(t)}${t === 'PAGAMENTO HE' && !permitido ? ' (bloqueado: BH negativo)' : ''}</option>`).join('')}</select></label>
        <label data-campo="tarefa">Tarefa da HE<select data-c="tarefa"><option value="">— escolha —</option>${TAREFAS_HE.map(t => `<option${atual.tarefa === t ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select></label>
        <label data-campo="chamado">Chamado<input data-c="chamado" value="${esc(atual.chamado || '')}" placeholder="Nº do chamado"></label>
        <label data-campo="local">Local de atendimento<input data-c="local" value="${esc(atual.local || '')}" placeholder="Ex.: SPO VMN, SMFRC12"></label>
        <label data-campo="horario">Horário para ajuste das marcações<input data-c="horario" value="${esc(atual.horario || '')}" placeholder="Ex.: entrada 06:00 / saída 14:20"></label>
        <label class="largo">Justificativa<textarea data-c="texto" rows="3" placeholder="Explique o motivo (obrigatório)">${esc(atual.texto || '')}</textarea></label>
      </div>
      <div class="bh-prev" id="bh-prev"></div>
      <div class="bh-modal-pe">
        ${FLUXO_JUSTIFICATIVA ? '' : '<span style="font-size:12px;color:#64748b;flex-basis:100%">Copie e cole no fim da aba JUSTIFICATIVAS_PAINEL da planilha — aparece aqui na próxima atualização.</span>'}
        <button type="button" class="bh-pd-btn${FLUXO_JUSTIFICATIVA ? '' : ' on'}" data-a="copiar" title="Copia a linha completa no formato da aba JUSTIFICATIVAS_PAINEL">📋 Copiar para Excel</button>
        <span style="flex:1"></span>
        <button type="button" class="bh-pd-btn" data-a="cancelar">Cancelar</button>
        ${FLUXO_JUSTIFICATIVA ? `<button type="button" class="bh-pd-btn on" data-a="enviar">Enviar justificativa${itens.length > 1 ? ' (' + itens.length + ')' : ''}</button>` : ''}
      </div>
    </div>`;
    document.getElementById(CONTAINER).appendChild(modal);

    const campo = c => modal.querySelector(`[data-c="${c}"]`);
    const ler = () => ({ trat: campo('trat').value, tarefa: campo('tarefa').value, chamado: campo('chamado').value.trim(),
      local: campo('local').value.trim(), horario: campo('horario').value.trim(), texto: campo('texto').value.trim() });
    const mostrar = () => {
      const j = ler();
      modal.querySelector('[data-campo="tarefa"]').hidden = !(ehHE && j.trat === 'PAGAMENTO HE');   // BANCO DE HORAS não pede tarefa
      modal.querySelector('[data-campo="chamado"]').hidden = !ehHE;
      modal.querySelector('[data-campo="local"]').hidden = tipo === 'ATRASO' || tipo === 'FALTA';
      modal.querySelector('[data-campo="horario"]').hidden = j.trat !== 'AJUSTAR PONTO';
      const r = itens.map(i => calcularStatus(j, i.hePerm, i.tipo));
      const ruim = r.find(x => x[1] !== 'Justificado');
      $('#bh-prev').innerHTML = ruim
        ? `<span style="color:${COR_SIT[ruim[1]]}">● ${esc(ruim[0])}</span>`
        : `<span style="color:${VERDE}">● ${esc(r[0][0])}</span>`;
      return { j, ok: !ruim };
    };
    modal.querySelectorAll('[data-c]').forEach(e => e.addEventListener(e.tagName === 'SELECT' ? 'change' : 'input', mostrar));
    mostrar();
    const fechar = () => modal.remove();
    modal.querySelector('.bh-x').onclick = fechar;
    modal.querySelector('[data-a="cancelar"]').onclick = fechar;
    modal.addEventListener('click', e => { if (e.target === modal) fechar(); });

    // mesmos nomes das colunas da tbJustificativasPainel (aba JUSTIFICATIVAS_PAINEL)
    const linha = (i, j) => {
      const [auto, status] = calcularStatus(j, i.hePerm, i.tipo);
      return { ID: i.id, Chave: i.id, ID_Pendencia: i.idOrig || '', Pendencia: i.tipo, Nome: i.n, Matricula_6: i.m, 'Centro de Custo': i.cc,
        Turno: i.turno, Data: i.d, Dia: i.dia, 'Detalhe / O que verificar': i.det, Supervisor: String(i.sup || '').trim(),
        Coordenador: i.co, Gerente: i.ge, 'Área Técnica': i.area, Status_Colaborador: i.st, BH_ATUAL_H: i.saldo, Horas_Pendencia_h: i.horas,
        Pagamento_HE_Permitido: i.hePerm, Tratativa_Gestor: j.trat,
        Tarefa_HE: j.trat === 'PAGAMENTO HE' ? j.tarefa : '', Chamado: j.chamado, Local_Atendimento: j.local,
        Horario_Ajuste_Marcacoes: j.trat === 'AJUSTAR PONTO' ? j.horario : '', Justificativa_Gestor: j.texto,
        Data_Resposta_Gestor: hojeIso(), Status_RH: '', Observacao_RH: '', Data_Ajuste_RH: '', Status_Automatico: auto, 'Status Pendência': status, Origem: 'PAINEL' };
    };

    modal.querySelector('[data-a="copiar"]').onclick = async () => {
      const { j, ok } = mostrar();
      if (!j.trat) { $('#bh-prev').innerHTML = `<span style="color:${VERMELHO}">● Escolha a tratativa.</span>`; return; }
      if (!ok) return;
      const tsv = itens.map(i => { const l = linha(i, j); l.Data_Resposta_Gestor = dataBr(l.Data_Resposta_Gestor); l.Data = dataBr(l.Data); return COLS_J.map(c => String(l[c] == null ? '' : l[c]).replace(/[\t\n\r]+/g, ' ')).join('\t'); }).join('\n');
      try { await navigator.clipboard.writeText(tsv); $('#bh-prev').insertAdjacentHTML('beforeend', ' · <b>copiado!</b> cole na primeira linha vazia da aba JUSTIFICATIVAS_PAINEL (coluna A)'); }
      catch (e) { window.prompt('Copie (Ctrl+C):', tsv); }
    };

    const btEnv = modal.querySelector('[data-a="enviar"]');
    if (btEnv) btEnv.onclick = async () => {
      const { j, ok } = mostrar();
      if (!j.trat) { $('#bh-prev').innerHTML = `<span style="color:${VERMELHO}">● Escolha a tratativa.</span>`; return; }
      if (!ok) return;   // a mensagem do problema já está na tela
      if (!FLUXO_JUSTIFICATIVA) { $('#bh-prev').innerHTML = `<span style="color:${VERMELHO}">● O envio ainda não está configurado. Use "Copiar para Excel".</span>`; return; }
      btEnv.disabled = true; btEnv.textContent = 'Enviando…';
      try {
        const corpo = { itens: itens.map(i => linha(i, j)) };
        const r = await fetch(FLUXO_JUSTIFICATIVA, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(corpo) });
        const txt = await r.text();
        let resp = null; try { resp = JSON.parse(txt); } catch (e) { }
        if (!r.ok || (resp && resp.ok === false)) throw new Error((resp && resp.erro) || ('o fluxo respondeu ' + r.status));
        // ALTERADO (21/10/2026): o script só grava quando recebe o campo "dados"; sem ele, devolve a exportação
        // (com "itens" e sem "gravadas") e nada é gravado. Confere a resposta antes de dar como enviado.
        if (resp && !('gravadas' in resp) && ('itens' in resp || 'geradoEm' in resp))
          throw new Error('o fluxo rodou mas NÃO gravou: no passo "Executar script", o campo "dados" precisa receber o Corpo do gatilho manual');
        if (resp && 'gravadas' in resp && !resp.gravadas) throw new Error('o fluxo não gravou nenhuma linha');
        const l = locais(), em = new Date().toISOString();
        corpo.itens.forEach(x => {
          l[x.ID] = { em, j: { trat: x.Tratativa_Gestor, tarefa: x.Tarefa_HE, chamado: x.Chamado, local: x.Local_Atendimento, horario: x.Horario_Ajuste_Marcacoes,
            texto: x.Justificativa_Gestor, dataResp: x.Data_Resposta_Gestor, auto: x.Status_Automatico, status: x['Status Pendência'] } };
        });
        salvarLocais(l);
        ids.forEach(id => pendSel.delete(id));
        fechar();
        renderPend();
      } catch (e) {
        btEnv.disabled = false; btEnv.textContent = 'Tentar de novo';
        $('#bh-prev').innerHTML = `<span style="color:${VERMELHO}">● Não foi possível enviar: ${esc(e.message)}. Se continuar, use "Copiar para Excel".</span>`;
      }
    };
    setTimeout(() => campo('trat').focus(), 50);
  }

  // =====================================================================
  // ABA: ANDAMENTO POR SUPERVISOR (mesma base das Pendências)
  // =====================================================================
  function renderAndamento() {
    if (!$('#bh-an-kpis')) return;
    if (!pend) { $('#bh-an-sub').textContent = 'Carregando…'; return; }
    const base0 = pendFiltrados();
    const base = base0.filter(i => !pendFiltro.tipo || i.tipo === pendFiltro.tipo);

    const g = {};
    base.forEach(i => {
      const k = String(i.sup || '').trim() || '(sem supervisor)';
      const s = sitDe(i), j = justDe(i);
      const o = g[k] = g[k] || { tot: 0, 'Justificado': 0, 'Erro na Justificativa': 0, 'Pendente Operação': 0, horasHE: 0, antiga: '', ultima: '' };
      o.tot++; o[s] = (o[s] || 0) + 1;
      if (ehATratar(s)) {
        if (!o.antiga || i.d < o.antiga) o.antiga = i.d;
        if (i.tipo === 'HORA EXTRA') o.horasHE += +i.horas || 0;
      }
      const dr = String(j.dataResp || '').slice(0, 10);
      if (/^\d{4}-\d{2}-\d{2}$/.test(dr) && dr > o.ultima) o.ultima = dr;
    });
    const aT = k => g[k]['Pendente Operação'] + g[k]['Erro na Justificativa'];
    const sups = Object.keys(g).sort((a, b) => aT(b) - aT(a) || g[b].tot - g[a].tot || a.localeCompare(b, 'pt-BR'));

    // cards
    const tot = base.length, just = base.filter(i => sitDe(i) === 'Justificado').length;
    const pendOp = base.filter(i => sitDe(i) === 'Pendente Operação').length, erro = base.filter(i => sitDe(i) === 'Erro na Justificativa').length;
    const zerados = sups.filter(k => !aT(k)).length;
    $('#bh-an-kpis').innerHTML = [
      ['#0b3a75', 'Tratado', pct(just, tot) + '%', `${nf(just)} de ${nf(tot)} pendências justificadas`],
      [COR_SIT['Pendente Operação'], 'Pendente Operação', nf(pendOp), 'sem resposta do gestor'],
      [COR_SIT['Erro na Justificativa'], 'Com erro', nf(erro), 'justificativa a corrigir'],
      [VERDE, 'Supervisores em dia', `${nf(zerados)}/${nf(sups.length)}`, 'sem nada a tratar'],
      ['#334155', 'HE a tratar', hMin(sups.reduce((s, k) => s + g[k].horasHE, 0)).replace('+', ''), 'horas de HE ainda sem justificativa ok']
    ].map(c => `<div class="bh-k" style="background:${c[0]};opacity:1;cursor:default"><small>${c[1]}</small><b>${c[2]}</b><span>${c[3]}</span></div>`).join('');
    $('#bh-an-sub').textContent = `${nf(sups.length)} supervisores · ordenado por quem tem mais pendências a tratar · clique numa barra para abrir as pendências`;

    // gráfico empilhado
    $('#bh-box-an').style.height = Math.max(220, sups.length * 30 + 70) + 'px';
    const SER = [['Justificado', 'Justificado'], ['Erro na Justificativa', 'Com erro'], ['Pendente Operação', 'Pendente Operação']];
    desenhar('bh-g-an', {
      type: 'bar',
      data: { labels: sups, datasets: SER.map(([k, r]) => ({ label: r, data: sups.map(s => g[s][k]), backgroundColor: COR_SIT[k], stack: 's', borderRadius: 3 })) },
      plugins: [{
        id: 'pctTratado', afterDatasetsDraw(ch) {
          const { ctx } = ch, ult = ch.getDatasetMeta(SER.length - 1);
          ctx.save(); ctx.font = '700 11px sans-serif'; ctx.textBaseline = 'middle'; ctx.fillStyle = '#0f172a';
          ult.data.forEach((el, k) => { const s = sups[k]; ctx.fillText(pct(g[s]['Justificado'], g[s].tot) + '% tratado', ch.scales.x.getPixelForValue(g[s].tot) + 6, el.y); });
          ctx.restore();
        }
      }],
      options: {
        indexAxis: 'y', layout: { padding: { right: 90 } },
        onClick: (e, els) => { if (els.length) verPendencias(sups[els[0].index]); },
        onHover: (e, els) => { if (e.native) e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        scales: { x: { stacked: true, grid: { color: '#f1f5f9' }, ticks: { precision: 0 } }, y: { stacked: true, grid: { display: false } } },
        plugins: {
          legend: { position: 'top', align: 'end' },
          tooltip: { callbacks: { label: c => `${c.dataset.label}: ${c.raw} (${pct(c.raw, g[sups[c.dataIndex]].tot)}%)` } },
          datalabels: { display: c => c.raw > 0, color: '#fff', font: { weight: '700', size: 10 }, formatter: v => v }
        }
      }
    });

    // tabela
    $('#bh-an-head').innerHTML = `<tr><th>Supervisor</th><th class="num">Total</th><th class="num">Justificado</th><th class="num">Com erro</th><th class="num">Pendente Op.</th><th style="min-width:160px">% tratado</th><th class="num">HE a tratar</th><th>Mais antiga a tratar</th><th>Última resposta</th><th></th></tr>`;
    $('#bh-an-body').innerHTML = sups.map(k => {
      const o = g[k], p = pct(o['Justificado'], o.tot);
      const cor = p >= 90 ? VERDE : p >= 60 ? '#f59e0b' : VERMELHO;
      return `<tr data-sup="${esc(k)}"><td><b>${esc(k)}</b></td><td class="num">${nf(o.tot)}</td>
        <td class="num pos">${nf(o['Justificado'])}</td><td class="num${o['Erro na Justificativa'] ? ' neg' : ''}">${nf(o['Erro na Justificativa'])}</td>
        <td class="num" style="${o['Pendente Operação'] ? 'color:#b45309;font-weight:700' : ''}">${nf(o['Pendente Operação'])}</td>
        <td><div class="bh-bar" style="grid-template-columns:1fr 44px;margin:0"><i><em style="width:${p}%;background:${cor}"></em></i><b>${p}%</b></div></td>
        <td class="num">${o.horasHE ? hMin(o.horasHE).replace('+', '') : '—'}</td>
        <td>${o.antiga ? dataBr(o.antiga) : '<span style="color:#16a34a">em dia ✓</span>'}</td>
        <td>${o.ultima ? dataBr(o.ultima) : '—'}</td>
        <td>${aT(k) ? '<button type="button" class="bh-pd-btn">Ver pendências</button>' : ''}</td></tr>`;
    }).join('') || '<tr><td colspan="10">Nenhuma pendência no filtro</td></tr>';
    $('#bh-an-body').querySelectorAll('tr[data-sup] .bh-pd-btn').forEach(b => b.onclick = () => verPendencias(b.closest('tr').dataset.sup));
  }

  // abre a aba Pendências já filtrada no supervisor
  function verPendencias(sup) {
    const sel = document.querySelector(`#${CONTAINER} select[data-f="sup"]`);
    let opt = sel && [...sel.options].find(o => o.value.trim() === String(sup).trim());
    if (sel && !opt) { sel.insertAdjacentHTML('beforeend', `<option value="${esc(sup)}">${esc(sup)}</option>`); opt = sel.options[sel.options.length - 1]; }
    filtro.sup = opt ? opt.value : sup;
    if (sel) sel.value = filtro.sup;
    pendFiltro.sit = 'tratar'; pendSel.clear();
    atualizar();
    irAba('pend');
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
    if (pend) { renderPend(); renderAndamento(); }
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
    // o dados_bh.json publicado é de pendências (script trocado no fluxo)? avisa em vez de quebrar a tela
    const pareceBH = i => i && typeof i === 'object' && 'h' in i && Array.isArray(i.bh);
    if (dados.itens.length && !dados.itens.some(pareceBH)) {
      el.innerHTML = `<div style="padding:14px 16px;border-radius:12px;background:#fee2e2;color:#991b1b;font-size:14px;line-height:1.5">
        ⚠ <b>O ${ARQUIVO} publicado não tem o saldo de banco de horas.</b><br>
        ${dados.itens[0] && 'tipo' in dados.itens[0] ? 'Ele está com a lista de <b>pendências</b> no lugar — o fluxo diário rodou o script de pendências e salvou com o nome ' + ARQUIVO + '.' : 'O conteúdo não está no formato esperado.'}<br>
        Restaure o script original do ${ARQUIVO} no fluxo (Histórico de versões do OneDrive) e rode o fluxo de novo.</div>`;
      return;
    }
    dados.itens.forEach(i => { if (!Array.isArray(i.bh)) i.bh = []; });
    css();
    layout(el);
    atualizar();
    if (!dados.itens.length) diagnostico(el, ARQUIVO, dadosBrutos);
    carregarPend();
    if (dados.geradoEm) $('#bh-rod').textContent = 'Atualizado em ' + new Date(dados.geradoEm).toLocaleString('pt-BR');
  }

  // se a seção estiver escondida ao carregar, chame bhResize() quando ela aparecer
  window.bhResize = () => Object.values(graficos).forEach(g => g.resize());

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
