/* ======================================================================
   gestao_combustivel.js — Gestão de Frotas
   Divide a página em duas abas:
   • Veículos: só quantidade de carros por área (contrato) e por tipo
   • Combustível: custos do contrato, combustível mês a mês (VALOR EMISSAO,
     por TIPO FROTA), Top 10 técnicos (só FROTA), gasto por contrato e ranking
   Lê dados_combustivel.json (gerado pelo script_combustivel.ts).
   Não altera nada do index.html: só envolve a função renderFrota().
   ====================================================================== */
(function () {
  // Contratos do cadastro de frota (coluna Operação da aba Placa) que a
  // planilha de combustível cobre. Se o nome mudar, ajuste aqui.
  const CONTRATOS_COM_COMBUSTIVEL = ['CLARO INFRA SP'];

  const CORES_TIPO = { FROTA: '#0e7c86', GMG: '#d97706', RAC: '#5e34b5', BENEFICIO: '#2563eb' };
  const MES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

  let COMB = null;          // conteúdo do dados_combustivel.json
  let mesComb = '';         // '' = todos os meses; senão 'AAAA-MM'
  let mesCombIniciado = false;
  let chartMeses = null, chartTop = null, chartGmgMeses = null, chartTopGmg = null;
  let abaFrota = 'veiculos';   // 'veiculos' | 'combustivel'

  const brl = v => (Number(v) || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  const brlCurto = v => {
    const n = Number(v) || 0;
    return n >= 1000 ? 'R$ ' + (n / 1000).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' mil' : brl(n);
  };
  const rotuloMes = ord => { const p = String(ord).split('-'); return MES_PT[Number(p[1]) - 1] + '/' + p[0].slice(2); };
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const chavePlaca = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

  async function carregarCombustivel() {
    try {
      const r = await fetch('dados_combustivel.json?_=' + Date.now());
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const j = await r.json();
      if (j && j.porMes) { COMB = j; console.log('[Combustível] dados carregados:', j.meses); }
    } catch (e) {
      console.warn('[Combustível] dados_combustivel.json não carregado:', e.message);
    }
    if (!mesCombIniciado && COMB && COMB.meses && COMB.meses.length) {
      mesComb = COMB.meses[COMB.meses.length - 1];   // começa no mês mais recente
      mesCombIniciado = true;
    }
    if (paginaFrotaVisivel()) renderFrota();
  }

  function paginaFrotaVisivel() {
    const p = document.getElementById('pagina-frota');
    return p && p.style.display !== 'none';
  }

  function blocoDoMes() {
    if (!COMB) return null;
    return mesComb ? COMB.porMes[mesComb] : COMB.geral;
  }

  function contratoAtual() { return typeof frotaOperacaoFiltroAtual === 'string' ? frotaOperacaoFiltroAtual : ''; }
  function contratoTemCombustivel(c) { return !c || CONTRATOS_COM_COMBUSTIVEL.indexOf(c) !== -1; }

  // Troca o "Valor Combustível" de cada placa pelo valor do mês escolhido
  function aplicarCombustivelNasPlacas() {
    if (!Array.isArray(FROTA)) return;
    const bloco = blocoDoMes();
    const mapa = {};
    // porPlacaTodos = todas as placas da planilha, de qualquer contrato
    if (bloco) (bloco.porPlacaTodos || bloco.porPlaca || []).forEach(p => { mapa[chavePlaca(p.Placa)] = p; });
    FROTA.forEach(v => {
      if (!('_combOriginal' in v)) v._combOriginal = v.valorCombustivel;
      const d = bloco ? mapa[chavePlaca(v.placa)] : null;
      v._combDet = d || null;
      v.valorCombustivel = bloco ? (d ? Number(d['Valor Total']) || 0 : 0) : v._combOriginal;
    });
  }

  // Ranking: todos os veículos do contrato filtrado, do maior para o menor,
  // com litros, VL/litro, hodômetro, km rodados e km/litro
  function renderRankingCompleto() {
    const tbody = document.getElementById('frotaRankingBody');
    if (!tbody) return;
    const tabela = tbody.closest('table');
    const painel = tbody.closest('.panel');
    if (tabela && !tabela.dataset.comb) {
      tabela.dataset.comb = '1';
      tabela.querySelector('thead').innerHTML = `<tr>
        <th class="num">#</th><th>Placa</th><th>Condutor</th><th>Operação</th>
        <th class="num">Aluguel/Compra</th><th class="num">Combustível</th><th class="num">Litros</th>
        <th class="num">VL/Litro</th><th class="num">Hodômetro ou Horímetro</th>
        <th class="num">KM Rodados ou Horas Trab.</th><th class="num">KM/Litro ou Litros/Hora</th><th class="num">Total</th></tr>`;
      const rolagem = document.createElement('div');
      rolagem.style.cssText = 'overflow:auto; max-height:640px; border-radius:8px;';
      tabela.parentElement.insertBefore(rolagem, tabela);
      rolagem.appendChild(tabela);
    }
    const c = contratoAtual();
    const lista = (c ? FROTA.filter(v => (canonOperacaoFrota(v.operacao) || 'Sem operação') === c) : FROTA)
      .map(v => {
        const aluguel = extrairNumeroValorFrota(v.valor), comb = extrairNumeroValorFrota(v.valorCombustivel);
        return { v, aluguel, comb, total: aluguel + comb, d: v._combDet };
      })
      .sort((a, b) => b.total - a.total);
    const exibidos = c ? lista : lista.slice(0, 10);

    const titulo = painel && painel.querySelector('h2');
    if (titulo) titulo.textContent = c
      ? 'Ranking de Gastos por Veículo — ' + c + ' (' + lista.length + ' veículos)' + (COMB ? ' · ' + (mesComb ? rotuloMes(mesComb) : 'todos os meses') : '')
      : 'Ranking de Gastos por Veículo — Top 10 (filtre um contrato para ver todos)';

    const n = (x, casas) => (Number(x) || 0).toLocaleString('pt-BR', { minimumFractionDigits: casas, maximumFractionDigits: casas });
    const ou = (x, f) => x ? f(x) : '—';
    tbody.innerHTML = exibidos.map((it, i) => {
      const d = it.d || {}, gmg = d.tipoFrota === 'GMG';
      return `<tr>
        <td class="num mono">${i + 1}</td>
        <td class="mono">${esc(it.v.placa) || '—'}</td>
        <td>${esc(it.v.condutor) || '—'}</td>
        <td>${esc(it.v.operacao) || '—'}</td>
        <td class="num mono">${brl(it.aluguel)}</td>
        <td class="num mono">${brl(it.comb)}</td>
        <td class="num mono">${ou(d.litros, x => n(x, 2))}</td>
        <td class="num mono">${ou(d.vlLitro, x => 'R$ ' + n(x, 3))}</td>
        <td class="num mono">${ou(d.hodometro, x => n(x, 0) + (gmg ? ' h' : ' km'))}</td>
        <td class="num mono">${ou(d.kmRodados, x => n(x, 0) + (gmg ? ' h' : ' km'))}</td>
        <td class="num mono">${ou(d.kmLitro, x => n(x, 2) + (gmg ? ' L/h' : ' km/L'))}</td>
        <td class="num mono" style="font-weight:700;">${brl(it.total)}</td>
      </tr>`;
    }).join('') || `<tr><td colspan="12" style="text-align:center; color:${cor.muted};">Nenhum veículo</td></tr>`;
  }

  // Aba Veículos: com filtro ativo, os cards mostram só o contrato escolhido
  function ajustarCardsContrato() {
    const row = document.getElementById('frotaOperacaoKpiRow');
    if (!row) return;
    const c = contratoAtual();
    if (!c) return;   // sem filtro: mantém os cards de todos os contratos
    const qtd = FROTA.filter(v => (canonOperacaoFrota(v.operacao) || 'Sem operação') === c).length;
    row.innerHTML = `
      <div class="kpi" style="cursor:pointer;" onclick="filtrarFrotaPorOperacaoEKpi('')" title="Voltar para todos os contratos">
        <div class="label">↺ Todos os contratos</div>
        <div class="value" style="color:${cor.muted}">${FROTA.length.toLocaleString('pt-BR')}</div>
        <div class="delta" style="color:${cor.muted};">clique para limpar o filtro</div>
      </div>
      <div class="kpi" style="border-color:${cor.accent}; box-shadow:0 0 0 2px ${cor.accent} inset;">
        <div class="label" style="color:${cor.accent};">🏢 ${esc(c)}</div>
        <div class="value" style="color:${cor.accent};">${qtd}</div>
        <div class="delta" style="color:${cor.accent};">veículos</div>
      </div>`;
  }

  // Aba Veículos: tabela Área (contrato) × Tipo
  function renderMatrizAreaTipo() {
    const alvo = document.getElementById('frotaMatrizAreaTipo');
    if (!alvo) return;
    const c = contratoAtual();
    const lista = c ? FROTA.filter(v => (canonOperacaoFrota(v.operacao) || 'Sem operação') === c) : FROTA;
    const tipos = Array.from(new Set(lista.map(v => normFrota(v.tipo) || 'Sem tipo'))).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const linhas = {};
    lista.forEach(v => {
      const a = canonOperacaoFrota(v.operacao) || 'Sem operação';
      const t = normFrota(v.tipo) || 'Sem tipo';
      if (!linhas[a]) linhas[a] = { total: 0 };
      linhas[a][t] = (linhas[a][t] || 0) + 1;
      linhas[a].total++;
    });
    const areas = Object.keys(linhas).sort((a, b) => linhas[b].total - linhas[a].total);
    const totTipo = {}; tipos.forEach(t => totTipo[t] = areas.reduce((s, a) => s + (linhas[a][t] || 0), 0));
    alvo.innerHTML = `<table>
      <thead><tr><th>Área / Contrato</th>${tipos.map(t => `<th class="num">${esc(t)}</th>`).join('')}<th class="num">Total</th></tr></thead>
      <tbody>
        ${areas.map(a => `<tr><td>${esc(a)}</td>${tipos.map(t => `<td class="num">${linhas[a][t] || '—'}</td>`).join('')}<td class="num" style="font-weight:700;">${linhas[a].total}</td></tr>`).join('')}
        <tr style="font-weight:700; border-top:2px solid ${cor.line};"><td>Total</td>${tipos.map(t => `<td class="num">${totTipo[t]}</td>`).join('')}<td class="num">${lista.length}</td></tr>
      </tbody></table>`;
  }

  // Aba Combustível: custos do contrato filtrado (ou de todos)
  function renderCustosContrato() {
    const row = document.getElementById('combCustoKpiRow');
    if (!row) return;
    const c = contratoAtual();
    const lista = c ? FROTA.filter(v => (canonOperacaoFrota(v.operacao) || 'Sem operação') === c) : FROTA;
    let aluguel = 0, combPlacas = 0;
    lista.forEach(v => { aluguel += extrairNumeroValorFrota(v.valor); combPlacas += extrairNumeroValorFrota(v.valorCombustivel); });
    const bloco = blocoDoMes();
    const usaRelatorio = bloco && c && contratoTemCombustivel(c);
    const frotaRel = usaRelatorio ? (bloco.porTipoFrota.find(t => t.tipo === 'FROTA') || { valor: 0 }).valor : 0;
    const comb = usaRelatorio ? frotaRel : combPlacas;
    const periodo = COMB ? (mesComb ? rotuloMes(mesComb) : 'todos os meses') : '';
    row.innerHTML = `
      <div class="kpi"><div class="label">🏢 ${esc(c || 'Todos os contratos')}</div>
        <div class="value">${lista.length}</div><div class="delta" style="color:${cor.muted};">veículos</div></div>
      <div class="kpi"><div class="label">🚗 Aluguel / Compra</div>
        <div class="value" style="font-size:20px;">${brl(aluguel)}</div></div>
      <div class="kpi"><div class="label">⛽ Combustível (Frota)</div>
        <div class="value" style="font-size:20px; color:${cor.warn};">${brl(comb)}</div>
        <div class="delta" style="color:${cor.muted};">${usaRelatorio ? periodo + ' · só TIPO FROTA = FROTA' : periodo + ' · soma das placas (Frota)'}</div></div>
      <div class="kpi"><div class="label">Σ Total</div>
        <div class="value" style="font-size:20px; color:${cor.accent};">${brl(aluguel + comb)}</div></div>`;
  }

  // Chips de mês (linha do filtro de contrato) — só na aba Combustível
  function renderChipsMes() {
    const strip = document.getElementById('combMesStrip');
    if (!strip) return;
    if (abaFrota !== 'combustivel' || !COMB) { strip.style.display = 'none'; return; }
    strip.style.display = '';
    strip.innerHTML = `<span class="mono" style="font-size:11px; color:${cor.muted}; text-transform:uppercase; letter-spacing:.05em; align-self:center; margin-right:2px;">Mês</span>` +
      COMB.meses.map(m => `<div class="chip${m === mesComb ? ' active' : ''}" data-mes="${m}">${rotuloMes(m)}</div>`).join('') +
      `<div class="chip${mesComb === '' ? ' active' : ''}" data-mes="">Todos</div>`;
    strip.querySelectorAll('.chip').forEach(ch => ch.onclick = () => { mesComb = ch.dataset.mes; renderFrota(); });
  }

  // Monta as duas abas uma única vez, reaproveitando os blocos que já existem
  function montarAbas() {
    const secao = document.getElementById('frotaSecao');
    if (!secao || document.getElementById('frotaAbas')) return;
    const painelDe = id => { const el = document.getElementById(id); return el ? el.closest('.panel') : null; };

    const filtroTopo = document.getElementById('frotaContratoTopo').parentElement;
    const kpiContratos = document.getElementById('frotaOperacaoKpiRow');
    const kpiStatus = document.getElementById('frotaKpiRow');
    const pGasto = painelDe('frotaGastoContratoBody');
    const pRanking = painelDe('frotaRankingBody');
    const pTipo = painelDe('frotaTipoChart');
    const gridGraficos = pTipo ? pTipo.parentElement : null;
    const pVeiculos = painelDe('frotaTable');

    const selectContrato = document.getElementById('frotaContratoTopo');
    const abas = document.createElement('div');
    abas.id = 'frotaAbas';
    abas.className = 'month-strip';
    abas.style.marginLeft = '12px';
    abas.innerHTML = `<div class="chip active" data-aba="veiculos">🚗 Veículos</div><div class="chip" data-aba="combustivel">⛽ Combustível</div>`;
    selectContrato.insertAdjacentElement('afterend', abas);

    const chipsMes = document.createElement('div');
    chipsMes.id = 'combMesStrip';
    chipsMes.className = 'month-strip';
    chipsMes.style.cssText = 'margin-left:auto; display:none;';
    filtroTopo.appendChild(chipsMes);
    filtroTopo.style.marginBottom = '16px';

    const abaV = document.createElement('div'); abaV.id = 'frotaAbaVeiculos';
    const abaC = document.createElement('div'); abaC.id = 'frotaAbaCombustivel'; abaC.style.display = 'none';
    const oculto = document.createElement('div'); oculto.id = 'frotaOculto'; oculto.style.display = 'none';
    filtroTopo.insertAdjacentElement('afterend', abaV);
    abaV.insertAdjacentElement('afterend', abaC);
    abaC.insertAdjacentElement('afterend', oculto);

    // Veículos: cards por contrato + Área × Tipo + gráfico por tipo
    abaV.appendChild(kpiContratos);
    // status dos veículos (Ativo, Disponível, Manutenção, Funilaria…) — segue o filtro de contrato
    if (kpiStatus) {
      const tituloStatus = document.createElement('h2');
      tituloStatus.textContent = 'Status dos veículos';
      tituloStatus.style.cssText = 'font-family:"Space Grotesk",sans-serif; font-size:15px; font-weight:600; margin:4px 0 10px;';
      abaV.appendChild(tituloStatus);
      abaV.appendChild(kpiStatus);
    }
    const grade = document.createElement('div');
    grade.className = 'grid';
    grade.style.gridTemplateColumns = '1.4fr 1fr';
    grade.innerHTML = `<div class="panel"><h2>Veículos por Área e Tipo</h2><div id="frotaMatrizAreaTipo" style="overflow:auto;"></div></div>`;
    abaV.appendChild(grade);
    if (pTipo) grade.appendChild(pTipo);

    // Combustível: custos + painel de combustível + gasto por contrato + ranking
    const custos = document.createElement('div');
    custos.className = 'kpi-row'; custos.id = 'combCustoKpiRow'; custos.style.marginBottom = '16px';
    abaC.appendChild(custos);
    const ancoraComb = document.createElement('div'); ancoraComb.id = 'combAncora';
    abaC.appendChild(ancoraComb);
    if (pGasto) abaC.appendChild(pGasto);
    if (pRanking) abaC.appendChild(pRanking);

    // O que não foi pedido fica escondido (o código original continua funcionando)
    [gridGraficos, pVeiculos].forEach(el => { if (el) oculto.appendChild(el); });

    abas.querySelectorAll('.chip').forEach(ch => ch.onclick = () => {
      abaFrota = ch.dataset.aba;
      abas.querySelectorAll('.chip').forEach(x => x.classList.toggle('active', x === ch));
      abaV.style.display = abaFrota === 'veiculos' ? '' : 'none';
      abaC.style.display = abaFrota === 'combustivel' ? '' : 'none';
      renderFrota();   // redesenha os gráficos da aba que ficou visível
    });
  }

  function garantirPainel() {
    let painel = document.getElementById('frotaCombustivelSecaoV8');
    if (!painel) {
      const ancora = document.getElementById('combAncora') || document.getElementById('frotaKpiRow');
      if (!ancora) return null;
      painel = document.createElement('div');
      painel.id = 'frotaCombustivelSecaoV8';
      painel.className = 'panel';
      painel.style.marginBottom = '16px';
      ancora.insertAdjacentElement('afterend', painel);
    }
    // (re)monta a estrutura se estiver faltando (ex.: depois do aviso de "sem dados")
    if (!painel.querySelector('#combTipoKpiRow')) painel.innerHTML = `
      <h2>Combustível — Claro Infra SPC</h2>
      <div class="kpi-row" id="combTipoKpiRow" style="margin:10px 0 16px;"></div>
      <div class="grid" style="grid-template-columns: 1fr 1.3fr;">
        <div class="panel">
          <h2>🚗 Frota — valor e litros mês a mês</h2>
          <div style="position:relative; height:300px;"><canvas id="combMesesChart"></canvas></div>
        </div>
        <div class="panel">
          <h2 id="combTopTitulo">Top 10 técnicos que mais usaram (Frota)</h2>
          <div style="position:relative; height:300px;"><canvas id="combTopChart"></canvas></div>
        </div>
      </div>
      <div class="grid" style="grid-template-columns: 1fr 1.3fr; margin-bottom:0;">
        <div class="panel">
          <h2>⚡ Gerador (GMG) — valor e litros mês a mês</h2>
          <div style="position:relative; height:300px;"><canvas id="combGmgMesesChart"></canvas></div>
        </div>
        <div class="panel">
          <h2 id="combTopGmgTitulo">Top 10 técnicos que mais abasteceram gerador (GMG)</h2>
          <div style="position:relative; height:300px;"><canvas id="combTopGmgChart"></canvas></div>
        </div>
      </div>
      <div id="combAviso" class="mono" style="font-size:11px; color:${cor.muted}; margin-top:10px;"></div>`;
    return painel;
  }

  function renderCombustivel() {
    const painel = garantirPainel();
    if (!painel) return;
    const c = contratoAtual();

    if (!COMB) {
      painel.style.display = '';
      painel.innerHTML = `<h2>Combustível — Claro Infra SPC</h2>
        <p style="font-size:13px; color:${cor.muted}; margin-top:8px;">Carregando dados_combustivel.json… Se esta mensagem não sumir, o arquivo não foi encontrado no repositório.</p>`;
      return;
    }
    if (!contratoTemCombustivel(c)) { painel.style.display = 'none'; return; }
    painel.style.display = '';

    const bloco = blocoDoMes();
    const periodo = mesComb ? rotuloMes(mesComb) : 'todos os meses';

    // cards por TIPO FROTA
    const kpi = document.getElementById('combTipoKpiRow');
    kpi.innerHTML = `<div class="kpi">
        <div class="label">⛽ Total ${periodo}</div>
        <div class="value" style="color:${cor.accent}; font-size:20px;">${brl(bloco.total)}</div>
        <div class="delta" style="color:${cor.muted};">${bloco.qtdAbastecimentos.toLocaleString('pt-BR')} abastecimentos</div>
      </div>` +
      bloco.porTipoFrota.map(t => {
        const pct = bloco.total > 0 ? (t.valor / bloco.total * 100).toFixed(1) : '0.0';
        const corT = CORES_TIPO[t.tipo] || cor.muted;
        return `<div class="kpi" style="border-color:${corT}55;">
          <div class="label" style="color:${corT};">${esc(t.tipo)}</div>
          <div class="value" style="font-size:20px; color:${corT};">${brl(t.valor)}</div>
          <div class="delta" style="color:${corT};">${pct}% · ${t.qtd} abast.</div>
        </div>`;
      }).join('');

    const tipoDoMes = (m, tipo) => {
      const t = COMB.porMes[m].porTipoFrota.find(x => x.tipo === tipo) || {};
      return { valor: Number(t.valor) || 0, litros: Number(t.litros) || 0 };
    };
    const variacao = (lista, i) => {
      const ant = i > 0 ? lista[i - 1] : null;
      if (!ant) return '';
      const p = (lista[i] - ant) / ant * 100;
      return ' (' + (p >= 0 ? '+' : '') + p.toFixed(1) + '%)';
    };
    const corMes = (base, m) => base + (mesComb && m !== mesComb ? '66' : 'dd');

    // Gráfico mês a mês: valor + variação em cima da barra, litros dentro da barra
    function graficoMesAMes(canvasId, tipo, corBase, chartAtual) {
      const valores = COMB.meses.map(m => tipoDoMes(m, tipo).valor);
      const litros = COMB.meses.map(m => tipoDoMes(m, tipo).litros);
      const varLitros = i => {
        const ant = i > 0 ? litros[i - 1] : 0;
        if (!ant) return '';
        const p = (litros[i] - ant) / ant * 100;
        return ' (' + (p >= 0 ? '+' : '') + p.toFixed(1) + '%)';
      };
      if (chartAtual) chartAtual.destroy();
      return new Chart(document.getElementById(canvasId), {
        type: 'bar',
        data: { labels: COMB.meses.map(rotuloMes), datasets: [{
          label: tipo, data: valores, borderRadius: 4,
          backgroundColor: COMB.meses.map(m => corMes(corBase, m)),
          datalabels: { labels: {
            valor: { anchor: 'end', align: 'top', color: '#000', font: { size: 11, family: 'JetBrains Mono', weight: '700' },
              formatter: (v, ctx) => brlCurto(v) + variacao(valores, ctx.dataIndex) },
            litros: { anchor: 'center', align: 'center', color: '#fff', textAlign: 'center',
              font: { size: 12, family: 'JetBrains Mono', weight: '700' },
              display: ctx => litros[ctx.dataIndex] > 0,
              formatter: (v, ctx) => ['⛽ ' + litros[ctx.dataIndex].toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + ' L', varLitros(ctx.dataIndex).trim()] }
          } }
        }] },
        plugins: [ChartDataLabels],
        options: {
          responsive: true, maintainAspectRatio: false, layout: { padding: { top: 24 } },
          plugins: { legend: { display: false }, datalabels: { display: true },
            tooltip: { callbacks: { label: ctx => brl(ctx.parsed.y) + ' · ' + litros[ctx.dataIndex].toLocaleString('pt-BR') + ' L' +
              (litros[ctx.dataIndex] ? ' · R$ ' + (ctx.parsed.y / litros[ctx.dataIndex]).toLocaleString('pt-BR', { maximumFractionDigits: 3 }) + '/L' : '') } } },
          scales: {
            x: { ticks: { color: cor.text, font: { weight: 'bold' } }, grid: { display: false } },
            y: { beginAtZero: true, ticks: { color: cor.text, callback: v => brlCurto(v) }, grid: { color: cor.line } }
          }
        }
      });
    }
    chartMeses = graficoMesAMes('combMesesChart', 'FROTA', CORES_TIPO.FROTA, chartMeses);
    chartGmgMeses = graficoMesAMes('combGmgMesesChart', 'GMG', CORES_TIPO.GMG, chartGmgMeses);

    // Top 10 (Frota e GMG) — em "Todos", barra empilhada mês a mês
    const PALETA_MESES = ['#0e7c86', '#5e34b5', '#d97706', '#059669', '#2563eb', '#dc2626'];
    function graficoTop(canvasId, tituloId, tituloBase, top, corBarra, chartAtual) {
      const mesAMes = !mesComb && top.length && top[0].porMes;
      document.getElementById(tituloId).textContent = tituloBase + ' — ' + (mesAMes ? 'mês a mês' : periodo);
      const rotuloBarra = { display: ctx => (ctx.dataset.data[ctx.dataIndex] || 0) > 0, color: '#fff', anchor: 'center', align: 'center',
        font: { size: 10, family: 'JetBrains Mono', weight: '700' }, formatter: v => brlCurto(v) };
      const datasets = mesAMes
        ? [
            ...COMB.meses.map((m, i) => ({
              label: rotuloMes(m), data: top.map(t => t.porMes[m] || 0),
              backgroundColor: PALETA_MESES[i % PALETA_MESES.length], borderRadius: 3, datalabels: rotuloBarra
            })),
            { label: 'Total', data: top.map(() => 0), backgroundColor: 'rgba(0,0,0,0)',
              datalabels: { anchor: 'end', align: 'right', offset: 4, color: '#000', font: { size: 10, family: 'JetBrains Mono', weight: '700' },
                formatter: (v, ctx) => brl(top[ctx.dataIndex].valor) } }
          ]
        : [{
            label: 'Valor', data: top.map(t => t.valor),
            backgroundColor: top.map((_, i) => i === 0 ? cor.bad : corBarra), borderRadius: 4,
            datalabels: { anchor: 'end', align: 'right', color: '#000', font: { size: 10, family: 'JetBrains Mono', weight: '700' },
              formatter: v => brl(v) }
          }];
      if (chartAtual) chartAtual.destroy();
      return new Chart(document.getElementById(canvasId), {
        type: 'bar',
        data: { labels: top.map(t => t.motorista), datasets },
        plugins: [ChartDataLabels],
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { right: 90 } },
          plugins: {
            legend: { display: !!mesAMes, labels: { color: cor.muted, font: { family: 'Inter', size: 11 }, filter: i => i.text !== 'Total' } },
            datalabels: { display: true },
            tooltip: { filter: i => i.dataset.label !== 'Total', callbacks: {
              label: ctx => (mesAMes ? ctx.dataset.label + ': ' : '') + brl(ctx.parsed.x),
              afterBody: items => {
                const t = top[items[0].dataIndex];
                return t.qtd + ' abastecimentos' + (t.litros ? ' · ' + t.litros.toLocaleString('pt-BR') + ' L' : '') +
                  (t.placas && t.placas.length ? ' · ' + t.placas.join(', ') : '');
              } } }
          },
          scales: {
            x: { stacked: !!mesAMes, beginAtZero: true, ticks: { display: false }, grid: { display: false } },
            y: { stacked: !!mesAMes, ticks: { color: cor.text, font: { size: 10, weight: 'bold' } }, grid: { display: false } }
          }
        }
      });
    }
    chartTop = graficoTop('combTopChart', 'combTopTitulo', 'Top 10 técnicos que mais usaram (Frota)',
      bloco.topMotoristasFrota || [], 'rgba(14,124,134,.8)', chartTop);
    chartTopGmg = graficoTop('combTopGmgChart', 'combTopGmgTitulo', 'Top 10 técnicos que mais abasteceram gerador (GMG)',
      bloco.topMotoristasGmg || [], 'rgba(217,119,6,.8)', chartTopGmg);
    if (!(bloco.topMotoristasGmg || []).length) console.warn('[Combustível] JSON sem "topMotoristasGmg" — atualize o script_combustivel.ts no Excel');

    const aviso = document.getElementById('combAviso');
    const jsonAntigo = !COMB.geral || !COMB.geral.topMotoristasGmg || !(COMB.geral.porPlacaTodos || [])[0] || !('litros' in COMB.geral.porPlacaTodos[0]);
    if (jsonAntigo) {
      aviso.innerHTML = `<span style="color:${cor.bad}; font-weight:700;">⚠ O dados_combustivel.json foi gerado pela versão ANTIGA do script (sem litros, km e ranking de gerador). Atualize o Office Script de combustível no Excel com o script_combustivel.ts novo e rode o fluxo.</span>`;
      console.warn('[Combustível] dados_combustivel.json no formato antigo — gerado em', COMB.atualizadoEm);
      return;
    }
    aviso.textContent = COMB.atualizadoEm
      ? 'Relatório de abastecimento atualizado em ' + new Date(COMB.atualizadoEm).toLocaleString('pt-BR') + '. O mês selecionado aqui também define o "Valor Combustível" de todas as placas na tabela e no ranking abaixo.'
      : '';
  }

  // Envolve a renderFrota original do index.html
  function instalar() {
    if (typeof renderFrota !== 'function' || renderFrota._combV8) return;
    if (renderFrota._comb) console.warn('[Combustível] Existe uma cópia ANTIGA deste código em outro arquivo (provavelmente gestao_preventiva.js). Restaure esse arquivo.');
    const original = renderFrota;
    const nova = function () {
      montarAbas();
      aplicarCombustivelNasPlacas();
      original.apply(this, arguments);
      ajustarCardsContrato();
      renderMatrizAreaTipo();
      renderCustosContrato();
      renderRankingCompleto();
      renderChipsMes();
      if (abaFrota === 'combustivel') renderCombustivel();
    };
    nova._comb = true;
    nova._combV8 = true;
    renderFrota = nova;   // mostrarPagina() e os filtros chamam pelo nome
  }

  instalar();
  console.log('[Combustível] v10 módulo carregado. renderFrota envolvida:', !!(window.renderFrota && renderFrota._combV8));
  carregarCombustivel();
  setInterval(carregarCombustivel, 5 * 60 * 1000);
})();
