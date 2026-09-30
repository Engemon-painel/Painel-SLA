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
    { k: 'x',  rotulo: 'Executor' },
    { k: 't',  rotulo: 'Tipo Preventiva' },
    { k: 'e',  rotulo: 'Equipe Responsável' }
  ];

  const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho',
                 'Agosto','Setembro','Outubro','Novembro','Dezembro'];

  let dados = null, grafico = null, raiz = null;
  const sel = { st: '', x: '', t: '', e: '' };
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
    const css = `
    .pv{--pv-azul:#0b1f8f;--pv-azul2:#1e88e5;--pv-borda:#e3e6ee;--pv-txt:#1b1f2a;--pv-sub:#5b6275;--pv-card:#fff;
        font-family:inherit;color:var(--pv-txt)}
    .pv-topo{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:12px;gap:12px;flex-wrap:wrap}
    .pv-topo h2{margin:0;font-size:20px}.pv-topo small{color:var(--pv-sub)}
    .pv-filtros{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:12px}
    .pv-f{background:var(--pv-card);border:1px solid var(--pv-borda);border-radius:8px;padding:8px 10px}
    .pv-f label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;color:var(--pv-sub);margin-bottom:4px}
    .pv-f select{width:100%;padding:6px;border:1px solid var(--pv-borda);border-radius:5px;background:#fff;font:inherit}
    .pv-kpis{display:grid;grid-template-columns:1fr 2fr 1fr 1fr;gap:10px;margin-bottom:12px}
    .pv-card{background:var(--pv-card);border:1px solid var(--pv-borda);border-radius:8px;padding:12px 14px;display:flex;flex-direction:column;gap:10px}
    .pv-num{border-left:3px solid var(--pv-txt);padding-left:10px}
    .pv-num b{display:block;font-size:24px;line-height:1.1}.pv-num span{font-size:13px;color:var(--pv-sub)}
    .pv-par{display:grid;grid-template-columns:1fr auto;gap:10px}
    .pv-barra{height:22px;background:#d0d0d0;border-radius:3px;overflow:hidden}
    .pv-barra i{display:block;height:100%;background:var(--pv-azul);transition:width .4s}
    .pv-tag{border-radius:4px;padding:8px 10px}.pv-tag span{font-size:13px;opacity:.75}.pv-tag b{display:block;font-size:16px}
    .pv-tag.azul{background:#e8f2fd}.pv-tag.amarelo{background:#fff8a8}
    .pv-tag.verde{background:#c8ecd0}.pv-tag.laranja{background:#ffe0b2}.pv-tag.vermelho{background:#ffb3b3}
    .pv-graf{background:var(--pv-card);border:1px solid var(--pv-borda);border-radius:8px;padding:12px 14px;margin-bottom:12px}
    .pv-graf h3,.pv-tab h3{margin:0 0 8px;font-size:15px;text-align:center}
    .pv-graf .pv-cv{position:relative;height:360px}
    .pv-tab{background:var(--pv-card);border:1px solid var(--pv-borda);border-radius:8px;padding:12px 14px}
    .pv-tab-topo{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:8px}
    .pv-tab-topo h3{margin:0;text-align:left}
    .pv-tab input{padding:6px 10px;border:1px solid var(--pv-borda);border-radius:5px;min-width:220px;font:inherit}
    .pv-rolagem{max-height:380px;overflow:auto}
    .pv-tab table{width:100%;border-collapse:collapse;font-size:13px}
    .pv-tab th{position:sticky;top:0;background:var(--pv-azul);color:#fff;text-align:left;padding:7px 8px;white-space:nowrap}
    .pv-tab td{padding:6px 8px;border-bottom:1px solid var(--pv-borda);white-space:nowrap}
    .pv-st{padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600}
    .pv-st.Pendente{background:#ffe0b2}.pv-st.Aberta{background:#fff8a8}
    @media (max-width:900px){.pv-kpis{grid-template-columns:1fr 1fr}.pv-kpis .pv-largo{grid-column:1/-1}}
    @media (max-width:520px){.pv-kpis{grid-template-columns:1fr}}`;
    const st = document.createElement('style'); st.id = 'pv-estilos'; st.textContent = css;
    document.head.appendChild(st);
  }

  function esqueleto() {
    const [a, m] = dados.mes.split('-').map(Number);
    raiz.innerHTML = `
    <div class="pv">
      <div class="pv-topo"><h2>Preventiva · ${MESES[m - 1]}/${a}</h2>
        <small>Atualizado em ${esc(dados.gerado_em || '')}</small></div>
      <div class="pv-filtros">${FILTROS.map(f => `
        <div class="pv-f"><label>${f.rotulo}</label><select data-k="${f.k}"></select></div>`).join('')}
      </div>
      <div class="pv-kpis">
        <div class="pv-card">
          <div class="pv-num"><b id="pv-plan"></b><span>Planejado no Mês</span></div>
          <div class="pv-tag azul"><span>Dias Úteis no Mês</span><b id="pv-du-mes"></b></div>
        </div>
        <div class="pv-card pv-largo">
          <div class="pv-par">
            <div class="pv-num"><b id="pv-prog"></b><span>Progresso em D-1</span></div>
            <div class="pv-num"><b id="pv-du"></b><span>Dias Úteis</span></div>
          </div>
          <div class="pv-barra"><i id="pv-barra"></i></div>
          <div class="pv-par">
            <div class="pv-num"><b id="pv-entrega"></b><span>% Entrega</span></div>
            <div class="pv-num"><b id="pv-pd1"></b><span>% Planejado em D-1</span></div>
          </div>
        </div>
        <div class="pv-card">
          <div class="pv-num"><b id="pv-saldo"></b><span>Saldo</span></div>
          <div class="pv-tag" id="pv-ader-box"><span>Aderência</span><b id="pv-ader"></b></div>
        </div>
        <div class="pv-card">
          <div class="pv-num"><b id="pv-pend"></b><span>Pendentes</span></div>
          <div class="pv-tag amarelo"><span>Em Aberto</span><b id="pv-aberto"></b></div>
        </div>
      </div>
      <div class="pv-graf"><h3>Realizado x Meta — Acumulado</h3><div class="pv-cv"><canvas id="pv-canvas"></canvas></div></div>
      <div class="pv-tab">
        <div class="pv-tab-topo"><h3>Preventivas não executadas (<span id="pv-qtd-tab">0</span>)</h3>
          <input id="pv-busca" type="search" placeholder="Buscar site, WO, município..."></div>
        <div class="pv-rolagem"><table>
          <thead><tr><th>Status</th><th>Site</th><th>Site Infratel</th><th>Município</th><th>Tipo</th>
          <th>Tipologia</th><th>Equipe</th><th>Executor</th><th>WO</th></tr></thead>
          <tbody id="pv-tbody"></tbody></table></div>
      </div>
    </div>`;

    raiz.querySelectorAll('.pv-f select').forEach(s =>
      s.addEventListener('change', () => { sel[s.dataset.k] = s.value; atualizar(); }));
    raiz.querySelector('#pv-busca').addEventListener('input', e => { busca = e.target.value.toLowerCase(); tabela(filtrar()); });
  }

  function filtrar(ignorar) {
    return dados.registros.filter(r => FILTROS.every(f => f.k === ignorar || !sel[f.k] || r[f.k] === sel[f.k]));
  }

  // Opções de cada filtro respeitam os demais filtros (como slicers do Power BI)
  function opcoes() {
    raiz.querySelectorAll('.pv-f select').forEach(s => {
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
    const dias = diasDoMes(dados.mes);
    const ref = dataReferencia(dados.mes);
    const duMes = dias.filter(d => d.util).length || 1;
    const duRef = dias.filter(d => d.util && d.data <= ref).length;

    const planejado = regs.length;
    const planD1 = Math.round(planejado * duRef / duMes);
    const executadas = regs.filter(r => r.st === 'Executada');
    const realTotal = executadas.length;   // igual ao Power BI: todo o realizado lançado
    const realD1 = realTotal;
    const pendentes = regs.filter(r => r.st !== 'Executada').length;
    const aberto = regs.filter(r => r.st === 'Aberta').length;
    const saldo = realD1 - planD1;
    const ader = planD1 ? realD1 / planD1 * 100 : 0;

    const $ = id => raiz.querySelector('#' + id);
    $('pv-plan').textContent = fmt(planejado);
    $('pv-du-mes').textContent = duMes;
    $('pv-prog').textContent = `${fmt(realD1)} / ${fmt(planD1)}`;
    $('pv-du').textContent = `${duRef} / ${duMes}`;
    $('pv-barra').style.width = Math.min(100, planD1 ? realD1 / planD1 * 100 : 0) + '%';
    $('pv-entrega').textContent = pct(planejado ? realTotal / planejado * 100 : 0);
    $('pv-pd1').textContent = pct(duRef / duMes * 100);
    $('pv-saldo').textContent = (saldo > 0 ? '+' : '') + fmt(saldo);
    $('pv-ader').textContent = pct(ader);
    $('pv-ader-box').className = 'pv-tag ' + (ader >= 95 ? 'verde' : ader >= 90 ? 'laranja' : 'vermelho');
    $('pv-pend').textContent = fmt(pendentes);
    $('pv-aberto').textContent = fmt(aberto);

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
        { label: 'Planejado Acumulado', data: plan, borderColor: '#1e88e5', backgroundColor: '#1e88e5',
          borderDash: [3, 4], borderWidth: 2, pointStyle: 'rect', pointRadius: 4, tension: 0,
          datalabels: { align: 'top', color: '#1b1f2a' } },
        { label: 'Realizado Acumulado', data: real, borderColor: '#0b1f8f', backgroundColor: '#0b1f8f',
          borderWidth: 3, pointRadius: 4, tension: 0.25, spanGaps: false,
          datalabels: { align: 'bottom', color: '#1b1f2a' } }
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
            font: { size: 11 }, formatter: v => fmt(v)
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
      [r.s, r.si, r.w, r.m, r.t, r.e, r.x, r.tp].some(v => String(v ?? '').toLowerCase().includes(busca)));
    raiz.querySelector('#pv-qtd-tab').textContent = fmt(lista.length);
    raiz.querySelector('#pv-tbody').innerHTML = lista.map(r => `<tr>
      <td><span class="pv-st ${esc(r.st)}">${esc(r.st)}</span></td><td>${esc(r.s)}</td><td>${esc(r.si)}</td>
      <td>${esc(r.m)}</td><td>${esc(r.t)}</td><td>${esc(r.tp)}</td><td>${esc(r.e)}</td><td>${esc(r.x)}</td><td>${esc(r.w)}</td>
    </tr>`).join('') || '<tr><td colspan="9" style="text-align:center;color:#5b6275">Nenhuma preventiva pendente</td></tr>';
  }

  async function montar(idContainer) {
    raiz = document.getElementById(idContainer || CONTAINER_PADRAO);
    if (!raiz) return;
    estilos();
    try {
      dados = window.DADOS_PREVENTIVA ||
        await fetch(ARQUIVO_JSON + '?v=' + Date.now(), { cache: 'no-store' }).then(r => { if (!r.ok) throw r.status; return r.json(); });
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
