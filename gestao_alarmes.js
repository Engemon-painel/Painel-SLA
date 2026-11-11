// ======================================================================
// ACOMPANHAMENTO DE ALARMES — Preventiva x Quedas (incidentes medidos no ESER)
// ----------------------------------------------------------------------
// POR SITE: cada dia em que o site teve preventiva EXECUTADA (qualquer tipo,
// dados_preventiva.json) vira uma linha, e compara as quedas
// (incidentes de infra do dados.json → incidentesPorSite, dia/hora do SUBMIT_DATE; base medida
// no ESER) do mesmo site:
//   • ANTES  = N dias antes da preventiva
//   • NO DIA = no dia da preventiva
//   • APÓS   = N dias depois (até o último dia da base)
// Resultado (por queda/dia, para não punir quem ainda tem poucos dias de após):
//   🟢 Melhorou (menos quedas depois) · 🔴 Piorou (mais quedas depois) ·
//   ➖ Igual · ⚪ Sem quedas · ⏳ Aguardando (ainda não há dias depois)
// Precisa do campo "data" em incidentesPorSite (Office Script do painel,
// 11/11/2026). Sem ele, usa só os incidentes FCA (fora do prazo) — parcial.
// No index.html: <script src="gestao_alarmes.js"></script> antes de </body>
// ======================================================================
(function(){
  const RES = {
    piorou:     { label:'Piorou',     icone:'🔴', cor:'#dc2626', bg:'rgba(220,38,38,.12)' },
    melhorou:   { label:'Melhorou',   icone:'🟢', cor:'#059669', bg:'rgba(5,150,105,.12)' },
    igual:      { label:'Igual',      icone:'➖', cor:'#d97706', bg:'rgba(217,119,6,.12)' },
    sem:        { label:'Sem quedas', icone:'⚪', cor:'#6b7590', bg:'rgba(107,117,144,.12)' },
    aguardando: { label:'Aguardando', icone:'⏳', cor:'#2563eb', bg:'rgba(37,99,235,.12)' }
  };
  const ORDEM_RES = ['piorou','melhorou','igual','sem','aguardando'];

  let PREV = [], VIS = [], INC = [], baseAte = '', fonteInc = '', erro = '';
  const filtro = { base: 'todas', janela: 7, mes: '', tipo: '', area: '', res: '', q: '' };
  let chartTipo = null, linhasAtuais = [];

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const normSite = s => String(s || '').toUpperCase().replace(/[\s\-_.]/g, '');
  const dataBr = iso => { if(!iso) return ''; const p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; };
  const diaNum = iso => { const p = String(iso).split('-'); return Date.UTC(+p[0], +p[1]-1, +p[2]) / 86400000; };
  const MES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const rotMes = m => { const p = m.split('-'); return MES_PT[+p[1]-1] + '/' + p[0].slice(2); };
  const horaDe = v => { const m = /(\d{1,2}):(\d{2})/.exec(String(v || '')); return m ? m[1].padStart(2,'0') + ':' + m[2] : ''; };
  function isoDe(v){
    const s = String(v || '').trim();
    let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); if(m) return m[1] + '-' + m[2] + '-' + m[3];
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s); if(m) return m[3] + '-' + m[2].padStart(2,'0') + '-' + m[1].padStart(2,'0');
    return '';
  }

  const CSS = `
    #pagina-alarmes .al-filtros{ margin:10px 0 16px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
    #pagina-alarmes .al-lbl{ font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em; font-family:'JetBrains Mono', monospace; }
    #pagina-alarmes select{ font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); font-family:'JetBrains Mono', monospace; }
    #alKpiRow{ grid-template-columns: repeat(auto-fit, minmax(130px,1fr)); }
    #alKpiRow .kpi{ cursor:pointer; }
    #alKpiRow .kpi.sel{ outline:2px solid var(--accent); }
    #alTable th, #alTable td{ white-space:nowrap; }
    #alTable td{ padding:7px 6px; }
    #alTable thead th{ position:sticky; top:0; background:#eef1f5; z-index:2; }
    #alTable td.n{ text-align:center; font-family:'JetBrains Mono', monospace; font-weight:600; }
    #alTable .pill-al{ display:inline-block; padding:3px 9px; border-radius:999px; font-size:12px; font-weight:600; }
    #alAviso{ padding:10px 12px; border-radius:10px; margin:8px 0 12px; font-size:13px; display:none; }
  `;

  const TEMPLATE = `
    <div class="panel">
      <h2>Acompanhamento de Alarmes — Site x Quedas após a preventiva</h2>
      <div class="hint">Por <b>site</b>: no dia em que o site teve preventiva (qualquer tipo — clima, energia, gerador...), <b>①</b> se teve queda de <b>indisponibilidade</b> (base do ESER) <b>no dia</b> ou <b>depois</b>; <b>②</b> quantas quedas tinha <b>antes</b>, <b>no dia</b> e <b>depois</b> (indisponibilidade + outros, ou só indisponibilidade). 🔴 piorou = mais quedas depois · 🟢 melhorou = menos quedas depois (comparado por queda/dia).</div>
      <div class="hint mono" id="alInfo">Carregando...</div>
      <div id="alAviso"></div>

      <div class="al-filtros">
        <label class="al-lbl" for="alBase">Comparar com</label>
        <select id="alBase"><option value="todas">Indisponibilidade + Outros</option><option value="I">Só Indisponibilidade (ESER)</option></select>
        <label class="al-lbl" for="alJanela" style="margin-left:8px;">Janela</label>
        <select id="alJanela"><option value="7">7 dias antes x 7 depois</option><option value="15">15 dias antes x 15 depois</option><option value="30">30 dias antes x 30 depois</option></select>
        <label class="al-lbl" for="alMes" style="margin-left:8px;">Mês da preventiva</label>
        <select id="alMes" style="min-width:110px;"></select>
        <label class="al-lbl" for="alTipo" style="margin-left:8px;">Teve preventiva de</label>
        <select id="alTipo" style="min-width:140px;"></select>
        <label class="al-lbl" for="alArea" style="margin-left:8px;">Área</label>
        <select id="alArea" style="min-width:120px;"></select>
        <label class="al-lbl" for="alRes" style="margin-left:8px;">Resultado</label>
        <select id="alRes" style="min-width:150px;"></select>
      </div>

      <div class="kpi-row" id="alKpiRow"></div>

      <div class="grid" style="grid-template-columns: 1.2fr 1fr;">
        <div class="panel">
          <h2>Resultado por mês</h2>
          <div class="hint">Sites com preventiva no mês que tiveram queda antes ou depois.</div>
          <div style="position:relative; height:290px;"><canvas id="alTipoChart"></canvas></div>
        </div>
        <div class="panel">
          <h2>Sites que mais pioraram</h2>
          <div class="hint">Diferença de quedas (depois − antes), somando as preventivas do site.</div>
          <div id="alPioresWrap" style="max-height:290px; overflow:auto;"></div>
        </div>
      </div>

      <div class="panel">
        <h2>Sites com preventiva</h2>
        <div class="al-filtros">
          <div class="search" title="Site, WO, município, equipe ou técnico">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
            <input id="alBusca" type="search" placeholder="Site, WO, município, técnico..." aria-label="Buscar" style="min-width:230px;" />
          </div>
          <span id="alContagem" class="mono" style="font-size:11px; color:var(--muted);"></span>
          <button type="button" class="link-base" id="alExportar" style="margin-left:auto; cursor:pointer; border:1px solid var(--accent); font-family:'JetBrains Mono', monospace;">⬇ Exportar (CSV)</button>
        </div>
        <div style="overflow:auto; max-height:70vh; border-radius:8px;">
          <table id="alTable">
            <thead><tr>
              <th>Data da preventiva</th><th>Site</th><th>Preventivas no dia</th>
              <th class="num" style="background:#fde8e8;" title="① Indisponibilidade (base do ESER) no dia da preventiva">① Indisp.<br>no dia</th><th class="num" style="background:#fde8e8;" id="alThIApos" title="① Indisponibilidade (base do ESER) depois da preventiva">① Indisp.<br>depois</th>
              <th class="num" id="alThAntes">② Antes</th><th class="num">② No dia</th><th class="num" id="alThApos">② Depois</th><th>Resultado</th>
              <th>Município</th><th>Área</th><th>Equipe / Técnico</th><th>WO</th>
            </tr></thead>
            <tbody id="alTbody"></tbody>
          </table>
          <div id="alVazio" class="no-results" style="display:none">Nenhuma preventiva encontrada.</div>
        </div>
      </div>
    </div>
  `;

  function montarEstrutura(){
    if(document.getElementById('pagina-alarmes')) return;
    const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);
    const nav = document.createElement('div');
    nav.className = 'nav-item'; nav.id = 'nav-alarmes';
    nav.innerHTML = '<span class="nav-icon">🚨</span> Acompanhamento de Alarmes';
    nav.onclick = () => mostrarPagina('alarmes');
    const ancora = document.getElementById('nav-preventiva');
    if(ancora) ancora.after(nav); else document.querySelector('.sidebar')?.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina'; pg.id = 'pagina-alarmes'; pg.style.display = 'none';
    pg.innerHTML = TEMPLATE;
    const irmao = document.getElementById('pagina-preventiva');
    if(irmao && irmao.parentNode) irmao.parentNode.appendChild(pg);
    else document.querySelector('main.main-content')?.appendChild(pg);

    if(typeof PAGINAS !== 'undefined' && PAGINAS.indexOf('alarmes') === -1) PAGINAS.push('alarmes');
    const mostrarOriginal = window.mostrarPagina;
    window.mostrarPagina = function(nome){
      mostrarOriginal(nome);
      if(nome === 'alarmes') render();
    };

    const liga = (id, campo, num) => document.getElementById(id).addEventListener('change', e => { filtro[campo] = num ? +e.target.value : e.target.value; render(); });
    liga('alJanela','janela',true); liga('alBase','base'); liga('alMes','mes'); liga('alTipo','tipo'); liga('alArea','area'); liga('alRes','res');
    let t;
    document.getElementById('alBusca').addEventListener('input', e => {
      clearTimeout(t); t = setTimeout(() => { filtro.q = e.target.value.trim().toLowerCase(); renderTabela(); }, 200);
    });
    document.getElementById('alExportar').addEventListener('click', exportarCsv);
  }

  // ---------------------------------------------------------------- dados
  async function carregar(){
    try{
      const [dPrev, dSla] = await Promise.all([
        fetch('dados_preventiva.json?v=' + Date.now(), { cache:'no-store' }).then(r => { if(!r.ok) throw new Error('dados_preventiva.json ' + r.status); return r.json(); }),
        fetch('dados.json?v=' + Date.now(), { cache:'no-store' }).then(r => { if(!r.ok) throw new Error('dados.json ' + r.status); return r.json(); })
      ]);
      // preventivas: mês atual + meses guardados (sem repetir WO)
      const vistos = new Set(); PREV = [];
      const add = lista => (lista || []).forEach(r => {
        if(!r || String(r.st || '').toLowerCase() !== 'executada') return;
        const d = isoDe(r.d); if(!d) return;
        const k = (r.w || '') + '|' + d + '|' + (r.s || '');
        if(vistos.has(k)) return; vistos.add(k);
        PREV.push({ wo:r.w || '', site:r.s || '', siteId:r.si || '', mun:r.m || '', tipo:r.t || '', area:r.x || '', equipe:r.e || '', tec:r.tec || '', data:d });
      });
      add(dPrev.registros);
      (dPrev.meses || []).forEach(m => add(m.registros));
      // POR SITE: todas as preventivas do mesmo site no mesmo dia viram UMA visita
      // (não importa se foi clima, energia, gerador...)
      const vis = {};
      PREV.forEach(p => {
        const k = normSite(p.site) + '|' + p.data;
        const v = vis[k] = vis[k] || { site:p.site, siteId:p.siteId, mun:p.mun, area:p.area, data:p.data, tipos:[], tecs:[], wos:[], equipes:[] };
        const pus = (arr, x) => { if(x && arr.indexOf(x) === -1) arr.push(x); };
        pus(v.tipos, p.tipo); pus(v.tecs, p.tec); pus(v.wos, p.wo); pus(v.equipes, p.equipe);
      });
      VIS = Object.values(vis).map(v => Object.assign(v, { tipo:v.tipos.join(', '), tec:v.tecs.join(', '), wo:v.wos.join(', '), equipe:v.equipes.join(', ') }));

      // quedas:
      //  1) quedasSite  → aba Historico inteira: Indisponibilidade (t="I") + Outros (t="O")
      //  2) incidentesPorSite com "data" → só Indisponibilidade (base do ESER)
      //  3) pendentesFca → só os fora do prazo (parcial)
      const lista = Array.isArray(dSla.incidentesPorSite) ? dSla.incidentesPorSite : [];
      const comData = lista.filter(x => x && x.data);
      const eserPorInc = {}; lista.forEach(x => { if(x && x.incidente) eserPorInc[x.incidente] = x.eser4h || ''; });
      if(Array.isArray(dSla.quedasSite) && dSla.quedasSite.length){
        fonteInc = 'todas';
        INC = dSla.quedasSite.map(x => ({ site:normSite(x.s), data:isoDe(x.d), hora:horaDe(x.d), inc:x.i || '', t:x.t === 'I' ? 'I' : 'O', eser:eserPorInc[x.i] || '' }));
      }else if(comData.length){
        fonteInc = 'indisp';
        INC = comData.map(x => ({ site:normSite(x.site), data:isoDe(x.data), hora:horaDe(x.data), inc:x.incidente || '', t:'I', eser:x.eser4h || '' }));
      }else{
        fonteInc = 'fca';
        INC = (dSla.pendentesFca || []).map(x => ({ site:normSite(x.site), data:isoDe(x.submit || x.reported || ''), hora:horaDe(x.submit || x.reported || ''), inc:x.incidente || '', t:'I', eser:'Fora' }));
      }
      const vi = new Set();
      INC = INC.filter(x => { if(!x.site || !x.data) return false; const k = x.inc || (x.site + x.data); if(vi.has(k)) return false; vi.add(k); return true; });
      baseAte = isoDe(dSla.dadosAte) || INC.reduce((m, x) => x.data > m ? x.data : m, '');
      erro = '';
      popularFiltros();
    }catch(e){ erro = e.message; console.error('[Alarmes]', e); }
    const pg = document.getElementById('pagina-alarmes');
    if(pg && pg.style.display !== 'none') render();
  }

  function popularFiltros(){
    const unicos = campo => Array.from(new Set(PREV.map(p => p[campo]).filter(Boolean))).sort((a,b) => a.localeCompare(b,'pt-BR'));
    const preencher = (id, valores, campo, rot) => {
      const sel = document.getElementById(id);
      if(filtro[campo] && valores.indexOf(filtro[campo]) === -1) filtro[campo] = '';
      sel.innerHTML = '<option value="">Todos</option>' + valores.map(v => `<option value="${esc(v)}">${esc(rot ? rot(v) : v)}</option>`).join('');
      sel.value = filtro[campo];
    };
    const meses = Array.from(new Set(PREV.map(p => p.data.slice(0,7)))).sort();
    preencher('alMes', meses, 'mes', rotMes);
    preencher('alTipo', unicos('tipo'), 'tipo');
    preencher('alArea', unicos('area'), 'area');
    const sr = document.getElementById('alRes');
    sr.innerHTML = '<option value="">Todos</option>' + ORDEM_RES.map(k => `<option value="${k}">${RES[k].icone} ${RES[k].label}</option>`).join('');
    sr.value = filtro.res;
    document.getElementById('alJanela').value = String(filtro.janela);
    document.getElementById('alBase').value = filtro.base;
  }

  // ---------------------------------------------------------------- cálculo
  function calcular(){
    const N = filtro.janela;
    const porSite = {};
    INC.forEach(x => { (porSite[x.site] = porSite[x.site] || []).push(x); });
    const limite = baseAte ? diaNum(baseAte) : null;
    const naBase = x => filtro.base === 'todas' || x.t === 'I';
    return VIS.map(p => {
      const dp = diaNum(p.data);
      const lst = (porSite[normSite(p.site)] || []).concat(p.siteId && normSite(p.siteId) !== normSite(p.site) ? (porSite[normSite(p.siteId)] || []) : []);
      const antes = [], dia = [], depois = [], iDia = [], iDepois = [];
      lst.forEach(x => {
        const d = diaNum(x.data), dif = d - dp;
        // 1) ESER – só Indisponibilidade: teve queda no dia / depois da preventiva?
        if(x.t === 'I'){ if(dif === 0) iDia.push(x); else if(dif > 0 && dif <= N) iDepois.push(x); }
        // 2) comparação antes x no dia x depois (Indisponibilidade + Outros, ou só Indisponibilidade)
        if(!naBase(x)) return;
        if(dif < 0 && dif >= -N) antes.push(x);
        else if(dif === 0) dia.push(x);
        else if(dif > 0 && dif <= N) depois.push(x);
      });
      const diasDepois = limite === null ? N : Math.max(0, Math.min(N, limite - dp));
      const tAntes = antes.length / N;
      const tDepois = diasDepois > 0 ? depois.length / diasDepois : 0;
      let res;
      if(diasDepois === 0) res = 'aguardando';
      else if(!antes.length && !depois.length) res = 'sem';
      else if(tDepois > tAntes + 1e-9) res = 'piorou';
      else if(tDepois < tAntes - 1e-9) res = 'melhorou';
      else res = 'igual';
      return Object.assign({}, p, { antes, dia, depois, iDia, iDepois, diasDepois, res });
    });
  }

  function filtradas(todas, ignorarRes){
    return todas.filter(l =>
      (!filtro.mes || l.data.slice(0,7) === filtro.mes) &&
      (!filtro.tipo || l.tipos.indexOf(filtro.tipo) !== -1) &&
      (!filtro.area || l.area === filtro.area) &&
      (ignorarRes || !filtro.res || l.res === filtro.res));
  }

  // ---------------------------------------------------------------- telas
  function render(){
    const info = document.getElementById('alInfo');
    const aviso = document.getElementById('alAviso');
    if(!info) return;
    if(erro){ info.textContent = 'Não foi possível carregar: ' + erro; return; }
    if(!PREV.length && !INC.length){ info.textContent = 'Carregando...'; return; }
    const N = filtro.janela;
    info.textContent = `${VIS.length} sites com preventiva (${PREV.length} preventivas executadas) · ${INC.length} quedas na base (${INC.filter(x => x.t === 'I').length} indisponibilidade, ${INC.filter(x => x.t === 'O').length} outros) · base de incidentes até ${dataBr(baseAte)} · janela de ${N} dias`;
    if(fonteInc === 'indisp'){
      aviso.style.display = '';
      aviso.style.background = 'rgba(37,99,235,.10)'; aviso.style.color = '#1e3a8a';
      aviso.innerHTML = 'ℹ O dados.json ainda não traz os incidentes <b>Outros</b> — por enquanto tudo é só <b>Indisponibilidade</b>. Atualize o Office Script do painel (versão de 11/11 com "quedasSite") e rode o fluxo.';
    }else if(fonteInc === 'fca'){
      aviso.style.display = '';
      aviso.style.background = 'rgba(217,119,6,.12)'; aviso.style.color = '#92400e';
      aviso.innerHTML = '⚠ <b>Parcial:</b> o dados.json ainda não traz o dia de cada incidente, então aqui entram só os incidentes <b>fora do prazo</b> (FCA). Atualize o Office Script do painel (versão de 11/11) e rode o fluxo para usar a base completa de incidentes medidos no ESER.';
    }else aviso.style.display = 'none';
    document.getElementById('alThAntes').textContent = `② Antes (${N}d)`;
    document.getElementById('alThApos').textContent = `② Depois (${N}d)`;
    document.getElementById('alThIApos').innerHTML = `① Indisp.<br>depois (${N}d)`;

    const todas = calcular();
    const semRes = filtradas(todas, true);
    linhasAtuais = filtradas(todas, false);

    const cont = {}; ORDEM_RES.forEach(k => cont[k] = 0);
    semRes.forEach(l => cont[l.res]++);
    const comQuedaDia = semRes.filter(l => l.iDia.length || l.iDepois.length).length;
    const kpi = (k, titulo, valor, cor, sub) => `
      <div class="kpi${filtro.res === k ? ' sel' : ''}" data-res="${k}" title="Clique para filtrar">
        <div class="label">${titulo}</div>
        <div class="value" style="color:${cor}">${valor}</div>
        ${sub ? `<div class="delta" style="color:var(--muted)">${sub}</div>` : ''}
      </div>`;
    const pct = n => semRes.length ? Math.round(n / semRes.length * 100) + '%' : '';
    document.getElementById('alKpiRow').innerHTML =
      kpi('', '📡 Sites com preventiva', semRes.length, 'var(--text)', '') +
      kpi('dia', '⚡ Indisp. no dia ou depois', comQuedaDia, '#dc2626', pct(comQuedaDia)) +
      ORDEM_RES.map(k => kpi(k, RES[k].icone + ' ' + RES[k].label, cont[k], RES[k].cor, pct(cont[k]))).join('');
    document.querySelectorAll('#alKpiRow .kpi').forEach(el => el.addEventListener('click', () => {
      const k = el.dataset.res;
      filtro.res = (k === 'dia' || filtro.res === k) ? (k === 'dia' && filtro.res !== 'dia' ? 'dia' : '') : k;
      document.getElementById('alRes').value = RES[filtro.res] ? filtro.res : '';
      render();
    }));
    if(filtro.res === 'dia') linhasAtuais = semRes.filter(l => l.iDia.length || l.iDepois.length);

    desenharGrafico(semRes);
    desenharPiores(semRes);
    renderTabela();
  }

  function desenharGrafico(lista){
    const cv = document.getElementById('alTipoChart');
    if(!cv || typeof Chart === 'undefined') return;
    const meses = Array.from(new Set(lista.map(l => l.data.slice(0,7)))).sort();
    const ks = ['piorou','igual','melhorou'];
    const ds = ks.map(k => ({
      label: RES[k].icone + ' ' + RES[k].label, backgroundColor: RES[k].cor, borderRadius: 4,
      data: meses.map(m => lista.filter(l => l.data.slice(0,7) === m && l.res === k).length)
    }));
    if(chartTipo) chartTipo.destroy();
    chartTipo = new Chart(cv, {
      type:'bar',
      data:{ labels: meses.map(rotMes), datasets: ds },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ position:'bottom' } },
        scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true, ticks:{ precision:0 } } } }
    });
  }

  function desenharPiores(lista){
    const porSite = {};
    lista.filter(l => l.res === 'piorou').forEach(l => {
      const s = porSite[l.site] = porSite[l.site] || { site:l.site, mun:l.mun, prev:0, antes:0, depois:0, tipos:new Set() };
      s.prev++; s.antes += l.antes.length; s.depois += l.depois.length; l.tipos.forEach(t => s.tipos.add(t));
    });
    const top = Object.values(porSite).sort((a,b) => (b.depois - b.antes) - (a.depois - a.antes) || b.depois - a.depois).slice(0, 15);
    const w = document.getElementById('alPioresWrap');
    if(!top.length){ w.innerHTML = '<div class="hint" style="padding:12px;">Nenhum site piorou no filtro atual. 🎉</div>'; return; }
    w.innerHTML = `<table style="width:100%;"><thead><tr><th>Site</th><th>Município</th><th>Tipo</th><th class="num">Antes</th><th class="num">Depois</th><th class="num">+</th></tr></thead><tbody>` +
      top.map(s => `<tr style="cursor:pointer;" data-site="${esc(s.site)}"><td><b>${esc(s.site)}</b></td><td>${esc(s.mun)}</td><td>${esc(Array.from(s.tipos).join(', '))}</td>
        <td class="num">${s.antes}</td><td class="num" style="color:#dc2626; font-weight:700;">${s.depois}</td><td class="num" style="color:#dc2626;">+${s.depois - s.antes}</td></tr>`).join('') +
      '</tbody></table>';
    w.querySelectorAll('tr[data-site]').forEach(tr => tr.addEventListener('click', () => {
      const b = document.getElementById('alBusca'); b.value = tr.dataset.site; filtro.q = tr.dataset.site.toLowerCase(); renderTabela();
      b.scrollIntoView({ behavior:'smooth', block:'center' });
    }));
  }

  function linhasTabela(){
    const q = filtro.q;
    return linhasAtuais.filter(l => !q || [l.site, l.siteId, l.wo, l.mun, l.equipe, l.tec].some(v => String(v || '').toLowerCase().indexOf(q) !== -1))
      .sort((a,b) => ORDEM_RES.indexOf(a.res) - ORDEM_RES.indexOf(b.res) || (b.depois.length - b.antes.length) - (a.depois.length - a.antes.length) || b.data.localeCompare(a.data));
  }

  function renderTabela(){
    const tb = document.getElementById('alTbody'); if(!tb) return;
    const ls = linhasTabela();
    document.getElementById('alContagem').textContent = ls.length + ' site(s) / dia(s) de preventiva';
    document.getElementById('alVazio').style.display = ls.length ? 'none' : '';
    const lista = arr => arr.map(x => dataBr(x.data) + (x.hora ? ' ' + x.hora : '') + ' · ' + x.inc + (x.t === 'O' ? ' (Outros)' : ' (Indisp.' + (x.eser ? ' · ' + x.eser : '') + ')')).join('\n');
    const cel = (arr, cor, extra) => `<td class="n" title="${esc(lista(arr))}" style="color:${arr.length ? cor : 'var(--muted)'};">${arr.length}${extra || ''}</td>`;
    tb.innerHTML = ls.slice(0, 1500).map(l => {
      const r = RES[l.res];
      const obs = l.res !== 'aguardando' && l.diasDepois < filtro.janela ? `<span style="font-size:10px; color:var(--muted); font-weight:400;"> /${l.diasDepois}d</span>` : '';
      return `<tr>
        <td>${dataBr(l.data)}</td><td><b>${esc(l.site)}</b>${l.siteId && l.siteId !== l.site ? `<span style="font-size:11px; color:var(--muted);"> ${esc(l.siteId)}</span>` : ''}</td>
        <td>${esc(l.tipo)}</td>
        ${cel(l.iDia, '#dc2626', l.iDia.length ? ' ⚡' : '')}${cel(l.iDepois, '#dc2626')}
        ${cel(l.antes, 'var(--text)')}${cel(l.dia, '#d97706')}${cel(l.depois, l.res === 'piorou' ? '#dc2626' : 'var(--text)', obs)}
        <td><span class="pill-al" style="background:${r.bg}; color:${r.cor};">${r.icone} ${r.label}</span></td>
        <td>${esc(l.mun)}</td><td>${esc(l.area)}</td><td>${esc(l.tec || l.equipe)}</td><td class="mono" style="font-size:11px;" title="${esc(l.wo)}">${esc(l.wos.length > 1 ? l.wos.length + ' WOs' : l.wo)}</td>
      </tr>`;
    }).join('') + (ls.length > 1500 ? `<tr><td colspan="13" class="hint">Mostrando 1500 de ${ls.length}. Use os filtros ou exporte o CSV.</td></tr>` : '');
  }

  function exportarCsv(){
    const N = filtro.janela;
    const cab = ['Data preventiva','Site','Site ID','Municipio','Preventivas no dia','Area','Equipe','Tecnico','WOs','Indisp. no dia',`Indisp. depois (${N}d)`,`Quedas antes (${N}d)`,'Quedas no dia',`Quedas depois (${N}d)`,'Dias depois na base','Resultado','Incidentes antes','Incidentes no dia','Incidentes depois'];
    const q = v => '"' + String(v ?? '').replace(/"/g,'""') + '"';
    const inc = arr => arr.map(x => dataBr(x.data) + (x.hora ? ' ' + x.hora : '') + ' ' + x.inc + (x.t === 'O' ? ' (Outros)' : '')).join(' | ');
    const linhas = linhasTabela().map(l => [dataBr(l.data), l.site, l.siteId, l.mun, l.tipo, l.area, l.equipe, l.tec, l.wo,
      l.iDia.length, l.iDepois.length, l.antes.length, l.dia.length, l.depois.length, l.diasDepois, RES[l.res].label, inc(l.antes), inc(l.dia), inc(l.depois)].map(q).join(';'));
    const blob = new Blob(['﻿' + cab.map(q).join(';') + '\n' + linhas.join('\n')], { type:'text/csv;charset=utf-8' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = 'acompanhamento_alarmes_por_site.csv'; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  function iniciar(){ montarEstrutura(); carregar(); }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar); else iniciar();
})();
