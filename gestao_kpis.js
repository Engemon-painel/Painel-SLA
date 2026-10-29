// =====================================================================
// gestao_kpis.js — página "KPIs de Pessoas" do Painel
//
// Junta em uma página só os três indicadores de pessoas:
//   • Turnover      → calculado com o Organograma (dados.json → colaboradores)
//                     Clássico     = ((Admissões + Desligamentos) / 2) / Headcount × 100
//                     Só desligados = Desligamentos / Headcount × 100
//                     Headcount = colaboradores ativos no último dia do mês
//                     (no mês atual, ativos hoje).
//   • Absenteísmo   → (Horas de faltas / Horas trabalháveis) × 100
//                     dados_pend_bh.json → absenteismo (folha atual, espelho de ponto)// =====================================================================
// gestao_kpis.js — página "KPIs de Pessoas" do Painel
//
// Junta em uma página só os três indicadores de pessoas:
//   • Turnover      → calculado com o Organograma (dados.json → colaboradores)
//                     Clássico     = ((Admissões + Desligamentos) / 2) / Headcount × 100
//                     Só desligados = Desligamentos / Headcount × 100
//                     Headcount = colaboradores ativos no último dia do mês
//                     (no mês atual, ativos hoje).
//   • Absenteísmo   → (Horas de faltas / Horas trabalháveis) × 100
//                     dados_pend_bh.json → absenteismo (folha atual, espelho de ponto)
//                     Horas de faltas = horas da jornada dos dias "Ausente"
//                     Horas trabalháveis = horas da jornada dos dias de escala já passados
//   • Banco de horas → dados_bh.json (saldo atual) + dados.json → bancoHoras (histórico)
//
// Instalação: no index.html, junto dos outros módulos:
//   <script src="gestao_kpis.js?v=1"></script>
// O item do menu, a página e o card da tela inicial aparecem sozinhos.
// =====================================================================
(function(){
  const C = (typeof cor !== 'undefined') ? cor : {
    good:'#059669', warn:'#d97706', bad:'#dc2626', accent:'#0e7c86', muted:'#6b7590', line:'#e1e5f0', text:'#1b2440'
  };
  const AZUL = '#0B2545', VERDE = '#059669', VERMELHO = '#dc2626', LARANJA = '#d97706';
  const MES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

  // colaboradores que estão em outro centro de custo, mas contam como Claro (igual ao Banco de Horas)
  const PROJ_FORCADO = { 'LIVIA MARIA VIEIRA YADA': 'CLARO', 'LEONARDO PERES FERRAZ': 'CLARO' };

  let COLAB = [], BH_HIST = [], BH_ATUAL = [], ABS = [], FOLHA = '', FOLHA_INI = '';
  const filtro = { pj: '', mes: '', modo: 'classico' };
  const charts = {};

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const pad2 = n => n < 10 ? '0' + n : String(n);
  const nomeChave = n => String(n || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
  function projForcado(n){
    const k = nomeChave(n); if (PROJ_FORCADO[k]) return PROJ_FORCADO[k];
    const f = k.length >= 15 && Object.keys(PROJ_FORCADO).find(x => x.startsWith(k) || k.startsWith(x));
    return f ? PROJ_FORCADO[f] : '';
  }
  const projDe = txt => { const t = nomeChave(txt); return t.includes('CIRION') ? 'CIRION' : (t.includes('CLARO') ? 'CLARO' : ''); };
  const passa = (nome, txtProj) => !filtro.pj || (projForcado(nome) || projDe(txtProj)) === filtro.pj;

  const pct = v => (isFinite(v) ? v : 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
  const num = v => Math.round(v).toLocaleString('pt-BR');
  function horas(v){
    const neg = v < 0, a = Math.abs(v), h = Math.floor(a), m = Math.round((a - h) * 60);
    const hh = m === 60 ? h + 1 : h, mm = m === 60 ? 0 : m;
    return (neg ? '-' : '') + hh.toLocaleString('pt-BR') + 'h' + (mm ? pad2(mm) : '');
  }
  const hojeIso = () => { const d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); };
  const fimDoMes = ym => { const [y, m] = ym.split('-').map(Number); return ym + '-' + pad2(new Date(y, m, 0).getDate()); };
  const rotMes = ym => { const [y, m] = ym.split('-'); return MES_PT[Number(m) - 1] + '/' + y.slice(2); };
  const brParaIso = br => { const m = String(br || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : ''; };
  const isoParaBr = iso => { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + '/' + m[2] + '/' + m[1] : ''; };

  // ---- METAS (troque aqui) · verde = dentro da meta · amarelo = até 50% acima · vermelho = acima disso
  const META_TURNOVER = 5;      // % ao mês
  const META_ABS = 3;           // %
  const LIMITE_BH = 180;        // horas: acima disso o saldo é crítico

  const semaforo = (v, meta) => v <= meta ? VERDE : (v <= meta * 1.5 ? LARANJA : VERMELHO);

  // ---------- Menu + página ----------
  function montarEstrutura(){
    const sidebar = document.querySelector('.sidebar');
    const main = document.querySelector('.main-content');
    if(!sidebar || !main || document.getElementById('nav-kpis')) return;

    const nav = document.createElement('div');
    nav.className = 'nav-item';
    nav.id = 'nav-kpis';
    nav.innerHTML = '<span class="nav-icon">📊</span> KPIs de Pessoas';
    nav.onclick = abrir;
    const ref = document.getElementById('nav-turnover') || document.getElementById('nav-organograma');
    if(ref && ref.parentNode === sidebar) sidebar.insertBefore(nav, ref.nextSibling);
    else sidebar.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina';
    pg.id = 'pagina-kpis';
    pg.style.display = 'none';
    pg.innerHTML = `
      <div class="kx">
        <div class="kx-topo">
          <div>
            <div class="kx-eyebrow">Pessoas · Resumo executivo</div>
            <h2 class="kx-titulo">KPIs de Pessoas</h2>
            <div class="kx-ref" id="kxRef"></div>
          </div>
          <div class="kx-filtros">
            <div class="kx-seg" id="kxProj">
              <button type="button" data-pj="">Todos</button><button type="button" data-pj="CLARO">Claro</button><button type="button" data-pj="CIRION">Cirion</button>
            </div>
            <select id="kxMes" class="kx-sel" title="Mês de referência do turnover"></select>
          </div>
        </div>

        <div class="kx-tiles" id="kxTiles"></div>

        <div class="kx-destaques" id="kxDestaques"></div>

        <div class="kx-grade">
          <div class="kx-card">
            <div class="kx-card-topo">
              <h3>Turnover · 12 meses</h3>
              <div class="kx-seg kx-seg-p" id="kxModo"><button type="button" data-modo="classico">Clássico</button><button type="button" data-modo="deslig">Só desligados</button></div>
            </div>
            <div class="kx-sub" id="kxToFormula"></div>
            <div style="position:relative; height:240px;"><canvas id="kxToChart"></canvas></div>
          </div>
          <div class="kx-card">
            <div class="kx-card-topo"><h3>Absenteísmo por supervisor</h3></div>
            <div class="kx-sub" id="kxAbsSub"></div>
            <div id="kxAbsRank"></div>
            <details class="kx-det"><summary>Ver todos os supervisores</summary><div style="overflow-x:auto;"><table class="kx-tab" id="kxAbsTab"></table></div></details>
          </div>
          <div class="kx-card">
            <div class="kx-card-topo"><h3>Banco de horas · saldo por competência</h3></div>
            <div class="kx-sub">Saldo positivo e negativo da equipe no fechamento de cada folha (15 a 14)</div>
            <div style="position:relative; height:240px;"><canvas id="kxBhChart"></canvas></div>
          </div>
        </div>
        <div class="kx-rodape" id="kxRodape"></div>
      </div>`;
    main.appendChild(pg);

    const st = document.createElement('style');
    st.textContent = `
      #pagina-kpis .kx{ display:flex; flex-direction:column; gap:14px; }
      #pagina-kpis .kx-topo{ background:${AZUL}; color:#fff; border-radius:16px; padding:18px 22px; display:flex; justify-content:space-between; align-items:flex-end; gap:14px; flex-wrap:wrap; }
      #pagina-kpis .kx-eyebrow{ font-family:'JetBrains Mono', monospace; font-size:10.5px; letter-spacing:.18em; text-transform:uppercase; color:#8fb3d9; }
      #pagina-kpis .kx-titulo{ font-family:'Space Grotesk', sans-serif; font-size:24px; font-weight:700; color:#fff; margin:2px 0 2px; }
      #pagina-kpis .kx-ref{ font-size:12px; color:#c4d4e6; }
      #pagina-kpis .kx-filtros{ display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
      #pagina-kpis .kx-seg{ display:inline-flex; background:rgba(255,255,255,.1); border-radius:999px; padding:3px; }
      #pagina-kpis .kx-seg button{ border:0; background:transparent; color:#dbe6f2; font:600 12px 'Inter', sans-serif; padding:6px 14px; border-radius:999px; cursor:pointer; }
      #pagina-kpis .kx-seg button.ativo{ background:#fff; color:${AZUL}; }
      #pagina-kpis .kx-seg-p{ background:#eef1f6; }
      #pagina-kpis .kx-seg-p button{ color:#5b6478; font-size:11px; padding:4px 11px; }
      #pagina-kpis .kx-seg-p button.ativo{ background:${AZUL}; color:#fff; }
      #pagina-kpis .kx-sel{ font:600 12px 'Inter', sans-serif; padding:7px 12px; border-radius:999px; border:0; background:#fff; color:${AZUL}; cursor:pointer; }
      #pagina-kpis .kx-tiles{ display:grid; grid-template-columns:repeat(6, minmax(0,1fr)); gap:12px; }
      @media (max-width:1200px){ #pagina-kpis .kx-tiles{ grid-template-columns:repeat(3, minmax(0,1fr)); } }
      @media (max-width:640px){ #pagina-kpis .kx-tiles{ grid-template-columns:repeat(2, minmax(0,1fr)); } }
      #pagina-kpis .kx-tile{ background:var(--panel,#fff); border:1px solid var(--line); border-radius:14px; padding:14px 16px 12px; position:relative; overflow:hidden; }
      #pagina-kpis .kx-tile::before{ content:''; position:absolute; left:0; top:0; bottom:0; width:4px; background:var(--kx-cor, ${AZUL}); }
      #pagina-kpis .kx-tile .l{ font-size:10.5px; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:var(--muted); }
      #pagina-kpis .kx-tile .v{ font-family:'Space Grotesk', sans-serif; font-size:30px; font-weight:700; line-height:1.1; margin:6px 0 4px; color:var(--kx-cor, ${AZUL}); font-variant-numeric:tabular-nums; }
      #pagina-kpis .kx-tile .d{ font-size:11.5px; color:var(--muted); line-height:1.35; }
      #pagina-kpis .kx-tile .d b{ font-weight:700; }
      #pagina-kpis .kx-tile .meta{ display:inline-block; margin-top:6px; font:700 10px 'JetBrains Mono', monospace; padding:2px 8px; border-radius:999px; background:color-mix(in srgb, var(--kx-cor) 12%, transparent); color:var(--kx-cor); }
      #pagina-kpis .kx-destaques{ background:var(--panel,#fff); border:1px solid var(--line); border-radius:14px; padding:12px 18px; display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:16px; }
      @media (max-width:900px){ #pagina-kpis .kx-destaques{ grid-template-columns:1fr; } }
      #pagina-kpis .kx-dest{ display:flex; gap:10px; align-items:flex-start; font-size:12.5px; line-height:1.45; color:var(--text); }
      #pagina-kpis .kx-dest i{ flex:none; width:8px; height:8px; border-radius:50%; margin-top:6px; background:var(--kx-cor); }
      #pagina-kpis .kx-grade{ display:grid; grid-template-columns:1.25fr 1fr 1.25fr; gap:14px; }
      @media (max-width:1100px){ #pagina-kpis .kx-grade{ grid-template-columns:1fr; } }
      #pagina-kpis .kx-card{ background:var(--panel,#fff); border:1px solid var(--line); border-radius:14px; padding:14px 16px; min-width:0; }
      #pagina-kpis .kx-card-topo{ display:flex; justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap; }
      #pagina-kpis .kx-card h3{ font-family:'Space Grotesk', sans-serif; font-size:14px; font-weight:700; color:var(--text); margin:0; }
      #pagina-kpis .kx-sub{ font-size:11px; color:var(--muted); margin:3px 0 10px; }
      #pagina-kpis .kx-rk{ display:grid; grid-template-columns:minmax(0,1fr) 90px 52px; align-items:center; gap:8px; padding:6px 0; border-bottom:1px solid var(--line); font-size:12px; }
      #pagina-kpis .kx-rk:last-child{ border-bottom:0; }
      #pagina-kpis .kx-rk .n{ white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
      #pagina-kpis .kx-rk .n small{ display:block; color:var(--muted); font-size:10.5px; }
      #pagina-kpis .kx-rk .b{ height:8px; border-radius:99px; background:#eef1f6; overflow:hidden; }
      #pagina-kpis .kx-rk .b span{ display:block; height:100%; border-radius:99px; }
      #pagina-kpis .kx-rk .p{ text-align:right; font-weight:700; font-variant-numeric:tabular-nums; }
      #pagina-kpis .kx-det{ margin-top:8px; font-size:12px; }
      #pagina-kpis .kx-det summary{ cursor:pointer; color:${AZUL}; font-weight:600; }
      #pagina-kpis .kx-tab{ width:100%; border-collapse:collapse; font-size:11.5px; margin-top:8px; }
      #pagina-kpis .kx-tab th{ text-align:left; font-size:10px; text-transform:uppercase; letter-spacing:.04em; color:var(--muted); padding:5px 6px; border-bottom:1px solid var(--line); white-space:nowrap; }
      #pagina-kpis .kx-tab td{ padding:5px 6px; border-bottom:1px solid var(--line); }
      #pagina-kpis .kx-tab .r{ text-align:right; font-variant-numeric:tabular-nums; }
      #pagina-kpis .kx-rodape{ font-size:11px; color:var(--muted); padding:0 4px; }`;
    document.head.appendChild(st);

    document.querySelectorAll('#kxProj button').forEach(b => b.onclick = () => { filtro.pj = b.dataset.pj; render(); });
    document.querySelectorAll('#kxModo button').forEach(b => b.onclick = () => { filtro.modo = b.dataset.modo; render(); });
    document.getElementById('kxMes').onchange = e => { filtro.mes = e.target.value; render(); };
  }

  function abrir(){
    document.querySelectorAll('[id^="pagina-"]').forEach(p => { p.style.display = 'none'; });
    document.querySelectorAll('.sidebar .nav-item').forEach(n => n.classList.remove('active'));
    document.getElementById('pagina-kpis').style.display = '';
    document.getElementById('nav-kpis').classList.add('active');
    render();
  }
  if(typeof window.mostrarPagina === 'function'){
    const original = window.mostrarPagina;
    window.mostrarPagina = function(){
      const pg = document.getElementById('pagina-kpis'); if(pg) pg.style.display = 'none';
      const nav = document.getElementById('nav-kpis'); if(nav) nav.classList.remove('active');
      return original.apply(this, arguments);
    };
  }

  function novoChart(id, cfg){
    if(charts[id]){ charts[id].destroy(); delete charts[id]; }
    const el = document.getElementById(id);
    if(!el || typeof Chart === 'undefined') return;
    charts[id] = new Chart(el, cfg);
  }
  const tile = (label, valor, detalhe, corV, meta) =>
    `<div class="kx-tile" style="--kx-cor:${corV || AZUL}"><div class="l">${label}</div><div class="v">${valor}</div><div class="d">${detalhe || ''}</div>${meta ? `<span class="meta">${meta}</span>` : ''}</div>`;
  const seta = (dif, bomQuandoCai, sufixo) => {
    if(dif == null || !isFinite(dif)) return '';
    if(Math.abs(dif) < 0.005) return '<b>= mês anterior</b>';
    const ruim = bomQuandoCai ? dif > 0 : dif < 0;
    return `<b style="color:${ruim ? VERMELHO : VERDE}">${dif > 0 ? '▲' : '▼'} ${Math.abs(dif).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}${sufixo}</b> vs mês anterior`;
  };

  // ---------- cálculos ----------
  function mesesDisponiveis(){
    const hoje = hojeIso(), lista = [];
    const d = new Date(Number(hoje.slice(0, 4)), Number(hoje.slice(5, 7)) - 1, 1);
    for(let i = 0; i < 12; i++){ lista.unshift(d.getFullYear() + '-' + pad2(d.getMonth() + 1)); d.setMonth(d.getMonth() - 1); }
    return lista;
  }
  function turnoverMes(ym){
    const ini = ym + '-01', fimM = fimDoMes(ym), hoje = hojeIso();
    const ref = fimM > hoje ? hoje : fimM;
    const base = COLAB.filter(x => passa(x.nome, x.contrato));
    const adm = base.filter(x => x.dataAdmis >= ini && x.dataAdmis <= fimM).length;
    const des = base.filter(x => x.dataDemissao && x.dataDemissao >= ini && x.dataDemissao <= fimM).length;
    const hc = base.filter(x => x.dataAdmis && x.dataAdmis <= ref && (!x.dataDemissao || x.dataDemissao > ref)).length;
    const tClass = hc ? ((adm + des) / 2) / hc * 100 : 0, tDes = hc ? des / hc * 100 : 0;
    return { ym, adm, des, hc, t: filtro.modo === 'classico' ? tClass : tDes };
  }
  function calcAbs(){
    const base = ABS.filter(a => passa(a.n, a.cc));
    const jornada = a => a.dPrev ? a.hPrev / a.dPrev : 0;
    let hF = 0, hT = 0, dF = 0, comFalta = 0;
    const porSup = {};
    base.forEach(a => {
      const f = a.hAusReg || 0, t = jornada(a) * (a.dEscP || 0), d = a.dAusReg || 0;
      hF += f; hT += t; dF += d; if(d > 0) comFalta++;
      const s = a.sup || '(sem supervisor)';
      const o = porSup[s] || (porSup[s] = { hF:0, hT:0, dF:0, n:0, comF:0 });
      o.hF += f; o.hT += t; o.dF += d; o.n++; if(d > 0) o.comF++;
    });
    const sups = Object.entries(porSup).map(([s, o]) => ({ s, ...o, t: o.hT ? o.hF / o.hT * 100 : 0 })).sort((a, b) => b.t - a.t);
    return { n: base.length, hF, hT, dF, comFalta, taxa: hT ? hF / hT * 100 : 0, sups };
  }
  function calcBH(){
    const atual = BH_ATUAL.filter(i => passa(i.n, i.projeto) && String(i.st || i.status || '').toUpperCase().startsWith('ATIVO'));
    let pos = 0, neg = 0, nPos = 0, nNeg = 0, nCrit = 0;
    atual.forEach(i => { const h = Number(i.h) || 0; if(h > 0){ pos += h; nPos++; } else if(h < 0){ neg += h; nNeg++; } if(h > LIMITE_BH) nCrit++; });
    const porComp = {};
    BH_HIST.forEach(b => {
      if(!passa(b.nome, b.descricaoCC)) return;
      const fim = brParaIso(String(b.periodo || '').split(' a ')[1] || ''); if(!fim) return;
      const o = porComp[fim] || (porComp[fim] = { pos:0, neg:0, n:0 });
      const h = Number(b.saldoAtualNum) || 0; if(h > 0) o.pos += h; else o.neg += h; o.n++;
    });
    const comps = Object.keys(porComp).sort().filter(k => porComp[k].n >= 50).slice(-12);
    return { n: atual.length, pos, neg, total: pos + neg, nPos, nNeg, nCrit, porComp, comps };
  }

  // ---------- desenho ----------
  function render(){
    const pg = document.getElementById('pagina-kpis');
    if(!pg || pg.style.display === 'none') return;
    document.querySelectorAll('#kxProj button').forEach(b => b.classList.toggle('ativo', b.dataset.pj === filtro.pj));
    document.querySelectorAll('#kxModo button').forEach(b => b.classList.toggle('ativo', b.dataset.modo === filtro.modo));

    const meses = mesesDisponiveis();
    if(!filtro.mes || meses.indexOf(filtro.mes) < 0) filtro.mes = meses[meses.length - 1];
    document.getElementById('kxMes').innerHTML = meses.slice().reverse().map(m => `<option value="${m}" ${m === filtro.mes ? 'selected' : ''}>${rotMes(m)}</option>`).join('');

    const serie = meses.map(turnoverMes);
    const i = meses.indexOf(filtro.mes), at = serie[i], ant = serie[i - 1];
    const media12 = serie.reduce((a, s) => a + s.t, 0) / serie.length;
    const A = calcAbs(), B = calcBH();
    const projTxt = filtro.pj ? (filtro.pj === 'CLARO' ? 'Claro' : 'Cirion') : 'Claro + Cirion';
    const perAbs = FOLHA_INI ? `${isoParaBr(FOLHA_INI).slice(0, 5)} a ${FOLHA ? FOLHA.slice(0, 5) : 'hoje'}` : 'folha atual';

    document.getElementById('kxRef').textContent = `${projTxt} · turnover de ${rotMes(filtro.mes)} · absenteísmo da folha ${perAbs} · banco de horas atual`;

    // ---- faixa executiva
    const corTo = semaforo(at.t, META_TURNOVER), corAbs = semaforo(A.taxa, META_ABS);
    document.getElementById('kxTiles').innerHTML =
      tile('Headcount', num(at.hc), seta(ant ? at.hc - ant.hc : null, false, '') || 'ativos', AZUL) +
      tile('Admitidos', num(at.adm), rotMes(filtro.mes) + (ant ? ` · anterior ${ant.adm}` : ''), VERDE) +
      tile('Desligados', num(at.des), rotMes(filtro.mes) + (ant ? ` · anterior ${ant.des}` : ''), VERMELHO) +
      tile('Turnover', pct(at.t), seta(ant ? at.t - ant.t : null, true, ' p.p.') + `<br>média 12m ${pct(media12)}`, corTo, `META ≤ ${META_TURNOVER}%`) +
      tile('Absenteísmo', pct(A.taxa), `${horas(A.hF)} de faltas<br>${A.comFalta} de ${A.n} colaboradores`, corAbs, `META ≤ ${META_ABS}%`) +
      tile('Banco de horas', horas(B.total), `<b style="color:${VERDE}">+${horas(B.pos)}</b> · <b style="color:${VERMELHO}">${horas(B.neg)}</b><br>${B.nCrit} acima de ${LIMITE_BH}h`, B.nCrit ? LARANJA : AZUL);

    // ---- destaques
    const top2 = A.sups.filter(s => s.hF > 0).slice(0, 2);
    const partTop = A.hF ? top2.reduce((a, s) => a + s.hF, 0) / A.hF * 100 : 0;
    const ultBH = B.comps[B.comps.length - 1], penBH = B.comps[B.comps.length - 2];
    const varBH = ultBH && penBH ? (B.porComp[ultBH].pos + B.porComp[ultBH].neg) - (B.porComp[penBH].pos + B.porComp[penBH].neg) : null;
    const dest = [
      [corTo, `<b>Turnover ${pct(at.t)}</b> em ${rotMes(filtro.mes)} (${at.adm} admissões, ${at.des} desligamentos) — ${at.t <= media12 ? 'abaixo' : 'acima'} da média de 12 meses (${pct(media12)}).`],
      [corAbs, `<b>Absenteísmo ${pct(A.taxa)}</b> na folha.` + (top2.length ? ` ${top2.map(s => esc(s.s.split(' ')[0]) + ' (' + pct(s.t) + ')').join(' e ')} concentram ${Math.round(partTop)}% das horas de falta.` : '')],
      [B.nCrit ? LARANJA : AZUL, `<b>Banco de horas ${horas(B.total)}</b>` + (varBH != null ? ` (${varBH >= 0 ? '+' : ''}${horas(varBH)} na última competência)` : '') + `. ${B.nNeg} colaboradores negativos e ${B.nCrit} acima de ${LIMITE_BH}h.`]
    ];
    document.getElementById('kxDestaques').innerHTML = dest.map(([c, t]) => `<div class="kx-dest" style="--kx-cor:${c}"><i></i><div>${t}</div></div>`).join('');

    // ---- turnover
    document.getElementById('kxToFormula').textContent = filtro.modo === 'classico'
      ? '((Admissões + Desligamentos) ÷ 2) ÷ Headcount × 100 · tracejado = meta'
      : 'Desligamentos ÷ Headcount × 100 · tracejado = meta';
    novoChart('kxToChart', {
      data: {
        labels: serie.map(s => rotMes(s.ym)),
        datasets: [
          { type:'line', label:'Turnover %', data: serie.map(s => +s.t.toFixed(2)), borderColor: AZUL, backgroundColor: AZUL, tension:.35, borderWidth:2.5,
            pointRadius: serie.map(s => s.ym === filtro.mes ? 6 : 3), pointBackgroundColor: serie.map(s => semaforo(s.t, META_TURNOVER)) },
          { type:'line', label:'Meta', data: serie.map(() => META_TURNOVER), borderColor:'#9aa5b8', borderDash:[5,5], borderWidth:1.5, pointRadius:0 }
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false, interaction:{ mode:'index', intersect:false },
        plugins:{ datalabels:{ display:false }, legend:{ display:false },
          tooltip:{ callbacks:{ label: c => c.datasetIndex ? ' Meta: ' + pct(c.raw) : ` Turnover: ${pct(c.raw)} · ${serie[c.dataIndex].adm} adm · ${serie[c.dataIndex].des} desl · HC ${serie[c.dataIndex].hc}` } } },
        scales:{ y:{ beginAtZero:true, ticks:{ callback: v => v + '%' }, grid:{ color:'rgba(0,0,0,.05)' } }, x:{ grid:{ display:false } } }
      }
    });

    // ---- absenteísmo: top 5
    const maxT = Math.max(1, ...A.sups.map(s => s.t));
    document.getElementById('kxAbsSub').textContent = `(Horas de faltas ÷ Horas trabalháveis) × 100 · ${horas(A.hF)} ÷ ${horas(A.hT)}`;
    document.getElementById('kxAbsRank').innerHTML = A.sups.slice(0, 5).map(s =>
      `<div class="kx-rk"><div class="n">${esc(s.s)}<small>${s.comF} de ${s.n} com falta · ${horas(s.hF)}</small></div>
        <div class="b"><span style="width:${(s.t / maxT * 100).toFixed(1)}%; background:${semaforo(s.t, META_ABS)}"></span></div>
        <div class="p" style="color:${semaforo(s.t, META_ABS)}">${s.t.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%</div></div>`).join('');
    document.getElementById('kxAbsTab').innerHTML =
      `<tr><th>Supervisor</th><th class="r">Colab.</th><th class="r">Com falta</th><th class="r">Dias</th><th class="r">H. faltas</th><th class="r">H. trabalháveis</th><th class="r">%</th></tr>` +
      A.sups.map(s => `<tr><td>${esc(s.s)}</td><td class="r">${s.n}</td><td class="r">${s.comF}</td><td class="r">${s.dF}</td><td class="r">${horas(s.hF)}</td><td class="r">${horas(s.hT)}</td><td class="r" style="font-weight:700; color:${semaforo(s.t, META_ABS)}">${pct(s.t)}</td></tr>`).join('');

    // ---- banco de horas
    const rot = k => MES_PT[Number(k.slice(5, 7)) - 1] + '/' + k.slice(2, 4);
    novoChart('kxBhChart', {
      data: {
        labels: B.comps.map(rot),
        datasets: [
          { type:'bar', label:'Positivo', data: B.comps.map(k => +B.porComp[k].pos.toFixed(1)), backgroundColor:'rgba(5,150,105,.55)', borderRadius:4, stack:'s' },
          { type:'bar', label:'Negativo', data: B.comps.map(k => +B.porComp[k].neg.toFixed(1)), backgroundColor:'rgba(220,38,38,.55)', borderRadius:4, stack:'s' },
          { type:'line', label:'Saldo', data: B.comps.map(k => +(B.porComp[k].pos + B.porComp[k].neg).toFixed(1)), borderColor: AZUL, backgroundColor: AZUL, tension:.35, pointRadius:3, borderWidth:2.5 }
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false, interaction:{ mode:'index', intersect:false },
        plugins:{ datalabels:{ display:false }, legend:{ position:'bottom', labels:{ boxWidth:10, font:{ size:11 } } }, tooltip:{ callbacks:{ label: c => ' ' + c.dataset.label + ': ' + horas(c.raw) } } },
        scales:{ y:{ ticks:{ callback: v => v.toLocaleString('pt-BR') + 'h' }, grid:{ color:'rgba(0,0,0,.05)' } }, x:{ stacked:true, grid:{ display:false } } }
      }
    });

    document.getElementById('kxRodape').textContent =
      `Turnover: Organograma (headcount = ativos no último dia do mês) · Absenteísmo: espelho de ponto · Banco de horas: base do RH · metas: turnover ≤ ${META_TURNOVER}% ao mês, absenteísmo ≤ ${META_ABS}%`;
  }
  window.renderKpisPessoas = render;

  // ---------- Dados ----------
  const pegar = arq => fetch(arq + '?_=' + Date.now()).then(r => r.ok ? r.json() : null).catch(() => null);
  async function carregar(){
    const [d, p, b] = await Promise.all([pegar('dados.json'), pegar('dados_pend_bh.json'), pegar('dados_bh.json')]);
    if(d){
      COLAB = (d.colaboradores || []).map(x => ({ nome: x.nome, contrato: String(x.contrato || ''),
        dataAdmis: String(x.dataAdmis || '').slice(0, 10), dataDemissao: String(x.dataDemissao || '').slice(0, 10) }))
        .filter(x => x.dataAdmis);
      BH_HIST = d.bancoHoras || [];
    }
    if(p){
      ABS = p.absenteismo || [];
      FOLHA_INI = p.folhaAtualIni || '';
      FOLHA = p.espelho && p.espelho.fim ? isoParaBr(p.espelho.fim) : '';
    }
    if(b) BH_ATUAL = b.itens || [];
    try{ render(); }catch(e){ console.error('[KPIs]', e); }
  }

  montarEstrutura();
  carregar();
  setInterval(carregar, 5 * 60 * 1000);
})();
//                     Horas de faltas = horas da jornada dos dias "Ausente"
//                     Horas trabalháveis = horas da jornada dos dias de escala já passados
//   • Banco de horas → dados_bh.json (saldo atual) + dados.json → bancoHoras (histórico)
//
// Instalação: no index.html, junto dos outros módulos:
//   <script src="gestao_kpis.js?v=1"></script>
// O item do menu, a página e o card da tela inicial aparecem sozinhos.
// =====================================================================
(function(){
  const C = (typeof cor !== 'undefined') ? cor : {
    good:'#059669', warn:'#d97706', bad:'#dc2626', accent:'#0e7c86', muted:'#6b7590', line:'#e1e5f0', text:'#1b2440'
  };
  const AZUL = '#0B2545', VERDE = '#059669', VERMELHO = '#dc2626', LARANJA = '#d97706';
  const MES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

  // colaboradores que estão em outro centro de custo, mas contam como Claro (igual ao Banco de Horas)
  const PROJ_FORCADO = { 'LIVIA MARIA VIEIRA YADA': 'CLARO', 'LEONARDO PERES FERRAZ': 'CLARO' };

  let COLAB = [], BH_HIST = [], BH_ATUAL = [], ABS = [], FOLHA = '', FOLHA_INI = '';
  const filtro = { pj: '', mes: '', modo: 'classico' };
  const charts = {};

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const pad2 = n => n < 10 ? '0' + n : String(n);
  const nomeChave = n => String(n || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
  function projForcado(n){
    const k = nomeChave(n); if (PROJ_FORCADO[k]) return PROJ_FORCADO[k];
    const f = k.length >= 15 && Object.keys(PROJ_FORCADO).find(x => x.startsWith(k) || k.startsWith(x));
    return f ? PROJ_FORCADO[f] : '';
  }
  const projDe = txt => { const t = nomeChave(txt); return t.includes('CIRION') ? 'CIRION' : (t.includes('CLARO') ? 'CLARO' : ''); };
  const passa = (nome, txtProj) => !filtro.pj || (projForcado(nome) || projDe(txtProj)) === filtro.pj;

  const pct = v => (isFinite(v) ? v : 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
  const num = v => Math.round(v).toLocaleString('pt-BR');
  function horas(v){
    const neg = v < 0, a = Math.abs(v), h = Math.floor(a), m = Math.round((a - h) * 60);
    const hh = m === 60 ? h + 1 : h, mm = m === 60 ? 0 : m;
    return (neg ? '-' : '') + hh.toLocaleString('pt-BR') + 'h' + (mm ? pad2(mm) : '');
  }
  const hojeIso = () => { const d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); };
  const fimDoMes = ym => { const [y, m] = ym.split('-').map(Number); return ym + '-' + pad2(new Date(y, m, 0).getDate()); };
  const rotMes = ym => { const [y, m] = ym.split('-'); return MES_PT[Number(m) - 1] + '/' + y.slice(2); };
  const brParaIso = br => { const m = String(br || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : ''; };
  const isoParaBr = iso => { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + '/' + m[2] + '/' + m[1] : ''; };

  const selStyle = 'font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); min-width:150px;';
  const lblStyle = 'font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em;';
  const formula = 'font-size:12px; color:var(--muted); margin:-4px 0 12px;';

  // ---------- Menu + página ----------
  function montarEstrutura(){
    const sidebar = document.querySelector('.sidebar');
    const main = document.querySelector('.main-content');
    if(!sidebar || !main || document.getElementById('nav-kpis')) return;

    const nav = document.createElement('div');
    nav.className = 'nav-item';
    nav.id = 'nav-kpis';
    nav.innerHTML = '<span class="nav-icon">📊</span> KPIs de Pessoas';
    nav.onclick = abrir;
    const ref = document.getElementById('nav-turnover') || document.getElementById('nav-organograma');
    if(ref && ref.parentNode === sidebar) sidebar.insertBefore(nav, ref.nextSibling);
    else sidebar.appendChild(nav);

    const pg = document.createElement('div');
    pg.className = 'pagina';
    pg.id = 'pagina-kpis';
    pg.style.display = 'none';
    pg.innerHTML = `
      <div class="panel">
        <h2>KPIs de Pessoas</h2>
        <div class="mono" style="font-size:11px; color:var(--muted); margin:2px 0 12px;">Turnover · Absenteísmo · Banco de horas</div>
        <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
          <label class="mono" style="${lblStyle}">Projeto</label>
          <select id="kpProj" class="mono" style="${selStyle}">
            <option value="">Todos</option><option value="CLARO">Claro</option><option value="CIRION">Cirion</option>
          </select>
          <label class="mono" style="${lblStyle} margin-left:8px;">Mês (turnover)</label>
          <select id="kpMes" class="mono" style="${selStyle}"></select>
        </div>
      </div>

      <div class="panel" style="margin-top:16px;">
        <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap;">
          <h2>🔄 Turnover</h2>
          <div id="kpModo" style="display:flex; gap:6px;">
            <button type="button" data-modo="classico" class="mono kp-btn">Clássico</button>
            <button type="button" data-modo="deslig" class="mono kp-btn">Só desligados</button>
          </div>
        </div>
        <div id="kpToFormula" style="${formula}"></div>
        <div class="kpi-row" id="kpToKpis" style="margin-bottom:14px;"></div>
        <div style="position:relative; height:300px;"><canvas id="kpToChart"></canvas></div>
        <div class="hint" style="margin-top:6px;">Linha = turnover do mês · barras = admissões e desligamentos · base: Organograma</div>
      </div>

      <div class="panel" style="margin-top:16px;">
        <h2>🩺 Absenteísmo</h2>
        <div id="kpAbsFormula" style="${formula}"></div>
        <div class="kpi-row" id="kpAbsKpis" style="margin-bottom:14px;"></div>
        <div style="overflow-x:auto;"><table class="tabela-kp" id="kpAbsTab"></table></div>
      </div>

      <div class="panel" style="margin-top:16px;">
        <h2>⏱️ Banco de horas</h2>
        <div id="kpBhFonte" style="${formula}"></div>
        <div class="kpi-row" id="kpBhKpis" style="margin-bottom:14px;"></div>
        <div style="position:relative; height:300px;"><canvas id="kpBhChart"></canvas></div>
        <div class="hint" style="margin-top:6px;">Saldo positivo e negativo somados por competência (folha de 15 a 14)</div>
      </div>`;
    main.appendChild(pg);

    const st = document.createElement('style');
    st.textContent = `
      #pagina-kpis .kp-btn{ font-size:11px; padding:5px 12px; border-radius:999px; border:1px solid var(--line); background:var(--panel); color:var(--text); cursor:pointer; }
      #pagina-kpis .kp-btn.ativo{ background:${AZUL}; border-color:${AZUL}; color:#fff; font-weight:700; }
      #pagina-kpis .tabela-kp{ width:100%; border-collapse:collapse; font-size:12px; }
      #pagina-kpis .tabela-kp th{ text-align:left; font-size:10.5px; text-transform:uppercase; letter-spacing:.04em; color:var(--muted); padding:7px 8px; border-bottom:1px solid var(--line); white-space:nowrap; }
      #pagina-kpis .tabela-kp td{ padding:7px 8px; border-bottom:1px solid var(--line); }
      #pagina-kpis .tabela-kp td.n, #pagina-kpis .tabela-kp th.n{ text-align:right; font-variant-numeric:tabular-nums; }
      #pagina-kpis .tabela-kp tr.tot td{ font-weight:700; background:rgba(11,37,69,.05); }`;
    document.head.appendChild(st);

    document.getElementById('kpProj').onchange = e => { filtro.pj = e.target.value; render(); };
    document.getElementById('kpMes').onchange = e => { filtro.mes = e.target.value; render(); };
    document.querySelectorAll('#kpModo .kp-btn').forEach(b => b.onclick = () => { filtro.modo = b.dataset.modo; render(); });
  }

  function abrir(){
    document.querySelectorAll('[id^="pagina-"]').forEach(p => { p.style.display = 'none'; });
    document.querySelectorAll('.sidebar .nav-item').forEach(n => n.classList.remove('active'));
    document.getElementById('pagina-kpis').style.display = '';
    document.getElementById('nav-kpis').classList.add('active');
    render();
  }
  if(typeof window.mostrarPagina === 'function'){
    const original = window.mostrarPagina;
    window.mostrarPagina = function(){
      const pg = document.getElementById('pagina-kpis'); if(pg) pg.style.display = 'none';
      const nav = document.getElementById('nav-kpis'); if(nav) nav.classList.remove('active');
      return original.apply(this, arguments);
    };
  }

  function novoChart(id, cfg){
    if(charts[id]){ charts[id].destroy(); delete charts[id]; }
    const el = document.getElementById(id);
    if(!el || typeof Chart === 'undefined') return;
    charts[id] = new Chart(el, cfg);
  }
  const kpi = (label, valor, delta, corV) =>
    `<div class="kpi"><div class="label">${label}</div><div class="value" style="color:${corV || AZUL}">${valor}</div>${delta ? `<div class="delta">${delta}</div>` : ''}</div>`;

  // ---------- Turnover ----------
  function mesesDisponiveis(){
    const hoje = hojeIso(), lista = [];
    const d = new Date(Number(hoje.slice(0, 4)), Number(hoje.slice(5, 7)) - 1, 1);
    for(let i = 0; i < 12; i++){
      lista.unshift(d.getFullYear() + '-' + pad2(d.getMonth() + 1));
      d.setMonth(d.getMonth() - 1);
    }
    return lista;
  }
  function turnoverMes(ym){
    const ini = ym + '-01', fimM = fimDoMes(ym), hoje = hojeIso();
    const ref = fimM > hoje ? hoje : fimM;
    const base = COLAB.filter(x => passa(x.nome, x.contrato));
    const adm = base.filter(x => x.dataAdmis >= ini && x.dataAdmis <= fimM).length;
    const des = base.filter(x => x.dataDemissao && x.dataDemissao >= ini && x.dataDemissao <= fimM).length;
    const hc = base.filter(x => x.dataAdmis && x.dataAdmis <= ref && (!x.dataDemissao || x.dataDemissao > ref)).length;
    const tClass = hc ? ((adm + des) / 2) / hc * 100 : 0;
    const tDes = hc ? des / hc * 100 : 0;
    return { ym, adm, des, hc, tClass, tDes, t: filtro.modo === 'classico' ? tClass : tDes };
  }
  function renderTurnover(){
    const meses = mesesDisponiveis();
    const sel = document.getElementById('kpMes');
    if(!filtro.mes || meses.indexOf(filtro.mes) < 0) filtro.mes = meses[meses.length - 1];
    sel.innerHTML = meses.slice().reverse().map(m => `<option value="${m}" ${m === filtro.mes ? 'selected' : ''}>${rotMes(m)}</option>`).join('');
    document.querySelectorAll('#kpModo .kp-btn').forEach(b => b.classList.toggle('ativo', b.dataset.modo === filtro.modo));

    const serie = meses.map(turnoverMes);
    const atual = serie.find(s => s.ym === filtro.mes);
    const ant = serie[serie.indexOf(atual) - 1];
    const somaA = serie.reduce((a, s) => a + s.adm, 0), somaD = serie.reduce((a, s) => a + s.des, 0);
    const hcMedio = serie.reduce((a, s) => a + s.hc, 0) / serie.length;
    const t12 = serie.reduce((a, s) => a + s.t, 0) / serie.length;   // média dos 12 turnovers mensais

    document.getElementById('kpToFormula').innerHTML = filtro.modo === 'classico'
      ? 'Turnover = ((Admissões + Desligamentos) ÷ 2) ÷ Headcount × 100'
      : 'Turnover = Desligamentos ÷ Headcount × 100';
    const conta = filtro.modo === 'classico'
      ? `((${atual.adm} + ${atual.des}) ÷ 2) ÷ ${atual.hc}`
      : `${atual.des} ÷ ${atual.hc}`;
    const dif = ant ? atual.t - ant.t : null;
    const ehMesAtual = fimDoMes(filtro.mes) > hojeIso();
    document.getElementById('kpToKpis').innerHTML =
      kpi('👥 Headcount', num(atual.hc), ehMesAtual ? 'ativos hoje' : 'ativos em ' + isoParaBr(fimDoMes(filtro.mes))) +
      kpi('➕ Admitidos', num(atual.adm), rotMes(filtro.mes), VERDE) +
      kpi('➖ Desligados', num(atual.des), rotMes(filtro.mes), VERMELHO) +
      kpi('🔄 Turnover ' + rotMes(filtro.mes), pct(atual.t),
        conta + (dif != null ? ` · <span style="color:${dif > 0 ? VERMELHO : VERDE}">${dif > 0 ? '▲' : '▼'} ${Math.abs(dif).toFixed(2).replace('.', ',')} p.p.</span>` : ''),
        atual.t > 10 ? VERMELHO : (atual.t > 5 ? LARANJA : VERDE)) +
      kpi('📅 Turnover médio 12 meses', pct(t12), `média mensal ${rotMes(meses[0])} → ${rotMes(meses[meses.length - 1])} · ${somaA} adm · ${somaD} desl · HC médio ${num(hcMedio)}`);

    novoChart('kpToChart', {
      data: {
        labels: serie.map(s => rotMes(s.ym)),
        datasets: [
          { type:'line', label:'Turnover %', data: serie.map(s => +s.t.toFixed(2)), yAxisID:'y1', borderColor: AZUL, backgroundColor: AZUL, tension:.3, pointRadius:4, borderWidth:2.5 },
          { type:'bar', label:'Admissões', data: serie.map(s => s.adm), yAxisID:'y', backgroundColor:'rgba(5,150,105,.55)', borderRadius:4 },
          { type:'bar', label:'Desligamentos', data: serie.map(s => s.des), yAxisID:'y', backgroundColor:'rgba(220,38,38,.55)', borderRadius:4 }
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false, interaction:{ mode:'index', intersect:false },
        plugins:{ datalabels:{ display:false }, legend:{ position:'bottom' }, tooltip:{ callbacks:{ label: c => c.dataset.yAxisID === 'y1' ? ' Turnover: ' + pct(c.raw) : ' ' + c.dataset.label + ': ' + c.raw } } },
        scales:{
          y:{ beginAtZero:true, title:{ display:true, text:'Pessoas' }, grid:{ color:'rgba(0,0,0,.05)' } },
          y1:{ beginAtZero:true, position:'right', title:{ display:true, text:'Turnover %' }, grid:{ display:false }, ticks:{ callback: v => v + '%' } },
          x:{ grid:{ display:false } }
        }
      }
    });
  }

  // ---------- Absenteísmo ----------
  function renderAbs(){
    const base = ABS.filter(a => passa(a.n, a.cc));
    const jornada = a => a.dPrev ? a.hPrev / a.dPrev : 0;
    const linha = a => ({ hF: a.hAusReg || 0, hT: jornada(a) * (a.dEscP || 0), dF: a.dAusReg || 0, dT: a.dEscP || 0 });
    let hF = 0, hT = 0, dF = 0, dT = 0, comFalta = 0;
    const porSup = {};
    base.forEach(a => {
      const l = linha(a);
      hF += l.hF; hT += l.hT; dF += l.dF; dT += l.dT; if(l.dF > 0) comFalta++;
      const s = a.sup || '(sem supervisor)';
      const o = porSup[s] || (porSup[s] = { hF:0, hT:0, dF:0, n:0, comF:0 });
      o.hF += l.hF; o.hT += l.hT; o.dF += l.dF; o.n++; if(l.dF > 0) o.comF++;
    });
    const taxa = hT ? hF / hT * 100 : 0;
    const per = FOLHA_INI ? `folha atual: ${isoParaBr(FOLHA_INI)} até ${FOLHA ? FOLHA : 'hoje'}` : 'folha atual';
    document.getElementById('kpAbsFormula').innerHTML =
      `Absenteísmo = (Horas de faltas ÷ Horas trabalháveis) × 100 · ${esc(per)}`;
    document.getElementById('kpAbsKpis').innerHTML =
      kpi('🩺 Absenteísmo', pct(taxa), `${horas(hF)} ÷ ${horas(hT)}`, taxa > 5 ? VERMELHO : (taxa > 3 ? LARANJA : VERDE)) +
      kpi('⏳ Horas de faltas', horas(hF), `${num(dF)} dias ausentes`, VERMELHO) +
      kpi('🗓️ Horas trabalháveis', horas(hT), `${num(dT)} dias de escala já passados`) +
      kpi('👤 Colaboradores com falta', num(comFalta), `de ${num(base.length)} no espelho`, LARANJA);

    const sups = Object.entries(porSup).sort((a, b) => (b[1].hT ? b[1].hF / b[1].hT : 0) - (a[1].hT ? a[1].hF / a[1].hT : 0));
    document.getElementById('kpAbsTab').innerHTML =
      `<tr><th>Supervisor</th><th class="n">Colaboradores</th><th class="n">Com falta</th><th class="n">Dias ausentes</th><th class="n">Horas de faltas</th><th class="n">Horas trabalháveis</th><th class="n">Absenteísmo</th></tr>` +
      sups.map(([s, o]) => {
        const t = o.hT ? o.hF / o.hT * 100 : 0;
        return `<tr><td>${esc(s)}</td><td class="n">${o.n}</td><td class="n">${o.comF}</td><td class="n">${o.dF}</td><td class="n">${horas(o.hF)}</td><td class="n">${horas(o.hT)}</td>
          <td class="n" style="font-weight:700; color:${t > 5 ? VERMELHO : (t > 3 ? LARANJA : VERDE)}">${pct(t)}</td></tr>`;
      }).join('') +
      `<tr class="tot"><td>Total</td><td class="n">${base.length}</td><td class="n">${comFalta}</td><td class="n">${dF}</td><td class="n">${horas(hF)}</td><td class="n">${horas(hT)}</td><td class="n">${pct(taxa)}</td></tr>`;
  }

  // ---------- Banco de horas ----------
  function renderBH(){
    const atual = BH_ATUAL.filter(i => passa(i.n, i.projeto) && String(i.st || i.status || '').toUpperCase().startsWith('ATIVO'));
    let pos = 0, neg = 0, nPos = 0, nNeg = 0, nCrit = 0;
    atual.forEach(i => {
      const h = Number(i.h) || 0;
      if(h > 0){ pos += h; nPos++; } else if(h < 0){ neg += h; nNeg++; }
      if(h > 180) nCrit++;
    });
    document.getElementById('kpBhFonte').innerHTML = 'Saldo atual dos colaboradores ativos · histórico pelo fechamento de cada competência';
    document.getElementById('kpBhKpis').innerHTML =
      kpi('⏱️ Saldo total', horas(pos + neg), `${num(atual.length)} colaboradores ativos`, pos + neg >= 0 ? AZUL : VERMELHO) +
      kpi('➕ Saldo positivo', horas(pos), `${nPos} colaboradores`, VERDE) +
      kpi('➖ Saldo negativo', horas(neg), `${nNeg} colaboradores`, VERMELHO) +
      kpi('⚠️ Acima de 180h', num(nCrit), 'muito crítico', nCrit ? VERMELHO : VERDE) +
      kpi('📐 Média por colaborador', horas(atual.length ? (pos + neg) / atual.length : 0), 'saldo total ÷ ativos');

    // histórico por competência
    const porComp = {};
    BH_HIST.forEach(b => {
      if(!passa(b.nome, b.descricaoCC)) return;
      const per = String(b.periodo || '');
      const fim = brParaIso(per.split(' a ')[1] || '');
      if(!fim) return;
      const o = porComp[fim] || (porComp[fim] = { pos:0, neg:0, n:0 });
      const h = Number(b.saldoAtualNum) || 0;
      if(h > 0) o.pos += h; else o.neg += h;
      o.n++;
    });
    const comps = Object.keys(porComp).sort().filter(k => porComp[k].n >= 50).slice(-12);
    const rot = k => MES_PT[Number(k.slice(5, 7)) - 1] + '/' + k.slice(2, 4);
    novoChart('kpBhChart', {
      data: {
        labels: comps.map(rot),
        datasets: [
          { type:'bar', label:'Saldo positivo', data: comps.map(k => +porComp[k].pos.toFixed(1)), backgroundColor:'rgba(5,150,105,.6)', borderRadius:4, stack:'s' },
          { type:'bar', label:'Saldo negativo', data: comps.map(k => +porComp[k].neg.toFixed(1)), backgroundColor:'rgba(220,38,38,.6)', borderRadius:4, stack:'s' },
          { type:'line', label:'Saldo total', data: comps.map(k => +(porComp[k].pos + porComp[k].neg).toFixed(1)), borderColor: AZUL, backgroundColor: AZUL, tension:.3, pointRadius:4, borderWidth:2.5 }
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false, interaction:{ mode:'index', intersect:false },
        plugins:{ datalabels:{ display:false }, legend:{ position:'bottom' }, tooltip:{ callbacks:{ label: c => ' ' + c.dataset.label + ': ' + horas(c.raw) } } },
        scales:{ y:{ title:{ display:true, text:'Horas' }, ticks:{ callback: v => v.toLocaleString('pt-BR') + 'h' } }, x:{ grid:{ display:false } } }
      }
    });
  }

  function render(){
    const pg = document.getElementById('pagina-kpis');
    if(!pg || pg.style.display === 'none') return;
    try{ renderTurnover(); }catch(e){ console.error('[KPIs] turnover', e); }
    try{ renderAbs(); }catch(e){ console.error('[KPIs] absenteísmo', e); }
    try{ renderBH(); }catch(e){ console.error('[KPIs] banco de horas', e); }
  }
  window.renderKpisPessoas = render;

  // ---------- Dados ----------
  const pegar = arq => fetch(arq + '?_=' + Date.now()).then(r => r.ok ? r.json() : null).catch(() => null);
  async function carregar(){
    const [d, p, b] = await Promise.all([pegar('dados.json'), pegar('dados_pend_bh.json'), pegar('dados_bh.json')]);
    if(d){
      COLAB = (d.colaboradores || []).map(x => ({ nome: x.nome, contrato: String(x.contrato || ''),
        dataAdmis: String(x.dataAdmis || '').slice(0, 10), dataDemissao: String(x.dataDemissao || '').slice(0, 10) }))
        .filter(x => x.dataAdmis);
      BH_HIST = d.bancoHoras || [];
    }
    if(p){
      ABS = p.absenteismo || [];
      FOLHA_INI = p.folhaAtualIni || '';
      FOLHA = p.espelho && p.espelho.fim ? isoParaBr(p.espelho.fim) : '';
    }
    if(b) BH_ATUAL = b.itens || [];
    render();
  }

  montarEstrutura();
  carregar();
  setInterval(carregar, 5 * 60 * 1000);
})();
