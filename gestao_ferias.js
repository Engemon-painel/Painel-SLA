/* =========================================================================
   Painel-SLA · Programação de Férias  (gestao_ferias.js)
   Aba nova dentro de "Organograma". Usa os colaboradores ATIVOS do dados.json
   (data de admissão, supervisor, coordenador) — os mesmos da base Organograma.
   Regra:
     · a cada aniversário de admissão o colaborador completa um período aquisitivo
       (ex.: admitido em 10/01/2026 → completa 1 ano em 10/01/2027);
     · 60 dias antes desse dia ele entra em "Programar agora";
     · limite para tirar as férias = 12 meses depois de completar o período (CLT, art. 134).
   O supervisor programa as datas no botão "📅 Programar" (grava pelo mesmo fluxo da
   Avaliação de Experiência, aba FERIAS_PROGRAMADAS da Base com Historico → dados.json
   "feriasProgramadas"). Até o próximo dados.json, a cópia fica salva neste navegador.
   ========================================================================= */
(function () {
  'use strict';

  const ANTECEDENCIA = 60;          // dias antes de completar o período para programar
  const MESES_GRADE = 12;           // meses mostrados no gráfico e na grade por supervisor
  const MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  // mesmo gatilho HTTP do fluxo "Avaliação Experiência – Gravar" (gestao_avaliacao.js)
  const FLUXO = 'https://defaultc54e138a79f54294aa9bb004cf8d7f.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/8ac7331ffa21496caf9fbb84763a05ae/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jnwHYpz7ZOuYh2QtvqebFXDA-H7ln30RcRt_Hvu6Pd8';
  const CHAVE_LOCAL = 'feriasProgLocais_v1';
  const filtro = { sup: '', mes: '', sit: '', q: '' };
  let montado = false, ultimaQtd = -1, listaExport = [], PROG = {}, progCarregado = false;

  const $ = (s, el) => (el || document).querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = n => Number(n || 0).toLocaleString('pt-BR');
  const titulo = s => String(s || '').toLowerCase().replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase())
    .replace(/\s(De|Da|Do|Dos|Das|E)\s/g, m => m.toLowerCase());

  // ---------- datas (sempre meia-noite local) ----------
  const hoje = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
  const parseISO = s => {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || ''));
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const b = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(String(s || ''));
    return b ? new Date(+b[3], +b[2] - 1, +b[1]) : null;
  };
  const addAnos = (d, n) => {           // 29/02 vira 28/02 em ano não bissexto
    const r = new Date(d.getFullYear() + n, d.getMonth(), d.getDate());
    if (r.getMonth() !== d.getMonth()) r.setDate(0);
    return r;
  };
  const addDias = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  const br = d => d ? String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear() : '—';
  const chaveMes = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  const rotMes = k => MES[+k.slice(5, 7) - 1] + '/' + k.slice(2, 4);
  const dif = (a, b) => Math.round((a - b) / 86400000);
  const iso = d => d ? d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') : '';
  const matNum = v => String(v == null ? '' : v).replace(/\.0+$/, '').replace(/^0+/, '').trim();
  const chaveProg = (mat, aniv) => matNum(mat) + '|' + iso(aniv);

  // ---------- férias já programadas (dados.json + cópia deste navegador) ----------
  const locais = () => { try { return JSON.parse(localStorage.getItem(CHAVE_LOCAL) || '{}'); } catch (e) { return {}; } };
  const salvarLocal = (k, r) => { try { const l = locais(); l[k] = r; localStorage.setItem(CHAVE_LOCAL, JSON.stringify(l)); } catch (e) { } };
  function carregarProg() {
    return fetch('dados.json?_=' + Date.now()).then(r => r.ok ? r.json() : null).catch(() => null).then(d => {
      PROG = {};
      (d && Array.isArray(d.feriasProgramadas) ? d.feriasProgramadas : []).forEach(x => {
        try { const o = typeof x === 'string' ? JSON.parse(x) : x; const r = typeof o.r === 'string' ? JSON.parse(o.r) : (o.r || o); r.__em = o.em || ''; PROG[matNum(r.matricula) + '|' + r.completaEm] = r; } catch (e) { }
      });
      const l = locais();
      Object.keys(l).forEach(k => { const b = PROG[k]; if (!b || String(b.__em || '') < String(l[k].__em || '')) PROG[k] = l[k]; });
      progCarregado = true;
      desenhar();
    });
  }
  const periodosDe = r => (r && Array.isArray(r.periodos) ? r.periodos : []).filter(p => p && p.inicio && Number(p.dias) > 0);

  // ---------- colaboradores → próximo período ----------
  function colaboradores() {
    let lista = [];
    try { lista = typeof COLABORADORES !== 'undefined' && Array.isArray(COLABORADORES) ? COLABORADORES : []; } catch (e) { lista = []; }
    const h = hoje();
    const vistos = new Set();
    const out = [];
    lista.forEach(c => {
      if (!c || c.ativo === false) return;
      const adm = parseISO(c.dataAdmis);
      if (!adm || adm > h) return;
      const chave = String(c.matricula || c.nome);
      if (vistos.has(chave)) return;
      vistos.add(chave);
      // período deste ano; se já passou antes do começo do mês atual, vale o do ano que vem
      // (quem completou no começo deste mês continua aparecendo, como "Já completou")
      const inicioMes = new Date(h.getFullYear(), h.getMonth(), 1);
      let anos = h.getFullYear() - adm.getFullYear();
      let aniv = addAnos(adm, anos);
      if (aniv < inicioMes) { anos++; aniv = addAnos(adm, anos); }
      if (anos < 1) { anos = 1; aniv = addAnos(adm, 1); }
      const programar = addDias(aniv, -ANTECEDENCIA);
      const limite = addDias(addAnos(aniv, 1), -1);
      const faltam = dif(aniv, h);
      const prog = PROG[chaveProg(c.matricula, aniv)];
      const periodos = periodosDe(prog);
      const sit = periodos.length ? 'Programada' : faltam < 0 ? 'Já completou' : faltam <= ANTECEDENCIA ? 'Programar agora' : (faltam <= ANTECEDENCIA + 30 ? 'Programar no próximo mês' : 'Aguardando');
      out.push({
        nome: c.nome, matricula: c.matricula, supervisor: c.supervisor || 'Sem supervisor', coordenador: c.coordenador || '',
        funcao: c.descFuncao || '', adm, anos, aniv, programar, limite, faltam, sit, mes: chaveMes(aniv), prog, periodos
      });
    });
    return out.sort((a, b) => a.aniv - b.aniv || String(a.nome).localeCompare(String(b.nome), 'pt-BR'));
  }

  const COR_SIT = { 'Já completou': '#7f1d1d', 'Programar agora': '#dc2626', 'Programar no próximo mês': '#d97706', 'Aguardando': '#64748b', 'Programada': '#15803d' };
  // cor de alerta de um mês (posição a partir do mês atual) quando ainda há gente sem programar
  const ALERTA = [
    { bg: '#b91c1c', fg: '#fff', r: 'este mês' },
    { bg: '#ea580c', fg: '#fff', r: 'próximo mês' },
    { bg: '#facc15', fg: '#422006', r: 'daqui a 2 meses' }
  ];
  const corMes = i => ALERTA[i] || { bg: '#e2e8f0', fg: '#0f172a', r: '3 meses ou mais' };
  const OK = { bg: '#15803d', fg: '#fff' };
  const resumoProg = periodos => periodos.map(p => br(parseISO(p.inicio)) + ' · ' + p.dias + ' d').join(' + ');

  // ---------- layout ----------
  function css() {
    if ($('#fer-css')) return;
    const st = document.createElement('style');
    st.id = 'fer-css';
    st.textContent = `
      #feriasSecao .fer-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:12px 0 16px}
      #feriasSecao .fer-kpi{border:1px solid #e2e8f0;border-left:5px solid #0e7c86;border-radius:10px;padding:12px 14px;background:#fff;cursor:pointer}
      #feriasSecao .fer-kpi.on{box-shadow:0 0 0 2px #0e7c86 inset}
      #feriasSecao .fer-kpi small{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#475569;font-weight:700}
      #feriasSecao .fer-kpi b{display:block;font-size:28px;color:#0f172a;margin:4px 0 2px}
      #feriasSecao .fer-kpi span{font-size:12px;color:#334155}
      #feriasSecao .fer-filtros{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:4px 0 12px}
      #feriasSecao .fer-filtros select,#feriasSecao .fer-filtros input{font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#0f172a}
      #feriasSecao .fer-filtros input{min-width:220px}
      #feriasSecao .fer-btn{font-size:12px;padding:6px 12px;border-radius:8px;border:1px solid #cbd5e1;background:#f8fafc;color:#0f172a;cursor:pointer;font-weight:600}
      #feriasSecao .fer-btn.exp{background:#0e7c86;color:#fff;border-color:#0e7c86;margin-left:auto}
      #feriasSecao .fer-meses{display:grid;grid-template-columns:repeat(${MESES_GRADE},minmax(58px,1fr));gap:6px;align-items:end;height:170px;margin:8px 0 4px}
      #feriasSecao .fer-mes{display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;cursor:pointer}
      #feriasSecao .fer-mes .bar{width:70%;background:#0e7c86;border-radius:6px 6px 0 0;min-height:2px}
      #feriasSecao .fer-mes.on .bar{background:#dc2626}
      #feriasSecao .fer-mes .v{font-size:13px;font-weight:800;color:#0f172a;margin-bottom:3px}
      #feriasSecao .fer-mes .r{font-size:11px;font-weight:700;color:#334155;margin-top:4px}
      #feriasSecao table{width:100%;border-collapse:collapse}
      #feriasSecao th{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:#334155;font-weight:700;padding:7px 8px;border-bottom:1px solid #cbd5e1;text-align:left;white-space:nowrap;background:#f8fafc}
      #feriasSecao td{font-size:13px;color:#0f172a;padding:7px 8px;border-bottom:1px solid #e2e8f0;white-space:nowrap}
      #feriasSecao td.c,#feriasSecao th.c{text-align:center}
      #feriasSecao .tag{display:inline-block;padding:2px 9px;border-radius:999px;color:#fff;font-size:11.5px;font-weight:700}
      #feriasSecao .grade td.c{font-weight:700}
      #feriasSecao .grade td.zero{color:#cbd5e1;font-weight:400}
      #feriasSecao .nota{font-size:12px;color:#475569;margin:2px 0 10px}
      #feriasSecao .fer-mes .v{font-size:14px;font-weight:900;padding:1px 8px;border-radius:999px}
      #feriasSecao .grade td{font-size:14px}
      #feriasSecao .grade .al{display:inline-block;min-width:30px;padding:3px 8px;border-radius:999px;font-weight:900;font-size:14px;line-height:1.2}
      #feriasSecao .grade .al small{font-size:10.5px;font-weight:800;opacity:.9;margin-left:3px}
      #feriasSecao .leg{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:#334155;margin:0 0 8px;align-items:center}
      #feriasSecao .leg i{display:inline-block;width:14px;height:14px;border-radius:4px;vertical-align:-2px;margin-right:4px}
      #feriasSecao .fer-prog{font-size:12px;padding:4px 10px;border-radius:8px;border:1px solid #0e7c86;background:#fff;color:#0e7c86;cursor:pointer;font-weight:700}
      #feriasSecao .fer-prog.urg{background:#b91c1c;border-color:#b91c1c;color:#fff}
      #feriasSecao .fer-prog.feito{border-color:#15803d;color:#15803d;background:#f0fdf4;font-weight:700}
      .fer-modal{position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9999;display:flex;align-items:center;justify-content:center;padding:16px}
      .fer-modal .cx{background:#fff;border-radius:14px;max-width:620px;width:100%;max-height:92vh;overflow:auto;padding:20px 22px;box-shadow:0 20px 60px rgba(0,0,0,.3);color:#0f172a}
      .fer-modal h3{margin:0 0 4px;font-size:18px}
      .fer-modal .info{font-size:13px;color:#334155;background:#f1f5f9;border-radius:10px;padding:10px 12px;margin:10px 0 14px;line-height:1.6}
      .fer-modal .per{display:grid;grid-template-columns:90px 1fr 110px 1fr;gap:8px;align-items:center;margin-bottom:8px;font-size:13px}
      .fer-modal input,.fer-modal select,.fer-modal textarea{font:inherit;font-size:13px;padding:7px 9px;border:1px solid #cbd5e1;border-radius:8px;width:100%;box-sizing:border-box;color:#0f172a;background:#fff}
      .fer-modal textarea{min-height:56px}
      .fer-modal label{font-size:12px;font-weight:700;color:#334155;display:block;margin:10px 0 4px}
      .fer-modal .fim{font-size:12.5px;color:#0f172a;font-weight:600}
      .fer-modal .rod{display:flex;gap:8px;align-items:center;margin-top:16px;flex-wrap:wrap}
      .fer-modal .rod button{font-size:13px;padding:8px 14px;border-radius:8px;border:1px solid #cbd5e1;background:#f8fafc;cursor:pointer;font-weight:700;color:#0f172a}
      .fer-modal .rod .pri{background:#0e7c86;color:#fff;border-color:#0e7c86}
      .fer-modal .rod .apagar{color:#b91c1c;border-color:#fca5a5;background:#fff}
      @media (max-width:560px){.fer-modal .per{grid-template-columns:1fr 1fr}}`;
    document.head.appendChild(st);
  }

  // A aba fica no menu lateral do Organograma (Estrutura · Avaliação de Experiência ·
  // Entrevista de Desligamento), montado pelo gestao_avaliacao.js: entra como o último item.
  function montar() {
    const pag = $('#pagina-organograma');
    const menu = pag && pag.querySelector('.org-menu .nav-group-items');
    const conteudo = pag && pag.querySelector('.org-conteudo');
    if (!menu || !conteudo || montado) return !!montado;
    css();
    const item = document.createElement('div');
    item.className = 'nav-item org-aba';
    item.dataset.aba = 'ferias';
    item.innerHTML = '<span class="nav-icon">🏖️</span> Programação de Férias';
    menu.appendChild(item);
    const caixa = document.createElement('div');
    caixa.id = 'org-aba-ferias';
    caixa.style.display = 'none';
    conteudo.appendChild(caixa);
    const sec = document.createElement('div');
    sec.className = 'panel';
    sec.id = 'feriasSecao';
    sec.innerHTML = `
      <h2>Programação de Férias</h2>
      <p class="nota">Colaboradores ativos do Organograma · cada um completa um período de férias no aniversário da admissão ·
        <b>${ANTECEDENCIA} dias antes</b> entra em "Programar agora" · limite para tirar = 12 meses depois de completar ·
        o supervisor programa as datas no botão <b>📅 Programar</b> da lista.</p>
      <div class="fer-kpis" id="fer-kpis"></div>
      <div class="panel" style="margin-bottom:16px">
        <h2 style="margin-bottom:2px">Quantos completam o período em cada mês</h2>
        <p class="nota">Próximos ${MESES_GRADE} meses · número colorido = ainda tem gente sem programar · clique num mês para filtrar a lista</p>
        <div class="leg" id="fer-leg"></div>
        <div class="fer-meses" id="fer-meses"></div>
      </div>
      <div class="panel" style="margin-bottom:16px">
        <h2 style="margin-bottom:2px">Por supervisor e mês</h2>
        <p class="nota">Quantos técnicos de cada supervisor completam o período em cada mês · o número pequeno é quantos <b>faltam programar</b> · clique numa célula para ver a lista</p>
        <div style="overflow:auto"><table class="grade" id="fer-grade"></table></div>
      </div>
      <div class="panel">
        <h2 style="margin-bottom:6px">Colaboradores</h2>
        <div class="fer-filtros">
          <input type="search" id="fer-q" placeholder="Buscar nome ou matrícula...">
          <select id="fer-sup"><option value="">Supervisor: todos</option></select>
          <select id="fer-mesf"><option value="">Mês: todos</option></select>
          <select id="fer-sit"><option value="">Situação: todas</option><option>Já completou</option><option>Programar agora</option><option>Programar no próximo mês</option><option>Aguardando</option><option>Programada</option></select>
          <button type="button" class="fer-btn" id="fer-limpar">Limpar filtros</button>
          <span id="fer-cont" style="font-size:12px;color:#334155"></span>
          <button type="button" class="fer-btn exp" id="fer-exp">⬇ Exportar (Excel)</button>
        </div>
        <div style="overflow:auto;max-height:560px"><table>
          <thead><tr><th>Colaborador</th><th>Matrícula</th><th>Supervisor</th><th class="c">Admissão</th><th class="c">Completa</th>
            <th class="c">Data que completa</th><th class="c">Programar a partir de</th><th class="c">Faltam</th><th class="c">Limite para tirar</th><th>Situação</th><th>Férias programadas</th></tr></thead>
          <tbody id="fer-tbody"></tbody></table></div>
      </div>`;
    caixa.appendChild(sec);

    // clique na nossa aba: esconde as outras; clique nas outras: esconde a nossa
    item.addEventListener('click', () => {
      pag.querySelectorAll('.org-aba').forEach(b => b.classList.toggle('active', b === item));
      ['org-aba-estrutura', 'org-aba-aval', 'org-aba-saida'].forEach(id => { const d = document.getElementById(id); if (d) d.style.display = 'none'; });
      caixa.style.display = '';
      desenhar();
    });
    pag.querySelectorAll('.org-aba').forEach(b => { if (b !== item) b.addEventListener('click', () => { caixa.style.display = 'none'; item.classList.remove('active'); }); });
    $('#fer-q').addEventListener('input', e => { filtro.q = e.target.value.trim(); lista(); });
    $('#fer-sup').addEventListener('change', e => { filtro.sup = e.target.value; desenhar(); });
    $('#fer-mesf').addEventListener('change', e => { filtro.mes = e.target.value; desenhar(); });
    $('#fer-sit').addEventListener('change', e => { filtro.sit = e.target.value; desenhar(); });
    $('#fer-limpar').addEventListener('click', () => {
      Object.keys(filtro).forEach(k => filtro[k] = '');
      ['#fer-q', '#fer-sup', '#fer-mesf', '#fer-sit'].forEach(s => { $(s).value = ''; });
      desenhar();
    });
    $('#fer-exp').addEventListener('click', exportar);
    montado = true;
    return true;
  }

  // ---------- desenho ----------
  function proximosMeses() {
    const h = hoje(), out = [];
    for (let i = 0; i < MESES_GRADE; i++) out.push(chaveMes(new Date(h.getFullYear(), h.getMonth() + i, 1)));
    return out;
  }

  function desenhar() {
    if (!montado || $('#org-aba-ferias').style.display === 'none') return;
    const todos = colaboradores();
    const meses = proximosMeses();
    // opções dos filtros
    const sups = [...new Set(todos.map(c => c.supervisor))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const selSup = $('#fer-sup');
    if (selSup.options.length !== sups.length + 1) {
      selSup.innerHTML = '<option value="">Supervisor: todos</option>' + sups.map(s => `<option value="${esc(s)}">${esc(titulo(s))}</option>`).join('');
      selSup.value = filtro.sup;
    }
    const selMes = $('#fer-mesf');
    if (selMes.options.length !== meses.length + 1) {
      selMes.innerHTML = '<option value="">Mês: todos</option>' + meses.map(m => `<option value="${m}">${rotMes(m)}</option>`).join('');
      selMes.value = filtro.mes;
    }
    const doSup = todos.filter(c => !filtro.sup || c.supervisor === filtro.sup);

    // KPIs
    const h = hoje(), mesAtual = chaveMes(h), mesProx = meses[1];
    const k = [
      { r: 'Já completou · sem programar', v: doSup.filter(c => c.sit === 'Já completou').length, s: 'completaram o período neste mês: programar já', cor: '#7f1d1d', f: { sit: 'Já completou', mes: '' }, al: true },
      { r: 'Programar agora', v: doSup.filter(c => c.sit === 'Programar agora').length, s: `completam nos próximos ${ANTECEDENCIA} dias e ainda sem data`, cor: '#dc2626', f: { sit: 'Programar agora', mes: '' }, al: true },
      { r: 'Já programadas', v: doSup.filter(c => c.sit === 'Programada').length, s: 'supervisor já preencheu as datas', cor: '#15803d', f: { sit: 'Programada', mes: '' } },
      { r: 'Completam em ' + rotMes(mesAtual), v: doSup.filter(c => c.mes === mesAtual).length, s: 'neste mês', cor: '#0e7c86', f: { mes: mesAtual, sit: '' } },
      { r: 'Completam em ' + rotMes(mesProx), v: doSup.filter(c => c.mes === mesProx).length, s: 'no próximo mês', cor: '#2563eb', f: { mes: mesProx, sit: '' } },
      { r: 'Fazem 1 ano', v: doSup.filter(c => c.anos === 1 && c.faltam <= ANTECEDENCIA + 30).length, s: 'primeiro período nos próximos 90 dias', cor: '#d97706', f: null },
      { r: 'Colaboradores ativos', v: doSup.length, s: filtro.sup ? titulo(filtro.sup) : 'todos os supervisores', cor: '#64748b', f: { mes: '', sit: '' } }
    ];
    $('#fer-kpis').innerHTML = k.map((x, i) => `<div class="fer-kpi${x.f && filtro.sit === x.f.sit && filtro.mes === x.f.mes && (x.f.sit || x.f.mes) ? ' on' : ''}" data-i="${i}" style="border-left-color:${x.cor}">
      <small>${esc(x.r)}</small><b style="${x.al && x.v ? 'color:' + x.cor : ''}">${nf(x.v)}</b><span>${esc(x.s)}</span></div>`).join('');
    $('#fer-kpis').querySelectorAll('.fer-kpi').forEach(el => el.addEventListener('click', () => {
      const f = k[+el.dataset.i].f;
      if (!f) return;
      filtro.sit = f.sit; filtro.mes = f.mes;
      $('#fer-sit').value = f.sit; $('#fer-mesf').value = f.mes;
      desenhar();
    }));

    // meses
    const porMes = meses.map(m => doSup.filter(c => c.mes === m).length);
    const pendMes = meses.map(m => doSup.filter(c => c.mes === m && c.sit !== 'Programada').length);
    const max = Math.max(1, ...porMes);
    $('#fer-leg').innerHTML = ALERTA.map(a => `<span><i style="background:${a.bg}"></i>${a.r}</span>`).join('') +
      `<span><i style="background:${OK.bg}"></i>todos programados</span>`;
    $('#fer-meses').innerHTML = meses.map((m, i) => {
      const cor = !porMes[i] ? null : pendMes[i] ? corMes(i) : OK;
      return `<div class="fer-mes${filtro.mes === m ? ' on' : ''}" data-m="${m}" title="${porMes[i]} completam em ${rotMes(m)} · ${pendMes[i]} sem programar">
      <div class="v" style="${cor ? `background:${cor.bg};color:${cor.fg}` : ''}">${porMes[i] || ''}</div><div class="bar" style="height:${porMes[i] / max * 120}px;${cor && cor !== OK && i < 3 ? 'background:' + cor.bg : ''}"></div><div class="r">${rotMes(m)}</div></div>`;
    }).join('');
    $('#fer-meses').querySelectorAll('.fer-mes').forEach(el => el.addEventListener('click', () => {
      filtro.mes = filtro.mes === el.dataset.m ? '' : el.dataset.m;
      $('#fer-mesf').value = filtro.mes;
      desenhar();
    }));

    // grade supervisor x mês
    const supsGrade = [...new Set(doSup.map(c => c.supervisor))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const cel = (v, pend, i, sup) => {
      if (!v) return '<td class="c zero">–</td>';
      const cor = pend ? corMes(i) : OK;
      const tip = `${v} completam em ${rotMes(meses[i])} · ${pend ? pend + ' sem programar' : 'todos programados'}`;
      return `<td class="c" data-sup="${esc(sup)}" data-m="${meses[i]}" style="cursor:pointer" title="${esc(tip)}"><span class="al" style="background:${cor.bg};color:${cor.fg}">${nf(v)}${pend && pend !== v ? `<small>(${pend})</small>` : pend ? '' : ' ✓'}</span></td>`;
    };
    $('#fer-grade').innerHTML = `<thead><tr><th>Supervisor</th>${meses.map(m => `<th class="c">${rotMes(m)}</th>`).join('')}<th class="c">Total</th><th class="c">Sem programar</th></tr></thead><tbody>` +
      supsGrade.map(s => {
        const g = doSup.filter(c => c.supervisor === s);
        const vals = meses.map(m => g.filter(c => c.mes === m).length);
        const pend = meses.map(m => g.filter(c => c.mes === m && c.sit !== 'Programada').length);
        const urg = g.filter(c => c.sit === 'Já completou' || c.sit === 'Programar agora').length;
        return `<tr><td style="font-weight:700">${esc(titulo(s))}</td>${vals.map((v, i) => cel(v, pend[i], i, s)).join('')}<td class="c" style="font-weight:900">${nf(vals.reduce((a, b) => a + b, 0))}</td>
          <td class="c">${urg ? `<span class="al" style="background:#b91c1c;color:#fff" title="já completaram ou completam em ${ANTECEDENCIA} dias e estão sem data">${nf(urg)}</span>` : '<span class="al" style="background:#15803d;color:#fff">✓</span>'}</td></tr>`;
      }).join('') +
      `<tr style="background:#f1f5f9"><td><b>Total</b></td>${porMes.map((v, i) => cel(v, pendMes[i], i, '')).join('')}<td class="c" style="font-weight:900">${nf(porMes.reduce((a, b) => a + b, 0))}</td>
        <td class="c" style="font-weight:900">${nf(doSup.filter(c => c.sit === 'Já completou' || c.sit === 'Programar agora').length)}</td></tr></tbody>`;
    $('#fer-grade').querySelectorAll('td[data-m]').forEach(td => td.addEventListener('click', () => {
      filtro.mes = td.dataset.m; $('#fer-mesf').value = filtro.mes;
      if (td.dataset.sup) { filtro.sup = td.dataset.sup; $('#fer-sup').value = filtro.sup; }
      desenhar();
      const t = $('#fer-tbody'); if (t) t.closest('.panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }));

    lista(doSup);
  }

  function lista(base) {
    const doSup = base || colaboradores().filter(c => !filtro.sup || c.supervisor === filtro.sup);
    const q = filtro.q.toLowerCase();
    const itens = doSup.filter(c => (!filtro.mes || c.mes === filtro.mes) && (!filtro.sit || c.sit === filtro.sit) &&
      (!q || String(c.nome).toLowerCase().includes(q) || String(c.matricula).includes(q)));
    listaExport = itens;
    $('#fer-cont').textContent = nf(itens.length) + ' colaborador(es)';
    $('#fer-tbody').innerHTML = itens.length ? itens.map(c => `<tr>
      <td style="font-weight:600">${esc(c.nome)}</td><td>${esc(c.matricula)}</td><td>${esc(titulo(c.supervisor))}</td>
      <td class="c">${br(c.adm)}</td><td class="c">${c.anos} ${c.anos === 1 ? 'ano' : 'anos'}</td>
      <td class="c" style="font-weight:700">${br(c.aniv)}</td><td class="c">${br(c.programar)}</td>
      <td class="c">${c.faltam < 0 ? 'há ' + nf(-c.faltam) + ' d' : nf(c.faltam) + ' d'}</td><td class="c">${br(c.limite)}</td>
      <td><span class="tag" style="background:${COR_SIT[c.sit]}">${esc(c.sit)}</span></td>
      <td>${c.periodos.length
        ? `<button type="button" class="fer-prog feito" data-k="${esc(chaveProg(c.matricula, c.aniv))}" title="Clique para alterar">✓ ${esc(resumoProg(c.periodos))}</button>`
        : `<button type="button" class="fer-prog${c.sit === 'Já completou' || c.sit === 'Programar agora' ? ' urg' : ''}" data-k="${esc(chaveProg(c.matricula, c.aniv))}">📅 Programar</button>`}</td></tr>`).join('')
      : '<tr><td colspan="11" style="text-align:center;color:#64748b">Nenhum colaborador com esses filtros.</td></tr>';
    $('#fer-tbody').querySelectorAll('.fer-prog').forEach(b => b.addEventListener('click', () => {
      const c = itens.find(x => chaveProg(x.matricula, x.aniv) === b.dataset.k);
      if (c) abrirProgramacao(c);
    }));
  }

  // ---------- formulário: supervisor preenche as datas ----------
  // CLT art. 134: até 3 períodos; um com 14 dias ou mais e os outros com 5 dias ou mais; total até 30.
  function abrirProgramacao(c) {
    const r = c.prog || {};
    const per = periodosDe(r);
    const opDias = v => '<option value="">—</option>' + Array.from({ length: 26 }, (_, i) => 30 - i).map(n => `<option value="${n}"${Number(v) === n ? ' selected' : ''}>${n} dias</option>`).join('');
    const linhaPer = i => `<div class="per"><b>${i + 1}º período</b><input type="date" id="fp-ini${i}" value="${esc(per[i] ? per[i].inicio : '')}" min="${iso(c.aniv)}" max="${iso(c.limite)}">
      <select id="fp-dias${i}">${opDias(per[i] ? per[i].dias : (i === 0 && !per.length ? 30 : ''))}</select><span class="fim" id="fp-fim${i}"></span></div>`;
    const fundo = document.createElement('div');
    fundo.className = 'fer-modal';
    fundo.innerHTML = `<div class="cx">
      <h3>📅 Programar férias</h3>
      <div style="font-size:14px;font-weight:700">${esc(c.nome)} <span style="font-weight:400;color:#475569">· matrícula ${esc(c.matricula)} · ${esc(titulo(c.supervisor))}</span></div>
      <div class="info">Admissão <b>${br(c.adm)}</b> · completa ${c.anos} ${c.anos === 1 ? 'ano' : 'anos'} em <b>${br(c.aniv)}</b><br>
        As férias devem começar a partir de <b>${br(c.aniv)}</b> e terminar até <b>${br(c.limite)}</b> (limite).<br>
        <span style="font-size:12px">Pode dividir em até 3 períodos: um com 14 dias ou mais e os outros com pelo menos 5 dias (total até 30).</span></div>
      ${[0, 1, 2].map(linhaPer).join('')}
      <label for="fp-por">Preenchido por</label><input id="fp-por" value="${esc(r.preenchidoPor || titulo(c.supervisor))}">
      <label for="fp-obs">Observação (opcional)</label><textarea id="fp-obs" placeholder="Ex.: vai vender 10 dias, combinar com a escala...">${esc(r.observacao || '')}</textarea>
      <div class="rod"><span id="fp-msg" style="font-size:12.5px;margin-right:auto"></span>
        ${per.length ? '<button type="button" class="apagar" data-apagar>Apagar programação</button>' : ''}
        <button type="button" data-x>Cancelar</button><button type="button" class="pri" data-salvar>Salvar</button></div></div>`;
    document.body.appendChild(fundo);
    const q = s => fundo.querySelector(s);
    const fechar = () => fundo.remove();
    fundo.addEventListener('click', e => { if (e.target === fundo) fechar(); });
    q('[data-x]').addEventListener('click', fechar);
    const ler = () => [0, 1, 2].map(i => {
      const ini = parseISO(q('#fp-ini' + i).value), dias = Number(q('#fp-dias' + i).value) || 0;
      return ini && dias ? { inicio: iso(ini), dias, fim: iso(addDias(ini, dias - 1)) } : null;
    }).filter(Boolean);
    const mostrarFins = () => [0, 1, 2].forEach(i => {
      const ini = parseISO(q('#fp-ini' + i).value), dias = Number(q('#fp-dias' + i).value) || 0;
      q('#fp-fim' + i).textContent = ini && dias ? 'até ' + br(addDias(ini, dias - 1)) : '';
    });
    fundo.addEventListener('input', mostrarFins); fundo.addEventListener('change', mostrarFins); mostrarFins();
    const msg = t => { q('#fp-msg').innerHTML = `<span style="color:#b91c1c">${esc(t)}</span>`; };
    const gravar = async (periodos, bt, rotulo) => {
      const o = { tipo: 'FERIAS', matricula: matNum(c.matricula), nome: c.nome, supervisor: c.supervisor, dataAdmissao: iso(c.adm),
        completaEm: iso(c.aniv), limite: iso(c.limite), periodos, observacao: q('#fp-obs').value.trim(), preenchidoPor: q('#fp-por').value.trim(),
        dataPreenchimento: iso(hoje()) };
      bt.disabled = true; const txt0 = bt.textContent; bt.textContent = rotulo;
      try {
        const resp0 = await fetch(FLUXO, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(o) });
        const t = await resp0.text();
        if (!resp0.ok) throw new Error('o fluxo respondeu ' + resp0.status);
        if (/"ok"\s*:\s*false/.test(t)) throw new Error(/Faltou a matr|etapa/i.test(t) ? 'o Excel ainda está com o script antigo: cole a versão nova do “Avaliação Experiência” (com Férias)' : 'o script recusou a gravação');
        o.__em = new Date().toISOString().replace('T', ' ').slice(0, 16);
        const k = chaveProg(c.matricula, c.aniv);
        PROG[k] = o; salvarLocal(k, o);
        fechar(); desenhar();
      } catch (e) {
        msg('Não gravou: ' + e.message);
        bt.disabled = false; bt.textContent = txt0;
      }
    };
    q('[data-salvar]').addEventListener('click', () => {
      const p = ler();
      if (!p.length) return msg('Preencha a data de início e os dias do 1º período.');
      const total = p.reduce((s, x) => s + x.dias, 0);
      if (total > 30) return msg(`O total passou de 30 dias (${total}).`);
      if (p.length > 1 && !p.some(x => x.dias >= 14)) return msg('Quando divide, um dos períodos precisa ter 14 dias ou mais.');
      if (p.length > 1 && p.some(x => x.dias < 5)) return msg('Cada período precisa ter pelo menos 5 dias.');
      const ord = p.slice().sort((a, b) => a.inicio < b.inicio ? -1 : 1);
      for (let i = 1; i < ord.length; i++) if (ord[i].inicio <= ord[i - 1].fim) return msg('Os períodos estão se sobrepondo.');
      if (p.some(x => x.inicio < iso(c.aniv))) return msg(`As férias só podem começar a partir de ${br(c.aniv)} (quando completa o período).`);
      if (p.some(x => x.fim > iso(c.limite))) return msg(`Tem período terminando depois do limite (${br(c.limite)}).`);
      gravar(ord, q('[data-salvar]'), 'Gravando…');
    });
    const ap = q('[data-apagar]');
    if (ap) ap.addEventListener('click', () => gravar([], ap, 'Apagando…'));
  }

  function exportar() {
    const cab = ['Colaborador', 'Matrícula', 'Supervisor', 'Coordenador', 'Função', 'Admissão', 'Completa (anos)', 'Data que completa', 'Programar a partir de', 'Faltam (dias)', 'Limite para tirar', 'Situação', 'Férias programadas', 'Preenchido por'];
    const linhas = listaExport.map(c => [c.nome, c.matricula, titulo(c.supervisor), titulo(c.coordenador), c.funcao, br(c.adm), c.anos, br(c.aniv), br(c.programar), c.faltam, br(c.limite), c.sit,
      c.periodos.map(p => br(parseISO(p.inicio)) + ' a ' + br(parseISO(p.fim)) + ' (' + p.dias + ' d)').join(' + '), c.prog && c.periodos.length ? c.prog.preenchidoPor || '' : '']);
    const csv = '﻿' + [cab].concat(linhas).map(l => l.map(v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"').join(';')).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = 'programacao_ferias_' + chaveMes(hoje()) + '.csv';
    document.body.appendChild(a); a.click(); a.remove();
  }

  // monta a aba quando a página existir e redesenha quando os colaboradores chegarem
  function vigiar() {
    montar();
    if (montado && !progCarregado) { progCarregado = true; carregarProg(); }
    let qtd = 0;
    try { qtd = typeof COLABORADORES !== 'undefined' && COLABORADORES ? COLABORADORES.length : 0; } catch (e) { qtd = 0; }
    if (qtd !== ultimaQtd) { ultimaQtd = qtd; desenhar(); }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', vigiar);
  else vigiar();
  setInterval(vigiar, 1500);
  window.PainelFerias = { redesenhar: desenhar };
})();
