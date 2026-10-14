// gestao_toa.js — Seção "Produtividade TOA" do Painel-SLA
//
// Lê "toa" e "colaboradores" do dados.json (gerados pelo Office Script, item 48).
//  - Produtividade: atividades com Status "Concluída" (sem Refeição), por técnico/dia
//    e por supervisor (coluna "Supervisor" da aba TOA).
//  - Aderência: % dos técnicos ATIVOS de cada supervisor (aba Organograma, Área Técnica
//    = Técnico) que lançaram pelo menos 1 atividade no TOA no dia (qualquer status,
//    sem Refeição). O cruzamento é pelo nome do técnico.
//
// Para ligar: no index.html, adicione a linha abaixo ANTES de <script src="tela_inicial.js">
//   <script src="gestao_toa.js"></script>
// O item do menu e a página são criados por este arquivo.

(function(){
  'use strict';

  const COR = { good:'#059669', warn:'#d97706', bad:'#dc2626', accent:'#0e7c86', muted:'#6b7590', line:'#e1e5f0', text:'#1b2440' };
  const META_ADERENCIA = 0.90;              // meta de aderência (90%) — ajuste aqui
  const RECARREGAR_A_CADA_MS = 5 * 60 * 1000;
  const DIAS_SEMANA = ['dom','seg','ter','qua','qui','sex','sáb'];
  const PALAVRAS_IGNORADAS = ['de','da','do','das','dos','e'];
  // Contratos que NÃO entram na seção TOA (comparação sem acento e sem maiúsculas).
  // Vale para Contrato_2 e Unidade de Negócio do Organograma.
  const CONTRATOS_EXCLUIDOS = ['cirion'];

  let TOA = [];      // [{ data:'aaaa-mm-dd', recurso, supervisor, total, concluidas }]
  let COLABS = [];
  let ultimaCarga = 0;
  let carregando = null;
  const estado = { inicio:'', fim:'', supervisor:'', ignorarFds:true, busca:'', aba:'d1',
    quadro: { d1: { dia:'', status:'', busca:'' }, d0: { dia:'', status:'', busca:'' } } };
  let chartSup = null, chartDia = null;

  // ---------- utilitários ----------
  function esc(s){ return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function semAcento(s){ return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\s+/g,' ').trim(); }
  function pad2(n){ return n < 10 ? '0' + n : String(n); }
  function isoDeData(d){ return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function dataDeIso(iso){ const p = iso.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function somarDias(iso, n){ const d = dataDeIso(iso); d.setDate(d.getDate() + n); return isoDeData(d); }
  function fmtData(iso){ const p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
  function fmtDiaCurto(iso){ const p = iso.split('-'); return p[2] + '/' + p[1]; }
  function diaSemana(iso){ return dataDeIso(iso).getDay(); }
  function fimDeSemana(iso){ const d = diaSemana(iso); return d === 0 || d === 6; }
  function listaDias(ini, fim){
    const out = [];
    if(!ini || !fim || ini > fim) return out;
    let d = dataDeIso(ini);
    const f = dataDeIso(fim);
    while(d <= f && out.length < 400){ out.push(isoDeData(d)); d.setDate(d.getDate() + 1); }
    return out;
  }
  function fmt1(n){ return n.toLocaleString('pt-BR', { minimumFractionDigits:1, maximumFractionDigits:1 }); }
  function pct(n){ return (n * 100).toFixed(0) + '%'; }
  function corAderencia(p){ return p >= META_ADERENCIA ? COR.good : (p >= META_ADERENCIA - 0.2 ? COR.warn : COR.bad); }
  function fundoAderencia(p){
    if(p >= META_ADERENCIA) return 'rgba(5,150,105,.14)';
    if(p >= META_ADERENCIA - 0.2) return 'rgba(217,119,6,.14)';
    return 'rgba(220,38,38,.14)';
  }

  // Chaves de nome para cruzar TOA x Organograma: nome completo e "primeiro + último".
  function chavesNome(s){
    const base = semAcento(String(s || '').replace(/\(.*?\)/g, ' ')).replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
    if(!base) return [];
    const partes = base.split(' ').filter(x => PALAVRAS_IGNORADAS.indexOf(x) === -1);
    const out = [base];
    if(partes.length >= 2) out.push(partes[0] + ' ' + partes[partes.length - 1]);
    return out;
  }
  function chavePrincipal(s){ const k = chavesNome(s); return k.length ? k[0] : ''; }
  function mesmoSupervisor(a, b){ return semAcento(a) === semAcento(b); }

  // ---------- abas D-1 / D-0 ----------
  function quadroHtml(q, titulo, dica){
    return `<div class="panel" style="margin-bottom:16px;">
      <h2>${titulo}</h2>
      <div class="hint">${dica} Técnicos ativos do Organograma, sem Cirion.</div>
      <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:4px 0 12px;">
        <label for="toaQ${q}Dia" class="mono" style="${ESTILO_LABEL}">Dia</label>
        <input type="date" id="toaQ${q}Dia" class="mono" style="${ESTILO_SELECT}">
        <div class="chip" id="toaQ${q}Padrao">${q === 'd0' ? 'Base mais recente' : 'Dia anterior'}</div>
        <label for="toaQ${q}Status" class="mono" style="${ESTILO_LABEL} margin-left:8px;">Mostrar</label>
        <select id="toaQ${q}Status" class="mono" style="${ESTILO_SELECT} min-width:200px;">
          <option value="">Todos</option>
          <option value="fez">Fez serviço</option>
          <option value="sem">Na base sem concluir</option>
          <option value="fora">Não apareceu na base</option>
        </select>
        <div class="search" style="margin-left:8px;">
          <input id="toaQ${q}Busca" type="search" placeholder="Buscar técnico..." aria-label="Buscar técnico">
        </div>
        <span id="toaQ${q}Contagem" class="mono" style="font-size:11px; color:var(--muted); margin-left:auto;"></span>
      </div>
      <div class="kpi-row" id="toaQ${q}Kpis" style="margin-bottom:14px;"></div>
      <div style="overflow:auto; max-height:520px; border-radius:8px;"><table id="toaQ${q}Table"></table></div>
    </div>`;
  }

  // D-0 = data mais recente da base; D-1 = dia anterior que tem dados (ou o dia de calendário anterior).
  function diaPadrao(q){
    const max = maiorData();
    if(!max) return '';
    if(q === 'd0') return max;
    const anteriores = Array.from(new Set(TOA.map(t => t.data))).filter(d => d < max).sort();
    return anteriores.length ? anteriores[anteriores.length - 1] : somarDias(max, -1);
  }

  function mostrarAba(){
    ['d1', 'd0'].forEach(a => {
      const el = document.getElementById('toaAba-' + a);
      if(el) el.style.display = estado.aba === a ? '' : 'none';
    });
    document.querySelectorAll('#pagina-toa [data-toa-aba]').forEach(b => b.classList.toggle('ativa', b.dataset.toaAba === estado.aba));
  }

  // ---------- estrutura (menu + página) ----------
  const ESTILO_SELECT = 'font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text);';
  const ESTILO_LABEL = 'font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.05em;';

  function montarEstrutura(){
    if(document.getElementById('pagina-toa')) return;

    const ref = document.getElementById('nav-afaz') || document.getElementById('nav-organograma');
    const nav = document.createElement('div');
    nav.className = 'nav-item';
    nav.id = 'nav-toa';
    nav.innerHTML = '<span class="nav-icon">📲</span> Produtividade TOA';
    nav.addEventListener('click', () => {
      if(typeof mostrarPagina === 'function') mostrarPagina('toa');
      redesenhar();
    });
    if(ref && ref.parentNode) ref.parentNode.insertBefore(nav, ref.nextSibling);

    if(typeof PAGINAS !== 'undefined' && Array.isArray(PAGINAS) && PAGINAS.indexOf('toa') === -1) PAGINAS.push('toa');

    const main = document.querySelector('main.main-content');
    if(!main) return;
    const pg = document.createElement('div');
    pg.className = 'pagina';
    pg.id = 'pagina-toa';
    pg.style.display = 'none';
    pg.innerHTML = `
      <style>
        #pagina-toa .toa-abas{ display:flex; gap:8px; flex-wrap:wrap; margin:6px 0 16px; }
        #pagina-toa .toa-aba{ display:flex; align-items:center; gap:8px; padding:10px 16px; border-radius:10px;
          border:1px solid var(--line); background:var(--panel-2); color:var(--text); font-family:'Inter', sans-serif;
          font-size:13px; font-weight:600; cursor:pointer; }
        #pagina-toa .toa-aba:hover{ background:rgba(14,124,134,.10); }
        #pagina-toa .toa-aba.ativa{ background:rgba(14,124,134,.14); color:var(--accent); border-color:rgba(14,124,134,.35); }
        #pagina-toa .toa-aba small{ font-weight:500; color:var(--muted); font-size:11px; }
      </style>
      <div class="panel">
        <h2>Produtividade TOA</h2>
        <div class="hint" id="toaInfo">Carregando dados do TOA...</div>

        <div class="toa-abas" role="tablist">
          <button type="button" class="toa-aba" data-toa-aba="d1" role="tab">📊 D-1 <small id="toaAbaD1Dia"></small></button>
          <button type="button" class="toa-aba" data-toa-aba="d0" role="tab">📲 D-0 <small id="toaAbaD0Dia"></small></button>
        </div>

        <div style="margin:0 0 14px; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
          <label for="toaSupervisor" class="mono" style="${ESTILO_LABEL}">Supervisor</label>
          <select id="toaSupervisor" class="mono" style="${ESTILO_SELECT} min-width:220px;"><option value="">Todos</option></select>
        </div>

        <div id="toaAba-d1">
          ${quadroHtml('d1', 'D-1 — técnicos na base do dia anterior', 'Fechamento do dia anterior à base mais recente: quem apareceu na base do TOA, quem fez serviço e quem da equipe não apareceu.')}

          <div style="margin:4px 0 14px; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
            <label for="toaInicio" class="mono" style="${ESTILO_LABEL}">Período de</label>
            <input type="date" id="toaInicio" class="mono" style="${ESTILO_SELECT}">
            <label for="toaFim" class="mono" style="${ESTILO_LABEL} margin-left:6px;">até</label>
            <input type="date" id="toaFim" class="mono" style="${ESTILO_SELECT}">
            <div class="chip" data-toa-atalho="7">7 dias</div>
            <div class="chip" data-toa-atalho="mes">Mês</div>
            <label class="mono" style="font-size:12px; color:var(--text); display:flex; align-items:center; gap:6px; margin-left:8px; cursor:pointer;">
              <input type="checkbox" id="toaIgnorarFds" checked> Ignorar sábado e domingo na aderência
            </label>
            <button type="button" class="link-base" style="margin-left:auto; cursor:pointer; border:1px solid var(--accent); font-family:'JetBrains Mono', monospace;" id="toaExportar">⬇ Exportar (CSV)</button>
          </div>

          <div class="kpi-row" id="toaKpis"></div>

          <div class="grid">
            <div class="panel">
              <h2>Média de atividades concluídas por técnico/dia — por supervisor</h2>
              <div class="hint">Concluídas ÷ (técnicos × dias com lançamento). Supervisor da aba TOA.</div>
              <div id="toaSupChartWrap" style="position:relative; height:300px;"><canvas id="toaSupChart"></canvas></div>
            </div>
            <div class="panel">
              <h2>Concluídas por dia</h2>
              <div class="hint">Barras: total de concluídas. Linha: média por técnico que lançou no dia.</div>
              <canvas id="toaDiaChart" height="190"></canvas>
            </div>
          </div>

          <div class="panel" style="margin-bottom:16px;">
            <h2>Aderência ao TOA por supervisor</h2>
            <div class="hint">% dos técnicos ativos da equipe (Organograma) que lançaram ao menos 1 atividade no dia. Passe o mouse para ver quem não lançou. Meta ${pct(META_ADERENCIA)}.</div>
            <div style="overflow-x:auto; border-radius:8px;"><table id="toaAderenciaTable"></table></div>
            <div id="toaNaoEncontrados" class="mono" style="font-size:11px; color:var(--muted); margin-top:10px;"></div>
          </div>

          <div class="panel">
            <h2>Produtividade por técnico por dia</h2>
            <div class="hint">Atividades concluídas por dia. Vazio = sem lançamento; 0 = lançou mas não concluiu nada.</div>
            <div style="display:flex; gap:8px; margin-bottom:10px; align-items:center; flex-wrap:wrap;">
              <div class="search">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden style="opacity:.6"><path d="M21 21l-4.35-4.35" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="11" cy="11" r="6" stroke="currentColor" stroke-width="2"/></svg>
                <input id="toaBusca" type="search" placeholder="Buscar técnico ou supervisor..." aria-label="Buscar técnico">
              </div>
              <span id="toaContagem" class="mono" style="font-size:11px; color:var(--muted); margin-left:auto;"></span>
            </div>
            <div style="overflow-x:auto; border-radius:8px;"><table id="toaTecTable"></table></div>
          </div>
        </div>

        <div id="toaAba-d0" style="display:none;">
          ${quadroHtml('d0', 'D-0 — técnicos na base de hoje', 'Base mais recente do TOA: quem já aparece, quem já fez serviço (concluiu ao menos 1 atividade) e quem da equipe ainda não apareceu.')}
        </div>
      </div>`;
    main.appendChild(pg);

    document.getElementById('toaInicio').addEventListener('change', e => { estado.inicio = e.target.value; render(); });
    document.getElementById('toaFim').addEventListener('change', e => { estado.fim = e.target.value; render(); });
    document.getElementById('toaSupervisor').addEventListener('change', e => { estado.supervisor = e.target.value; render(); });
    document.getElementById('toaIgnorarFds').addEventListener('change', e => { estado.ignorarFds = e.target.checked; render(); });
    document.getElementById('toaExportar').addEventListener('click', exportarCsv);
    ['d1', 'd0'].forEach(q => {
      const st = estado.quadro[q];
      document.getElementById('toaQ' + q + 'Dia').addEventListener('change', e => { st.dia = e.target.value; renderQuadro(q); });
      document.getElementById('toaQ' + q + 'Status').addEventListener('change', e => { st.status = e.target.value; renderQuadro(q); });
      document.getElementById('toaQ' + q + 'Padrao').addEventListener('click', () => { st.dia = diaPadrao(q); renderQuadro(q); });
      let tq;
      document.getElementById('toaQ' + q + 'Busca').addEventListener('input', e => {
        clearTimeout(tq);
        tq = setTimeout(() => { st.busca = e.target.value; renderQuadro(q); }, 200);
      });
    });
    pg.querySelectorAll('[data-toa-aba]').forEach(b => {
      b.addEventListener('click', () => { estado.aba = b.dataset.toaAba; mostrarAba(); render(); });
    });
    mostrarAba();
    let t;
    document.getElementById('toaBusca').addEventListener('input', e => {
      clearTimeout(t);
      t = setTimeout(() => { estado.busca = e.target.value; renderTabelaTecnicos(); }, 200);
    });
    pg.querySelectorAll('[data-toa-atalho]').forEach(ch => {
      ch.addEventListener('click', () => {
        const fim = diaPadrao('d1');
        if(!fim) return;
        estado.fim = fim;
        estado.inicio = ch.dataset.toaAtalho === 'mes' ? fim.slice(0, 8) + '01' : somarDias(fim, -6);
        render();
      });
    });
  }

  // ---------- cruzamento de nomes TOA x Organograma ----------
  // O Organograma é a referência: cada nome do TOA é trocado pelo nome do
  // Organograma correspondente. Aceita acento diferente, sobrenome a mais/a menos,
  // ordem diferente e 1 letra trocada em nomes com 5+ letras.
  let AJUSTES = [];   // [{ toa, organograma }]

  function tokensNome(s){
    return semAcento(String(s || '').replace(/\(.*?\)/g, ' '))
      .replace(/[^a-z ]/g, ' ').split(/\s+/)
      .filter(x => x && PALAVRAS_IGNORADAS.indexOf(x) === -1);
  }
  function distancia(a, b){
    let ant = Array.from({ length: b.length + 1 }, (_, j) => j);
    for(let i = 1; i <= a.length; i++){
      const atual = [i];
      for(let j = 1; j <= b.length; j++){
        atual[j] = Math.min(ant[j] + 1, atual[j - 1] + 1, ant[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      }
      ant = atual;
    }
    return ant[b.length];
  }
  function tokIgual(a, b){
    if(a === b) return true;
    return a.length >= 5 && b.length >= 5 && Math.abs(a.length - b.length) <= 1 && distancia(a, b) <= 1;
  }
  // Pontua o quanto o nome do TOA (tt) bate com o do Organograma (to). 0 = não bate.
  function pontuar(tt, to){
    if(!tt.length || !to.length || !tokIgual(tt[0], to[0])) return 0;   // primeiro nome tem que bater
    let comuns = 0;
    tt.forEach(x => { if(to.some(y => tokIgual(x, y))) comuns++; });
    if(comuns < 2) return 0;
    const ultimoOk = to.some(y => tokIgual(tt[tt.length - 1], y)) ? 5 : 0;
    return comuns * 10 + ultimoOk - Math.abs(tt.length - to.length);
  }

  function ajustarNomesToa(){
    AJUSTES = [];
    const org = COLABS.filter(c => c.ativo && c.nome).map(c => ({ nome: c.nome, tokens: tokensNome(c.nome), sup: semAcento(c.supervisor) }));
    const porChave = {};
    org.forEach(o => chavesNome(o.nome).forEach(k => { if(!porChave[k]) porChave[k] = o.nome; }));

    const resolvido = {};   // nome TOA -> nome Organograma
    const nomesToa = {};
    TOA.forEach(t => { if(!nomesToa[t.recurso]) nomesToa[t.recurso] = semAcento(t.supervisor); });

    Object.keys(nomesToa).forEach(nomeToa => {
      // 1) igual (sem acento) ou primeiro + último nome
      const exato = chavesNome(nomeToa).map(k => porChave[k]).find(Boolean);
      if(exato){ resolvido[nomeToa] = exato; return; }

      // 2) melhor pontuação, desde que não empate com outro nome
      const tt = tokensNome(nomeToa);
      let melhor = null, pMelhor = 0, pSegundo = 0;
      org.forEach(o => {
        const p = pontuar(tt, o.tokens);
        if(p > pMelhor){ pSegundo = pMelhor; pMelhor = p; melhor = o; }
        else if(p > pSegundo){ pSegundo = p; }
      });
      if(melhor && pMelhor > pSegundo){ resolvido[nomeToa] = melhor.nome; return; }

      // 3) mesmo primeiro nome e mesmo supervisor, e só um assim na equipe
      const sup = nomesToa[nomeToa];
      if(sup && tt.length){
        const candidatos = org.filter(o => o.sup === sup && o.tokens.length && tokIgual(o.tokens[0], tt[0]));
        if(candidatos.length === 1){ resolvido[nomeToa] = candidatos[0].nome; }
      }
    });

    TOA.forEach(t => {
      const novo = resolvido[t.recurso];
      if(novo && novo !== t.recurso){
        if(!AJUSTES.some(a => a.toa === t.recurso)) AJUSTES.push({ toa: t.recurso, organograma: novo });
        t.recursoToa = t.recurso;
        t.recurso = novo;
      }
    });
    AJUSTES.sort((a, b) => a.organograma.localeCompare(b.organograma, 'pt-BR'));
  }

  // ---------- dados ----------
  function maiorData(){
    let m = '';
    TOA.forEach(t => { if(t.data > m) m = t.data; });
    return m;
  }
  function menorData(){
    let m = '';
    TOA.forEach(t => { if(!m || t.data < m) m = t.data; });
    return m;
  }

  function carregar(forcar){
    if(!forcar && ultimaCarga > 0 && Date.now() - ultimaCarga < RECARREGAR_A_CADA_MS) return Promise.resolve();
    if(carregando) return carregando;
    carregando = fetch('dados.json?_=' + Date.now())
      .then(r => { if(!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(d => {
        const todosColabs = Array.isArray(d.colaboradores) ? d.colaboradores : [];
        const excluido = c => CONTRATOS_EXCLUIDOS.some(x =>
          semAcento(c.contrato).includes(x) || semAcento(c.unidadeNegocio).includes(x));
        // Organograma sem Cirion (equipes, supervisores e aderência)
        COLABS = todosColabs.filter(c => !excluido(c));
        // TOA sem os técnicos que são da Cirion no Organograma
        const nomesExcluidos = new Set();
        todosColabs.filter(excluido).forEach(c => chavesNome(c.nome).forEach(k => nomesExcluidos.add(k)));
        const supsExcluidos = new Set(todosColabs.filter(excluido).map(c => semAcento(c.supervisor)).filter(Boolean));
        const supsValidos = new Set(COLABS.map(c => semAcento(c.supervisor)).filter(Boolean));
        TOA = (Array.isArray(d.toa) ? d.toa : []).filter(t => {
          if(!t || !t.data || !t.recurso) return false;
          if(chavesNome(t.recurso).some(k => nomesExcluidos.has(k))) return false;
          const s = semAcento(t.supervisor);
          if(s && supsExcluidos.has(s) && !supsValidos.has(s)) return false; // supervisor só da Cirion
          return true;
        });
        ajustarNomesToa();
        ultimaCarga = Date.now();
        if(!estado.fim && TOA.length){
          estado.fim = diaPadrao('d1');
          estado.inicio = somarDias(estado.fim, -6);
        }
      })
      .catch(e => {
        console.error('[TOA]', e);
        const info = document.getElementById('toaInfo');
        if(info){ info.textContent = 'Não foi possível carregar o dados.json (' + e.message + ').'; info.style.color = COR.bad; }
      })
      .finally(() => { carregando = null; });
    return carregando;
  }

  // Calcula tudo que as visões usam, a partir dos filtros atuais.
  function calcular(){
    const dias = listaDias(estado.inicio, estado.fim);
    const diasAderencia = estado.ignorarFds ? dias.filter(d => !fimDeSemana(d)) : dias;

    const noPeriodo = TOA.filter(t => t.data >= estado.inicio && t.data <= estado.fim);
    const toaF = estado.supervisor ? noPeriodo.filter(t => mesmoSupervisor(t.supervisor, estado.supervisor)) : noPeriodo;

    const tecnicos = COLABS.filter(c => c.ativo && semAcento(c.areaTecnica) === 'tecnico' && c.nome);
    const tecF = estado.supervisor ? tecnicos.filter(c => mesmoSupervisor(c.supervisor, estado.supervisor)) : tecnicos;

    // Quem lançou em cada dia (todas as equipes; o cruzamento é pelo nome).
    const lancouPorDia = {};
    noPeriodo.forEach(t => {
      if(!(Number(t.total) > 0)) return;
      if(!lancouPorDia[t.data]) lancouPorDia[t.data] = new Set();
      chavesNome(t.recurso).forEach(k => lancouPorDia[t.data].add(k));
    });
    function lancou(c, dia){
      const set = lancouPorDia[dia];
      return !!set && chavesNome(c.nome).some(k => set.has(k));
    }

    // Aderência por supervisor (Organograma).
    const equipes = {};
    tecF.forEach(c => {
      const label = String(c.supervisor || '').trim() || 'Sem supervisor';
      const k = semAcento(label);
      if(!equipes[k]) equipes[k] = { label, membros: [] };
      equipes[k].membros.push(c);
    });
    let somaLancaram = 0, somaEsperados = 0;
    const aderencia = Object.values(equipes).map(eq => {
      let lanc = 0, esp = 0;
      const porDia = diasAderencia.map(dia => {
        const semLanc = eq.membros.filter(c => !lancou(c, dia));
        const n = eq.membros.length - semLanc.length;
        lanc += n; esp += eq.membros.length;
        return { dia, lancaram: n, equipe: eq.membros.length, semLanc: semLanc.map(c => c.nome) };
      });
      const nunca = eq.membros.filter(c => !diasAderencia.some(dia => lancou(c, dia))).map(c => c.nome);
      somaLancaram += lanc; somaEsperados += esp;
      return { label: eq.label, equipe: eq.membros.length, porDia, media: esp > 0 ? lanc / esp : 0, nunca };
    }).sort((a, b) => b.media - a.media);

    // Técnicos do TOA que não existem no Organograma (ativos) — ajuda a achar nomes diferentes.
    const chavesOrg = new Set();
    COLABS.filter(c => c.ativo && c.nome).forEach(c => chavesNome(c.nome).forEach(k => chavesOrg.add(k)));
    const naoEncontrados = Array.from(new Set(
      toaF.filter(t => !chavesNome(t.recurso).some(k => chavesOrg.has(k))).map(t => t.recurso)
    )).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    // Produtividade por técnico.
    const porTec = {};
    toaF.forEach(t => {
      const k = chavePrincipal(t.recurso);
      if(!k) return;
      if(!porTec[k]) porTec[k] = { nome: t.recurso, supervisor: t.supervisor || '', dias: {}, concluidas: 0, total: 0, diasLanc: 0 };
      const p = porTec[k];
      if(!p.supervisor && t.supervisor) p.supervisor = t.supervisor;
      const c = Number(t.concluidas) || 0;
      const tot = Number(t.total) || 0;
      if(!p.dias[t.data]) p.dias[t.data] = { concluidas: 0, total: 0 };
      p.dias[t.data].concluidas += c;
      p.dias[t.data].total += tot;
      p.concluidas += c;
      p.total += tot;
    });
    Object.values(porTec).forEach(p => {
      p.diasLanc = Object.values(p.dias).filter(d => d.total > 0).length;
      p.media = p.diasLanc > 0 ? p.concluidas / p.diasLanc : 0;
    });

    // Produtividade por supervisor (coluna Supervisor do TOA).
    const porSup = {};
    toaF.forEach(t => {
      const label = String(t.supervisor || '').trim() || 'Sem supervisor';
      const k = semAcento(label);
      if(!porSup[k]) porSup[k] = { label, concluidas: 0, pares: 0, tecnicos: new Set() };
      porSup[k].concluidas += Number(t.concluidas) || 0;
      if(Number(t.total) > 0){ porSup[k].pares++; porSup[k].tecnicos.add(chavePrincipal(t.recurso)); }
    });
    const supervisores = Object.values(porSup)
      .map(s => ({ label: s.label, concluidas: s.concluidas, tecnicos: s.tecnicos.size, media: s.pares > 0 ? s.concluidas / s.pares : 0 }))
      .sort((a, b) => b.media - a.media);

    // Por dia.
    const porDia = dias.map(dia => {
      const doDia = toaF.filter(t => t.data === dia);
      const concl = doDia.reduce((s, t) => s + (Number(t.concluidas) || 0), 0);
      const tecs = new Set(doDia.filter(t => Number(t.total) > 0).map(t => chavePrincipal(t.recurso))).size;
      return { dia, concluidas: concl, tecnicos: tecs, media: tecs > 0 ? concl / tecs : null };
    });

    const totalConcluidas = toaF.reduce((s, t) => s + (Number(t.concluidas) || 0), 0);
    const totalLancadas = toaF.reduce((s, t) => s + (Number(t.total) || 0), 0);
    const pares = toaF.filter(t => Number(t.total) > 0).length;

    return {
      dias, diasAderencia, toaF, aderencia, naoEncontrados, porTec, supervisores, porDia,
      kpi: {
        concluidas: totalConcluidas,
        lancadas: totalLancadas,
        tecnicos: Object.keys(porTec).length,
        mediaTecDia: pares > 0 ? totalConcluidas / pares : 0,
        aderencia: somaEsperados > 0 ? somaLancaram / somaEsperados : null,
        equipe: tecF.length
      }
    };
  }

  // ---------- render ----------
  let ultimoCalculo = null;

  function render(){
    const pg = document.getElementById('pagina-toa');
    if(!pg) return;
    const info = document.getElementById('toaInfo');

    if(!TOA.length){
      info.textContent = 'Sem dados do TOA no dados.json. Rode o Office Script atualizado (item 48) com a aba "TOA" preenchida.';
      ['toaKpis','toaAderenciaTable','toaTecTable','toaNaoEncontrados','toaQd1Kpis','toaQd1Table','toaQd0Kpis','toaQd0Table'].forEach(id => { const el = document.getElementById(id); if(el) el.innerHTML = ''; });
      return;
    }

    const min = menorData(), max = maiorData();
    if(!estado.fim) estado.fim = diaPadrao('d1');
    if(!estado.inicio) estado.inicio = somarDias(estado.fim, -6);
    if(estado.inicio > estado.fim){ const x = estado.inicio; estado.inicio = estado.fim; estado.fim = x; }
    info.style.color = '';
    info.textContent = 'Base TOA de ' + fmtData(min) + ' a ' + fmtData(max) + ' · período: ' + fmtData(estado.inicio) + ' a ' + fmtData(estado.fim) + '.';

    const inpIni = document.getElementById('toaInicio');
    const inpFim = document.getElementById('toaFim');
    inpIni.value = estado.inicio; inpFim.value = estado.fim;
    inpIni.min = inpFim.min = min; inpIni.max = inpFim.max = max;

    // opções de supervisor (TOA + Organograma)
    const sel = document.getElementById('toaSupervisor');
    const sups = {};
    TOA.forEach(t => { const s = String(t.supervisor || '').trim(); if(s && !sups[semAcento(s)]) sups[semAcento(s)] = s; });
    COLABS.filter(c => c.ativo && semAcento(c.areaTecnica) === 'tecnico').forEach(c => {
      const s = String(c.supervisor || '').trim(); if(s && !sups[semAcento(s)]) sups[semAcento(s)] = s;
    });
    const listaSups = Object.values(sups).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    sel.innerHTML = '<option value="">Todos</option>' + listaSups.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
    sel.value = listaSups.some(s => s === estado.supervisor) ? estado.supervisor : '';
    if(!sel.value) estado.supervisor = '';
    document.getElementById('toaIgnorarFds').checked = estado.ignorarFds;

    const r = calcular();
    ultimoCalculo = r;

    renderQuadro('d1');
    renderQuadro('d0');
    const a1 = document.getElementById('toaAbaD1Dia'), a0 = document.getElementById('toaAbaD0Dia');
    if(a1) a1.textContent = fmtDiaCurto(estado.quadro.d1.dia || diaPadrao('d1'));
    if(a0) a0.textContent = fmtDiaCurto(estado.quadro.d0.dia || diaPadrao('d0'));
    renderKpis(r);
    renderGraficos(r);
    renderAderencia(r);
    renderTabelaTecnicos();
  }

  // ---------- quadro do dia (D-1 e D-0): quem está na base e quem fez serviço ----------
  function renderQuadro(q){
    const st = estado.quadro[q];
    const tabela = document.getElementById('toaQ' + q + 'Table');
    if(!tabela || !TOA.length) return;
    const max = maiorData();
    if(!st.dia || st.dia > max) st.dia = diaPadrao(q);
    const inp = document.getElementById('toaQ' + q + 'Dia');
    inp.value = st.dia; inp.min = menorData(); inp.max = max;
    document.getElementById('toaQ' + q + 'Status').value = st.status;

    const dia = st.dia;
    const doDia = TOA.filter(t => t.data === dia && (!estado.supervisor || mesmoSupervisor(t.supervisor, estado.supervisor)));
    const tecnicos = COLABS.filter(c => c.ativo && c.nome && semAcento(c.areaTecnica) === 'tecnico' &&
      (!estado.supervisor || mesmoSupervisor(c.supervisor, estado.supervisor)));

    // Técnicos que aparecem na base do dia
    const linhas = {};
    doDia.forEach(t => {
      const k = chavePrincipal(t.recurso);
      if(!k) return;
      if(!linhas[k]) linhas[k] = { nome: t.recurso, supervisor: t.supervisor || '', total: 0, concluidas: 0, noOrg: false, ausente: false };
      linhas[k].total += Number(t.total) || 0;
      linhas[k].concluidas += Number(t.concluidas) || 0;
    });
    // Equipe do Organograma: marca quem apareceu e inclui quem não apareceu
    tecnicos.forEach(c => {
      const k = chavesNome(c.nome).find(x => linhas[x]);
      if(k){
        linhas[k].noOrg = true;
        if(!linhas[k].supervisor) linhas[k].supervisor = c.supervisor || '';
      } else {
        const kn = chavePrincipal(c.nome);
        if(kn && !linhas[kn]) linhas[kn] = { nome: c.nome, supervisor: c.supervisor || '', total: 0, concluidas: 0, noOrg: true, ausente: true };
      }
    });

    const todas = Object.values(linhas).map(l => {
      l.situacao = l.ausente ? 'fora' : (l.concluidas > 0 ? 'fez' : 'sem');
      return l;
    });
    const naBase = todas.filter(l => !l.ausente);
    const fez = todas.filter(l => l.situacao === 'fez');
    const sem = todas.filter(l => l.situacao === 'sem');
    const fora = todas.filter(l => l.situacao === 'fora');
    const concluidas = naBase.reduce((s, l) => s + l.concluidas, 0);

    const kpi = (rotulo, valor, cor, extra) => `<div class="kpi"><div class="label">${rotulo}</div><div class="value" style="color:${cor}">${valor}</div>${extra ? `<div class="delta" style="color:${COR.muted};">${extra}</div>` : ''}</div>`;
    document.getElementById('toaQ' + q + 'Kpis').innerHTML =
      kpi('📅 Dia', fmtData(dia).slice(0, 5) + ' ' + DIAS_SEMANA[diaSemana(dia)], COR.text) +
      kpi('📲 Na base do TOA', naBase.length, COR.accent, 'técnicos com lançamento no dia') +
      kpi('✅ Fizeram serviço', fez.length, COR.good, concluidas + ' atividades concluídas') +
      kpi('⏳ Na base sem concluir', sem.length, COR.warn) +
      kpi('🚫 Não apareceram', fora.length, COR.bad, 'técnicos da equipe sem lançamento');

    const ordem = { fez: 0, sem: 1, fora: 2 };
    const busca = semAcento(st.busca);
    const lista = todas
      .filter(l => !st.status || l.situacao === st.status)
      .filter(l => !busca || semAcento(l.nome + ' ' + l.supervisor).includes(busca))
      .sort((a, b) => String(a.supervisor).localeCompare(String(b.supervisor), 'pt-BR') || ordem[a.situacao] - ordem[b.situacao] || b.concluidas - a.concluidas || a.nome.localeCompare(b.nome, 'pt-BR'));

    document.getElementById('toaQ' + q + 'Contagem').textContent = lista.length + ' técnico' + (lista.length === 1 ? '' : 's');

    if(!lista.length){
      tabela.innerHTML = `<tbody><tr><td style="text-align:center; color:${COR.muted};">Nenhum técnico para esse filtro.</td></tr></tbody>`;
      return;
    }
    const pill = s => s === 'fez' ? '<span class="pill good">Fez serviço</span>'
      : (s === 'sem' ? '<span class="pill warn">Na base sem concluir</span>' : '<span class="pill bad">Não apareceu</span>');
    tabela.innerHTML = '<thead><tr><th>Supervisor</th><th>Técnico</th><th class="num">Lançadas</th><th class="num">Concluídas</th><th>Situação</th></tr></thead><tbody>' +
      lista.map(l => `<tr>
        <td style="white-space:nowrap;">${esc(l.supervisor) || '—'}</td>
        <td style="white-space:nowrap;">${esc(l.nome)}${!l.noOrg ? ` <span class="mono" style="font-size:10px; color:${COR.muted};" title="Nome não encontrado entre os técnicos ativos do Organograma">(fora do Organograma)</span>` : ''}</td>
        <td class="num">${l.ausente ? '' : l.total}</td>
        <td class="num">${l.ausente ? '' : l.concluidas}</td>
        <td>${pill(l.situacao)}</td>
      </tr>`).join('') + '</tbody>';
  }

  function renderKpis(r){
    const k = r.kpi;
    const pctConcl = k.lancadas > 0 ? k.concluidas / k.lancadas : 0;
    document.getElementById('toaKpis').innerHTML = `
      <div class="kpi"><div class="label">✅ Atividades concluídas</div><div class="value" style="color:${COR.accent}">${k.concluidas.toLocaleString('pt-BR')}</div>
        <div class="delta" style="color:${COR.muted};">${pct(pctConcl)} das ${k.lancadas.toLocaleString('pt-BR')} lançadas</div></div>
      <div class="kpi"><div class="label">👷 Técnicos com lançamento</div><div class="value">${k.tecnicos.toLocaleString('pt-BR')}</div></div>
      <div class="kpi"><div class="label">📈 Concluídas por técnico/dia</div><div class="value">${fmt1(k.mediaTecDia)}</div>
        <div class="delta" style="color:${COR.muted};">só dias em que o técnico lançou</div></div>
      <div class="kpi"><div class="label">📲 Aderência ao TOA</div>
        <div class="value" style="color:${k.aderencia === null ? COR.muted : corAderencia(k.aderencia)}">${k.aderencia === null ? '—' : pct(k.aderencia)}</div>
        <div class="delta" style="color:${COR.muted};">meta ${pct(META_ADERENCIA)} · ${r.diasAderencia.length} dia(s)</div></div>`;
  }

  function renderGraficos(r){
    if(typeof Chart === 'undefined') return;
    const plugins = typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [];

    const wrap = document.getElementById('toaSupChartWrap');
    if(wrap) wrap.style.height = Math.max(240, r.supervisores.length * 34 + 40) + 'px';
    if(chartSup) chartSup.destroy();
    chartSup = new Chart(document.getElementById('toaSupChart'), {
      type: 'bar',
      data: {
        labels: r.supervisores.map(s => s.label),
        datasets: [{
          label: 'Concluídas por técnico/dia',
          data: r.supervisores.map(s => Number(s.media.toFixed(2))),
          backgroundColor: 'rgba(14,124,134,.80)', borderRadius: 4,
          datalabels: { color:'#000000', anchor:'end', align:'right', font:{ size:10, family:'JetBrains Mono', weight:'700' }, formatter: v => fmt1(v) }
        }]
      },
      plugins,
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        layout: { padding: { right: 36 } },
        plugins: {
          legend: { display:false }, datalabels: { display:true },
          tooltip: { callbacks: { afterLabel: ctx => {
            const s = r.supervisores[ctx.dataIndex];
            return s.concluidas + ' concluídas · ' + s.tecnicos + ' técnico(s)';
          } } }
        },
        scales: {
          x: { beginAtZero:true, ticks:{ color:COR.text, font:{ weight:'bold' } }, grid:{ color:COR.line } },
          y: { ticks:{ color:COR.text, font:{ size:11, weight:'bold' } }, grid:{ display:false } }
        }
      }
    });

    if(chartDia) chartDia.destroy();
    chartDia = new Chart(document.getElementById('toaDiaChart'), {
      data: {
        labels: r.porDia.map(d => fmtDiaCurto(d.dia) + ' ' + DIAS_SEMANA[diaSemana(d.dia)]),
        datasets: [
          { type:'bar', label:'Concluídas', data: r.porDia.map(d => d.concluidas), yAxisID:'y', order:2,
            backgroundColor: r.porDia.map(d => fimDeSemana(d.dia) ? 'rgba(107,117,144,.35)' : 'rgba(107,117,144,.65)'), borderRadius:4,
            datalabels:{ color:'#000000', anchor:'end', align:'top', font:{ size:9, family:'JetBrains Mono', weight:'700' }, formatter: v => v > 0 ? v : '' } },
          { type:'line', label:'Média por técnico', data: r.porDia.map(d => d.media === null ? null : Number(d.media.toFixed(2))), yAxisID:'y1', order:1,
            borderColor: COR.warn, backgroundColor: COR.warn, pointBackgroundColor: COR.warn, tension:.3, pointRadius:3, spanGaps:true,
            datalabels:{ color:COR.warn, align:'top', offset:6, font:{ size:9, family:'JetBrains Mono', weight:'700' }, formatter: v => v === null ? '' : fmt1(v) } }
        ]
      },
      plugins,
      options: {
        responsive: true,
        plugins: { legend:{ labels:{ color:COR.muted, font:{ family:'Inter', size:11 } } }, datalabels:{ display:true } },
        scales: {
          x: { ticks:{ color:COR.text, font:{ size:10, weight:'bold' } }, grid:{ display:false } },
          y: { beginAtZero:true, ticks:{ color:COR.text, font:{ weight:'bold' } }, grid:{ display:false } },
          y1: { position:'right', beginAtZero:true, ticks:{ color:COR.text }, grid:{ display:false } }
        }
      }
    });
  }

  function renderAderencia(r){
    const tabela = document.getElementById('toaAderenciaTable');
    if(!r.aderencia.length){
      tabela.innerHTML = `<tbody><tr><td style="text-align:center; color:${COR.muted};">Nenhum técnico ativo encontrado no Organograma para esse filtro.</td></tr></tbody>`;
    } else if(!r.diasAderencia.length){
      tabela.innerHTML = `<tbody><tr><td style="text-align:center; color:${COR.muted};">Nenhum dia útil no período selecionado.</td></tr></tbody>`;
    } else {
      const cab = '<thead><tr><th>Supervisor</th><th class="num">Equipe</th>' +
        r.diasAderencia.map(d => `<th class="num" style="white-space:nowrap;">${fmtDiaCurto(d)}<br><span style="font-weight:400;">${DIAS_SEMANA[diaSemana(d)]}</span></th>`).join('') +
        '<th class="num">Média</th></tr></thead>';
      const corpo = r.aderencia.map(s => {
        const celulas = s.porDia.map(x => {
          const p = x.equipe > 0 ? x.lancaram / x.equipe : 0;
          const titulo = x.lancaram + ' de ' + x.equipe + ' lançaram' + (x.semLanc.length ? '\nSem lançamento: ' + x.semLanc.join(', ') : '');
          return `<td class="num" title="${esc(titulo)}" style="background:${fundoAderencia(p)}; color:${corAderencia(p)} !important; white-space:nowrap;">${pct(p)}</td>`;
        }).join('');
        const tituloNunca = s.nunca.length ? 'Sem nenhum lançamento no período: ' + s.nunca.join(', ') : 'Todos lançaram em pelo menos um dia';
        return `<tr>
          <td style="white-space:nowrap;">${esc(s.label)}</td>
          <td class="num">${s.equipe}</td>
          ${celulas}
          <td class="num" title="${esc(tituloNunca)}"><span class="pill ${s.media >= META_ADERENCIA ? 'good' : (s.media >= META_ADERENCIA - 0.2 ? 'warn' : 'bad')}">${pct(s.media)}</span>${s.nunca.length ? `<div style="font-size:10px; color:${COR.bad}; font-weight:600;">${s.nunca.length} sem lançar</div>` : ''}</td>
        </tr>`;
      }).join('');
      tabela.innerHTML = cab + '<tbody>' + corpo + '</tbody>';
    }

    const ne = document.getElementById('toaNaoEncontrados');
    const partes = [];
    if(AJUSTES.length){
      const lista = AJUSTES.map(a => a.toa + ' → ' + a.organograma).join('\n');
      partes.push(`✔ ${AJUSTES.length} nome(s) do TOA ajustados para o nome do Organograma: <span title="${esc(lista)}" style="text-decoration:underline dotted; cursor:help;">ver ajustes</span>`);
    }
    if(r.naoEncontrados.length){
      partes.push(`⚠ ${r.naoEncontrados.length} nome(s) do TOA não encontrados entre os ativos do Organograma (não entram na aderência): <span title="${esc(r.naoEncontrados.join('\n'))}" style="text-decoration:underline dotted; cursor:help;">ver nomes</span>`);
    }
    ne.innerHTML = partes.join('<br>');
  }

  function renderTabelaTecnicos(){
    const r = ultimoCalculo;
    const tabela = document.getElementById('toaTecTable');
    const contagem = document.getElementById('toaContagem');
    if(!r || !tabela) return;

    const q = semAcento(estado.busca);
    const linhas = Object.values(r.porTec)
      .filter(p => !q || semAcento(p.nome + ' ' + p.supervisor).includes(q))
      .sort((a, b) => b.media - a.media || b.concluidas - a.concluidas);

    if(contagem) contagem.textContent = linhas.length + ' técnico' + (linhas.length === 1 ? '' : 's');

    if(!linhas.length){
      tabela.innerHTML = `<tbody><tr><td style="text-align:center; color:${COR.muted};">Nenhum lançamento no período para esse filtro.</td></tr></tbody>`;
      return;
    }

    let maxDia = 1;
    linhas.forEach(p => Object.values(p.dias).forEach(d => { if(d.concluidas > maxDia) maxDia = d.concluidas; }));

    const cab = '<thead><tr><th>Técnico</th><th>Supervisor</th>' +
      r.dias.map(d => `<th class="num" style="white-space:nowrap;${fimDeSemana(d) ? ' opacity:.6;' : ''}">${fmtDiaCurto(d)}<br><span style="font-weight:400;">${DIAS_SEMANA[diaSemana(d)]}</span></th>`).join('') +
      '<th class="num">Total</th><th class="num">Dias</th><th class="num">Média/dia</th></tr></thead>';

    const corpo = linhas.map(p => {
      const celulas = r.dias.map(d => {
        const v = p.dias[d];
        if(!v || v.total === 0) return `<td class="num" style="${fimDeSemana(d) ? 'background:rgba(107,117,144,.06);' : ''}"></td>`;
        const alpha = (0.08 + 0.4 * (v.concluidas / maxDia)).toFixed(2);
        const estilo = v.concluidas === 0 ? `color:${COR.warn} !important;` : `background:rgba(14,124,134,${alpha});`;
        return `<td class="num" title="${v.concluidas} concluída(s) de ${v.total} lançada(s)" style="${estilo}">${v.concluidas}</td>`;
      }).join('');
      return `<tr>
        <td style="white-space:nowrap;">${esc(p.nome)}</td>
        <td style="white-space:nowrap;">${esc(p.supervisor) || '—'}</td>
        ${celulas}
        <td class="num">${p.concluidas}</td>
        <td class="num">${p.diasLanc}</td>
        <td class="num">${fmt1(p.media)}</td>
      </tr>`;
    }).join('');

    tabela.innerHTML = cab + '<tbody>' + corpo + '</tbody>';
  }

  function exportarCsv(){
    const r = ultimoCalculo;
    if(!r || !r.toaF.length){ alert('Nenhum lançamento do TOA para exportar com os filtros atuais.'); return; }
    function csvEsc(v){ const s = String(v ?? ''); return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
    const linhas = [['Data','Tecnico','Supervisor','Concluidas','Lancadas (sem refeicao)'].join(';')];
    r.toaF.slice().sort((a, b) => a.data.localeCompare(b.data) || String(a.recurso).localeCompare(String(b.recurso), 'pt-BR'))
      .forEach(t => linhas.push([fmtData(t.data), t.recurso, t.supervisor, t.concluidas, t.total].map(csvEsc).join(';')));
    const blob = new Blob(['\uFEFF' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'TOA_Produtividade_' + estado.inicio + '_a_' + estado.fim + '.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function redesenhar(){
    carregar(false).then(render);
  }

  montarEstrutura();
  window.PainelTOA = { redesenhar, recarregar: () => carregar(true).then(render) };
})();
