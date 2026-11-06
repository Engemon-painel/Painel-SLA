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
  // Expediente: qualquer atividade com status no TOA (Pendente, Iniciada, em rota,
  // Concluída, Não Concluída) ou Refeição = iniciou. Na base sem nenhuma atividade
  // (linha sem status) = não iniciou.

  let PRES = [];
  let PRES_INFO = '';   // diagnóstico da carga da presença (mostrado quando a tabela fica vazia)     // [{ data, recurso }] técnicos com bonequinho no TOA (rota no dia), mesmo sem atividade
  let TOA = [];      // [{ data:'aaaa-mm-dd', recurso, supervisor, total, concluidas }]
  let COLABS = [];
  let ultimaCarga = 0;
  let carregando = null;
  const estado = { inicio:'', fim:'', supervisor:'', ignorarFds:true, busca:'', aba:'d1',
    quadro: { d1: { dia:'', status:'', busca:'', log:'nao', sup:'' }, d0: { dia:'', status:'', busca:'', log:'nao', sup:'' } } };
  let chartSup = null, chartDia = null;

  // ---------- utilitários ----------
  function esc(s){ return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function semAcento(s){ return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\s+/g,' ').trim(); }
  function pad2(n){ return n < 10 ? '0' + n : String(n); }
  function isoDeData(d){ return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function dataDeIso(iso){ const p = iso.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function somarDias(iso, n){ const d = dataDeIso(iso); d.setDate(d.getDate() + n); return isoDeData(d); }
  // ALTERADO (06/11/2026): a data da presença pode vir em vários formatos (2026-11-06T10:00:00,
  // 06/11/2026, 06/11/2026 10:00, número serial do Excel). Converte tudo para aaaa-mm-dd.
  function dataIsoPres(v){
    if(v == null || v === '') return '';
    if(typeof v === 'number' || /^\d{5}(\.\d+)?$/.test(String(v).trim())){
      const d = new Date(Math.round((Number(v) - 25569) * 86400000));
      return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
    }
    const t = String(v).trim();
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if(m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
    const d = new Date(t);
    return isNaN(d) ? '' : isoDeData(d);
  }
  // ALTERADO (06/11/2026): o dados_toa_presenca.json pode vir em vários formatos, conforme o fluxo:
  //   { toaPresenca: [...] }  ·  [ ... ] (lista solta)  ·  { value: [...] } (Listar linhas do Excel)
  //   ·  { body: { value: [...] } }  ·  ou qualquer outra chave com uma lista de linhas.
  function listaPresenca(pr){
    if(!pr) return [];
    if(typeof pr === 'string'){ try { pr = JSON.parse(pr); } catch(e){ return []; } }
    if(Array.isArray(pr)) return pr;
    for(const k of ['toaPresenca', 'TOA_PRESENCA', 'presenca', 'value', 'body', 'dados', 'itens', 'linhas', 'rows']){
      if(pr[k] != null){ const l = listaPresenca(pr[k]); if(l.length) return l; }
    }
    const qualquer = Object.values(pr).find(v => Array.isArray(v) && v.length && typeof v[0] === 'object');
    return qualquer || [];
  }
  // resumo do arquivo, para o aviso da tabela quando não acha presença
  function infoPresenca(pr){
    if(pr == null) return 'dados_toa_presenca.json não encontrado (ou não é um JSON válido)';
    const l = listaPresenca(pr);
    if(pr && pr.aviso) return 'dados_toa_presenca.json (gerado em ' + String(pr.geradoEm || '?').replace('T', ' ').slice(0, 16) + ' UTC): ' + pr.aviso + (l.length ? ' · ' + l.length + ' linhas' : '');
    if(!l.length) return 'dados_toa_presenca.json sem linhas (gerado em ' + String((pr && pr.geradoEm) || '?').replace('T', ' ').slice(0, 16) + ' UTC) · chaves do arquivo: ' + (Array.isArray(pr) ? '(lista vazia)' : Object.keys(pr).join(', ') || '(vazio)');
    return 'dados_toa_presenca.json com ' + l.length + ' linhas · colunas: ' + Object.keys(l[0] || {}).join(', ');
  }
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
  // Presença TOA: o console do TOA traz o nome completo e o Organograma às vezes corta o nome
  // (30 letras) ou tem sobrenome a menos. Se um nome começa com o outro (e só existe 1 candidato),
  // usa o nome do Organograma. Acentos já são ignorados em chavesNome().
  let AJUSTES_PRES = [];
  function ajustarNomesPres(){
    AJUSTES_PRES = [];
    const org = COLABS.filter(c => c.ativo && c.nome).map(c => ({ nome: c.nome, k: chavePrincipal(c.nome) })).filter(o => o.k);
    const exatas = new Set(org.map(o => o.k));
    const vistos = {};
    PRES = PRES.map(p => {
      const kp = chavePrincipal(p.recurso);
      if(!kp || exatas.has(kp)) return p;
      if(!(kp in vistos)){
        const cand = org.filter(o => Math.min(o.k.length, kp.length) >= 15 && (kp.indexOf(o.k) === 0 || o.k.indexOf(kp) === 0));
        const unicos = [...new Set(cand.map(o => o.nome))];
        vistos[kp] = unicos.length === 1 ? unicos[0] : '';
        if(vistos[kp]) AJUSTES_PRES.push({ toa: p.recurso, organograma: vistos[kp] });
      }
      return vistos[kp] ? { data: p.data, recurso: vistos[kp] } : p;
    });
  }
  // Técnicos SEM ACESSO ao TOA (restrição) — por matrícula. Não entram como "não logou" nem na % de aderência.
  const RESTRICAO_TOA = ['6702', '8769', '8087', '6665', '7147', '6627', '8782', '6858', '6914', '8796', '6769', '6780'];
  function matNum(m){ const x = String(m == null ? '' : m).split('-').pop().replace(/\D/g, '').replace(/^0+/, ''); return x; }
  function temRestricao(c){ return !!c && RESTRICAO_TOA.indexOf(matNum(c.matricula)) !== -1; }
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
        <div class="chip" id="toaQ${q}Padrao">${q === 'd0' ? 'Hoje' : 'Ontem'}</div>
        <label for="toaQ${q}Status" class="mono" style="${ESTILO_LABEL} margin-left:8px; display:none;">Mostrar</label>
        <select id="toaQ${q}Status" class="mono" style="${ESTILO_SELECT} min-width:200px; display:none;">
          <option value="">Todos</option>
          <option value="fez">Fez serviço</option>
          <option value="sem">Na base sem concluir</option>
          <option value="ini">🔵 Iniciou</option>
          <option value="naoini">🟠 Não iniciou</option>
        </select>
        <div class="search" style="margin-left:8px;">
          <input id="toaQ${q}Busca" type="search" placeholder="Buscar técnico..." aria-label="Buscar técnico">
        </div>
        <span id="toaQ${q}Contagem" class="mono" style="font-size:11px; color:var(--muted); margin-left:auto;"></span>
      </div>
      <div class="kpi-row" id="toaQ${q}Kpis" style="margin-bottom:14px;"></div>

      <h2 style="font-size:14px; margin:6px 0 8px;">Resumo por supervisor — Organograma x TOA</h2>
      <div style="margin-bottom:18px;">
        <div style="overflow:auto; border-radius:8px; margin-bottom:14px;"><table id="toaQ${q}Sup" class="toa-resumo"></table></div>
        <div style="max-width:760px;">
          <div class="mono" style="${ESTILO_LABEL} margin-bottom:6px;">🏆 Ranking de aderência por supervisor</div>
          <div id="toaQ${q}AderWrap" style="position:relative; height:260px;"><canvas id="toaQ${q}AderChart"></canvas></div>
        </div>
      </div>

      <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin:6px 0 4px;">
        <h2 style="font-size:14px; margin:0;">📲 Presença no TOA — técnicos do Organograma</h2>
        <select id="toaQ${q}SupF" class="mono" style="${ESTILO_SELECT} margin-left:auto; min-width:220px;" aria-label="Supervisor"><option value="">Supervisor: todos</option></select>
        <select id="toaQ${q}Log" class="mono" style="${ESTILO_SELECT}">
          <option value="nao">🔴 Não logaram</option>
          <option value="sim">🟢 Logaram</option>
          <option value="ag">⏳ Entram depois (horário)</option>
          <option value="folga">📅 De folga (escala)</option>
          <option value="restr">⛔ Sem acesso (restrição)</option>
          <option value="">Todos</option>
        </select>
      </div>
      <div class="hint">Técnicos ativos do Organograma com ou sem o bonequinho no console do TOA no dia (aba TOA_PRESENCA, lida pelo robô).</div>
      <div style="overflow:auto; max-height:520px; border-radius:8px;"><table id="toaQ${q}Table"></table></div>
    </div>`;
  }

  // D-0 = hoje (data do computador); D-1 = ontem.
  function hojeIsoToa(){ return isoDeData(new Date()); }
  function diaPadrao(q){
    const hoje = hojeIsoToa();
    return q === 'd0' ? hoje : somarDias(hoje, -1);
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
          ${quadroHtml('d1', 'D-1 — técnicos na base do dia anterior', 'Fechamento de ontem: quem apareceu na base do TOA, quem fez serviço e quem da equipe não apareceu.')}

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

          <div class="panel" style="margin-bottom:16px;">
            <h2>Presença no TOA — técnicos logados por supervisor</h2>
            <div class="hint">Técnicos com o bonequinho no console do TOA (logados no dia), lidos pelo robô na aba TOA_PRESENCA. Não depende da base de atividades do TOA. Passe o mouse para ver quem logou e quem não logou.</div>
            <div style="overflow-x:auto; border-radius:8px;"><table id="toaPresencaTable"></table></div>
            <div id="toaPresencaFora" class="mono" style="font-size:11px; color:var(--muted); margin-top:10px;"></div>
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
          ${quadroHtml('d0', 'D-0 — técnicos na base de hoje', 'Dia de hoje: quem já aparece na base do TOA, quem já fez serviço (concluiu ao menos 1 atividade) e quem da equipe ainda não apareceu.')}
        </div>
      </div>`;
    main.appendChild(pg);
    if(!document.getElementById('toa-resumo-css')){
      const css = document.createElement('style');
      css.id = 'toa-resumo-css';
      css.textContent = `
        table.toa-resumo{ width:auto; min-width:100%; }
        table.toa-resumo th{ padding:6px 8px !important; font-size:10px !important; color:var(--text) !important; white-space:normal; max-width:96px; line-height:1.25; vertical-align:bottom; }
        table.toa-resumo td{ padding:5px 8px !important; font-size:13px !important; color:var(--text) !important; font-weight:600; }
        table.toa-resumo td.num{ font-weight:700; }
        .pill.forte{ font-size:12px; padding:2px 10px; color:#fff !important; }
        .pill.forte.good{ background:#059669; } .pill.forte.warn{ background:#d97706; } .pill.forte.bad{ background:#dc2626; }`;
      document.head.appendChild(css);
    }

    document.getElementById('toaInicio').addEventListener('change', e => { estado.inicio = e.target.value; render(); });
    document.getElementById('toaFim').addEventListener('change', e => { estado.fim = e.target.value; render(); });
    document.getElementById('toaSupervisor').addEventListener('change', e => { estado.supervisor = e.target.value; render(); });
    document.getElementById('toaIgnorarFds').addEventListener('change', e => { estado.ignorarFds = e.target.checked; render(); });
    document.getElementById('toaExportar').addEventListener('click', exportarCsv);
    ['d1', 'd0'].forEach(q => {
      const st = estado.quadro[q];
      document.getElementById('toaQ' + q + 'Dia').addEventListener('change', e => { st.dia = e.target.value; renderQuadro(q); });
      document.getElementById('toaQ' + q + 'Status').addEventListener('change', e => { st.status = e.target.value; renderQuadro(q); });
      document.getElementById('toaQ' + q + 'Log').addEventListener('change', e => { st.log = e.target.value; renderQuadro(q); });
      document.getElementById('toaQ' + q + 'SupF').addEventListener('change', e => { st.sup = e.target.value; renderQuadro(q); });
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

  // ---------- Gestão de Escala (dados_escala.json) ----------
  // Para os supervisores que têm escala (ex.: Marcelo, Nayane, Ycaro, Esaie), a escala do
  // dia diz quem devia trabalhar: escalado (X) sem atividade no TOA = 🟠 Não iniciou.
  let ESCALA = null;   // { datas: [...], tecnicos: [...], porChave: {} }
  function juntarEscalas(a, b){
    const lista = [a, b].filter(e => e && Array.isArray(e.datas) && Array.isArray(e.tecnicos));
    if(!lista.length) return null;
    if(lista.length === 1) return lista[0];
    const datas = [...new Set(lista[0].datas.concat(lista[1].datas))].sort();
    const idx = {}; datas.forEach((d, i) => idx[d] = i);
    const tecnicos = [];
    lista.forEach(e => e.tecnicos.forEach(t => {
      const cod = new Array(datas.length).fill('');
      e.datas.forEach((d, i) => { cod[idx[d]] = (t.codigos || [])[i] || ''; });
      tecnicos.push(Object.assign({}, t, { codigos: cod }));
    }));
    return { datas, tecnicos };
  }
  function prepararEscala(e){
    ESCALA = null;
    if(!e || !Array.isArray(e.datas) || !Array.isArray(e.tecnicos)) return;
    const porChave = {};
    e.tecnicos.forEach(t => chavesNome(t.nome).forEach(k => { if(!porChave[k]) porChave[k] = t; }));
    ESCALA = { datas: e.datas, tecnicos: e.tecnicos, porChave };
  }
  function escalaDoTecnico(nome){
    if(!ESCALA) return null;
    const direto = chavesNome(nome).map(k => ESCALA.porChave[k]).find(Boolean);
    if(direto) return direto;
    const tt = tokensNome(nome);
    let melhor = null, p1 = 0, p2 = 0;
    ESCALA.tecnicos.forEach(t => {
      const p = pontuar(tt, tokensNome(t.nome));
      if(p > p1){ p2 = p1; p1 = p; melhor = t; } else if(p > p2){ p2 = p; }
    });
    return melhor && p1 > p2 ? melhor : null;
  }
  function codigoEscala(tec, dia){
    if(!tec || !ESCALA) return '';
    const i = ESCALA.datas.indexOf(dia);
    return i === -1 ? '' : String((tec.codigos || [])[i] || '').trim().toUpperCase();
  }
  function rotuloEscala(c){
    const n = semAcento(c);
    if(!n) return { txt: 'Sem escala no dia', cls: '', cor: 'rgba(107,117,144,.14)', tcor: '#6b7590' };
    if(n === 'x') return { txt: 'Escalado', cls: '', cor: 'rgba(37,99,235,.12)', tcor: '#1d4ed8' };
    if(n === 'fo') return { txt: 'Folga', cls: '', cor: 'rgba(107,117,144,.14)', tcor: '#6b7590' };
    if(n === 'fbh') return { txt: 'Folga BH', cls: '', cor: 'rgba(107,117,144,.14)', tcor: '#6b7590' };
    if(n === 'falta' || n === 'f') return { txt: 'Falta', cls: 'bad' };
    if(n === 'atm' || n.indexOf('atestado') === 0) return { txt: 'Atestado', cls: 'warn' };
    if(n === 'des') return { txt: 'Desligado', cls: 'bad' };
    if(n.indexOf('treinamento') === 0) return { txt: 'Treinamento', cls: 'warn' };
    return { txt: c, cls: 'warn' };
  }
  function pillEscala(c){
    const r = rotuloEscala(c);
    return r.cls ? `<span class="pill ${r.cls}">${esc(r.txt)}</span>`
      : `<span class="pill" style="background:${r.cor}; color:${r.tcor};">${esc(r.txt)}</span>`;
  }


  // ---------- escala + horário de entrada (Gestão de Escala; se o técnico não estiver lá, usa Escala/Horário do Organograma) ----------
  const HORA_LEITURA_ROBO = '18:00';   // dias anteriores: horário em que o robô lê a presença no TOA
  function horaEntrada(txt){ const m = /(\d{1,2})\s*[:hH]\s*(\d{2})/.exec(String(txt || '')); return m ? pad2(+m[1]) + ':' + m[2] : ''; }
  function escaladoPeloOrg(escala, dia){
    const n = semAcento(escala).replace(/[\s.]/g, '');
    const dw = diaSemana(dia);
    if(/5x2|segasex|segundaasexta/.test(n)) return dw >= 1 && dw <= 5;
    if(/terasab|tercaasabado/.test(n)) return dw >= 2 && dw <= 6;
    if(/6x1|segasab|segundaasabado/.test(n)) return dw >= 1 && dw <= 6;
    return null;   // 12x36, espanhola etc. sem Gestão de Escala: não dá para saber o dia
  }
  // Escala padrão por supervisor, para equipes que não estão na Gestão de Escala nem têm Escala/Horário no Organograma
  const ESCALA_PADRAO_SUP = {
    'cleiton de souza cardoso': { escala: '5X2', horario: '08:00–18:00' }
  };
  function infoEscala(c, dia){
    const pad = ESCALA_PADRAO_SUP[semAcento(c.supervisor)];
    if(pad) c = Object.assign({}, c, { escala: c.escala || pad.escala, horario: c.horario || pad.horario });
    const tec = escalaDoTecnico(c.nome);
    const cod = tec ? codigoEscala(tec, dia) : '';
    let escalado = false, fonte = '';
    if(tec && cod){ escalado = cod === 'X'; fonte = 'escala'; }
    else {
      const e = escaladoPeloOrg(c.escala, dia);
      if(e !== null){ escalado = e; fonte = 'org'; }
      else if(tec){ fonte = 'escala'; }
    }
    const horTxt = String((tec && tec.horario) || c.horario || '').trim();
    return { tec, cod, escalado, fonte, hora: horaEntrada(horTxt), horario: horTxt.replace(/\s*(AS|ÀS|as|às|-|–)\s*/, ' – '), nomeEscala: String((tec && tec.escala) || c.escala || '').trim() };
  }
  function horaRef(dia){
    const hoje = hojeIsoToa();
    if(dia === hoje){ const d = new Date(); return pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }
    return dia < hoje ? HORA_LEITURA_ROBO : '00:00';
  }
  function celEscala(l){
    const base = (l.fonte === 'escala' && !l.cod) ? `<span class="pill" style="background:rgba(107,117,144,.14); color:#6b7590;">Folga</span>` : l.fonte === 'escala' ? pillEscala(l.cod)
      : l.fonte === 'org' ? (l.escalado ? `<span class="pill" style="background:rgba(37,99,235,.12); color:#1d4ed8;">Escalado</span>` : `<span class="pill" style="background:rgba(107,117,144,.14); color:#6b7590;">Folga</span>`)
      : `<span style="color:${COR.muted};">—</span>`;
    const det = l.nomeEscala;
    return base + (det ? ` <span class="mono" style="font-size:11px; color:${COR.text};">${esc(det)}</span>` : '');
  }

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
    const pegarEscalaBase = fetch('dados_escala.json?_=' + Date.now())
      .then(r => r.ok ? r.json() : null).catch(() => null);
    // escala das equipes Residente (Rodolfo) e Volante (Thiago) — planilha Controle de Faltas e Atestados
    const pegarEscalaExtra = fetch('dados_escala_extra.json?_=' + Date.now())
      .then(r => r.ok ? r.json() : null).catch(() => null);
    const pegarEscala = Promise.all([pegarEscalaBase, pegarEscalaExtra]).then(([a, b]) => juntarEscalas(a, b));
    // Presença TOA também pode vir de um arquivo separado (fluxo pequeno, sem rodar o fluxo grande do painel)
    const pegarPresenca = fetch('dados_toa_presenca.json?_=' + Date.now())
      .then(r => r.ok ? r.json() : null).catch(() => null);
    carregando = fetch('dados.json?_=' + Date.now())
      .then(r => { if(!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(d => pegarEscala.then(e => { prepararEscala(e); return d; }))
      .then(d => pegarPresenca.then(pr => { d.__presExtra = listaPresenca(pr); d.__presInfo = infoPresenca(pr); return d; }))
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
          const tecnicoValido = COLABS.some(c => chavesNome(c.nome).some(k => chavesNome(t.recurso).indexOf(k) !== -1));
          if(!tecnicoValido && s && supsExcluidos.has(s) && !supsValidos.has(s)) return false; // supervisor só da Cirion
          return true;
        });
        const presJunta = {};
        let presLidas = 0, presNaoLogou = 0, presSemDado = 0;
        listaPresenca(d.toaPresenca).concat(d.__presExtra || []).forEach(p => {
          if(!p) return;
          presLidas++;
          const data = dataIsoPres(p.data != null ? p.data : (p.Data != null ? p.Data : p.dia));
          const recurso = String(p.recurso || p.Recurso || p.tecnico || p.Tecnico || p['Técnico'] || p.nome || p.Nome || '').trim();
          // Ícone do TOA: "res-ic res_type_10 ..." = bonequinho (logado). "ic ptnon_working_cal" (fora do
          // calendário de trabalho) e "ic ptnon_configured_cal" (sem calendário) = NÃO logado.
          // Sem coluna de ícone (formato antigo) = conta como logado, como antes.
          const icone = String(p.icone != null ? p.icone : (p.Icone != null ? p.Icone : (p['Ícone'] != null ? p['Ícone'] : ''))).toLowerCase();
          const logou = p.logou === false ? false : p.logou === true ? true : (!icone || (/res_type_/.test(icone) && !/ptnon_/.test(icone)));
          if(!data || !recurso){ presSemDado++; return; }
          if(!logou){ presNaoLogou++; return; }
          presJunta[data + '|' + semAcento(recurso)] = { data, recurso };
        });
        // diagnóstico: vem do dados.json (debugToaPresenca, gerado pelo script do painel). O dados_toa_presenca.json
        // é opcional (só existe se houver um fluxo separado); quando não existe, não aparece no aviso.
        const dbg = d.debugToaPresenca;
        PRES_INFO = dbg
          ? `Script do painel (aba ${dbg.aba || 'TOA_PRESENCA'}): ${dbg.linhas || 0} linhas, ${dbg.logaram || 0} logaram, ${dbg.naoLogaram || 0} sem bonequinho` + (dbg.aviso ? ' · ' + dbg.aviso : '')
          : (!Array.isArray(d.toaPresenca) ? 'O dados.json não tem toaPresenca: atualize o Office Script do painel e rode o fluxo.'
            : `dados.json → toaPresenca com ${d.toaPresenca.length} linhas (atualize o Office Script do painel para ver o motivo).`);
        if((d.__presExtra || []).length) PRES_INFO += ' · ' + d.__presInfo;
        PRES = Object.values(presJunta).filter(p => p && p.data && p.recurso && !chavesNome(p.recurso).some(k => nomesExcluidos.has(k)));
        ajustarNomesPres();
        ajustarNomesToa();
        ultimaCarga = Date.now();
        if(!estado.fim && (TOA.length || PRES.length)){
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
    const diasUteisSel = estado.ignorarFds ? dias.filter(d => !fimDeSemana(d)) : dias;

    const noPeriodo = TOA.filter(t => t.data >= estado.inicio && t.data <= estado.fim);
    // Dias sem base do TOA ainda (ex.: a base vai só até ontem/anteontem) não entram como 0% na aderência.
    const diasComBase = new Set(noPeriodo.map(t => t.data));
    const diasAderencia = diasUteisSel.filter(d => diasComBase.has(d));
    const diasSemBase = diasUteisSel.filter(d => !diasComBase.has(d));
    // Presença: vai do início do período até o último dia que o robô gravou (inclui hoje, mesmo se o período termina ontem).
    const fimPres = PRES.reduce((m, p) => p.data > m ? p.data : m, estado.fim);
    const diasPres = [...new Set(PRES.filter(p => p.data >= estado.inicio && p.data <= fimPres).map(p => p.data))]
      .filter(d => !estado.ignorarFds || !fimDeSemana(d)).sort();
    const toaF = estado.supervisor ? noPeriodo.filter(t => mesmoSupervisor(t.supervisor, estado.supervisor)) : noPeriodo;

    const tecnicos = COLABS.filter(c => c.ativo && semAcento(c.areaTecnica) === 'tecnico' && c.nome);
    const tecF = estado.supervisor ? tecnicos.filter(c => mesmoSupervisor(c.supervisor, estado.supervisor)) : tecnicos;

    // Quem lançou em cada dia (todas as equipes; o cruzamento é pelo nome).
    const lancouPorDia = {};
    noPeriodo.forEach(t => {
      if(!((Number(t.total) || 0) + (Number(t.refeicao) || 0) > 0)) return;
      if(!lancouPorDia[t.data]) lancouPorDia[t.data] = new Set();
      chavesNome(t.recurso).forEach(k => lancouPorDia[t.data].add(k));
    });
    // Presença TOA: técnico logado (bonequinho no console do TOA, aba TOA_PRESENCA).
    const rotaPorDia = {};
    PRES.forEach(p => {
      if(p.data < estado.inicio || p.data > fimPres) return;
      if(!rotaPorDia[p.data]) rotaPorDia[p.data] = new Set();
      chavesNome(p.recurso).forEach(k => rotaPorDia[p.data].add(k));
    });
    function logado(c, dia){
      const r = rotaPorDia[dia];
      return !!r && chavesNome(c.nome).some(k => r.has(k));
    }
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

    // Presença TOA por supervisor (Organograma), mesmos dias da aderência.
    const presenca = Object.values(equipes).map(eq => {
      let soma = 0;
      const porDia = diasPres.map(dia => {
        const logados = eq.membros.filter(c => logado(c, dia));
        soma += logados.length;
        return { dia, n: logados.length, equipe: eq.membros.length, nomes: logados.map(c => c.nome), naoLog: eq.membros.filter(c => !logado(c, dia)).map(c => c.nome) };
      });
      return { label: eq.label, equipe: eq.membros.length, porDia, media: diasPres.length ? soma / diasPres.length : 0 };
    }).sort((a, b) => b.media - a.media);
    const chavesTec = new Set();
    tecF.forEach(c => chavesNome(c.nome).forEach(k => chavesTec.add(k)));
    const presFora = [...new Set(PRES.filter(p => p.data >= estado.inicio && p.data <= fimPres && !chavesNome(p.recurso).some(k => chavesTec.has(k))).map(p => p.recurso))].sort();
    const temPresenca = diasPres.length > 0;

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
      dias, diasAderencia, diasSemBase, diasPres, toaF, aderencia, presenca, presFora, temPresenca, naoEncontrados, porTec, supervisores, porDia,
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

    if(!TOA.length && !PRES.length){
      info.textContent = 'Sem dados do TOA no dados.json. Rode o Office Script atualizado (item 48) com a aba "TOA" preenchida.';
      ['toaKpis','toaAderenciaTable','toaTecTable','toaNaoEncontrados','toaQd1Kpis','toaQd1Table','toaQd1Sup','toaQd1Fora','toaQd0Kpis','toaQd0Table','toaQd0Sup','toaQd0Fora'].forEach(id => { const el = document.getElementById(id); if(el) el.innerHTML = ''; });
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
    renderPresenca(r);
    renderTabelaTecnicos();
  }

  // ---------- quadro do dia (D-1 e D-0): quem está na base e quem fez serviço ----------
  function renderQuadro(q){
    const st = estado.quadro[q];
    const tabela = document.getElementById('toaQ' + q + 'Table');
    if(!tabela || (!TOA.length && !PRES.length)) return;
    const max = [maiorData(), hojeIsoToa()].sort().pop();
    if(!st.dia || st.dia > max) st.dia = diaPadrao(q);
    const inp = document.getElementById('toaQ' + q + 'Dia');
    inp.value = st.dia; inp.min = menorData(); inp.max = max;
    document.getElementById('toaQ' + q + 'Status').value = st.status;

    const dia = st.dia;
    const doDia = TOA.filter(t => t.data === dia);
    const tecnicos = COLABS.filter(c => c.ativo && c.nome && semAcento(c.areaTecnica) === 'tecnico');

    // Técnicos que aparecem na base do dia (qualquer atividade, inclusive Refeição)
    const linhas = {};
    doDia.forEach(t => {
      const k = chavePrincipal(t.recurso);
      if(!k) return;
      if(!linhas[k]) linhas[k] = { nome: t.recurso, supervisor: t.supervisor || '', total: 0, concluidas: 0, refeicao: 0, status: {}, noOrg: false, ausente: false };
      const l = linhas[k];
      l.total += Number(t.total) || 0;
      l.concluidas += Number(t.concluidas) || 0;
      l.refeicao += Number(t.refeicao) || 0;
      Object.keys(t.status || {}).forEach(x => { l.status[x] = (l.status[x] || 0) + (Number(t.status[x]) || 0); });
    });
    // Equipe do Organograma: marca quem apareceu (supervisor do Organograma) e inclui quem não apareceu
    tecnicos.forEach(c => {
      const k = chavesNome(c.nome).find(x => linhas[x]);
      if(k){
        linhas[k].noOrg = true;
        linhas[k].restrito = temRestricao(c);
        linhas[k].org = c;
        if(c.supervisor) linhas[k].supervisor = c.supervisor;
      } else {
        const kn = chavePrincipal(c.nome);
        if(kn && !linhas[kn]) linhas[kn] = { nome: c.nome, supervisor: c.supervisor || '', total: 0, concluidas: 0, refeicao: 0, status: {}, noOrg: true, ausente: true, restrito: temRestricao(c), org: c };
      }
    });
    if(estado.supervisor){
      Object.keys(linhas).forEach(k => { if(!mesmoSupervisor(linhas[k].supervisor, estado.supervisor)) delete linhas[k]; });
    }
    // Escala do dia (só técnicos que existem na Gestão de Escala)
    const refHora = horaRef(dia);
    Object.values(linhas).forEach(l => {
      const ie = infoEscala(l.org || { nome: l.nome }, dia);
      l.temEscala = !!ie.tec;
      l.codEscala = ie.cod;
      l.escalado = ie.escalado;
      l.horaEnt = ie.hora;
      l.folga = !!ie.fonte && !ie.escalado;   // escala conhecida e não escalado no dia
    });
    // Último lançamento de quem não apareceu (ajuda a conferir data/nome)
    const ultimoDia = {};
    TOA.forEach(t => { const k = chavePrincipal(t.recurso); if(k && (!ultimoDia[k] || t.data > ultimoDia[k])) ultimoDia[k] = t.data; });

    // Presença TOA do dia (bonequinho, aba TOA_PRESENCA)
    const presDia = new Set();
    PRES.filter(p => p.data === dia).forEach(p => chavesNome(p.recurso).forEach(k => presDia.add(k)));
    const temPresDia = presDia.size > 0;
    const todas = Object.values(linhas).map(l => {
      l.logou = chavesNome(l.nome).some(k => presDia.has(k));
      l.aguardando = l.escalado && !l.logou && !!l.horaEnt && l.horaEnt > refHora;   // ainda não chegou o horário de entrada
      l.situacao = l.ausente ? 'fora' : (l.concluidas > 0 ? 'fez' : 'sem');
      l.iniciou = !l.ausente && (l.concluidas > 0 || l.refeicao > 0 ||
        Object.keys(l.status).some(x => (Number(l.status[x]) || 0) > 0 && semAcento(x) !== 'sem status'));
      return l;
    });
    const naBase = todas.filter(l => !l.ausente);
    // Tabela da base: quem está no TOA + escalados (X) que não apareceram no TOA (🟠)
    const listaBase = todas.filter(l => !l.ausente || l.escalado);
    const fora = todas.filter(l => l.ausente && !l.escalado);
    const iniciaram = naBase.filter(l => l.iniciou).length;
    const naoIniciaram = listaBase.filter(l => !l.iniciou).length;
    const escaladosTodos = todas.filter(l => l.escalado && !l.restrito);
    const aguardandoDia = escaladosTodos.filter(l => l.aguardando).length;
    const escalados = escaladosTodos.filter(l => !l.aguardando);
    const escaladosIni = escalados.filter(l => l.logou).length;
    const orgTodos = todas.filter(l => l.noOrg);
    const restritos = orgTodos.filter(l => l.restrito && !l.logou).length;
    const logaram = orgTodos.filter(l => l.logou).length;
    const naoLogaram = orgTodos.filter(l => !l.logou && !l.restrito && !l.folga && !l.aguardando).length;
    const baseAder = orgTodos.length - restritos;
    const fez = naBase.filter(l => l.situacao === 'fez');
    const concluidas = naBase.reduce((s, l) => s + l.concluidas, 0);
    const equipe = todas.filter(l => l.noOrg).length;

    const semDados = !TOA.some(t => t.data === dia);
    const aviso = semDados
      ? `<div class="mono" style="flex-basis:100%; font-size:12px; color:${COR.bad}; background:rgba(220,38,38,.08); border:1px solid rgba(220,38,38,.25); border-radius:10px; padding:10px 12px;">⚠ Nenhum lançamento do TOA com data ${fmtData(dia)} no dados.json. A base mais recente que chegou ao painel é de ${fmtData(maiorData())}. Confira se a aba TOA tem as linhas desse dia e rode o fluxo de novo.</div>`
      : '';
    const kpi = (rotulo, valor, cor, extra) => `<div class="kpi"><div class="label">${rotulo}</div><div class="value" style="color:${cor}">${valor}</div>${extra ? `<div class="delta" style="color:${COR.muted};">${extra}</div>` : ''}</div>`;
    document.getElementById('toaQ' + q + 'Kpis').innerHTML = aviso +
      kpi('📅 Dia', fmtData(dia).slice(0, 5) + ' ' + DIAS_SEMANA[diaSemana(dia)], COR.text) +
      kpi('👥 Equipe (Organograma)', equipe, COR.text, 'técnicos ativos, sem Cirion') +
      kpi('🟢 Logaram (Presença TOA)', temPresDia ? logaram : '—', COR.good, temPresDia ? 'aderência ' + pct(baseAder ? logaram / baseAder : 0) + ' da equipe com acesso' : 'sem presença no dia') +
      kpi('🔴 Não logaram', temPresDia ? naoLogaram : '—', COR.bad, temPresDia ? 'sem o bonequinho no TOA (fora folga e quem entra depois)' : 'sem presença no dia') +
      kpi('⛔ Restrição (sem acesso ao TOA)', restritos, COR.muted, 'fora da conta de aderência') +
      (ESCALA ? kpi('📋 Escalados (escala)', escalados.length, '#1d4ed8', (escalados.length && temPresDia ? escaladosIni + ' logaram · aderência ' + pct(escaladosIni / escalados.length) : (escalados.length ? 'sem presença no dia' : 'nenhum escalado no horário')) + (aguardandoDia ? ' · ⏳ ' + aguardandoDia + ' entram depois' : '')) : '') +
      kpi('✅ Fizeram serviço', fez.length, COR.good, concluidas + ' atividades concluídas');

    // ----- resumo por supervisor -----
    const porSup = {};
    todas.forEach(l => {
      const lab = String(l.supervisor || '').trim() || 'Sem supervisor';
      const k = semAcento(lab);
      if(!porSup[k]) porSup[k] = { label: lab, equipe: 0, base: 0, ini: 0, naoIni: 0, fora: 0, fez: 0, esc: 0, escIni: 0, log: 0, naoLog: 0, restr: 0, ag: 0, folga: 0 };
      const g = porSup[k];
      if(l.noOrg){ g.equipe++; if(l.logou) g.log++; else if(l.restrito) g.restr++; else if(l.folga) g.folga++; else if(!l.aguardando) g.naoLog++; }
      if(l.escalado && !l.restrito){ if(l.aguardando) g.ag++; else { g.esc++; if(l.logou) g.escIni++; } }
      if(l.ausente && !l.escalado){ g.fora++; return; }
      if(!l.ausente) g.base++;
      if(l.iniciou) g.ini++; else g.naoIni++;
      if(l.situacao === 'fez') g.fez++;
    });
    const sups = Object.values(porSup).sort((a, b) => a.label.localeCompare(b.label, 'pt-BR'));
    const tot = sups.reduce((t, g) => { ['equipe','base','ini','naoIni','fora','fez','esc','escIni','log','naoLog','restr','ag','folga'].forEach(c => t[c] += g[c]); return t; }, { equipe:0, base:0, ini:0, naoIni:0, fora:0, fez:0, esc:0, escIni:0, log:0, naoLog:0, restr:0, ag:0, folga:0 });
    const sp = v => temPresDia ? v : '—';
    const pctCel = (n, d) => d > 0 ? `<span class="pill forte ${n / d >= META_ADERENCIA ? 'good' : (n / d >= META_ADERENCIA - 0.2 ? 'warn' : 'bad')}">${pct(n / d)}</span>` : '—';
    const linhaSup = (g, negrito) => `<tr${negrito ? ` style="font-weight:700; border-top:2px solid ${COR.line};"` : ''}>
      <td style="white-space:nowrap;">${esc(g.label)}</td>
      <td class="num">${g.equipe}</td>
      <td class="num" style="color:${COR.good} !important;">${sp(g.log)}</td>
      <td class="num" style="color:${COR.bad} !important;">${sp(g.naoLog)}</td>
      <td class="num">${g.folga || '—'}</td>
      <td class="num">${g.restr || '—'}</td>
      <td class="num">${g.fez}</td>
      <td class="num">${temPresDia ? pctCel(g.log, g.equipe - g.restr) : '—'}</td>
      ${ESCALA ? `<td class="num" title="${g.ag ? g.ag + ' escalado(s) entram depois do horário de leitura' : ''}">${g.esc || '—'}${g.ag ? ` <span style="font-size:10px; color:${COR.muted};">+${g.ag}⏳</span>` : ''}</td><td class="num">${g.esc && temPresDia ? pctCel(g.escIni, g.esc) : '—'}</td>` : ''}
    </tr>`;
    document.getElementById('toaQ' + q + 'Sup').innerHTML =
      '<thead><tr><th>Supervisor</th><th class="num">Equipe</th><th class="num">🟢 Logaram</th><th class="num">🔴 Não logaram</th><th class="num" title="não logaram, mas estavam de folga pela escala">📅 Folga</th><th class="num">⛔ Restrição</th><th class="num">✅ Fizeram serviço</th><th class="num" title="logados ÷ equipe com acesso ao TOA">% logados da equipe</th>' + (ESCALA ? '<th class="num" title="escalados no dia, sem restrição, cujo horário de entrada já passou (⏳ = entram depois)">📋 Escalados</th><th class="num">Aderência (escalados logados)</th>' : '') + '</tr></thead><tbody>' +
      sups.map(g => linhaSup(g, false)).join('') + linhaSup(Object.assign({ label: 'Total' }, tot), true) + '</tbody>';

    // ----- gráfico: ranking de aderência (escalados sem restrição que logaram) -----
    desenharRankingAder(q, sups, temPresDia);

    // ----- tabela: na base do TOA -----
    const ordem = { fez: 0, sem: 1 };
    const busca = semAcento(st.busca);
    const bate = l => !busca || semAcento(l.nome + ' ' + l.supervisor).includes(busca);
    const lista = listaBase
      .filter(l => !st.status || (st.status === 'ini' ? l.iniciou : st.status === 'naoini' ? !l.iniciou : l.situacao === st.status))
      .filter(bate)
      .sort((a, b) => String(a.supervisor).localeCompare(String(b.supervisor), 'pt-BR') || (b.iniciou - a.iniciou) || ordem[a.situacao] - ordem[b.situacao] || b.concluidas - a.concluidas || a.nome.localeCompare(b.nome, 'pt-BR'));
    const listaFora = fora.filter(bate)
      .sort((a, b) => String(a.supervisor).localeCompare(String(b.supervisor), 'pt-BR') || a.nome.localeCompare(b.nome, 'pt-BR'));


    function corStatus(nome){
      const n = semAcento(nome);
      if(n.indexOf('conclu') === 0) return 'good';
      if(n.indexOf('cancel') === 0 || n.indexOf('nao conclu') === 0 || n.indexOf('suspens') === 0) return 'bad';
      return 'warn';
    }
    function situacaoToa(l){
      if(l.ausente) return `<span class="pill warn">Escalado — sem atividade no TOA</span>` +
        (ultimoDia[chavePrincipal(l.nome)] ? ` <span class="mono" style="font-size:10px; color:${COR.muted};">último lançamento ${fmtDiaCurto(ultimoDia[chavePrincipal(l.nome)])}</span>` : '');
      const itens = Object.keys(l.status).filter(x => semAcento(x) !== 'sem status').sort((x, y) => l.status[y] - l.status[x])
        .map(x => `<span class="pill ${corStatus(x)}" style="margin:1px 4px 1px 0;">${esc(x)} ${l.status[x]}</span>`);
      if(l.refeicao) itens.push(`<span class="pill" style="margin:1px 4px 1px 0; background:rgba(107,117,144,.14); color:${COR.muted};">Refeição ${l.refeicao}</span>`);
      return itens.length ? itens.join('') : `<span class="mono" style="font-size:11px; color:${COR.muted};">sem atividade</span>`;
    }
    // ----- tabela: técnicos do Organograma que NÃO logaram no TOA (Presença TOA) -----
    const ultLogin = {};
    PRES.forEach(p => { const k = chavePrincipal(p.recurso); if(k && p.data <= dia && (!ultLogin[k] || p.data > ultLogin[k])) ultLogin[k] = p.data; });
    const todosPres = tecnicos
      .filter(c => !estado.supervisor || mesmoSupervisor(c.supervisor, estado.supervisor))
      .map(c => { const ie = infoEscala(c, dia); const logou = chavesNome(c.nome).some(k => presDia.has(k)); return { nome: c.nome, supervisor: c.supervisor || '', logou, restrito: temRestricao(c), fonte: ie.fonte, cod: ie.cod, escalado: ie.escalado, hora: ie.hora, horario: ie.horario, nomeEscala: ie.nomeEscala, aguardando: ie.escalado && !logou && !!ie.hora && ie.hora > refHora, folga: !!ie.fonte && !ie.escalado && !logou }; })
      .filter(bate);
    const selSup = document.getElementById('toaQ' + q + 'SupF');
    const supsTab = [...new Set(todosPres.map(l => l.supervisor).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    if(st.sup && supsTab.indexOf(st.sup) === -1) st.sup = '';
    selSup.innerHTML = '<option value="">Supervisor: todos</option>' + supsTab.map(x => `<option value="${esc(x)}"${x === st.sup ? ' selected' : ''}>${esc(x)}</option>`).join('');
    const todosSup = todosPres.filter(l => !st.sup || mesmoSupervisor(l.supervisor, st.sup));
    const nSim = todosSup.filter(l => l.logou).length, nRes = todosSup.filter(l => !l.logou && l.restrito).length, nAg = todosSup.filter(l => l.aguardando && !l.restrito).length, nFol = todosSup.filter(l => l.folga && !l.restrito).length, nNao = todosSup.length - nSim - nRes - nAg - nFol;
    document.getElementById('toaQ' + q + 'Log').value = st.log;
    const naoLog = todosSup.filter(l => !st.log || (st.log === 'sim' ? l.logou : st.log === 'restr' ? (!l.logou && l.restrito) : st.log === 'ag' ? (l.aguardando && !l.restrito) : st.log === 'folga' ? (l.folga && !l.restrito) : (!l.logou && !l.restrito && !l.aguardando && !l.folga)))
      .sort((a, b) => String(a.supervisor).localeCompare(String(b.supervisor), 'pt-BR') || (a.logou - b.logou) || a.nome.localeCompare(b.nome, 'pt-BR'));
    document.getElementById('toaQ' + q + 'Contagem').textContent = temPresDia ? '🟢 ' + nSim + ' logaram · 🔴 ' + nNao + ' não logaram' + (nAg ? ' · ⏳ ' + nAg + ' entram depois' : '') + (nFol ? ' · 📅 ' + nFol + ' de folga' : '') + (nRes ? ' · ⛔ ' + nRes + ' sem acesso' : '') : 'sem presença no dia';
    tabela.innerHTML = !temPresDia
      ? `<tbody><tr><td style="text-align:center; color:${COR.muted};">Sem dados de presença para ${fmtData(dia)}. O robô grava a aba TOA_PRESENCA no fim do dia.${PRES.length ? '<br>Dias com presença na base: ' + [...new Set(PRES.map(p => p.data))].sort().slice(-7).map(fmtData).join(', ') : '<br>Nenhum técnico logado na base de presença.'}<br><span style="font-size:11px">${esc(PRES_INFO)}</span></td></tr></tbody>`
      : !naoLog.length
        ? `<tbody><tr><td style="text-align:center; color:${COR.muted};">Nenhum técnico nesse filtro.</td></tr></tbody>`
        : '<thead><tr><th>Supervisor</th><th>Técnico</th><th>Presença</th><th>Escala</th><th>Horário</th><th>Restrição</th></tr></thead><tbody>' +
          naoLog.map(l => {
            return `<tr>
              <td style="white-space:nowrap;">${esc(l.supervisor) || '—'}</td>
              <td style="white-space:nowrap;">${esc(l.nome)}</td>
              <td style="white-space:nowrap;">${l.logou ? '<span class="pill good">🟢 Logou</span>' : (l.aguardando && !l.restrito) ? `<span class="pill" style="background:rgba(37,99,235,.12); color:#1d4ed8;">⏳ Entra às ${esc(l.hora)}</span>` : (l.folga && !l.restrito) ? '<span class="pill" style="background:rgba(107,117,144,.16); color:#374151;">📅 Folga</span>' : l.restrito ? '<span class="pill" style="background:rgba(107,117,144,.16);">⛔ Sem acesso (restrição)</span>' : '<span class="pill bad">🔴 Não logou</span>'}</td>
              <td style="white-space:nowrap;">${celEscala(l)}</td>
              <td class="mono" style="white-space:nowrap; font-size:12px; color:${COR.text};">${l.horario ? esc(l.horario) : `<span style="color:${COR.muted};">—</span>`}</td>
              <td>${l.restrito ? '<span class="pill bad">Sim</span>' : '<span class="pill good">Não</span>'}</td>
            </tr>`;
          }).join('') + '</tbody>';
  }

  const chartsAder = {};
  function desenharRankingAder(q, sups, temPresDia){
    const wrap = document.getElementById('toaQ' + q + 'AderWrap');
    if(!wrap) return;
    if(chartsAder[q]){ chartsAder[q].destroy(); chartsAder[q] = null; }
    const dados = sups.filter(g => g.esc > 0).map(g => ({ label: g.label, v: g.escIni / g.esc, n: g.escIni, d: g.esc }))
      .sort((a, b) => b.v - a.v || b.d - a.d);
    if(!temPresDia || !dados.length || typeof Chart === 'undefined'){
      wrap.innerHTML = `<div class="mono" style="font-size:12px; color:${COR.muted}; padding:20px 0;">${!temPresDia ? 'Sem presença do TOA no dia.' : 'Nenhum escalado (sem restrição) no dia.'}</div><canvas id="toaQ${q}AderChart" style="display:none;"></canvas>`;
      return;
    }
    wrap.innerHTML = `<canvas id="toaQ${q}AderChart"></canvas>`;
    wrap.style.height = Math.max(180, dados.length * 30 + 40) + 'px';
    const cor = v => v >= META_ADERENCIA ? '#059669' : (v >= META_ADERENCIA - 0.2 ? '#d97706' : '#dc2626');
    const plugins = typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [];
    chartsAder[q] = new Chart(document.getElementById('toaQ' + q + 'AderChart'), {
      type: 'bar',
      data: { labels: dados.map((x, i) => (i === 0 ? '🏆 ' : '') + x.label),
        datasets: [{ label: 'Aderência', data: dados.map(x => Math.round(x.v * 1000) / 10), backgroundColor: dados.map(x => cor(x.v)), borderRadius: 4, barThickness: 18,
          datalabels: { color: '#111827', anchor: 'end', align: 'right', font: { size: 11, family: 'JetBrains Mono', weight: '700' }, formatter: (v, c) => v.toFixed(0) + '% (' + dados[c.dataIndex].n + '/' + dados[c.dataIndex].d + ')' } }] },
      plugins,
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { right: 70 } },
        plugins: { legend: { display: false }, datalabels: { display: true },
          tooltip: { callbacks: { label: c => ' ' + dados[c.dataIndex].n + ' de ' + dados[c.dataIndex].d + ' escalados logaram (' + c.raw + '%)' } } },
        scales: { x: { min: 0, max: 100, ticks: { color: COR.text, callback: v => v + '%' }, grid: { color: COR.line } },
                  y: { ticks: { color: COR.text, font: { size: 11, weight: 'bold' } }, grid: { display: false } } } }
    });
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

  function renderPresenca(r){
    const tabela = document.getElementById('toaPresencaTable');
    const fora = document.getElementById('toaPresencaFora');
    if(!tabela) return;
    fora.textContent = '';
    if(!r.temPresenca){
      tabela.innerHTML = `<tbody><tr><td style="text-align:center; color:${COR.muted};">Sem dados de presença no período. O robô do Power Automate grava na aba TOA_PRESENCA (a partir de 05/11/2026).</td></tr></tbody>`;
      return;
    }
    if(!r.presenca.length || !r.diasPres.length){
      tabela.innerHTML = `<tbody><tr><td style="text-align:center; color:${COR.muted};">Nenhum técnico para esse filtro.</td></tr></tbody>`;
      return;
    }
    const tot = r.diasPres.map((d, i) => r.presenca.reduce((a, s) => a + s.porDia[i].n, 0));
    const totEq = r.presenca.reduce((a, s) => a + s.equipe, 0);
    const cab = '<thead><tr><th>Supervisor</th><th class="num">Equipe</th>' +
      r.diasPres.map(d => `<th class="num" style="white-space:nowrap;">${fmtDiaCurto(d)}<br><span style="font-weight:400;">${DIAS_SEMANA[diaSemana(d)]}</span></th>`).join('') +
      '<th class="num">Média logados/dia</th></tr></thead>';
    const corpo = r.presenca.map(s => {
      const cel = s.porDia.map(x => {
        const t = x.n + ' de ' + x.equipe + ' logados' + (x.nomes.length ? '\nLogados: ' + x.nomes.join(', ') : '') + (x.naoLog.length ? '\nNão logados: ' + x.naoLog.join(', ') : '');
        return `<td class="num" title="${esc(t)}" style="white-space:nowrap;"><b>${x.n}</b></td>`;
      }).join('');
      return `<tr><td style="white-space:nowrap;">${esc(s.label)}</td><td class="num">${s.equipe}</td>${cel}
        <td class="num"><b>${s.media.toFixed(1).replace('.', ',')}</b></td></tr>`;
    }).join('');
    const rod = `<tfoot><tr style="font-weight:700;"><td>Total</td><td class="num">${totEq}</td>${tot.map(n => `<td class="num">${n}</td>`).join('')}<td class="num">${(tot.reduce((a, b) => a + b, 0) / Math.max(1, tot.length)).toFixed(1).replace('.', ',')}</td></tr></tfoot>`;
    tabela.innerHTML = cab + '<tbody>' + corpo + '</tbody>' + rod;
    const notas = [];
    if(AJUSTES_PRES.length) notas.push(`✔ ${AJUSTES_PRES.length} nome(s) do TOA ajustados para o nome do Organograma (acento ou sobrenome cortado): <span title="${esc(AJUSTES_PRES.map(a => a.toa + ' → ' + a.organograma).join('\n'))}" style="text-decoration:underline dotted; cursor:help;">ver ajustes</span>`);
    if(r.presFora.length){
      notas.push(`ℹ ${r.presFora.length} técnico(s) logado(s) no TOA que não estão como técnicos ativos desta equipe no Organograma: <span title="${esc(r.presFora.join('\n'))}" style="text-decoration:underline dotted; cursor:help;">ver nomes</span>`);
    }
    fora.innerHTML = notas.join('<br>');
    if(false){
      fora.innerHTML = `ℹ ${r.presFora.length} técnico(s) logado(s) no TOA que não estão como técnicos ativos desta equipe no Organograma: <span title="${esc(r.presFora.join('\n'))}" style="text-decoration:underline dotted; cursor:help;">ver nomes</span>`;
    }
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
    if(r.diasSemBase && r.diasSemBase.length) partes.push(`⏳ Sem base do TOA ainda (fora da aderência): ${r.diasSemBase.map(d => fmtDiaCurto(d)).join(', ')}`);
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
