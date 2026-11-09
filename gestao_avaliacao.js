// gestao_avaliacao.js — Avaliação de Experiência (30 / 50 / 75 dias) — aba lateral dentro da página Organograma
//
// Lista os técnicos ativos admitidos a partir de ADMISSAO_DESDE (dados.json → colaboradores) e
// deixa o supervisor preencher a avaliação de cada etapa. A gravação vai para o fluxo do
// Power Automate FLUXO_AVALIACAO, que roda o Office Script "Avaliação Experiência" na planilha
// do Organograma (aba AVALIACAO_EXPERIENCIA). O dados.json devolve as avaliações gravadas em
// "avaliacoesExperiencia"; até o próximo dados.json, a cópia salva fica neste navegador.
//
// Para ligar: no index.html, antes de tela_inicial.js:
//   <script src="gestao_avaliacao.js?v=1"></script>

(function () {
  'use strict';

  // >>> COLE AQUI a URL do gatilho HTTP do fluxo "Avaliação Experiência – Gravar" <<<
  const FLUXO_AVALIACAO = 'https://defaultc54e138a79f54294aa9bb004cf8d7f.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/8ac7331ffa21496caf9fbb84763a05ae/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jnwHYpz7ZOuYh2QtvqebFXDA-H7ln30RcRt_Hvu6Pd8';
  const ADMISSAO_DESDE = '2026-10-01';   // técnicos admitidos a partir desta data
  const CHAVE_LOCAL = 'avalExpLocais_v1';
  const DESLIG_DESDE = '2026-07-01';    // desligamentos a partir desta data (mesma base do Turnover)
  const CHAVE_LOCAL_SAIDA = 'saidaLocais_v1';

  const ESCALA = [
    [1, 'Não Satisfatório'], [2, 'Em Desenvolvimento'], [3, 'Satisfatório'], [4, 'Bom'], [5, 'Excepcional']
  ];
  const SECOES = [
    { titulo: 'Competências Técnicas e Operacionais (Hard Skills)', grupos: [
      { nome: 'Conhecimento Técnico Teórico', itens: [
        ['h1', 'Interpretação de Diagramas Elétricos e Esquemas', 'Consegue analisar o diagrama e identificar os componentes no painel.'],
        ['h2', 'Normas de Segurança (NR-10, NR-12, NR-18, NR-35)', 'Sabe os procedimentos, utiliza EPI corretamente e demonstra preocupação constante.'],
        ['h3', 'Princípios de Operação de Equipamentos (Ex.: CLPs, Inversores)', 'Identificou a falha no módulo de entrada do CLP X.'] ] },
      { nome: 'Manutenção e Operação', itens: [
        ['h4', 'Execução de Manutenção Preventiva', 'Segue o checklist à risca, não pula etapas e identifica pontos de atenção.'],
        ['h5', 'Diagnóstico e Resolução de Falhas (Man. Corretiva)', 'Consegue isolar o problema de forma lógica e eficiente, usando multímetro e analisador.'],
        ['h6', 'Procedimentos de Partida e Parada de Sistemas', 'Executou o procedimento de partida do sistema Y conforme o manual, com segurança.'] ] },
      { nome: 'Uso de Ferramentas e Instrumentos', itens: [
        ['h7', 'Ferramentas Manuais (Chaves, Alicates, etc.)', 'Usa as ferramentas corretas para cada tarefa e as mantém organizadas e em bom estado.'],
        ['h8', 'Instrumentos de Medição (Multímetro, Megômetro, etc.)', 'Sabe configurar o multímetro para medir tensão, corrente e continuidade com segurança.'] ] },
      { nome: 'Competências Administrativas', itens: [
        ['h9', 'Preenchimento de Relatórios de Serviço (OS)', 'Os relatórios são claros, objetivos e com todas as informações preenchidas (horas, peças, causa raiz).'],
        ['h10', 'Registro de Atividades no Sistema', 'Atualiza o status da OS no sistema em tempo real, melhorando o rastreamento.'],
        ['h11', 'Gestão de Peças e Materiais', 'Controla o estoque da van, solicita peças com antecedência e presta contas corretamente.'] ] } ] },
    { titulo: 'Competências Comportamentais (Soft Skills)', grupos: [
      { nome: '', itens: [
        ['s1', 'Proatividade', 'Antes de finalizar a OS, verificou outros pontos do sistema que poderiam gerar falhas futuras.'],
        ['s2', 'Comunicação', 'Comunica-se bem com o cliente no local e avisa a central sobre atrasos ou necessidades.'],
        ['s3', 'Trabalho em Equipe', 'Disponibilizou-se para ajudar um colega que estava com dificuldade em uma corretiva complexa.'],
        ['s4', 'Organização e Planejamento', 'Planeja a rota do dia para otimizar deslocamento e carrega a van com o material necessário.'],
        ['s5', 'Comprometimento com a Segurança', 'Não hesita em parar uma atividade se identificar uma condição insegura.'],
        ['s6', 'Resolução de Problemas sob Pressão', 'Manteve a calma e resolveu a falha crítica no cliente Z minimizando o tempo de parada.'] ] },
      // NOVO (06/11/2026): itens do modelo de avaliação por valores que ainda não estavam no formulário
      // (Comunicação, Colaboração e Segurança já existiam acima e não foram repetidos)
      { nome: 'Valores e Conduta', itens: [
        ['s7', 'Ética, Integridade e Confidencialidade', 'Respeita as normas internas, trata bem as pessoas e protege informações da empresa, clientes e colegas.'],
        ['s8', 'Foco no Cliente Interno e Externo', 'Cuida de prazo, qualidade e postura no atendimento e entende o impacto do seu trabalho na operação.'],
        ['s9', 'Senso de Dono e Responsabilidade', 'Cumpre horários, acompanha o que assume até o fim e avisa a tempo quando há impedimento.'],
        ['s10', 'Disciplina na Gestão de Informações', 'Checklists, OS, sistemas e relatórios completos, corretos, no prazo e rastreáveis.'],
        ['s11', 'Melhoria Contínua e Abertura a Feedback', 'Corrige erros apontados, aplica as orientações recebidas e busca aprender por conta própria.'],
        ['s12', 'Orientação para Resultados', 'Entrega o combinado com qualidade e no prazo, reduz retrabalho e contribui para as metas da área.'] ] },
      { nome: 'Aderência ao Cargo', itens: [
        ['s13', 'Aprendizagem, Adaptação e Aderência à Função', 'Está se adaptando à rotina, entende suas responsabilidades e evolui no ritmo esperado para o cargo.'] ] } ] }
  ];
  const TODOS = SECOES.flatMap(s => s.grupos.flatMap(g => g.itens));
  const HARD = TODOS.filter(i => i[0][0] === 'h').map(i => i[0]);
  const SOFT = TODOS.filter(i => i[0][0] === 's').map(i => i[0]);
  const DECISOES = [
    ['APROVADO', '✅ Aprovado — segue na empresa', '#047857', 'rgba(4,120,87,.12)'],
    ['ATENCAO', '⚠️ Em atenção — acompanhar / plano de ação', '#b45309', 'rgba(180,83,9,.12)'],
    ['REPROVADO', '❌ Reprovado — não efetivar', '#b91c1c', 'rgba(185,28,28,.12)']
  ];
  const ETAPAS = ['30', '50', '75'];   // 3 períodos: 30, 50 e 75 dias (janela abre 7 dias antes do prazo)

  let COLABS = [], AVAL = {}, carregado = false;
  let DESLIG = [], ATIVOS_TEC = [], AREAS = [], SAIDA = {};
  const estadoS = { sup: '', busca: '', st: '' };
  const estado = { sup: '', busca: '' };

  // ---------- util ----------
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const semAcento = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const pad2 = n => (n < 10 ? '0' : '') + n;
  const isoDe = d => d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  const hojeIso = () => isoDe(new Date());
  const somar = (iso, n) => { const p = iso.split('-').map(Number); const d = new Date(p[0], p[1] - 1, p[2]); d.setDate(d.getDate() + n); return isoDe(d); };
  const dias = (a, b) => { const pa = a.split('-').map(Number), pb = b.split('-').map(Number); return Math.round((new Date(pb[0], pb[1] - 1, pb[2]) - new Date(pa[0], pa[1] - 1, pa[2])) / 86400000); };
  const br = iso => iso ? iso.split('-').reverse().join('/') : '';
  const matNum = m => String(m == null ? '' : m).split('-').pop().trim();
  const media = (r, ids) => { const v = ids.map(i => Number((r.notas || {})[i])).filter(n => n >= 1 && n <= 5); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
  const fmtMedia = m => m == null ? '—' : m.toFixed(1).replace('.', ',');

  function locais() { try { return JSON.parse(localStorage.getItem(CHAVE_LOCAL) || '{}'); } catch (e) { return {}; } }
  function salvarLocal(chave, r) { try { const l = locais(); l[chave] = r; localStorage.setItem(CHAVE_LOCAL, JSON.stringify(l)); } catch (e) { } }

  function lerResposta(txt) {
    let d = null;
    try { d = JSON.parse(String(txt || '').replace(/^﻿/, '')); } catch (e) { return null; }
    for (let i = 0; i < 4; i++) {
      if (typeof d === 'string') { try { d = JSON.parse(d); continue; } catch (e) { return null; } }
      if (d && typeof d === 'object' && !Array.isArray(d) && !('gravadas' in d) && !('ok' in d)) {
        const emb = ['result', 'body', 'Result', 'Body', 'resultado', 'content', 'value', 'outputs'].find(k => k in d);
        if (emb) { d = d[emb]; continue; }
      }
      break;
    }
    return d && typeof d === 'object' ? d : null;
  }

  // ---------- dados ----------
  function carregar() {
    return fetch('dados.json?_=' + Date.now()).then(r => r.ok ? r.json() : null).catch(() => null).then(d => {
      const cols = (d && Array.isArray(d.colaboradores)) ? d.colaboradores : [];
      COLABS = cols.filter(c => c && c.ativo && c.nome && String(c.dataAdmis || '') >= ADMISSAO_DESDE && /tecnic/.test(semAcento(c.areaTecnica)))
        .map(c => ({ nome: String(c.nome).trim(), mat: matNum(c.matricula), sup: String(c.supervisor || '').trim(), adm: String(c.dataAdmis).slice(0, 10),
          area: String(c.areaTecnica || ''), un: String(c.unidadeNegocio || '') }));
      AVAL = {};
      (d && Array.isArray(d.avaliacoesExperiencia) ? d.avaliacoesExperiencia : []).forEach(x => {
        try { const o = typeof x === 'string' ? JSON.parse(x) : x; const r = typeof o.r === 'string' ? JSON.parse(o.r) : (o.r || o); r.__em = o.em || ''; r.__base = true; AVAL[matNum(r.matricula) + '|' + r.etapa] = r; } catch (e) { }
      });
      // cópia salva neste navegador vale até o dados.json trazer uma versão igual ou mais nova
      const l = locais();
      Object.keys(l).forEach(k => { const b = AVAL[k]; if (!b || String(b.__em || '') < String(l[k].__em || '')) AVAL[k] = l[k]; });
      // ----- saídas (desligamentos) -----
      const mapC = c => ({ nome: String(c.nome).trim(), mat: matNum(c.matricula), sup: String(c.supervisor || '').trim(), coord: String(c.coordenador || '').trim(),
        contrato: String(c.contrato || '').trim(), un: String(c.unidadeNegocio || '').trim(), cargo: String(c.descFuncao || c.areaTecnica || '').trim(),
        area: String(c.areaTecnica || ''), adm: String(c.dataAdmis || '').slice(0, 10), dem: String(c.dataDemissao || '').slice(0, 10) });
      // ALTERADO (09/11/2026): quem é desligado costuma sair da aba D x P e fica sem "Área Técnica" no Organograma;
      // por isso também vale a função do cadastro (Desc.Funcao do SRA), ex.: "TECNICO DE MANUTENCAO".
      const ehTec = c => /tecnic|lid/.test(semAcento(c.areaTecnica)) || /tecnic|lider/.test(semAcento(c.descFuncao));
      const porMat = {};
      cols.forEach(c => { if (c && c.matricula != null) porMat[matNum(c.matricula)] = c; });
      // 1) desligamentos do Turnover (dados.json → turnover): tipo, causa e "o que entra" do RH
      DESLIG = [];
      (d && Array.isArray(d.turnover) ? d.turnover : []).forEach(t => {
        if (!t || !t.nome || String(t.dataDemissao || '') < DESLIG_DESDE) return;
        const mat = matNum(t.matricula), c = porMat[mat];
        const b = c ? mapC(c) : { nome: String(t.nome).trim(), mat, sup: '', coord: '', contrato: '', un: '', cargo: '', area: '' };
        DESLIG.push(Object.assign(b, { nome: String(t.nome).trim(), mat, sup: String(t.supervisor || b.sup || '').trim(), adm: String(t.dataAdmis || b.adm || '').slice(0, 10),
          dem: String(t.dataDemissao || '').slice(0, 10), tipoRH: String(t.tipoDesligamento || '').trim(), causaRH: String(t.causaPareto || '').trim(), oqRH: String(t.oQueEntra || '').trim() }));
      });
      // 2) técnicos desligados no Organograma que ainda não estão no Turnover
      cols.filter(c => c && c.nome && ehTec(c) && String(c.dataDemissao || '') >= DESLIG_DESDE && !DESLIG.some(x => x.mat === matNum(c.matricula)))
        .forEach(c => DESLIG.push(mapC(c)));
      ATIVOS_TEC = cols.filter(c => c && c.nome && c.ativo && ehTec(c)).map(mapC).sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
      AREAS = [...new Set(cols.map(c => String(c.unidadeNegocio || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
      SAIDA = {};
      (d && Array.isArray(d.entrevistasDesligamento) ? d.entrevistasDesligamento : []).forEach(x => {
        try { const o = typeof x === 'string' ? JSON.parse(x) : x; const r = typeof o.r === 'string' ? JSON.parse(o.r) : (o.r || o); r.__em = o.em || ''; SAIDA[matNum(r.matricula)] = r; } catch (e) { }
      });
      const ls = locaisSaida();
      Object.keys(ls).forEach(k => { const b = SAIDA[k]; if (!b || String(b.__em || '') < String(ls[k].__em || '')) SAIDA[k] = ls[k]; });
      // quem preencheu a saída de alguém que ainda não está como desligado no Organograma também aparece
      Object.values(SAIDA).forEach(r => { if (!DESLIG.some(x => x.mat === matNum(r.matricula))) DESLIG.push({ nome: r.nome, mat: matNum(r.matricula), sup: r.supervisor, coord: r.coordenador, contrato: r.contrato, un: r.areaProjeto, cargo: r.cargo, adm: r.dataAdmissao, dem: r.dataDesligamento }); });
      carregado = true;
    });
  }
  function locaisSaida() { try { return JSON.parse(localStorage.getItem(CHAVE_LOCAL_SAIDA) || '{}'); } catch (e) { return {}; } }
  function salvarLocalSaida(chave, r) { try { const l = locaisSaida(); l[chave] = r; localStorage.setItem(CHAVE_LOCAL_SAIDA, JSON.stringify(l)); } catch (e) { } }

  // ---------- estrutura ----------
  // Abas do lado, dentro da página Organograma (igual à Preventiva):
  //   👥 Estrutura (tudo o que já existia na página)  ·  📝 Avaliação de Experiência
  let abaOrg = 'estrutura';
  function montarPagina() {
    const pg = document.getElementById('pagina-organograma');
    if (!pg) return null;
    if (document.getElementById('org-aba-aval')) return document.getElementById('org-aba-aval');
    if (!document.getElementById('org-abas-css')) {
      const css = document.createElement('style'); css.id = 'org-abas-css';
      css.textContent = `
        .org-layout{display:flex;gap:20px;align-items:flex-start}
        .org-menu{width:200px;flex:0 0 200px;background:var(--panel,#fff);border:1px solid var(--line,#e1e5f0);border-radius:14px;padding:10px;position:sticky;top:20px;box-shadow:0 2px 8px rgba(27,36,64,.06)}
        .org-menu .nav-group{margin-bottom:0}
        .org-menu .nav-group-label{cursor:default}
        .org-conteudo{flex:1;min-width:0}
        @media (max-width:860px){.org-layout{flex-direction:column}.org-menu{width:100%;position:static}}`;
      document.head.appendChild(css);
    }
    const estrutura = document.createElement('div'); estrutura.id = 'org-aba-estrutura';
    while (pg.firstChild) estrutura.appendChild(pg.firstChild);
    const aval = document.createElement('div'); aval.id = 'org-aba-aval'; aval.style.display = 'none';
    const layout = document.createElement('div'); layout.className = 'org-layout';
    layout.innerHTML = `<nav class="org-menu"><div class="nav-group"><div class="nav-group-label"><span>Organograma</span></div><div class="nav-group-items">
        <div class="nav-item org-aba active" data-aba="estrutura"><span class="nav-icon">👥</span> Estrutura</div>
        <div class="nav-item org-aba" data-aba="aval"><span class="nav-icon">📝</span> Avaliação de Experiência</div>
        <div class="nav-item org-aba" data-aba="saida"><span class="nav-icon">🗣️</span> Entrevista de Desligamento</div>
      </div></div></nav><div class="org-conteudo"></div>`;
    layout.querySelector('.org-conteudo').appendChild(estrutura);
    layout.querySelector('.org-conteudo').appendChild(aval);
    const saida = document.createElement('div'); saida.id = 'org-aba-saida'; saida.style.display = 'none';
    layout.querySelector('.org-conteudo').appendChild(saida);
    pg.appendChild(layout);
    layout.querySelectorAll('.org-aba').forEach(b => b.addEventListener('click', () => trocarAbaOrg(b.dataset.aba)));
    return aval;
  }
  function trocarAbaOrg(aba) {
    abaOrg = aba;
    document.querySelectorAll('#pagina-organograma .org-aba').forEach(b => b.classList.toggle('active', b.dataset.aba === aba));
    document.getElementById('org-aba-estrutura').style.display = aba === 'estrutura' ? '' : 'none';
    document.getElementById('org-aba-aval').style.display = aba === 'aval' ? '' : 'none';
    document.getElementById('org-aba-saida').style.display = aba === 'saida' ? '' : 'none';
    if (aba === 'aval') iniciar();
    else if (aba === 'saida') iniciarSaida();
    else if (typeof renderHeadcount === 'function') renderHeadcount();
  }

  function montar() {
    const pg = montarPagina();
    if (!pg || document.getElementById('avalExpSecao')) return;
    if (!document.getElementById('aval-css')) {
      const css = document.createElement('style'); css.id = 'aval-css';
      css.textContent = `
        #avalExpSecao table{ width:100%; border-collapse:collapse; }
        #avalExpSecao th{ font-size:10.5px; text-transform:uppercase; letter-spacing:.04em; color:var(--text); text-align:left; padding:7px 8px; border-bottom:1px solid var(--line); }
        #avalExpSecao td{ font-size:13px; color:var(--text); padding:7px 8px; border-bottom:1px solid var(--line); vertical-align:middle; }
        .av-et{ display:inline-flex; flex-direction:column; align-items:flex-start; gap:3px; border:1px solid var(--line); border-radius:10px; padding:6px 9px; min-width:118px; cursor:pointer; background:var(--panel); font-size:11.5px; line-height:1.25; }
        .av-et:hover{ border-color:#1e3a8a; }
        .av-et b{ font-size:12px; }
        .av-modal-fundo{ position:fixed; inset:0; background:rgba(15,23,42,.55); z-index:9999; display:flex; align-items:flex-start; justify-content:center; overflow:auto; padding:24px 12px; }
        .av-modal{ background:#fff; color:#111827; width:min(980px,100%); border-radius:14px; box-shadow:0 20px 60px rgba(0,0,0,.35); }
        .av-modal h3{ margin:0; font-size:18px; }
        .av-modal .av-sec{ padding:14px 22px; border-top:1px solid #e5e7eb; }
        .av-modal .av-sec h4{ margin:0 0 4px; font-size:14px; color:#1e3a8a; text-transform:uppercase; letter-spacing:.04em; }
        .av-modal .av-grupo{ font-size:12px; font-weight:700; color:#374151; margin:12px 0 4px; }
        .av-item{ display:grid; grid-template-columns: minmax(200px,1.1fr) auto minmax(220px,1.4fr); gap:10px; align-items:start; padding:7px 0; border-bottom:1px dashed #e5e7eb; }
        .av-item .av-nome{ font-size:13px; font-weight:600; }
        .av-notas{ display:flex; gap:4px; }
        .av-notas label{ width:30px; height:30px; border:1px solid #cbd5e1; border-radius:8px; display:flex; align-items:center; justify-content:center; font-weight:700; cursor:pointer; font-size:13px; user-select:none; }
        .av-notas input{ display:none; }
        .av-notas input:checked + span{ color:#fff; }
        .av-notas label:has(input:checked){ background:#1e3a8a; border-color:#1e3a8a; color:#fff; }
        .av-modal textarea, .av-modal input[type=text], .av-modal input[type=date], .av-modal select{ width:100%; box-sizing:border-box; border:1px solid #cbd5e1; border-radius:8px; padding:7px 9px; font:inherit; font-size:13px; color:#111827; background:#fff; }
        .av-modal textarea{ min-height:52px; resize:vertical; }
        .av-modal .av-l{ font-size:11px; font-weight:700; color:#374151; text-transform:uppercase; letter-spacing:.04em; margin:10px 0 4px; display:block; }
        .av-dec{ display:flex; gap:8px; flex-wrap:wrap; }
        .av-dec label{ border:2px solid #e5e7eb; border-radius:10px; padding:8px 12px; cursor:pointer; font-size:13px; font-weight:600; }
        .av-dec input{ margin-right:6px; }
        .av-rodape{ display:flex; gap:10px; justify-content:flex-end; align-items:center; padding:14px 22px; border-top:1px solid #e5e7eb; position:sticky; bottom:0; background:#fff; border-radius:0 0 14px 14px; }
        .av-bt{ border:0; border-radius:9px; padding:9px 16px; font-weight:700; cursor:pointer; font-size:13px; }
        .av-bt.pri{ background:#1e3a8a; color:#fff; } .av-bt.sec{ background:#e5e7eb; color:#111827; }
        @media (max-width:720px){ .av-item{ grid-template-columns:1fr; } }
        @media print{ body *{ visibility:hidden; } .av-modal, .av-modal *{ visibility:visible; } .av-modal{ position:absolute; left:0; top:0; box-shadow:none; } .av-rodape{ display:none; } }`;
      document.head.appendChild(css);
    }
    const sec = document.createElement('div');
    sec.className = 'panel'; sec.id = 'avalExpSecao';
    sec.innerHTML = `
      <h2>🧭 Avaliação de Experiência — técnicos novos (${ETAPAS.join(' / ')} dias)</h2>
      <div class="hint">Técnicos ativos admitidos a partir de ${br(ADMISSAO_DESDE)}. Clique na etapa para o supervisor preencher a avaliação; o RH acompanha aqui o que já foi avaliado e a decisão.</div>
      <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:10px 0 12px;">
        <select id="avSup" class="mono" style="font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); min-width:220px;"><option value="">Supervisor: todos</option></select>
        <div class="search"><input id="avBusca" type="search" placeholder="Buscar técnico..." aria-label="Buscar técnico"></div>
        <span id="avResumo" class="mono" style="font-size:12px; color:var(--text); margin-left:auto;"></span>
      </div>
      <div style="overflow:auto; border-radius:8px;"><table id="avTabela"></table></div>
      ${FLUXO_AVALIACAO ? '' : '<div class="mono" style="font-size:12px; color:#b45309; margin-top:10px;">⚠ O envio ainda não está ligado: falta colar a URL do fluxo “Avaliação Experiência – Gravar” em FLUXO_AVALIACAO (gestao_avaliacao.js).</div>'}`;
    pg.appendChild(sec);
    document.getElementById('avSup').addEventListener('change', e => { estado.sup = e.target.value; render(); });
    document.getElementById('avBusca').addEventListener('input', e => { estado.busca = e.target.value; render(); });
    sec.addEventListener('click', e => {
      const b = e.target.closest('[data-av]');
      if (b) abrir(b.dataset.av.split('|')[0], b.dataset.av.split('|')[1]);
    });
  }

  // ---------- lista ----------
  function situacao(c, etapa) {
    const r = AVAL[c.mat + '|' + etapa];
    const prazo = somar(c.adm, Number(etapa));
    const hoje = hojeIso();
    if (r) return { r, prazo, txt: 'Avaliado', cor: '#047857' };
    if (hoje > prazo) return { prazo, txt: 'Atrasada · ' + dias(prazo, hoje) + 'd', cor: '#b91c1c' };
    if (dias(hoje, prazo) <= 7) return { prazo, txt: 'Vence em ' + dias(hoje, prazo) + 'd', cor: '#b45309' };
    return { prazo, txt: 'A partir de ' + br(somar(prazo, -7)).slice(0, 5), cor: '#6b7280' };
  }
  function render() {
    const tb = document.getElementById('avTabela');
    if (!tb) return;
    const sups = [...new Set(COLABS.map(c => c.sup).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const sel = document.getElementById('avSup');
    sel.innerHTML = '<option value="">Supervisor: todos</option>' + sups.map(s => `<option value="${esc(s)}"${s === estado.sup ? ' selected' : ''}>${esc(s)}</option>`).join('');
    const q = semAcento(estado.busca);
    const lista = COLABS.filter(c => (!estado.sup || c.sup === estado.sup) && (!q || semAcento(c.nome + ' ' + c.sup).includes(q)))
      .sort((a, b) => a.sup.localeCompare(b.sup, 'pt-BR') || a.adm.localeCompare(b.adm) || a.nome.localeCompare(b.nome, 'pt-BR'));
    let feitas = 0, atrasadas = 0;
    lista.forEach(c => ETAPAS.forEach(e => { const s = situacao(c, e); if (s.r) feitas++; else if (/Atrasada/.test(s.txt)) atrasadas++; }));
    document.getElementById('avResumo').innerHTML = `${lista.length} técnico(s) · ✅ ${feitas} avaliações feitas · <span style="color:#b91c1c; font-weight:700;">🔴 ${atrasadas} atrasada(s)</span>`;
    if (!carregado) { tb.innerHTML = '<tbody><tr><td>Carregando…</td></tr></tbody>'; return; }
    if (!lista.length) { tb.innerHTML = `<tbody><tr><td style="color:var(--muted);">Nenhum técnico admitido a partir de ${br(ADMISSAO_DESDE)} nesse filtro.</td></tr></tbody>`; return; }
    const hoje = hojeIso();
    tb.innerHTML = '<thead><tr><th>Técnico</th><th>Supervisor</th><th>Admissão</th><th>Dias de casa</th>' + ETAPAS.map(e => '<th>' + e + ' dias</th>').join('') + '<th>Última decisão</th></tr></thead><tbody>' +
      lista.map(c => {
        const cel = ETAPAS.map(e => {
          const s = situacao(c, e);
          const m = s.r ? media(s.r, HARD.concat(SOFT)) : null;
          return `<td><div class="av-et" data-av="${esc(c.mat)}|${e}" title="Prazo: ${br(s.prazo)}">
            <b style="color:${s.cor};">${s.r ? '✅ ' : ''}${esc(s.txt)}</b>
            <span>${s.r ? 'Média ' + fmtMedia(m) + ' · ' + br(s.r.dataAvaliacao).slice(0, 5) : 'Prazo ' + br(s.prazo).slice(0, 5)}</span></div></td>`;
        }).join('');
        const ult = ETAPAS.map(e => AVAL[c.mat + '|' + e]).filter(Boolean).pop();
        const dec = ult ? DECISOES.find(d => d[0] === ult.decisao) : null;
        return `<tr><td><b>${esc(c.nome)}</b><div style="font-size:11px; color:var(--muted);">${esc(c.mat)} · ${esc(c.un || c.area)}</div></td>
          <td>${esc(c.sup) || '—'}</td><td class="mono">${br(c.adm)}</td><td class="mono">${Math.max(0, dias(c.adm, hoje))}</td>${cel}
          <td>${dec ? `<span style="display:inline-block; padding:3px 9px; border-radius:999px; font-size:12px; font-weight:700; color:${dec[2]}; background:${dec[3]};">${esc(dec[1].split(' — ')[0])}</span><div style="font-size:11px; color:var(--muted);">${ult.etapa} dias${ult.informacao ? ' · ' + esc(String(ult.informacao).slice(0, 60)) : ''}</div>` : '<span style="color:var(--muted);">—</span>'}</td></tr>`;
      }).join('') + '</tbody>';
  }

  // ---------- formulário ----------
  function abrir(mat, etapa) {
    const c = COLABS.find(x => x.mat === mat);
    if (!c) return;
    const r = AVAL[mat + '|' + etapa] || {};
    const n = r.notas || {}, cm = r.comentarios || {}, pl = r.planos || [];
    const fundo = document.createElement('div');
    fundo.className = 'av-modal-fundo';
    const itemHtml = it => `<div class="av-item">
        <div class="av-nome">${esc(it[1])}</div>
        <div class="av-notas">${ESCALA.map(([v, t]) => `<label title="${v} - ${t}"><input type="radio" name="n_${it[0]}" value="${v}"${String(n[it[0]]) === String(v) ? ' checked' : ''}>${v}</label>`).join('')}</div>
        <textarea data-c="${it[0]}" placeholder="Ex.: ${esc(it[2])}">${esc(cm[it[0]] || '')}</textarea>
      </div>`;
    fundo.innerHTML = `<div class="av-modal" role="dialog" aria-label="Avaliação de experiência">
      <div style="padding:18px 22px; display:flex; justify-content:space-between; gap:12px; align-items:flex-start;">
        <div><h3>Avaliação de Experiência — ${etapa} dias</h3>
          <div style="font-size:13px; color:#374151; margin-top:4px;"><b>${esc(c.nome)}</b> · matrícula ${esc(c.mat)} · supervisor ${esc(c.sup)} · admissão ${br(c.adm)} · prazo ${br(somar(c.adm, Number(etapa)))}</div></div>
        <button type="button" class="av-bt sec" data-x>✕</button>
      </div>
      <div class="av-sec" style="background:#f8fafc;">
        <div style="font-size:12px; color:#374151;"><b>Escala:</b> ${ESCALA.map(([v, t]) => v + ' - ' + t).join(' &nbsp;|&nbsp; ')}</div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
          <div><span class="av-l">Avaliador</span><input type="text" id="avAvaliador" value="${esc(r.avaliador || c.sup)}"></div>
          <div><span class="av-l">Data da avaliação</span><input type="date" id="avData" value="${esc(r.dataAvaliacao || hojeIso())}"></div>
        </div>
      </div>
      ${SECOES.map(s => `<div class="av-sec"><h4>${esc(s.titulo)}</h4>
        <div style="font-size:11px; color:#6b7280;">Nota (1-5) · Comentários / evidências observáveis (crucial para o feedback)</div>
        ${s.grupos.map(g => (g.nome ? `<div class="av-grupo">${esc(g.nome)}</div>` : '') + g.itens.map(itemHtml).join('')).join('')}</div>`).join('')}
      <div class="av-sec"><h4>Avaliação Geral e Planos de Ação</h4>
        <div id="avMedias" style="font-size:13px; font-weight:700; color:#1e3a8a; margin-bottom:6px;"></div>
        <span class="av-l">Principais pontos fortes</span><textarea id="avFortes" placeholder="Ex.: Excelente capacidade de diagnóstico, proatividade e comprometimento com a segurança.">${esc(r.pontosFortes || '')}</textarea>
        <span class="av-l">Áreas de desenvolvimento prioritárias</span><textarea id="avAreas" placeholder="Ex.: Aprofundar o conhecimento na plataforma de CLP 'ABC' e melhorar a gestão de tempo em deslocamentos.">${esc(r.areasDesenvolvimento || '')}</textarea>
        <span class="av-l">Planos de ação / metas SMART (próximos 3 meses)</span>
        <input type="text" id="avP1" placeholder="1. Ex.: Concluir o treinamento online sobre CLPs até 01/08." value="${esc(pl[0] || '')}" style="margin-bottom:6px;">
        <input type="text" id="avP2" placeholder="2. Ex.: Reduzir o tempo médio de deslocamento entre OSs em 15%." value="${esc(pl[1] || '')}" style="margin-bottom:6px;">
        <input type="text" id="avP3" placeholder="3. Ex.: Realizar uma preventiva complexa no sistema 'X' sob supervisão até 01/09." value="${esc(pl[2] || '')}">
        <span class="av-l">Feedback do colaborador (autoavaliação)</span><textarea id="avFeedback" placeholder="Visão do técnico sobre o período, acertos e dificuldades.">${esc(r.feedbackColaborador || '')}</textarea>
      </div>
      <div class="av-sec"><h4>Aprovação</h4>
        <div class="av-dec">${DECISOES.map(d => `<label style="color:${d[2]};"><input type="radio" name="avDec" value="${d[0]}"${r.decisao === d[0] ? ' checked' : ''}>${esc(d[1])}</label>`).join('')}</div>
        <span class="av-l">Informação</span><textarea id="avInfo" placeholder="Informação para o RH (ex.: motivo, próximos passos, data de reavaliação).">${esc(r.informacao || '')}</textarea>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:24px; margin-top:26px; font-size:12px; color:#374151;">
          <div style="border-top:1px solid #111827; padding-top:4px;">Assinatura do Colaborador</div>
          <div style="border-top:1px solid #111827; padding-top:4px;">Assinatura do Gestor</div>
        </div>
      </div>
      <div class="av-rodape">
        <span id="avMsg" style="font-size:12.5px; margin-right:auto;"></span>
        <button type="button" class="av-bt sec" data-print>🖨 Imprimir</button>
        <button type="button" class="av-bt sec" data-x>Cancelar</button>
        <button type="button" class="av-bt pri" data-salvar>Salvar avaliação</button>
      </div></div>`;
    document.body.appendChild(fundo);
    const $ = s => fundo.querySelector(s);
    const coletar = () => {
      const notas = {}, comentarios = {};
      TODOS.forEach(it => { const v = fundo.querySelector(`input[name="n_${it[0]}"]:checked`); if (v) notas[it[0]] = Number(v.value); const t = fundo.querySelector(`textarea[data-c="${it[0]}"]`).value.trim(); if (t) comentarios[it[0]] = t; });
      const dec = fundo.querySelector('input[name="avDec"]:checked');
      return { matricula: c.mat, nome: c.nome, supervisor: c.sup, etapa, dataAdmissao: c.adm, dataAvaliacao: $('#avData').value, avaliador: $('#avAvaliador').value.trim(),
        notas, comentarios, pontosFortes: $('#avFortes').value.trim(), areasDesenvolvimento: $('#avAreas').value.trim(),
        planos: [$('#avP1').value.trim(), $('#avP2').value.trim(), $('#avP3').value.trim()], feedbackColaborador: $('#avFeedback').value.trim(),
        decisao: dec ? dec.value : '', informacao: $('#avInfo').value.trim() };
    };
    const atualizarMedias = () => { const o = coletar(); $('#avMedias').textContent = `Média Hard Skills: ${fmtMedia(media(o, HARD))} · Média Soft Skills: ${fmtMedia(media(o, SOFT))} · Média geral: ${fmtMedia(media(o, HARD.concat(SOFT)))} · ${Object.keys(o.notas).length} de ${TODOS.length} notas`; };
    fundo.addEventListener('change', atualizarMedias);
    atualizarMedias();
    const fechar = () => fundo.remove();
    fundo.querySelectorAll('[data-x]').forEach(b => b.addEventListener('click', fechar));
    fundo.addEventListener('click', e => { if (e.target === fundo) fechar(); });
    $('[data-print]').addEventListener('click', () => window.print());
    $('[data-salvar]').addEventListener('click', async () => {
      const o = coletar(), msg = $('#avMsg'), bt = $('[data-salvar]');
      const faltam = TODOS.length - Object.keys(o.notas).length;
      if (faltam) { msg.innerHTML = `<span style="color:#b91c1c;">Falta dar nota em ${faltam} competência(s).</span>`; return; }
      if (!o.decisao) { msg.innerHTML = '<span style="color:#b91c1c;">Escolha a aprovação (Aprovado / Em atenção / Reprovado).</span>'; return; }
      if (!FLUXO_AVALIACAO) { msg.innerHTML = '<span style="color:#b91c1c;">O envio ainda não está ligado (falta a URL do fluxo).</span>'; return; }
      bt.disabled = true; bt.textContent = 'Gravando…';
      try {
        const r = await fetch(FLUXO_AVALIACAO, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(o) });
        const resp = lerResposta(await r.text());
        if (!r.ok) throw new Error('o fluxo respondeu ' + r.status);
        if (resp && resp.ok === false) throw new Error(resp.erro || 'o script recusou a gravação');
        o.__em = new Date().toISOString().replace('T', ' ').slice(0, 16);
        AVAL[c.mat + '|' + etapa] = o; salvarLocal(c.mat + '|' + etapa, o);
        render(); fechar();
      } catch (e) {
        msg.innerHTML = `<span style="color:#b91c1c;">Não gravou: ${esc(e.message)}</span>`;
        bt.disabled = false; bt.textContent = 'Salvar avaliação';
      }
    });
  }


  // =====================================================================
  // 🚪 SAÍDA DO TÉCNICO — entrevista de desligamento preenchida pelo supervisor
  // =====================================================================
  const PERGUNTAS_SAIDA = [
    { id: 'q1', txt: '1. Qual foi o tipo de desligamento?', ops: ['Pedido de demissão', 'Demissão sem justa causa', 'Demissão por justa causa', 'Abandono'] },
    { id: 'q2', txt: '2. O supervisor tinha conhecimento da intenção de saída do colaborador?', ops: ['Sim, já havia sinalizado anteriormente', 'Sim, comunicou próximo ao desligamento', 'Não, foi uma decisão inesperada'] },
    { id: 'q3', txt: '3. O colaborador apresentou algum motivo para sua saída?', ops: ['Sim', 'Não', 'Não foi possível identificar'] },
    { id: 'q4', txt: '4. Qual foi o principal motivo informado para o desligamento?', outros: true, ops: ['Proposta de outra empresa', 'Remuneração / benefícios', 'Oportunidade de crescimento profissional', 'Mudança de cidade ou região', 'Questões pessoais', 'Insatisfação com a liderança', 'Insatisfação com as atividades exercidas', 'Condições de trabalho', 'Jornada / escala de trabalho', 'Desempenho ou adaptação à função', 'Outros'] },
    { id: 'q5', txt: '5. Como você avalia o desempenho do colaborador durante o período na empresa?', ops: ['Excelente', 'Bom', 'Regular', 'Abaixo do esperado'] },
    { id: 'q6', txt: '6. Como era o relacionamento do colaborador com a equipe?', ops: ['Muito bom', 'Bom', 'Regular', 'Apresentava dificuldades de relacionamento'] },
    { id: 'q7', txt: '7. O colaborador demonstrava interesse em permanecer na empresa?', ops: ['Sim', 'Não', 'Parcialmente', 'Não foi possível avaliar'] },
    { id: 'q8', txt: '8. Na sua avaliação, o desligamento poderia ter sido evitado?', ops: ['Sim', 'Não', 'Talvez'] },
    { id: 'q9', txt: '9. Caso pudesse ter sido evitado, quais ações poderiam contribuir para a permanência?', multi: true, outros: true, ops: ['Revisão salarial', 'Oportunidade de promoção', 'Mudança de equipe ou projeto', 'Ajuste de escala', 'Outros'] }
  ];
  const OBRIG_SAIDA = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8'];

  function montarSaida() {
    montar();   // garante o CSS do formulário
    const box = document.getElementById('org-aba-saida');
    if (!box || document.getElementById('saidaSecao')) return;
    const sec = document.createElement('div');
    sec.className = 'panel'; sec.id = 'saidaSecao';
    sec.innerHTML = `
      <h2>🗣️ Entrevista de Desligamento</h2>
      <div class="hint">Desligamentos a partir de ${br(DESLIG_DESDE)} (base do Turnover + Organograma). O supervisor responde a entrevista; o RH acompanha aqui as causas e os motivos de saída.</div>
      <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:10px 0 12px;">
        <select id="sdSup" class="mono" style="font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); min-width:220px;"><option value="">Supervisor: todos</option></select>
        <select id="sdSt" class="mono" style="font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text);">
          <option value="">Situação: todas</option><option value="pend">🔴 Pendentes</option><option value="ok">✅ Preenchidas</option></select>
        <div class="search"><input id="sdBusca" type="search" placeholder="Buscar colaborador..." aria-label="Buscar colaborador"></div>
        <button type="button" id="sdNovo" class="av-bt sec" style="margin-left:auto;">+ Registrar outro desligamento</button>
      </div>
      <div class="kpi-row" id="sdKpis" style="margin-bottom:16px;"></div>
      <div class="grid" style="grid-template-columns:1fr 1fr; gap:16px; margin-bottom:16px;">
        <div class="panel" style="margin:0;"><h2>Pareto — Causas de Desligamento (RH)</h2>
          <div class="hint">Barras = quantidade por causa · linha = % acumulado · tracejado = corte de 80%</div>
          <div id="sdPar1Wrap" style="position:relative; height:320px;"><canvas id="sdPar1"></canvas></div></div>
        <div class="panel" style="margin:0;"><h2>Tipo de Desligamento</h2>
          <div class="hint">Resposta 1 da entrevista; sem entrevista, vem do tipo do Turnover (Desligamento Empresa = sem justa causa)</div>
          <div id="sdTipoWrap" style="position:relative; height:320px;"><canvas id="sdTipo"></canvas></div></div>
      </div>
      <div id="sdResumo" class="mono" style="font-size:12px; color:var(--text); margin-bottom:8px;"></div>
      <div style="overflow:auto; border-radius:8px; max-height:620px;"><table id="sdTabela"></table></div>
      <div id="sdMotivos" style="margin-top:16px;"></div>`;
    box.appendChild(sec);
    document.getElementById('sdSup').addEventListener('change', e => { estadoS.sup = e.target.value; renderSaida(); });
    document.getElementById('sdSt').addEventListener('change', e => { estadoS.st = e.target.value; renderSaida(); });
    document.getElementById('sdBusca').addEventListener('input', e => { estadoS.busca = e.target.value; renderSaida(); });
    document.getElementById('sdNovo').addEventListener('click', escolherOutro);
    sec.addEventListener('click', e => { const b = e.target.closest('[data-sd]'); if (b) abrirSaida(b.dataset.sd); });
  }

  function renderSaida() {
    const tb = document.getElementById('sdTabela');
    if (!tb) return;
    const sups = [...new Set(DESLIG.map(c => c.sup).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    document.getElementById('sdSup').innerHTML = '<option value="">Supervisor: todos</option>' + sups.map(x => `<option value="${esc(x)}"${x === estadoS.sup ? ' selected' : ''}>${esc(x)}</option>`).join('');
    if (!carregado) { tb.innerHTML = '<tbody><tr><td>Carregando…</td></tr></tbody>'; return; }
    const q = semAcento(estadoS.busca);
    const base = DESLIG.filter(c => (!estadoS.sup || c.sup === estadoS.sup) && (!q || semAcento(c.nome + ' ' + c.sup).includes(q)));
    const lista = base.filter(c => !estadoS.st || (estadoS.st === 'ok' ? feito(c) : !feito(c)))
      .sort((a, b) => (b.dem || '').localeCompare(a.dem || '') || a.nome.localeCompare(b.nome, 'pt-BR'));
    const feitos = base.filter(feito).length, antigos = base.filter(c => !SAIDA[c.mat] && c.causaRH).length;
    document.getElementById('sdResumo').innerHTML = `${base.length} desligamento(s) · ✅ ${feitos} feito(s)${antigos ? ` (${antigos} no modelo antigo)` : ''} · <span style="color:#b91c1c; font-weight:700;">🔴 ${base.length - feitos} pendente(s)</span>`;
    tb.innerHTML = !lista.length ? '<tbody><tr><td style="color:var(--muted);">Nenhum desligamento nesse filtro.</td></tr></tbody>' :
      '<thead><tr><th>Colaborador</th><th>Supervisor</th><th>Área / Projeto</th><th>Admissão</th><th>Desligamento</th><th>Tipo</th><th>Causa (RH)</th><th>Motivo (entrevista)</th><th>Poderia evitar?</th><th></th></tr></thead><tbody>' +
      lista.map(c => {
        const r = SAIDA[c.mat];
        const motivo = r ? (r.q4 === 'Outros' && r.q4_outros ? 'Outros: ' + r.q4_outros : r.q4) : '';
        return `<tr><td><b>${esc(c.nome)}</b><div style="font-size:11px; color:var(--muted);">${esc(c.mat)} · ${esc(c.cargo)}</div></td>
          <td>${esc(c.sup) || '—'}</td><td>${esc(c.un) || '—'}</td><td class="mono">${br(c.adm)}</td><td class="mono">${br(c.dem) || '—'}</td>
          <td>${esc(r ? r.q1 : c.tipoRH) || '—'}</td><td>${c.causaRH ? esc(c.causaRH) + (c.oqRH ? `<div style="font-size:11px; color:var(--muted);">${esc(c.oqRH)}</div>` : '') : '—'}</td><td>${r ? esc(motivo) : '—'}</td><td>${r ? esc(r.q8) : '—'}</td>
          <td><button type="button" class="av-bt ${feito(c) ? 'sec' : 'pri'}" data-sd="${esc(c.mat)}" title="${!r && c.causaRH ? 'Feito no modelo antigo (causa do Turnover). Clique se quiser completar a entrevista nova.' : ''}">${r ? '✅ Ver / editar' : c.causaRH ? '✅ Modelo antigo' : '🔴 Preencher'}</button></td></tr>`;
      }).join('') + '</tbody>';
    // resumo dos motivos (o que o RH quer ver)
    const resp = base.map(c => SAIDA[c.mat]).filter(Boolean);
    const cont = id => { const m = {}; resp.forEach(r => { const v = r[id]; (Array.isArray(v) ? v : [v]).filter(Boolean).forEach(x => m[x] = (m[x] || 0) + 1); }); return Object.entries(m).sort((a, b) => b[1] - a[1]); };
    const bloco = (tit, id) => { const l = cont(id); return `<div style="flex:1 1 260px; border:1px solid var(--line); border-radius:12px; padding:10px 12px; background:var(--panel);">
        <div style="font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:var(--text); margin-bottom:6px;">${tit}</div>
        ${l.length ? l.map(([k, n]) => `<div style="display:flex; gap:8px; align-items:center; font-size:12.5px; margin:3px 0;"><span style="flex:1;">${esc(k)}</span><b>${n}</b>
          <span style="display:inline-block; height:8px; width:${Math.round(n / resp.length * 90)}px; background:#1e3a8a; border-radius:4px;"></span></div>`).join('') : '<span style="font-size:12px; color:var(--muted);">Sem respostas ainda.</span>'}</div>`; };
    desenharKpisParetos(base);
    document.getElementById('sdMotivos').innerHTML = resp.length ? `<div style="display:flex; gap:12px; flex-wrap:wrap;">${bloco('Tipo de desligamento', 'q1')}${bloco('Principal motivo', 'q4')}${bloco('Poderia ter sido evitado?', 'q8')}${bloco('Ações para permanência', 'q9')}</div>` : '';
  }


  // sugestão das respostas a partir do Turnover (o supervisor confere / corrige)
  function sugestaoRH(c) {
    const o = {};
    const t = semAcento(c.tipoRH), k = semAcento(c.causaRH + ' ' + c.oqRH);
    if (/pedido/.test(t)) o.q1 = 'Pedido de demissão';
    else if (/abandono/.test(t)) o.q1 = 'Abandono';
    else if (/justa causa/.test(k) && !/sem justa/.test(k)) o.q1 = 'Demissão por justa causa';
    else if (/empresa/.test(t)) o.q1 = 'Demissão sem justa causa';
    if (/proposta|outra empresa|pj/.test(k)) o.q4 = 'Proposta de outra empresa';
    else if (/salario|remunera|beneficio/.test(k)) o.q4 = 'Remuneração / benefícios';
    else if (/promoc|cresciment/.test(k)) o.q4 = 'Oportunidade de crescimento profissional';
    else if (/familia|pessoa/.test(k)) o.q4 = 'Questões pessoais';
    else if (/horario|escala|jornada/.test(k)) o.q4 = 'Jornada / escala de trabalho';
    else if (/desempenho|performance|adapt|norma|desistiu/.test(k)) o.q4 = 'Desempenho ou adaptação à função';
    else if (/conduta|indisciplina|comport|falta|atraso|insubord|veiculo/.test(k)) { o.q4 = 'Outros'; o.q4_outros = c.oqRH || c.causaRH; }
    if (o.q4) o.q3 = 'Sim';
    return o;
  }

  // ---------- quadrinhos + Pareto (mesmo visual do Turnover) ----------
  const chartsSd = {};
  function contarPor(lista, f) { const m = {}; lista.forEach(x => { const v = String(f(x) || '').trim() || 'Não informado'; m[v] = (m[v] || 0) + 1; }); return Object.entries(m).sort((a, b) => b[1] - a[1]); }
  function desenharKpisParetos(base) {
    const total = base.length;
    const porTipo = contarPor(base, tipo4);
    const casa = base.map(c => c.adm && c.dem ? dias(c.adm, c.dem) : null).filter(v => v !== null && v >= 0);
    const media = casa.length ? casa.reduce((a, b) => a + b, 0) / casa.length : null;
    const ate30 = casa.filter(v => v <= 30).length, ate90 = casa.filter(v => v <= 90).length;
    const causas = contarPor(base.filter(c => c.causaRH), c => c.causaRH);
    const feitos = base.filter(feito).length;
    const fmtCasa = v => v == null ? '—' : (v >= 60 ? (v / 30).toFixed(1).replace('.', ',') + ' meses' : Math.round(v) + ' dias');
    const PAL = ['#2563eb', '#d97706', '#059669', '#5e34b5', '#be185d'];
    const kpi = (lab, val, cor, sub) => `<div class="kpi"><div class="label">${lab}</div><div class="value" style="color:${cor}; ${String(val).length > 8 ? 'font-size:18px;' : ''}">${val}</div>${sub ? `<div class="delta">${sub}</div>` : ''}</div>`;
    document.getElementById('sdKpis').innerHTML =
      kpi('🔄 Desligamentos', total, '#0e7c86') +
      porTipo.slice(0, 3).map(([k, n], i) => kpi(esc(k), n, PAL[i], (total ? (n / total * 100).toFixed(1) : 0) + '% do total')).join('') +
      kpi('⏳ Tempo médio de casa', fmtCasa(media), '#1b2440', 'entre admissão e demissão') +
      kpi('🚨 Saíram em até 30 dias', ate30, ate30 ? '#dc2626' : '#059669', (total ? (ate30 / total * 100).toFixed(1) : 0) + '% do total') +
      kpi('⚠️ Saíram em até 90 dias', ate90, ate90 ? '#d97706' : '#059669', (total ? (ate90 / total * 100).toFixed(1) : 0) + '% do total') +
      (causas[0] ? kpi('🥇 Principal causa', esc(causas[0][0]), '#1b2440', causas[0][1] + ' (' + (causas[0][1] / Math.max(1, causas.reduce((a, x) => a + x[1], 0)) * 100).toFixed(1) + '%)') : '') +
      kpi('🗣️ Entrevistas feitas', feitos + ' de ' + total, feitos === total ? '#059669' : '#dc2626', total ? (feitos / total * 100).toFixed(0) + '% respondidas' : '');
    pareto('sdPar1', causas, 'Desligamentos');
    graficoTipo(base);
  }

  // feito = entrevista nova preenchida OU já classificado no modelo antigo (causa no Turnover)
  function feito(c) { return !!SAIDA[c.mat] || !!c.causaRH; }
  const TIPOS4 = [['Pedido de demissão', '#2563eb'], ['Demissão sem justa causa', '#d97706'], ['Demissão por justa causa', '#dc2626'], ['Abandono', '#5e34b5']];
  function tipo4(c) { const r = SAIDA[c.mat]; return (r && r.q1) || sugestaoRH(c).q1 || 'Não informado'; }
  function graficoTipo(base) {
    const cv = document.getElementById('sdTipo');
    if (!cv || typeof Chart === 'undefined') return;
    if (chartsSd.sdTipo) { chartsSd.sdTipo.destroy(); chartsSd.sdTipo = null; }
    const total = base.length || 1;
    const cats = TIPOS4.slice();
    const sem = base.filter(c => tipo4(c) === 'Não informado').length;
    if (sem) cats.push(['Não informado', '#94a3b8']);
    const vals = cats.map(([k]) => base.filter(c => tipo4(c) === k).length);
    const plugins = typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [];
    chartsSd.sdTipo = new Chart(cv, {
      type: 'bar',
      data: { labels: cats.map(c => c[0]), datasets: [{ label: 'Desligamentos', data: vals, backgroundColor: cats.map(c => c[1]), borderRadius: 4, barThickness: 34,
        datalabels: { anchor: 'end', align: 'right', color: '#1b2440', font: { size: 11, family: 'JetBrains Mono', weight: '700' }, formatter: v => v ? v + ' (' + Math.round(v / total * 100) + '%)' : '' } }] },
      plugins,
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { right: 70 } },
        plugins: { legend: { display: false }, datalabels: { display: true } },
        scales: { x: { beginAtZero: true, ticks: { color: '#1b2440', font: { weight: 'bold' } }, grid: { display: false } },
          y: { ticks: { color: '#1b2440', font: { size: 11, weight: 'bold' } }, grid: { display: false } } } }
    });
  }
  function pareto(id, itens, rot) {
    const cv = document.getElementById(id);
    if (!cv) return;
    if (chartsSd[id]) { chartsSd[id].destroy(); chartsSd[id] = null; }
    const wrap = cv.parentNode;
    let aviso = wrap.querySelector('.sd-vazio');
    if (!itens.length || typeof Chart === 'undefined') {
      cv.style.display = 'none';
      if (!aviso) { aviso = document.createElement('div'); aviso.className = 'sd-vazio mono'; aviso.style.cssText = 'font-size:12px; color:var(--muted); padding:20px 0;'; wrap.appendChild(aviso); }
      aviso.textContent = itens.length ? '' : 'Sem respostas ainda.';
      return;
    }
    cv.style.display = ''; if (aviso) aviso.remove();
    const total = itens.reduce((a, x) => a + x[1], 0);
    let acum = 0;
    const acumPct = itens.map(([, q]) => { acum += q; return acum / total * 100; });
    const corte = acumPct.findIndex(v => v >= 80);
    const plugins = typeof ChartDataLabels !== 'undefined' ? [ChartDataLabels] : [];
    chartsSd[id] = new Chart(cv, {
      data: { labels: itens.map(c => c[0]), datasets: [
        { type: 'bar', label: rot, data: itens.map(c => c[1]), order: 2, yAxisID: 'y', borderRadius: 4,
          backgroundColor: itens.map((_, i) => (corte === -1 || i <= corte) ? '#dc2626' : 'rgba(107,117,144,.6)'),
          datalabels: { anchor: 'end', align: 'end', color: '#1b2440', font: { size: 11, family: 'JetBrains Mono', weight: '700' } } },
        { type: 'line', label: '% Acumulado', data: acumPct, order: 1, yAxisID: 'y1', borderColor: '#0e7c86', backgroundColor: '#0e7c86', pointRadius: 4, tension: .2,
          datalabels: { color: '#0e7c86', align: 'top', offset: 6, font: { size: 9, family: 'JetBrains Mono', weight: '700' }, formatter: v => v.toFixed(0) + '%' } },
        { type: 'line', label: '80%', data: itens.map(() => 80), order: 1, yAxisID: 'y1', borderColor: '#000000', borderDash: [5, 5], pointRadius: 0, datalabels: { display: false } } ] },
      plugins,
      options: { responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#6b7590', font: { family: 'Inter', size: 11 } } }, datalabels: { display: true } },
        scales: { x: { ticks: { color: '#1b2440', font: { size: 10, weight: 'bold' }, maxRotation: 45, minRotation: 0 }, grid: { display: false } },
          y: { ticks: { color: '#1b2440', font: { weight: 'bold' } }, grid: { display: false }, beginAtZero: true },
          y1: { position: 'right', min: 0, max: 110, ticks: { color: '#1b2440', font: { weight: 'bold' }, callback: v => v <= 100 ? v + '%' : '' }, grid: { display: false } } } }
    });
  }

  function escolherOutro() {
    const fundo = document.createElement('div');
    fundo.className = 'av-modal-fundo';
    fundo.innerHTML = `<div class="av-modal" style="width:min(560px,100%);"><div style="padding:18px 22px;"><h3>Registrar outro desligamento</h3>
      <div style="font-size:12.5px; color:#374151; margin:6px 0 10px;">Para quem ainda não aparece como desligado no Organograma.</div>
      <select id="sdEsc">${ATIVOS_TEC.map(c => `<option value="${esc(c.mat)}">${esc(c.nome)} — ${esc(c.sup)}</option>`).join('')}</select></div>
      <div class="av-rodape"><button type="button" class="av-bt sec" data-x>Cancelar</button><button type="button" class="av-bt pri" data-ok>Continuar</button></div></div>`;
    document.body.appendChild(fundo);
    fundo.querySelector('[data-x]').addEventListener('click', () => fundo.remove());
    fundo.querySelector('[data-ok]').addEventListener('click', () => { const m = fundo.querySelector('#sdEsc').value; fundo.remove(); abrirSaida(m); });
  }

  function abrirSaida(mat) {
    const c = DESLIG.find(x => x.mat === mat) || ATIVOS_TEC.find(x => x.mat === mat);
    if (!c) return;
    const r = Object.assign({}, sugestaoRH(c), SAIDA[mat] || {});
    const contr = r.contrato || (/cirion/i.test(c.contrato) ? 'Cirion' : /claro/i.test(c.contrato) ? 'Claro' : (c.contrato ? 'Outros' : 'Claro'));
    const areaSel = r.areaProjeto || c.un || '';
    const areas = AREAS.indexOf(areaSel) === -1 && areaSel ? AREAS.concat([areaSel]) : AREAS;
    const campo = (id, rot, val, tipo) => `<div><span class="av-l">${rot}</span><input type="${tipo || 'text'}" id="${id}" value="${esc(val || '')}"></div>`;
    const pergunta = p => {
      const v = r[p.id];
      const marcado = o => p.multi ? (Array.isArray(v) && v.indexOf(o) !== -1) : v === o;
      return `<div style="margin:12px 0 4px; font-size:13.5px; font-weight:700;">${esc(p.txt)}${p.multi ? ' <span style="font-weight:400; color:#6b7280;">(pode marcar mais de uma)</span>' : ''}</div>
        <div style="display:flex; flex-direction:column; gap:4px;">${p.ops.map(o => `<label style="font-size:13px; cursor:pointer;"><input type="${p.multi ? 'checkbox' : 'radio'}" name="sd_${p.id}" value="${esc(o)}"${marcado(o) ? ' checked' : ''}> ${esc(o)}</label>`).join('')}</div>
        ${p.outros ? `<input type="text" id="sd_${p.id}_outros" placeholder="Se marcou Outros, descreva" value="${esc(r[p.id + '_outros'] || '')}" style="margin-top:6px;">` : ''}`;
    };
    const fundo = document.createElement('div');
    fundo.className = 'av-modal-fundo';
    fundo.innerHTML = `<div class="av-modal" role="dialog" aria-label="Entrevista de desligamento">
      <div style="padding:18px 22px; display:flex; justify-content:space-between; gap:12px; align-items:flex-start;">
        <div><h3>Entrevista de Desligamento</h3><div style="font-size:13px; color:#374151; margin-top:4px;"><b>${esc(c.nome)}</b> · matrícula ${esc(c.mat)}</div></div>
        <button type="button" class="av-bt sec" data-x>✕</button></div>
      ${c.tipoRH || c.causaRH ? `<div class="av-sec" style="background:#f8fafc; font-size:12.5px;"><b>Registro do RH (Turnover):</b> ${esc(c.tipoRH || '—')} · causa <b>${esc(c.causaRH || '—')}</b>${c.oqRH ? ' · ' + esc(c.oqRH) : ''}${SAIDA[mat] ? '' : ' <span style="color:#6b7280;">(as respostas 1 e 4 vieram sugeridas daqui — confira)</span>'}</div>` : ''}
      <div class="av-sec"><h4>1. Identificação do colaborador</h4>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
          ${campo('sdNome', 'Nome completo do colaborador', r.nome || c.nome)}
          ${campo('sdMat', 'Matrícula', r.matricula || c.mat)}
          <div><span class="av-l">Contrato</span><select id="sdContrato">${['Claro', 'Cirion', 'Outros'].map(o => `<option${o === contr ? ' selected' : ''}>${o}</option>`).join('')}</select></div>
          <div><span class="av-l">Área / Projeto</span><select id="sdArea"><option value=""></option>${areas.map(o => `<option${o === areaSel ? ' selected' : ''}>${esc(o)}</option>`).join('')}<option${areaSel === 'Outros' ? ' selected' : ''}>Outros</option></select></div>
          ${campo('sdCargo', 'Cargo', r.cargo || c.cargo)}
          ${campo('sdSupv', 'Supervisor responsável', r.supervisor || c.sup)}
          ${campo('sdCoord', 'Coordenador responsável', r.coordenador || c.coord)}
          <div></div>
          ${campo('sdAdm', 'Data de admissão', r.dataAdmissao || c.adm, 'date')}
          ${campo('sdDem', 'Data do desligamento', r.dataDesligamento || c.dem, 'date')}
        </div>
      </div>
      <div class="av-sec"><h4>2. Informações sobre o desligamento</h4>${PERGUNTAS_SAIDA.slice(0, 4).map(pergunta).join('')}</div>
      <div class="av-sec"><h4>3. Avaliação do supervisor</h4>${PERGUNTAS_SAIDA.slice(4).map(pergunta).join('')}</div>
      <div class="av-sec"><span class="av-l">Observações (opcional)</span><textarea id="sdObs" placeholder="Algo mais que o RH precise saber sobre essa saída.">${esc(r.observacoes || '')}</textarea>
        ${campo('sdPor', 'Preenchido por', r.preenchidoPor || c.sup)}</div>
      <div class="av-rodape"><span id="sdMsg" style="font-size:12.5px; margin-right:auto;"></span>
        <button type="button" class="av-bt sec" data-print>🖨 Imprimir</button>
        <button type="button" class="av-bt sec" data-x>Cancelar</button>
        <button type="button" class="av-bt pri" data-salvar>Salvar</button></div></div>`;
    document.body.appendChild(fundo);
    const $ = q => fundo.querySelector(q);
    const fechar = () => fundo.remove();
    fundo.querySelectorAll('[data-x]').forEach(b => b.addEventListener('click', fechar));
    fundo.addEventListener('click', e => { if (e.target === fundo) fechar(); });
    $('[data-print]').addEventListener('click', () => window.print());
    $('[data-salvar]').addEventListener('click', async () => {
      const o = { tipo: 'SAIDA', nome: $('#sdNome').value.trim(), matricula: $('#sdMat').value.trim() || c.mat, contrato: $('#sdContrato').value, areaProjeto: $('#sdArea').value,
        cargo: $('#sdCargo').value.trim(), supervisor: $('#sdSupv').value.trim(), coordenador: $('#sdCoord').value.trim(), dataAdmissao: $('#sdAdm').value, dataDesligamento: $('#sdDem').value,
        observacoes: $('#sdObs').value.trim(), preenchidoPor: $('#sdPor').value.trim(), dataPreenchimento: hojeIso() };
      PERGUNTAS_SAIDA.forEach(p => {
        const sel = [...fundo.querySelectorAll(`input[name="sd_${p.id}"]:checked`)].map(x => x.value);
        o[p.id] = p.multi ? sel : (sel[0] || '');
        if (p.outros) o[p.id + '_outros'] = ($('#sd_' + p.id + '_outros').value || '').trim();
      });
      const msg = $('#sdMsg'), bt = $('[data-salvar]');
      const falta = OBRIG_SAIDA.filter(id => !o[id]);
      if (falta.length) { msg.innerHTML = `<span style="color:#b91c1c;">Responda a(s) pergunta(s): ${falta.map(x => x.slice(1)).join(', ')}.</span>`; return; }
      if (!o.dataDesligamento) { msg.innerHTML = '<span style="color:#b91c1c;">Informe a data do desligamento.</span>'; return; }
      if (!FLUXO_AVALIACAO) { msg.innerHTML = '<span style="color:#b91c1c;">O envio ainda não está ligado (falta a URL do fluxo).</span>'; return; }
      bt.disabled = true; bt.textContent = 'Gravando…';
      try {
        const resp0 = await fetch(FLUXO_AVALIACAO, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(o) });
        const resp = lerResposta(await resp0.text());
        if (!resp0.ok) throw new Error('o fluxo respondeu ' + resp0.status);
        if (resp && resp.ok === false) throw new Error(/etapa/i.test(resp.erro || '') ? 'o Excel ainda está com o script antigo: cole a versão nova do “Avaliação Experiência” (com a Saída)' : (resp.erro || 'o script recusou a gravação'));
        o.__em = new Date().toISOString().replace('T', ' ').slice(0, 16);
        const k = matNum(o.matricula);
        SAIDA[k] = o; salvarLocalSaida(k, o);
        if (!DESLIG.some(x => x.mat === k)) DESLIG.push({ nome: o.nome, mat: k, sup: o.supervisor, coord: o.coordenador, contrato: o.contrato, un: o.areaProjeto, cargo: o.cargo, adm: o.dataAdmissao, dem: o.dataDesligamento });
        renderSaida(); fechar();
      } catch (e) {
        msg.innerHTML = `<span style="color:#b91c1c;">Não gravou: ${esc(e.message)}</span>`;
        bt.disabled = false; bt.textContent = 'Salvar';
      }
    });
  }

  function iniciarSaida() {
    montarSaida();
    renderSaida();
    if (!carregado) carregar().then(() => { renderSaida(); render(); });
  }

  // ---------- página própria no menu, logo abaixo de Organograma ----------
  function iniciar() {
    montar();
    render();
    if (!carregado) carregar().then(() => { render(); renderSaida(); });
  }
  const prep = () => montarPagina();
  // o script fica no fim do <body>: a página Organograma já existe → monta as abas já
  if (document.getElementById('pagina-organograma')) prep(); else document.addEventListener('DOMContentLoaded', prep);
})();
