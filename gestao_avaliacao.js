// gestao_avaliacao.js — Avaliação de Experiência (30 / 60 / 90 dias) na página Organograma
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
  const FLUXO_AVALIACAO = '';
  const ADMISSAO_DESDE = '2026-10-01';   // técnicos admitidos a partir desta data
  const CHAVE_LOCAL = 'avalExpLocais_v1';

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
        ['s6', 'Resolução de Problemas sob Pressão', 'Manteve a calma e resolveu a falha crítica no cliente Z minimizando o tempo de parada.'] ] } ] }
  ];
  const TODOS = SECOES.flatMap(s => s.grupos.flatMap(g => g.itens));
  const HARD = TODOS.filter(i => i[0][0] === 'h').map(i => i[0]);
  const SOFT = TODOS.filter(i => i[0][0] === 's').map(i => i[0]);
  const DECISOES = [
    ['APROVADO', '✅ Aprovado — segue na empresa', '#047857', 'rgba(4,120,87,.12)'],
    ['ATENCAO', '⚠️ Em atenção — acompanhar / plano de ação', '#b45309', 'rgba(180,83,9,.12)'],
    ['REPROVADO', '❌ Reprovado — não efetivar', '#b91c1c', 'rgba(185,28,28,.12)']
  ];
  const ETAPAS = ['30', '60', '90'];

  let COLABS = [], AVAL = {}, carregado = false;
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
      carregado = true;
    });
  }

  // ---------- estrutura ----------
  function montar() {
    const pg = document.getElementById('pagina-organograma');
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
    sec.className = 'panel'; sec.id = 'avalExpSecao'; sec.style.marginTop = '16px';
    sec.innerHTML = `
      <h2>🧭 Avaliação de Experiência — técnicos novos (30 / 60 / 90 dias)</h2>
      <div class="hint">Técnicos ativos admitidos a partir de ${br(ADMISSAO_DESDE)}. Clique na etapa para o supervisor preencher a avaliação; o RH acompanha aqui o que já foi avaliado e a decisão.</div>
      <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:10px 0 12px;">
        <select id="avSup" class="mono" style="font-size:12px; padding:6px 10px; border-radius:8px; border:1px solid var(--line); background:var(--panel); color:var(--text); min-width:220px;"><option value="">Supervisor: todos</option></select>
        <div class="search"><input id="avBusca" type="search" placeholder="Buscar técnico..." aria-label="Buscar técnico"></div>
        <span id="avResumo" class="mono" style="font-size:12px; color:var(--text); margin-left:auto;"></span>
      </div>
      <div style="overflow:auto; border-radius:8px;"><table id="avTabela"></table></div>
      ${FLUXO_AVALIACAO ? '' : '<div class="mono" style="font-size:12px; color:#b45309; margin-top:10px;">⚠ O envio ainda não está ligado: falta colar a URL do fluxo “Avaliação Experiência – Gravar” em FLUXO_AVALIACAO (gestao_avaliacao.js).</div>'}`;
    const ref = document.getElementById('headcountSecao');
    if (ref && ref.parentNode === pg) pg.insertBefore(sec, ref.nextSibling); else pg.appendChild(sec);
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
    tb.innerHTML = '<thead><tr><th>Técnico</th><th>Supervisor</th><th>Admissão</th><th>Dias de casa</th><th>30 dias</th><th>60 dias</th><th>90 dias</th><th>Última decisão</th></tr></thead><tbody>' +
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

  // ---------- liga na página Organograma ----------
  function iniciar() {
    montar();
    render();
    carregar().then(render);
  }
  const antes = window.mostrarPagina;
  if (typeof antes === 'function') {
    window.mostrarPagina = function (p) {
      const r = antes.apply(this, arguments);
      if (p === 'organograma') { if (!document.getElementById('avalExpSecao')) iniciar(); else if (!carregado) carregar().then(render); }
      return r;
    };
  }
  if (document.readyState !== 'loading') setTimeout(() => { if (document.getElementById('pagina-organograma') && document.getElementById('pagina-organograma').style.display !== 'none') iniciar(); }, 0);
})();
