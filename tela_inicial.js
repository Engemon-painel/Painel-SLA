// ======================================================================
// TELA INICIAL — Central de Operações (Painel do ESER)
// ----------------------------------------------------------------------
// Abre o painel numa tela de entrada com cards grandes (ESER, Organograma,
// Gestão de Vagas, Frotas, NR, Exames...). Clicou no card, vai para a página.
// Os cards são montados a partir do menu lateral: página nova no menu
// aparece aqui sozinha. No menu lateral ganha o item "🏠 Início".
// No index.html: <script src="tela_inicial.js"></script> (depois dos outros)
//
// ALTERADO (07/10/2026): a Gestão de Vagas também passou a pedir senha,
// igual à Gestão de Frotas (cada área libera separadamente).
// ALTERADO (18/10/2026): cards mais compactos (4 por linha em tela larga,
// ícones, textos e etiquetas menores).
// ======================================================================
(function(){

  // ---- NOME DO PAINEL: troque aqui se quiser outro ----
  const NOME_PAINEL = 'Central de Operações';
  const SUBTITULO   = 'Engemon OP Services · Gestão Operacional';

  // ---- TRAVAS POR SENHA (trava visual, não é segurança forte) ----
  // A senha fica guardada só como "impressão digital" (SHA-256). Para trocar a
  // senha, peça o código novo e substitua o valor abaixo.
  const FROTA_SENHA_SHA256 = '8b94527a1c8f5579868ed765d2a280015bdc1cfaeb5d79c58fb3a8d817aa5631';
  // Senha da Gestão de Vagas (por enquanto, a mesma da Frota).
  const VAGAS_SENHA_SHA256 = '8b94527a1c8f5579868ed765d2a280015bdc1cfaeb5d79c58fb3a8d817aa5631';

  // Repositório do GitHub de onde o painel é publicado (dono/nome). Vazio =
  // descobre sozinho pelo endereço (ex.: engemon-painel.github.io/Painel-SLA).
  // Usado para mostrar a hora em que o fluxo (script) atualizou cada arquivo.
  const REPO_GITHUB = '';

  // >>> COLE AQUI A LINHA "const BANNER = 'data:image/jpeg;base64,...';" DO SEU ARQUIVO ATUAL <<<
  const BANNER = 'COLE_AQUI_O_BANNER_BASE64';

  // Descrição curta de cada card (pelo id do item do menu). Página nova sem
  // descrição aqui aparece só com o nome.
  const DESCRICOES = {
    'nav-consolidado': 'Visão mensal, diária e semanal do SLA',
    'nav-area':        'Indicadores e metas por área (SPC1–4)',
    'nav-site':        'Cidades, top sites e incidentes',
    'nav-fca':         'Fora do prazo e justificativas',
    'nav-infra':       'Tipo de infra e unidade responsável',
    'nav-organograma': 'Headcount, admissões e equipe',
    'nav-preventiva':  'Realizado x meta das preventivas do mês',
    'nav-afaz':        'Backlog, entrantes e prioritários pendentes',
    'nav-vagas':       'Requisições e recrutamento',
    'nav-frota':       'Veículos, status e gastos',
    'nav-bancohoras':  'Banco de horas e extras',
    'nav-nr':          'Cursos NR dos técnicos',
    'nav-aso':         'Vencimento dos ASOs',
    'nav-escala':      'Técnicos por dia, horário e supervisão'
  };

  const CSS = `
    body.modo-inicio{ padding:0 !important; background:#001c39 !important; min-height:100vh; }
    body.modo-inicio > header, body.modo-inicio > .app-shell, body.modo-inicio > footer{ display:none !important; }
    #telaInicial{ display:none; min-height:100vh; color:#fff; font-family:'Inter', sans-serif;
      background: radial-gradient(1200px 500px at 85% 110%, rgba(40,110,170,.35), transparent 60%),
                  linear-gradient(180deg, #00233f 0%, #001c39 55%, #00152c 100%); }
    body.modo-inicio #telaInicial{ display:block; }
    #telaInicial .ti-banner{ position:relative; width:100%; aspect-ratio:1654/282; min-height:150px;
      background:url(${BANNER}) center/cover no-repeat; }
    #telaInicial .ti-banner::after{ content:''; position:absolute; inset:auto 0 0 0; height:22%;
      background:linear-gradient(180deg, transparent, #00233f); }
    #telaInicial .ti-corpo{ max-width:1100px; margin:0 auto; padding:10px clamp(18px,4vw,48px) 50px; position:relative; }
    #telaInicial .ti-topo{ display:flex; justify-content:space-between; align-items:flex-end; gap:16px; flex-wrap:wrap; margin-bottom:18px; }
    #telaInicial .ti-eyebrow{ font-family:'JetBrains Mono', monospace; font-size:11px; letter-spacing:.2em; text-transform:uppercase; color:#7fb8e6; margin-bottom:6px; }
    #telaInicial h1{ font-family:'Space Grotesk', sans-serif; font-size:clamp(24px,3.2vw,34px); font-weight:700; color:#fff; }
    #telaInicial .ti-status{ font-family:'JetBrains Mono', monospace; font-size:12px; color:#9fb7cc; }
    #telaInicial .ti-grade{ display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:14px; }
    @media (max-width:1000px){ #telaInicial .ti-grade{ grid-template-columns:repeat(3, minmax(0,1fr)); } }
    @media (max-width:760px){ #telaInicial .ti-grade{ grid-template-columns:repeat(2, minmax(0,1fr)); } }
    #telaInicial .ti-card{ position:relative; background:#fff; color:#1b2440; border-radius:14px; padding:16px 12px 14px;
      min-height:130px; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; gap:6px;
      cursor:pointer; border:0; font:inherit;
      box-shadow:0 6px 18px rgba(0,0,0,.3), 0 2px 0 rgba(255,255,255,.6) inset;
      transition:transform .18s ease, box-shadow .18s ease;
      animation: tiEntra .45s ease both; }
    #telaInicial .ti-card:hover, #telaInicial .ti-card:focus-visible{ transform:translateY(-4px);
      box-shadow:0 12px 26px rgba(0,0,0,.4); outline:none; }
    #telaInicial .ti-card:focus-visible{ box-shadow:0 0 0 3px #7fb8e6, 0 12px 26px rgba(0,0,0,.4); }
    #telaInicial .ti-icone{ font-size:26px; line-height:1; width:48px; height:48px; border-radius:12px;
      display:flex; align-items:center; justify-content:center; background:#eef3f8; }
    #telaInicial .ti-nome{ font-weight:700; font-size:13.5px; }
    #telaInicial .ti-desc{ font-size:10.5px; color:#6b7590; line-height:1.3; }
    #telaInicial .ti-data{ margin-top:4px; display:inline-flex; align-items:center; gap:5px; max-width:100%; white-space:nowrap;
      font-family:'JetBrains Mono', monospace; font-size:10px; color:#0e7c86; background:rgba(14,124,134,.08);
      padding:3px 8px; border-radius:999px; }
    #telaInicial .ti-data .rot{ font-family:'Inter', sans-serif; font-size:9.5px; color:#6b7590; font-weight:500; }
    #telaInicial .ti-data b{ font-weight:700; }
    #telaInicial .ti-data.velho{ color:#b45309; background:rgba(217,119,6,.12); }
    #telaInicial .ti-data.velho .rot{ color:#b45309; }
    #telaInicial .ti-badge{ position:absolute; top:8px; right:8px; background:#dc2626; color:#fff; font-family:'JetBrains Mono', monospace;
      font-size:10px; font-weight:700; padding:2px 7px; border-radius:999px; }
    #telaInicial .ti-card.ti-grupo{ background:linear-gradient(160deg,#ffffff 0%, #e6f2f3 100%); }
    #telaInicial .ti-card.ti-grupo .ti-icone{ background:#0e7c86; color:#fff; font-family:'Space Grotesk', sans-serif; font-size:15px; font-weight:700; letter-spacing:.04em; }
    #telaInicial .ti-voltar{ display:inline-flex; align-items:center; gap:8px; margin-bottom:18px; background:transparent; color:#cfe3f3;
      border:1px solid rgba(207,227,243,.35); border-radius:999px; padding:8px 16px; font-family:'JetBrains Mono', monospace; font-size:12px; cursor:pointer; }
    #telaInicial .ti-voltar:hover{ background:rgba(255,255,255,.08); }
    #telaInicial .ti-secao{ font-family:'Space Grotesk', sans-serif; font-size:18px; font-weight:600; color:#fff; margin:0 0 16px; }
    /* modo foco: ao entrar por um card, mostra só aquela área */
    body.foco-unico .sidebar{ display:none !important; }
    body.foco-unico #monthStrip, body.foco-unico #btnSla4h, body.foco-unico #btnSla6h{ display:none !important; }
    body.foco-eser .sidebar > .nav-item:not(#nav-inicio){ display:none !important; }
    body.foco-eser #navGroupEser.collapsed .nav-group-items{ max-height:400px; opacity:1; }
    #tiFoco{ display:none; align-items:center; gap:12px; margin-bottom:16px; }
    body.foco-unico #tiFoco{ display:flex; }
    #tiFoco button{ display:inline-flex; align-items:center; gap:6px; font-family:'JetBrains Mono', monospace; font-size:12px; font-weight:600;
      padding:8px 14px; border-radius:999px; border:1px solid var(--line); background:var(--panel); color:var(--text); cursor:pointer; }
    #tiFoco button:hover{ border-color:var(--accent); color:var(--accent); }
    #tiFoco .ti-foco-nome{ font-family:'Space Grotesk', sans-serif; font-size:20px; font-weight:700; color:var(--text); }
    .trava-area{ max-width:420px; margin:40px auto; background:var(--panel); border:1px solid var(--line); border-radius:16px;
      padding:28px 26px; text-align:center; box-shadow:0 6px 20px rgba(27,36,64,.08); }
    .trava-area .tf-icone{ font-size:40px; margin-bottom:8px; }
    .trava-area h2{ font-family:'Space Grotesk', sans-serif; font-size:18px; margin-bottom:6px; }
    .trava-area p{ font-size:13px; color:var(--muted); margin-bottom:16px; }
    .trava-area input{ width:100%; font-size:14px; padding:10px 12px; border-radius:10px; border:1px solid var(--line); margin-bottom:10px; }
    .trava-area input:focus{ outline:2px solid var(--accent); border-color:transparent; }
    .trava-area button{ width:100%; font-family:'JetBrains Mono', monospace; font-weight:700; font-size:13px; padding:10px; border-radius:999px;
      border:0; background:var(--accent); color:#fff; cursor:pointer; }
    .trava-area .tf-erro{ color:var(--bad); font-size:12px; min-height:16px; margin-top:8px; }
    @keyframes tiEntra{ from{ opacity:0; transform:translateY(14px); } to{ opacity:1; transform:none; } }
    @media (max-width:600px){ #telaInicial .ti-grade{ grid-template-columns:repeat(2, 1fr); gap:10px; } #telaInicial .ti-card{ min-height:110px; } }
  `;

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  function lerItem(nav){
    const icone = (nav.querySelector('.nav-icon')?.textContent || '').trim();
    const badge = (nav.querySelector('#fcaCountNav')?.textContent || '').replace(/[()]/g,'').trim();
    const clone = nav.cloneNode(true);
    clone.querySelectorAll('.nav-icon, #fcaCountNav').forEach(e => e.remove());
    return { id: nav.id, icone, nome: clone.textContent.trim(), badge, desc: DESCRICOES[nav.id] || '' };
  }

  function itensDoMenu(){
    const sidebar = document.querySelector('.sidebar');
    if(!sidebar) return { grupo: [], soltos: [] };
    const grupo = Array.from(sidebar.querySelectorAll('#navGroupEserItems .nav-item')).map(lerItem);
    const soltos = Array.from(sidebar.children)
      .filter(el => el.classList.contains('nav-item') && el.id && el.id !== 'nav-inicio')
      .map(lerItem);
    return { grupo, soltos };
  }

  // ---------------- data da última atualização de cada área ----------------
  const pad2 = n => String(n).padStart(2, '0');
  const isoParaBr = iso => { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + '/' + m[2] + '/' + m[1] : ''; };
  const brParaIso = br => { const m = String(br || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : ''; };
  function hojeIso(){ const d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  const maiorAteHoje = lista => lista.filter(x => x && x <= hojeIso()).sort().pop() || '';
  let datasNrAso = { nr:'', aso:'' };

  // ---- hora exata em que o script/fluxo atualizou os arquivos (último commit no GitHub)
  let HORARIOS = {};
  function repoAtual(){
    if(REPO_GITHUB) return REPO_GITHUB;
    const h = location.hostname, partes = location.pathname.split('/').filter(Boolean);
    if(!/\.github\.io$/i.test(h) || !partes.length) return '';
    return h.split('.')[0] + '/' + partes[0];
  }
  function fmtDataHora(iso){
    const d = new Date(iso); if(isNaN(d)) return '';
    return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear() + ' às ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }
  function buscarHorarios(){
    const repo = repoAtual(); if(!repo) return;
    const CHAVE = 'tiHorarios2:' + repo, VALIDADE = 10 * 60 * 1000;
    try{
      const c = JSON.parse(sessionStorage.getItem(CHAVE) || 'null');
      if(c && Date.now() - c.em < VALIDADE){ HORARIOS = c.h; return; }
    }catch(e){}
    const arquivos = ['dados.json', 'base_nr.json', 'dados_escala.json', 'dados_preventiva.json'];
    Promise.all(arquivos.map(a =>
      fetch('https://api.github.com/repos/' + repo + '/commits?per_page=1&path=' + encodeURIComponent(a))
        .then(r => r.ok ? r.json() : [])
        .then(j => (j && j[0] && j[0].commit && j[0].commit.committer && j[0].commit.committer.date) || '')
        .catch(() => '')
    )).then(datas => {
      const h = {}; arquivos.forEach((a, i) => { if(datas[i]) h[a] = datas[i]; });
      if(!Object.keys(h).length) return;
      HORARIOS = h;
      try{ sessionStorage.setItem(CHAVE, JSON.stringify({ em: Date.now(), h })); }catch(e){}
      if(document.body.classList.contains('modo-inicio')) desenhar();
    });
  }
  // ---- rótulo de data de cada card: { rot, val, dica, velho }
  const dmCurto = iso => iso ? iso.slice(8,10) + '/' + iso.slice(5,7) : '';
  const diasAtras = iso => { if(!iso) return 0; const p = iso.split('-'); return Math.floor((new Date(hojeIso() + 'T00:00:00') - new Date(+p[0], +p[1]-1, +p[2])) / 86400000); };
  const LIMITE_DIAS = 2;   // mais velho que isso = etiqueta laranja
  function etiqueta(rot, dataIso, hora, dicaExtra){
    if(!dataIso) return null;
    const dias = diasAtras(dataIso);
    const velho = dias > LIMITE_DIAS;
    const val = dmCurto(dataIso) + (hora ? ' às ' + hora : '');
    let dica = rot + ' ' + isoParaBr(dataIso) + (hora ? ' às ' + hora : '') + (dicaExtra ? ' · ' + dicaExtra : '');
    if(velho) dica += ' · sem atualização há ' + dias + ' dias — confira o fluxo';
    return { rot, val, dica, velho };
  }
  function doCommit(arq){
    if(!HORARIOS[arq]) return null;
    const d = new Date(HORARIOS[arq]); if(isNaN(d)) return null;
    const iso = d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    return etiqueta('Atualizado', iso, pad2(d.getHours()) + ':' + pad2(d.getMinutes()), 'horário em que o script rodou');
  }

  function datasDasAreas(){
    const txt = document.getElementById('loadStatus')?.textContent || '';
    const geral = (txt.match(/atualizados em (\d{4}-\d{2}-\d{2})/) || [])[1] || '';
    const baseAte = brParaIso((txt.match(/base até (\d{2}\/\d{2}\/\d{4})/) || [])[1]);
    const atualizado = etiqueta('Atualizado', geral);

    const colab = (typeof COLABORADORES !== 'undefined' && Array.isArray(COLABORADORES)) ? COLABORADORES : [];
    const movimento = maiorAteHoje(colab.flatMap(c => [c.dataAdmis, c.dataDemissao]));

    const eserEt = etiqueta('Incidentes até', baseAte);
    if(eserEt){ eserEt.velho = false; eserEt.dica = 'Incidentes até ' + isoParaBr(baseAte); }   // sem aviso de atraso no ESER
    const eser = eserEt || atualizado;
    const escala = window.EscalaInfo && window.EscalaInfo.atualizado_em;
    return {
      eser,
      'nav-consolidado': eser, 'nav-area': eser, 'nav-site': eser, 'nav-fca': eser, 'nav-infra': eser,
      'nav-organograma': movimento ? { rot:'Última movimentação', val: dmCurto(movimento), dica:'Última admissão/demissão em ' + isoParaBr(movimento), velho:false } : atualizado,
      'nav-vagas': doCommit('dados.json') || atualizado,
      'nav-frota': doCommit('dados.json') || atualizado,
      'nav-bancohoras': atualizado,
      'nav-nr': doCommit('base_nr.json') || etiqueta('Atualizado', datasNrAso.nr),
      'nav-aso': doCommit('base_nr.json') || etiqueta('Atualizado', datasNrAso.aso),
      'nav-escala': doCommit('dados_escala.json') || etiqueta('Atualizado', escala),
      'nav-preventiva': doCommit('dados_preventiva.json'),
      'nav-afaz': doCommit('dados_afaz.json')
    };
  }
  function buscarDatasNrAso(){
    if(!window.NRConversor) return;
    NRConversor.carregar().then(d => {
      const novo = { nr: (d.nr && d.nr.atualizado_em) || '', aso: (d.aso && d.aso.atualizado_em) || '' };
      if(novo.nr !== datasNrAso.nr || novo.aso !== datasNrAso.aso){
        datasNrAso = novo;
        if(document.body.classList.contains('modo-inicio')) desenhar();
      }
    }).catch(() => {});
  }
  let DATAS = {};
  const linhaData = e => e ? `<span class="ti-data${e.velho ? ' velho' : ''}" title="${esc(e.dica)}">${e.velho ? '⚠️' : '🕒'} <span class="rot">${esc(e.rot)}</span> <b>${esc(e.val)}</b></span>` : '';

  function cardHtml(it, i){
    return `<button type="button" class="ti-card" data-nav="${esc(it.id)}" style="animation-delay:${i*45}ms">
      ${it.badge ? `<span class="ti-badge" title="Pendentes">${esc(it.badge)}</span>` : ''}
      ${TRAVAS[it.id] && !liberada(TRAVAS[it.id].chave) ? '<span class="ti-badge" style="background:#374151;" title="Precisa de senha">🔒</span>' : ''}
      <span class="ti-icone">${esc(it.icone || '•')}</span>
      <span class="ti-nome">${esc(it.nome)}</span>
      ${it.desc ? `<span class="ti-desc">${esc(it.desc)}</span>` : ''}
      ${linhaData(DATAS[it.id])}
    </button>`;
  }

  let vista = 'principal';

  // ---------------- travas por senha (Frota e Gestão de Vagas) ----------------
  // Cada área: página, seção que fica escondida, título, chave da sessão,
  // senha (SHA-256) e função que redesenha a página depois de liberar.
  const TRAVAS = {
    'nav-frota': { pagina:'pagina-frota', secao:'frotaSecao', titulo:'Gestão de Frotas',
      texto:'os dados da frota', chave:'frotaLiberada', hash: FROTA_SENHA_SHA256, render:'renderFrota' },
    'nav-vagas': { pagina:'pagina-vagas', secao:'vagasSecao', titulo:'Gestão de Vagas',
      texto:'as requisições e vagas', chave:'vagasLiberada', hash: VAGAS_SENHA_SHA256, render:'renderVagas' }
  };
  const liberada = chave => { try{ return sessionStorage.getItem(chave) === '1'; }catch(e){ return false; } };
  async function sha256(txt){
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(txt));
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
  }
  function aplicarTrava(navId){
    const t = TRAVAS[navId]; if(!t) return;
    const pg = document.getElementById(t.pagina);
    const secao = document.getElementById(t.secao);
    if(!pg || !secao) return;
    const idTrava = 'trava-' + t.pagina;
    let trava = document.getElementById(idTrava);
    if(liberada(t.chave)){ secao.style.display = ''; if(trava) trava.remove(); return; }
    secao.style.display = 'none';
    if(trava) return;
    trava = document.createElement('div');
    trava.id = idTrava;
    trava.className = 'trava-area';
    trava.innerHTML = `<div class="tf-icone">🔒</div>
      <h2>${esc(t.titulo)}</h2>
      <p>Área restrita. Digite a senha para ver ${esc(t.texto)}.</p>
      <input type="password" class="tf-senha" placeholder="Senha" autocomplete="off">
      <button type="button" class="tf-entrar">Entrar</button>
      <div class="tf-erro"></div>`;
    pg.appendChild(trava);
    const campo = trava.querySelector('.tf-senha');
    const entrar = async () => {
      if(campo.value && await sha256(campo.value) === t.hash){
        try{ sessionStorage.setItem(t.chave, '1'); }catch(e){}
        aplicarTrava(navId);
        if(typeof window[t.render] === 'function') window[t.render]();
      } else {
        trava.querySelector('.tf-erro').textContent = 'Senha incorreta.';
        campo.select();
      }
    };
    trava.querySelector('.tf-entrar').onclick = entrar;
    campo.onkeydown = e => { if(e.key === 'Enter') entrar(); };
    setTimeout(() => campo.focus(), 50);
  }

  function desenhar(){
    const alvo = document.getElementById('tiConteudo');
    if(!alvo) return;
    const { grupo, soltos } = itensDoMenu();
    DATAS = datasDasAreas();
    const status = document.getElementById('loadStatus');
    document.getElementById('tiStatus').textContent = status ? status.textContent : '';

    if(vista === 'eser'){
      alvo.innerHTML = `<button type="button" class="ti-voltar" id="tiVoltar">← Voltar</button>
        <div class="ti-secao">ESER — escolha a visão</div>
        <div class="ti-grade">${grupo.map(cardHtml).join('')}</div>`;
      document.getElementById('tiVoltar').onclick = () => { vista = 'principal'; desenhar(); };
    } else {
      const badgeEser = grupo.map(g => g.badge).filter(Boolean)[0] || '';
      const cardEser = grupo.length ? `<button type="button" class="ti-card ti-grupo" data-grupo="eser">
          ${badgeEser ? `<span class="ti-badge" title="FCA pendentes">${esc(badgeEser)}</span>` : ''}
          <span class="ti-icone">ESER</span>
          <span class="ti-nome">ESER</span>
          <span class="ti-desc">${esc(grupo.map(g => g.nome).join(' · '))}</span>
          ${linhaData(DATAS.eser)}
        </button>` : '';
      alvo.innerHTML = `<div class="ti-grade">${cardEser}${soltos.map((it, i) => cardHtml(it, i + 1)).join('')}</div>`;
    }

    alvo.querySelectorAll('.ti-card').forEach(card => {
      card.onclick = () => {
        if(card.dataset.grupo === 'eser'){ vista = 'eser'; desenhar(); return; }
        irPara(card.dataset.nav);
      };
    });
  }

  function mostrarInicio(){
    vista = 'principal';
    document.body.classList.remove('foco-unico', 'foco-eser');
    document.body.classList.add('modo-inicio');
    desenhar();
    buscarDatasNrAso();
    buscarHorarios();
    window.scrollTo(0, 0);
  }

  function irPara(navId){
    document.body.classList.remove('modo-inicio');
    // abre o grupo ESER no menu se a página for dele
    const nav = document.getElementById(navId);
    const grupo = nav && nav.closest('.nav-group');
    if(grupo) grupo.classList.remove('collapsed');
    // ESER: menu lateral só com as visões do ESER · outras áreas: sem menu, só a página
    document.body.classList.toggle('foco-eser', !!grupo);
    document.body.classList.toggle('foco-unico', !grupo);
    const barra = document.getElementById('tiFoco');
    if(barra && nav){
      const it = lerItem(nav);
      barra.querySelector('.ti-foco-nome').textContent = (it.icone ? it.icone + '  ' : '') + it.nome;
    }
    if(nav) nav.click();
    window.scrollTo(0, 0);
  }

  function montar(){
    if(document.getElementById('telaInicial')) return;
    const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);

    const tela = document.createElement('div');
    tela.id = 'telaInicial';
    tela.innerHTML = `
      <div class="ti-banner" role="img" aria-label="Engemon OP Services"></div>
      <div class="ti-corpo">
        <div class="ti-topo">
          <div>
            <div class="ti-eyebrow">${esc(SUBTITULO)}</div>
            <h1>${esc(NOME_PAINEL)}</h1>
          </div>
          <div class="ti-status" id="tiStatus"></div>
        </div>
        <div id="tiConteudo"></div>
      </div>`;
    document.body.insertBefore(tela, document.body.firstChild);

    // item "Início" no topo do menu lateral
    const sidebar = document.querySelector('.sidebar');
    if(sidebar && !document.getElementById('nav-inicio')){
      const nav = document.createElement('div');
      nav.className = 'nav-item'; nav.id = 'nav-inicio';
      nav.innerHTML = '<span class="nav-icon">🏠</span> Início';
      nav.onclick = mostrarInicio;
      sidebar.insertBefore(nav, sidebar.firstChild);
    }
    // barra "← Início" no topo das páginas abertas pelos cards
    const main = document.querySelector('main.main-content');
    if(main && !document.getElementById('tiFoco')){
      const barra = document.createElement('div');
      barra.id = 'tiFoco';
      barra.innerHTML = '<button type="button">← Início</button><span class="ti-foco-nome"></span>';
      barra.querySelector('button').onclick = mostrarInicio;
      main.insertBefore(barra, main.firstChild);
    }

    // nome do painel no cabeçalho e na aba do navegador
    const eyebrowTopo = document.querySelector('header .brand-eyebrow');
    if(eyebrowTopo) eyebrowTopo.textContent = SUBTITULO;
    const h1Topo = document.querySelector('header h1');
    if(h1Topo) h1Topo.textContent = NOME_PAINEL;
    document.title = NOME_PAINEL + ' · Engemon OP Services';

    // o título do cabeçalho também volta para o início
    const h1 = document.querySelector('header h1');
    if(h1){ h1.style.cursor = 'pointer'; h1.title = 'Voltar para o início'; h1.onclick = mostrarInicio; }

    // atualiza o texto de "dados atualizados" e o contador do FCA quando os dados recarregam
    const status = document.getElementById('loadStatus');
    if(status) new MutationObserver(() => { if(document.body.classList.contains('modo-inicio')) desenhar(); })
      .observe(status, { childList:true, characterData:true, subtree:true });
    const badge = document.getElementById('fcaCountNav');
    if(badge) new MutationObserver(() => { if(document.body.classList.contains('modo-inicio')) desenhar(); })
      .observe(badge, { childList:true, characterData:true, subtree:true });

    // travas da Frota e da Gestão de Vagas: valem para qualquer caminho (card, menu ou atalho)
    const mostrarAntes = window.mostrarPagina;
    window.mostrarPagina = function(nome){
      mostrarAntes(nome);
      if(nome === 'frota') aplicarTrava('nav-frota');
      if(nome === 'vagas') aplicarTrava('nav-vagas');
    };
    aplicarTrava('nav-frota');
    aplicarTrava('nav-vagas');

    // outras páginas avisam quando terminam de carregar os dados (ex.: escala)
    window.addEventListener('painel-dados', () => { if(document.body.classList.contains('modo-inicio')) desenhar(); });

    mostrarInicio();
  }

  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', montar);
  else montar();
})();
