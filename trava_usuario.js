// ======================================================================
// TRAVA COM USUÁRIO E SENHA — Banco de Horas
// ----------------------------------------------------------------------
// A página do Banco de Horas só aparece depois de entrar com usuário e senha.
// Vale até fechar a aba do navegador (ou clicar em "Sair").
// No index.html: <script src="trava_usuario.js?v=1"></script> (depois do trava_senha.js)
//
// COMO CADASTRAR UM USUÁRIO
//   1) Abra o painel, aperte F12 → aba Console
//   2) Digite:  gerarUsuarioBH('nome@engemon.com.br', 'senha')   e aperte Enter
//   3) Copie a linha que aparecer e cole dentro de USUARIOS, abaixo
//   4) Suba o arquivo no GitHub (troque o ?v= no index.html)
// Para tirar alguém, apague a linha dele.
//
// IMPORTANTE: é uma trava visual. O painel fica no GitHub Pages e os arquivos
// de dados (dados_bh.json, dados_pend_bh.json) continuam públicos para quem
// souber o endereço. A senha fica guardada só como "impressão digital" (SHA-256).
// ======================================================================
(function(){

  // ---- USUÁRIOS (usuário: impressão digital de "usuario:senha") ----
  // ALTERADO (03/11/2026): usuário = e-mail. A chave é a impressão digital do e-mail (o e-mail não fica exposto)
  // e o valor é a impressão digital de "email:senha". O comentário no fim da linha é só para saber de quem é.
  const USUARIOS = {
    '00edc0c94fbeaf0e952dbc7082264b8f151eec82591bd711806fe63b9cad884a': '3270354c7676e105f2dc9e300277d39ade3b6e6e6859b798ec812fe0fe071615',   // rogerio.bandeira
    '04888ea54b4ee1a4a85cc9ff1b00b60e1d5c3d55267e78b6949f10e669678329': 'ff1a8bf245985434e30041a6651c26fba9617bbd97e163e8c2e5bb1b9f347bb4',   // marcelo.rosa
    '3ed841342ea8194c0b3a6157b77844d70f98057f14426f6c7fb469ceb0eb5079': '3c1eb6a1655d0c8409acd8fa6664a14276278086052036daea0f8480f40f408b',   // clayton.lima
    '316dcfdcd45e362457a24c917586fb4b44829fce3d22df762b3a92a4f28ab4d5': '912f933526316d38bee164908bff297a0ab2df34b0c9e0b951c9c57eb4eeeba6',   // ycaro.souza
    '7bac18468f21fb16618b95c5c1922a61c115847bc79a5ba0c4d7d666850568ff': '9764cf080e1d953f21013f4166d01f3b582e7077ec8575c78ba6107e9b44b0e6',   // thiago.patricio
    'c63b6330278a74bcb611eca8ef2910afce816ac5344b5620cffd7bcd24076366': '53936395dbcba83cb370f4bba6cc487c85f3416a455932d8e2c63ddc3c16b33b',   // moacir.mello
    '1fa53c822a70d7ec5b7b9b829fe791da0b477c29a0e8261c0c2c6139303383fe': '536bf93425513467e48032df6c135a7d84e989aa0b80ad9815895ced239e6337',   // rodolfo.alonso
    '4e2f36fb07e7a1197282e2dd94bd19a3851a0e0807fafd0b9a04c3d277a50305': 'ec315b57acce557113ecebc3204e2dae790fa98f71c13af5edb08d49542687c7',   // alexandre.menezes
    '7d506cfb4758ff8f885e381cd7f6d6271643d697be6892551f1e46539224b1cb': '4e7ac637bb6f14857a86eb720ae26e133c175780bec5fa888e9fb02ed36aef57',   // damares.penillo
    '90e3df3bc093ff83d9fc97df5e361941b53d209fb8568513f4041b765c739e1f': '10e1e39c45816b7cff382ac639f747837a6a3dec91eb478fd9db6daa80766692',   // marcia.nascimento
    '754a73ade635c55be72a7ab3e41923c3dbc4b70051fc94b84562e8f0a7f19a6f': '48fea2ce2cbd13ce52244f87e87db680f2a64b3f918b70c7def8964501a0869f',   // jose.lemos
    '0c07470cf164de5b90928e350da66e67a349b8bc559aeab8e82bc2bfa022857b': '3171c81174381eef5e83066f5848836ba30035ca5215b5595351c8c2b10d8bab',   // cleiton.cardoso
    'e02c522356e6f2f5c0b15265d510156408eeace43ae128c5ae211b1a8935868a': 'fcac78ff1af6f64a9db587d84716a166f2536c716eb6cd0f6951b0cf44e77f90',   // nayane.hilario
    'f8054a0b500d6cff0fb2846b1aec2ce38c9c99ce9167448dffbbfce5f2ab8df4': 'afd0caf2e2786d876e8cb7730bf702e1b93b4e0cd74b04df603f31d2cee07b1e',   // keven.felix
    '7ded7755277227cf25b40489bb477bd6af9868a7bd9018274cdb2136896a7317': '863350986f5e78a4f49ebbb364e98722999a9005b57ad5422958a59923fd4b5b',   // karen.lewis
    '6388c426cb0022ac9f920873009432596cc4764ad30f715a5cfcd61515f98056': 'be523d8cf9264d2bd32a287d9f1d74938eeaf7515a947e77a92d4dbca4359405',   // esaie.geneste
    '00229abc250349f0b207ab05a30c9840f761bf4a6cc09daaef78b7c362951a4c': 'a41aabd31643143e829c48a563155198d58c8a0b366bd1f81097746c9a6bfdff',   // alipio.cardoso
    'ca102ef8c6f2e6b9b176d67dcef66b789c0ab3c400f83a9ecdd8287947d723b1': 'c4b6cbb7ec1cf170e747e7c5f9af9d61c3ea0bedd7d6017abce8928bd0236724',   // anna.vieira
    '29ce113e461a42134b126707387ed6a7cd6f89aab0ad7e649a0b2294eb115014': '1842bf97abf3e646d4e4d48944ea40d9345f07d5d5f80f55f48174273a445e13',   // guilherme.rocha
    '225ee3c551cf4124ea4e064c20634e7c08ab21f6d7e907fbd9c9e85be6d73a4d': '11760f8ce5c532d3deb96b85571904d8a42c7aade83eef577bd93648a42222f1',   // ewerton.silva
    'c2473941adddc037ac0b0ddcae8ca5f25c9360f47bc0e400d25ec3958fe3851f': '15c8c0a0259cff6f2550ce3b84209a0a2cbf5cbe6438b0466a22396dc16fa424',   // marcelo.remorino
  };

  const PAGINA = 'pagina-bancohoras';
  const NAV = 'nav-bancohoras';
  const TITULO = 'Banco de Horas';
  const CHAVE = 'bhUsuarioLogado';
  // ALTERADO (03/11/2026): DERRUBAR TODO MUNDO — o painel lê o sessao_bh.json a cada minuto.
  // Para tirar todos que estão na página, troque o número de "versao" nesse arquivo (ex.: 1 → 2) e suba no GitHub:
  // em até 1 minuto todos saem e precisam entrar de novo com e-mail e senha.
  const ARQ_SESSAO = 'sessao_bh.json';
  let versaoSessao = null;   // versão lida do sessao_bh.json

  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const normUser = u => String(u || '').trim().toLowerCase();
  async function sha256(txt){
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(txt));
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
  }
  const logado = () => { try{ const u = sessionStorage.getItem(CHAVE) || ''; const k = sessionStorage.getItem(CHAVE + 'K') || ''; return u && USUARIOS[k] ? u : ''; }catch(e){ return ''; } };

  // para cadastrar: gerarUsuarioBH('usuario', 'senha') no Console (F12)
  window.gerarUsuarioBH = async (usuario, senha) => {
    const u = normUser(usuario);
    if(!u || !senha) return 'Use: gerarUsuarioBH("usuario", "senha")';
    const linha = `    '${await sha256(u)}': '${await sha256(u + ':' + senha)}',   // ${u.split('@')[0]}`;
    console.log('Cole esta linha dentro de USUARIOS no trava_usuario.js:\n' + linha);
    return linha;
  };
  window.usuarioBH = logado;   // quem está logado (para outros módulos usarem)

  function estilo(){
    if(document.getElementById('estiloTravaUsuario')) return;
    const st = document.createElement('style'); st.id = 'estiloTravaUsuario';
    st.textContent = `
      .trava-usuario{ max-width:400px; margin:40px auto; background:var(--panel,#fff); border:1px solid var(--line,#e1e5f0); border-radius:16px;
        padding:28px 26px; text-align:center; box-shadow:0 6px 20px rgba(27,36,64,.08); }
      .trava-usuario .tu-icone{ font-size:40px; margin-bottom:8px; }
      .trava-usuario h2{ font-family:'Space Grotesk', sans-serif; font-size:18px; margin-bottom:6px; }
      .trava-usuario p{ font-size:13px; color:var(--muted,#6b7590); margin-bottom:16px; }
      .trava-usuario input{ width:100%; font-size:14px; padding:10px 12px; border-radius:10px; border:1px solid var(--line,#e1e5f0); margin-bottom:10px; box-sizing:border-box; }
      .trava-usuario input:focus{ outline:2px solid var(--accent,#0e7c86); border-color:transparent; }
      .trava-usuario button{ width:100%; font-family:'JetBrains Mono', monospace; font-weight:700; font-size:13px; padding:10px; border-radius:999px;
        border:0; background:var(--accent,#0e7c86); color:#fff; cursor:pointer; }
      .trava-usuario .tu-erro{ color:var(--bad,#dc2626); font-size:12px; min-height:16px; margin-top:8px; }
      .tu-sessao{ display:flex; justify-content:flex-end; align-items:center; gap:10px; font-size:12px; color:var(--muted,#6b7590); margin:0 0 8px; }
      .tu-sessao button{ border:1px solid var(--line,#e1e5f0); background:var(--panel,#fff); border-radius:999px; padding:4px 12px; font-size:12px; cursor:pointer; }`;
    document.head.appendChild(st);
  }

  let observador = null;
  function aplicar(){
    const pg = document.getElementById(PAGINA);
    if(!pg) return false;
    estilo();
    const esconder = () => Array.from(pg.children).forEach(el => {
      if(el.classList.contains('trava-usuario')) return;
      if(el.style.display !== 'none'){ el.style.display = 'none'; el.dataset.tuOculto = '1'; }
    });
    let trava = pg.querySelector(':scope > .trava-usuario');
    let sessao = pg.querySelector(':scope > .tu-sessao');
    const u = logado();

    if(u){
      if(observador){ observador.disconnect(); observador = null; }
      Array.from(pg.children).forEach(el => { if(el.dataset.tuOculto){ el.style.display = ''; delete el.dataset.tuOculto; } });
      if(trava) trava.remove();
      if(!sessao){
        sessao = document.createElement('div');
        sessao.className = 'tu-sessao';
        pg.prepend(sessao);
      }
      sessao.innerHTML = `<span>👤 ${esc(u)}</span><button type="button">Sair</button>`;
      sessao.querySelector('button').onclick = () => { try{ sessionStorage.removeItem(CHAVE); sessionStorage.removeItem(CHAVE + 'K'); }catch(e){} aplicar(); marcarCard(); };
      if(typeof window.bhResize === 'function') setTimeout(() => window.bhResize(), 50);
      return true;
    }

    if(sessao) sessao.remove();
    esconder();
    if(!observador){   // se a página montar conteúdo depois, esconde também
      observador = new MutationObserver(() => { if(!logado()) esconder(); });
      observador.observe(pg, { childList: true });
    }
    if(trava) return true;

    trava = document.createElement('div');
    trava.className = 'trava-usuario';
    trava.innerHTML = `<div class="tu-icone">🔒</div>
      <h2>${esc(TITULO)}</h2>
      <p>Área restrita. Entre com seu e-mail e senha.</p>
      <input type="text" class="tu-user" placeholder="E-mail" autocomplete="username">
      <input type="password" class="tu-senha" placeholder="Senha" autocomplete="current-password">
      <button type="button">Entrar</button>
      <div class="tu-erro"></div>`;
    pg.appendChild(trava);
    const iu = trava.querySelector('.tu-user'), is = trava.querySelector('.tu-senha'), erro = trava.querySelector('.tu-erro');
    const entrar = async () => {
      const usr = normUser(iu.value);
      const kUsr = usr ? await sha256(usr) : '';
      if(usr && is.value && USUARIOS[kUsr] && await sha256(usr + ':' + is.value) === USUARIOS[kUsr]){
        try{ sessionStorage.setItem(CHAVE, usr); sessionStorage.setItem(CHAVE + 'K', kUsr); sessionStorage.setItem(CHAVE + 'V', String(versaoSessao ?? '')); }catch(e){}
        is.value = '';
        aplicar(); marcarCard();
        window.dispatchEvent(new Event('resize'));
      } else {
        erro.textContent = 'E-mail ou senha incorretos.';
        is.select();
      }
    };
    trava.querySelector('button').onclick = entrar;
    iu.onkeydown = e => { if(e.key === 'Enter') is.focus(); };
    is.onkeydown = e => { if(e.key === 'Enter') entrar(); };
    return true;
  }

  // cadeado no card da tela inicial
  function marcarCard(){
    document.querySelectorAll(`#tiConteudo .ti-card[data-nav*="bancohoras"]`).forEach(card => {
      const tem = card.querySelector('.tu-cadeado');
      if(logado()){ if(tem) tem.remove(); return; }
      if(tem) return;
      const b = document.createElement('span');
      b.className = 'ti-badge tu-cadeado'; b.style.background = '#374151'; b.title = 'Precisa de usuário e senha'; b.textContent = '🔒';
      card.prepend(b);
    });
  }

  const mostrarAntes = window.mostrarPagina;
  if(typeof mostrarAntes === 'function'){
    window.mostrarPagina = function(nome){
      mostrarAntes.apply(this, arguments);
      if(String(nome).includes('bancohoras') && !logado()){
        aplicar();
        setTimeout(() => document.querySelector('#' + PAGINA + ' .trava-usuario .tu-user')?.focus(), 50);
      }
    };
  }

  async function conferirSessao(){
    try{
      const r = await fetch(ARQ_SESSAO + '?_=' + Date.now(), { cache: 'no-store' });
      if(!r.ok) return;
      const j = await r.json();
      versaoSessao = String(j && j.versao != null ? j.versao : '');
      if(!logado()) return;
      const minha = sessionStorage.getItem(CHAVE + 'V');
      if(minha !== versaoSessao){
        sessionStorage.removeItem(CHAVE); sessionStorage.removeItem(CHAVE + 'K'); sessionStorage.removeItem(CHAVE + 'V');
        aplicar(); marcarCard();
        const er = document.querySelector('#' + PAGINA + ' .trava-usuario .tu-erro');
        if(er){ er.style.color = '#64748b'; er.textContent = 'Sessão encerrada. Entre de novo com seu e-mail e senha.'; }
      }
    }catch(e){}
  }

  function iniciar(){
    conferirSessao();
    setInterval(conferirSessao, 60 * 1000);
    let n = 0;
    const tentar = () => { const ok = aplicar(); marcarCard(); return ok; };
    if(!tentar()){ const t = setInterval(() => { if(tentar() || ++n > 40) clearInterval(t); }, 250); }
    const tela = document.getElementById('tiConteudo');
    if(tela) new MutationObserver(marcarCard).observe(tela, { childList: true, subtree: true });
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
