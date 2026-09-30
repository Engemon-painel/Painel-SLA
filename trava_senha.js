// ======================================================================
// TRAVA DE SENHA — outras páginas (mesma senha da Gestão de Frotas)
// ----------------------------------------------------------------------
// A Frota continua travada pelo tela_inicial.js. Este arquivo trava as
// páginas listadas abaixo com a MESMA senha. Digitou a senha em qualquer
// uma, libera todas até fechar a aba.
// No index.html: <script src="trava_senha.js"></script> (depois do tela_inicial.js)
// ======================================================================
(function(){

  // ---- páginas travadas: palavra que aparece no id da página : título da trava ----
  // (ex.: 'turnover' encontra o elemento id="pagina-turnover" e o card id="nav-turnover")
  const PAGINAS_COM_SENHA = {
    turnover: 'Turnover'
    // vagas: 'Gestão de Vagas',
  };

  // mesma "impressão digital" da senha usada no tela_inicial.js
  const SENHA_SHA256 = '8b94527a1c8f5579868ed765d2a280015bdc1cfaeb5d79c58fb3a8d817aa5631';
  const CHAVE = 'frotaLiberada';   // mesma chave da Frota → uma senha libera tudo

  const liberado = () => { try{ return sessionStorage.getItem(CHAVE) === '1'; }catch(e){ return false; } };
  const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  async function sha256(txt){
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(txt));
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
  }
  const acharPagina = chave => document.getElementById('pagina-' + chave) ||
    document.querySelector(`[id^="pagina-"][id*="${chave}"]`);

  function estilo(){
    if(document.getElementById('estiloTravaSenha')) return;
    const st = document.createElement('style'); st.id = 'estiloTravaSenha';
    st.textContent = `
      .trava-senha{ max-width:420px; margin:40px auto; background:var(--panel,#fff); border:1px solid var(--line,#e1e5f0); border-radius:16px;
        padding:28px 26px; text-align:center; box-shadow:0 6px 20px rgba(27,36,64,.08); }
      .trava-senha .tf-icone{ font-size:40px; margin-bottom:8px; }
      .trava-senha h2{ font-family:'Space Grotesk', sans-serif; font-size:18px; margin-bottom:6px; }
      .trava-senha p{ font-size:13px; color:var(--muted,#6b7590); margin-bottom:16px; }
      .trava-senha input{ width:100%; font-size:14px; padding:10px 12px; border-radius:10px; border:1px solid var(--line,#e1e5f0); margin-bottom:10px; }
      .trava-senha input:focus{ outline:2px solid var(--accent,#0e7c86); border-color:transparent; }
      .trava-senha button{ width:100%; font-family:'JetBrains Mono', monospace; font-weight:700; font-size:13px; padding:10px; border-radius:999px;
        border:0; background:var(--accent,#0e7c86); color:#fff; cursor:pointer; }
      .trava-senha .tf-erro{ color:var(--bad,#dc2626); font-size:12px; min-height:16px; margin-top:8px; }`;
    document.head.appendChild(st);
  }

  const observadores = {};
  function aplicar(chave){
    const titulo = PAGINAS_COM_SENHA[chave];
    const pg = acharPagina(chave);
    if(!titulo || !pg) return false;
    estilo();
    const esconder = () => Array.from(pg.children).forEach(el => {
      if(el.classList.contains('trava-senha')) return;
      if(el.style.display !== 'none'){ el.style.display = 'none'; el.dataset.travaOculto = '1'; }
    });
    let trava = pg.querySelector(':scope > .trava-senha');

    if(liberado()){
      if(observadores[chave]){ observadores[chave].disconnect(); delete observadores[chave]; }
      Array.from(pg.children).forEach(el => { if(el.dataset.travaOculto){ el.style.display = ''; delete el.dataset.travaOculto; } });
      if(trava) trava.remove();
      return true;
    }

    esconder();
    if(!observadores[chave]){   // se a página montar conteúdo depois, esconde também
      observadores[chave] = new MutationObserver(() => { if(!liberado()) esconder(); });
      observadores[chave].observe(pg, { childList: true });
    }
    if(trava) return true;

    trava = document.createElement('div');
    trava.className = 'trava-senha';
    trava.innerHTML = `<div class="tf-icone">🔒</div>
      <h2>${esc(titulo)}</h2>
      <p>Área restrita. Digite a senha para ver os dados.</p>
      <input type="password" placeholder="Senha" autocomplete="off">
      <button type="button">Entrar</button>
      <div class="tf-erro"></div>`;
    pg.appendChild(trava);
    const inp = trava.querySelector('input'), erro = trava.querySelector('.tf-erro');
    const entrar = async () => {
      if(inp.value && await sha256(inp.value) === SENHA_SHA256){
        try{ sessionStorage.setItem(CHAVE, '1'); }catch(e){}
        Object.keys(PAGINAS_COM_SENHA).forEach(aplicar);
        marcarCards();
        window.dispatchEvent(new Event('resize'));   // redesenha gráficos que estavam escondidos
      } else {
        erro.textContent = 'Senha incorreta.';
        inp.select();
      }
    };
    trava.querySelector('button').onclick = entrar;
    inp.onkeydown = e => { if(e.key === 'Enter') entrar(); };
    return true;
  }

  // cadeado nos cards da tela inicial
  function marcarCards(){
    Object.keys(PAGINAS_COM_SENHA).forEach(chave => {
      document.querySelectorAll(`#tiConteudo .ti-card[data-nav*="${chave}"]`).forEach(card => {
        const tem = card.querySelector('.ti-cadeado');
        if(liberado()){ if(tem) tem.remove(); return; }
        if(tem) return;
        const b = document.createElement('span');
        b.className = 'ti-badge ti-cadeado'; b.style.background = '#374151'; b.title = 'Precisa de senha'; b.textContent = '🔒';
        card.prepend(b);
      });
    });
  }

  // foca a senha quando a página travada é aberta
  const mostrarAntes = window.mostrarPagina;
  if(typeof mostrarAntes === 'function'){
    window.mostrarPagina = function(nome){
      mostrarAntes.apply(this, arguments);
      Object.keys(PAGINAS_COM_SENHA).forEach(chave => {
        if(String(nome).includes(chave) && !liberado()){
          aplicar(chave);
          setTimeout(() => acharPagina(chave)?.querySelector('.trava-senha input')?.focus(), 50);
        }
      });
    };
  }

  function iniciar(){
    // as páginas podem ser criadas pelos outros scripts depois: tenta por alguns segundos
    const pendentes = new Set(Object.keys(PAGINAS_COM_SENHA));
    const tentar = () => { pendentes.forEach(c => { if(aplicar(c)) pendentes.delete(c); }); marcarCards(); };
    tentar();
    let n = 0;
    const t = setInterval(() => { tentar(); if(!pendentes.size || ++n > 40) clearInterval(t); }, 250);
    const tela = document.getElementById('tiConteudo');
    if(tela) new MutationObserver(marcarCards).observe(tela, { childList: true, subtree: true });
    else setTimeout(() => { const t2 = document.getElementById('tiConteudo'); if(t2) new MutationObserver(marcarCards).observe(t2, { childList: true }); }, 1000);
  }

  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
