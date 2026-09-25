// ======================================================================
// CONVERSOR NR / ASO — Painel do ESER
// ----------------------------------------------------------------------
// Transforma as abas da Base_NR.xlsx (BASE, Base_Supervisores,
// Preenchimento) nos dados das páginas "Gestão de NR" e "Exames
// Periódicos". Usado pelo painel (lendo base_nr.json, que o Power
// Automate atualiza) e pela página atualizar.html (planilha manual).
// No index.html precisa vir ANTES de gestao_nr.js e gestao_aso.js.
// ======================================================================
(function(){

// Mesmas regras usadas para gerar os arquivos do painel.
const CURSOS = ['NR6','NR12','NR17','NR20','NR18','NR26','SEP','NR10','NR35','DIREÇÃO','FICHA DE EPI','INTEGRAÇÃO','O.S'];
const PART = new Set(['DA','DE','DO','DOS','DAS','E']);
const IGNORAR = new Set(['TECNICO DE OUTRO SETOR','DEMISSAO']);

const semAcento = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const n = s => semAcento(s).toUpperCase().replace(/\s+/g, ' ').trim();
const t = s => n(s).split(' ').filter(w => w && !PART.has(w)).join(' ');
function mesmo(a, b){
  const x = t(a), y = t(b);
  if(!x || !y) return false;
  return x === y || (Math.min(x.length, y.length) >= 20 && (x.startsWith(y) || y.startsWith(x)));
}
const espacos = s => String(s ?? '').replace(/\s+/g, ' ').trim();
const pad = v => String(v).padStart(2, '0');
function serialParaIso(v){
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
  return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
}
function dataDeTexto(s){
  const m = String(s).match(/(\d{2})[\/.](\d{2})[\/.](\d{4})/);
  if(!m) return null;
  const d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  if(d.getUTCDate() !== +m[1] || d.getUTCMonth() !== +m[2] - 1) return null;
  return m[3] + '-' + m[2] + '-' + m[1];
}
function paraIso(v){
  if(v === null || v === undefined || v === '') return null;
  if(typeof v === 'number') return serialParaIso(v);
  if(v instanceof Date) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
  return dataDeTexto(v);
}
function classificar(v, desconhecidos){
  if(v === null || v === undefined || String(v).trim() === '') return { status:'SEM REGISTRO', data:null, obs:null };
  if(typeof v === 'number' || v instanceof Date) return { status:'OK', data:paraIso(v), obs:null };
  const s = String(v).trim(), u = s.toUpperCase(), d = dataDeTexto(s);
  if(u === 'OK') return { status:'OK', data:null, obs:null };
  if(u === 'PENDENTE') return { status:'PENDENTE', data:null, obs:null };
  if(u === 'N/A' || u === 'NA' || u.endsWith('- NA')) return { status:'N/A', data:null, obs:null };
  if(/DATA ERRADA|DIVERGENT|MESMA DATA|SEM DATA|- DATA$/.test(u)) return { status:'DIVERGÊNCIA', data:d, obs:s };
  if(/ASSIN|VERSO|FRENTE/.test(u)) return { status:'SEM ASSINATURA', data:d, obs:s };
  if(u === 'CONCLUIR' || u === 'CONTINUAR') return { status:'EM ANDAMENTO', data:null, obs:s };
  if(u === 'AVALIAR' || u === 'SOLICITAR') return { status:'AVALIAR', data:null, obs:s };
  if(d) return { status:'OK', data:d, obs:s };           // texto que é só uma data
  desconhecidos.add(s);
  return { status:'AVALIAR', data:null, obs:s };           // texto novo: marca para avaliar
}

// Lê uma aba como lista de objetos (cabeçalho = 1ª linha, sem espaços sobrando).
// obterLinhas(nomeAba) devolve a aba como matriz [[cabeçalho...], [linha...], ...] ou null.
function lerAba(obterLinhas, nomeAba){
  const linhas = obterLinhas(nomeAba);
  if(!linhas) throw new Error('Não encontrei a aba "' + nomeAba + '" na planilha.');
  const cab = (linhas[0] || []).map(h => espacos(h));
  return linhas.slice(1).map(l => { const o = {}; cab.forEach((h, i) => { if(h) o[h] = l[i]; }); return o; });
}
// pega a coluna ignorando maiúsculas/acentos no cabeçalho
function col(obj, nome){
  if(nome in obj) return obj[nome];
  const k = Object.keys(obj).find(x => n(x) === n(nome));
  return k ? obj[k] : null;
}

function converter(obterLinhas, org, hojeIso){
  org = org || [];
  const desconhecidos = new Set();
  const cursosDe = rec => { const o = {}; CURSOS.forEach(c => o[c] = classificar(col(rec, c), desconhecidos)); return o; };

  // ---------------- NR: técnicos da aba BASE
  // Coluna Status da BASE: "Ativo" / "Desligado" (se não existir, todos contam como ativos)
  const baseTodos = lerAba(obterLinhas, 'BASE').filter(r => col(r, 'Nome'));
  const ativo = r => { const st = n(col(r, 'Status')); return !st || st === 'ATIVO'; };
  const base = baseTodos.filter(ativo);
  const colaboradores = [];
  base.forEach(r => {
    if(n(col(r, 'Área Técnica')) !== 'TECNICO') return;
    colaboradores.push({
      matricula: col(r, 'Matricula'), nome: String(col(r, 'Nome')).trim(),
      admissao: paraIso(col(r, 'Data Admis.')), depto: col(r, 'Desc. Depto'), area: col(r, 'Área Técnica'),
      supervisor: String(col(r, 'Supervisor') ?? '').trim(), coordenador: String(col(r, 'Coordenador') ?? '').trim(),
      situacao_nr: col(r, 'Situação do Nr'), observacao: col(r, 'OBSERVAÇÃO') || null, cursos: cursosDe(r), saiu: false
    });
  });

  // ---------------- NR: quem está na Base_Supervisores e não está ATIVO na BASE = saiu da empresa
  const nomesBase = base.map(r => col(r, 'Nome'));
  const desligados = baseTodos.filter(r => !ativo(r));
  const SUPS = Array.from(new Set(org.map(c => espacos(c.supervisor)).filter(Boolean)));
  function supCompleto(g){
    if(!g || ['TBD','OK'].includes(n(g))) return '';
    const k = t(g).split(' ');
    const achou = SUPS.find(s => { const w = t(s).split(' '); return k.every(x => w.includes(x)); });
    return achou || espacos(g);
  }
  const saidas = [];
  lerAba(obterLinhas, 'Base_Supervisores').forEach(r => {
    const nome = col(r, 'NOME');
    if(!nome || IGNORAR.has(n(nome))) return;
    if(nomesBase.some(b => mesmo(nome, b))) return;
    const achados = org.filter(c => mesmo(nome, c.nome));
    const o = achados.find(c => !c.ativo) || achados[0] || {};
    const naBase = desligados.find(b => mesmo(nome, col(b, 'Nome')));
    const sup = o.supervisor ? espacos(o.supervisor) : supCompleto(col(r, 'GESTÃO'));
    const item = {
      matricula: o.matricula || (naBase ? col(naBase, 'Matricula') : '') || '', nome: espacos(nome),
      admissao: o.dataAdmis || (naBase ? paraIso(col(naBase, 'Data Admis.')) : null), depto: '', area: 'TÉCNICO',
      supervisor: sup, coordenador: espacos(o.coordenador), situacao_nr: 'SAIU DA EMPRESA',
      observacao: col(r, 'OBSERVAÇÃO') || null, cursos: cursosDe(r), saiu: true, demissao: o.dataDemissao || null
    };
    colaboradores.push(item);
    saidas.push({ nome: item.nome, ativoNoOrganograma: achados.some(c => c.ativo), noOrganograma: achados.length > 0, desligadoNaBase: !!naBase });
  });

  // ---------------- ASO: aba Preenchimento (sem CPF)
  const SUPS_N = {}; SUPS.forEach(s => SUPS_N[n(s)] = s);
  const semMatricula = [];
  const exames = lerAba(obterLinhas, 'Preenchimento').filter(r => col(r, 'Colaborador')).map(r => {
    const nome = col(r, 'Colaborador');
    const o = org.find(c => c.ativo && mesmo(nome, c.nome)) || {};
    if(!o.matricula) semMatricula.push(espacos(nome));
    let sup = espacos(col(r, 'Supervisor'));
    sup = SUPS_N[n(sup)] || sup || espacos(o.supervisor);
    return {
      matricula: o.matricula || '', nome: espacos(nome).toUpperCase(), funcao: o.descFuncao || '', area: o.areaTecnica || '',
      coordenador: espacos(o.coordenador), supervisor: sup,
      vencimento: paraIso(col(r, 'Vencimento do Aso')), data_exame: paraIso(col(r, 'Data do Exame')),
      status: String(col(r, 'Status') ?? '').trim(), obs: col(r, 'Observações') ? String(col(r, 'Observações')) : ''
    };
  });

  return {
    nr: { atualizado_em: hojeIso, cursos: CURSOS, colaboradores },
    aso: { atualizado_em: hojeIso, exames },
    resumo: { tecnicos: colaboradores.filter(c => !c.saiu).length, saidas, exames: exames.length,
              semMatricula, desconhecidos: Array.from(desconhecidos) }
  };
}

  // Espera a lista de colaboradores do dados.json (Organograma) carregar no painel.
  function esperarColaboradores(limiteMs){
    return new Promise(resolve => {
      const ini = Date.now();
      (function checa(){
        const lista = (typeof COLABORADORES !== 'undefined' && Array.isArray(COLABORADORES)) ? COLABORADORES : [];
        if(lista.length || Date.now() - ini > limiteMs) return resolve(lista);
        setTimeout(checa, 200);
      })();
    });
  }

  async function buscarJson(arq){
    const r = await fetch(arq + '?_=' + Date.now());
    if(!r.ok) throw new Error(arq + ': HTTP ' + r.status);
    return r.json();
  }

  // 1º tenta base_nr.json (automático, vem do Power Automate);
  // se não existir, usa dados_nr.json + dados_aso.json (gerados na mão).
  let cache = null, cacheEm = 0;
  function carregar(){
    if(cache && Date.now() - cacheEm < 60000) return cache;
    cacheEm = Date.now();
    cache = (async () => {
      try{
        const bruto = await buscarJson('base_nr.json');
        if(!bruto || !bruto.abas) throw new Error('base_nr.json sem abas');
        const org = await esperarColaboradores(20000);
        const r = converter(nome => {
          const k = Object.keys(bruto.abas).find(x => n(x) === n(nome));
          return k ? bruto.abas[k] : null;
        }, org, bruto.atualizado_em);
        r.fonte = 'base_nr.json';
        return r;
      }catch(e){
        console.warn('[NR] base_nr.json indisponível, usando dados_nr.json / dados_aso.json —', e.message);
        const [nr, aso] = await Promise.all([
          buscarJson('dados_nr.json').catch(() => null),
          buscarJson('dados_aso.json').catch(() => null)
        ]);
        if(!nr && !aso) throw new Error('não encontrei base_nr.json nem dados_nr.json / dados_aso.json');
        return { nr, aso, fonte:'dados_*.json' };
      }
    })();
    cache.catch(() => { cache = null; });
    return cache;
  }

  window.NRConversor = { converter, carregar, CURSOS };
})();
