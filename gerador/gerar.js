// =====================================================================
// GERADOR DO PAINEL — roda no GitHub (Actions), sem Power Automate.
// Para cada planilha enviada pela macro, roda o MESMO Office Script de
// sempre (pasta scripts/) e grava o JSON do painel.
//
//   node gerar.js --entrada <pasta com os .xlsx> --saida <raiz do site>
//                 [--arquivos todos|a.xlsx,b.xlsx] [--modo teste|producao]
//
// modo teste    → grava em _teste/ e mostra a comparação com o que está no ar
// modo producao → grava por cima dos JSONs do painel
// =====================================================================
"use strict";
const fs = require("fs");
const path = require("path");
const ts = require("typescript");
const { abrirPasta, ExcelScript } = require("./excel");

// ---------------------------------------------------------------------
// Qual planilha (nome no GitHub) gera qual arquivo do painel
// ---------------------------------------------------------------------
const TAREFAS = [
  // dados.json: Base com Historico + Organograma/PJ (planilha Organograma) + Controle (planilha de Vagas), se forem enviadas
  { planilha: "base_historico.xlsx", saida: "dados.json", script: "dados.ts", tipo: "dados",
    extras: [{ planilha: "organograma.xlsx", abas: ["Organograma", "PJ"] },
             { planilha: "vagas.xlsx", abas: ["Controle"] },
             { planilha: "fca.xlsx", abas: ["FCA"] },
             // Frota: planilha da Tamires (CONTROLE DE FROTA - CONSOLIDADA). Vira a aba "Placa" que o script lê;
             // pega a aba que tem a coluna "Placa" e começa a tabela na linha desse cabeçalho.
             { planilha: "frota.xlsx", abas: ["Placa"], cabecalho: "Placa", preferida: "Base ativa" }] },
  // NOVO (25/11/2026): incidentes do mês atual com TODAS as colunas da aba Historico (botão "Exportar base do mês" do ESER)
  { planilha: "base_historico.xlsx", saida: "historico_mes.json", script: "historico_mes.ts" },
  { planilha: "afaz.xlsx", saida: "dados_afaz.json", script: "afaz.ts" },
  { planilha: "preventiva.xlsx", saida: "dados_preventiva.json", script: "preventiva.ts", anterior: true },
  { planilha: "nr.xlsx", saida: "base_nr.json", script: "nr.ts" },
  { planilha: "bh.xlsx", saida: "dados_bh.json", script: "bh.ts" },
  { planilha: "bh.xlsx", saida: "dados_pend_bh.json", script: "pend_bh.ts" },
  { planilha: "combustivel.xlsx", saida: "dados_combustivel.json", script: "combustivel.ts" },
  { planilha: "escala.xlsx", saida: "dados_escala.json", script: "escala.ts" },
  // Faltas/Atestados: aba VOLANTE (planilha do Thiago) + aba RESIDENTE (planilha dos residentes, se for enviada)
  { planilha: "faltas.xlsx", saida: "dados_escala_extra.json", script: "escala_extra.ts",
    extras: [{ planilha: "faltas_residente.xlsx", abas: ["RESIDENTE"], semPrimeira: true }] }
];

const sem = v => String(v == null ? "" : v).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toLowerCase();

function argumento(nome, padrao) {
  const i = process.argv.indexOf("--" + nome);
  return i >= 0 && process.argv[i + 1] !== undefined ? process.argv[i + 1] : padrao;
}

// ---------------------------------------------------------------------
// Carrega um Office Script (.ts) e devolve a função main
// ---------------------------------------------------------------------
const cacheScripts = {};
function carregarScript(arquivo) {
  if (cacheScripts[arquivo]) return cacheScripts[arquivo];
  const fonte = fs.readFileSync(path.join(__dirname, "scripts", arquivo), "utf8");
  const js = ts.transpileModule(fonte, {
    compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.None, removeComments: false }
  }).outputText;
  const fabrica = new Function("ExcelScript", "console", js + "\n;return { main: typeof main === 'function' ? main : undefined };");
  const mod = fabrica(ExcelScript, console);
  if (!mod.main) throw new Error("O script " + arquivo + " não tem a função main");
  cacheScripts[arquivo] = mod.main;
  return mod.main;
}

function lerJsonAtual(saida, arquivo) {
  try { return fs.readFileSync(path.join(saida, arquivo), "utf8"); } catch (e) { return ""; }
}

// ---------------------------------------------------------------------
// dados.json: o script principal recebe FCA, Rateio, Banco de Horas e
// Turnover de outras planilhas. Aqui:
//  - FCA = justificativas que já estão no painel + aba "FCA" da base (se existir)
//  - Rateio, Banco de Horas e Turnover = mantidos do dados.json que já está no ar
// ---------------------------------------------------------------------
function argsDados(wb, atualTxt) {
  let vivo = {};
  try { vivo = JSON.parse(atualTxt || "{}"); } catch (e) { vivo = {}; }
  const fca = [];
  (vivo.pendentesFca || []).forEach(x => { if (x && x.incidente && x.causa) fca.push({ incidente: x.incidente, causa: x.causa, acao: x.acao || "", acaoTomada: x.acaoTomada || "" }); });
  (vivo.fcaTipoDivergente || []).forEach(x => { if (x && x.incidente && x.causa) fca.push({ incidente: x.incidente, causa: x.causa, acao: x.acao || "", acaoTomada: "" }); });
  const sem = s => String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
  const aba = wb.getWorksheets().find(w => sem(w.getName()) === "fca");
  let daAba = 0;
  if (aba) {
    const u = aba.getUsedRange(true);
    if (u) {
      const v = u.getValues();
      // cabeçalhos do Forms ("Número do incidente", "Qual a causa?", "Ação tomada"...) → campos do script
      const campo = h => {
        const t = sem(h);
        if (/tomada/.test(t)) return "acaoTomada";
        if (/incidente|incident/.test(t)) return "incidente";
        if (/\bfato\b/.test(t)) return "fato";
        if (/\bcausa\b/.test(t)) return "causa";
        if (/\bacao\b|\bacoes\b/.test(t)) return "acao";
        return String(h).trim();
      };
      const cab = v[0].map(x => campo(x));
      for (let i = 1; i < v.length; i++) {
        const item = {}; let tem = false;
        cab.forEach((c, j) => { if (!c) return; const val = String(v[i][j] == null ? "" : v[i][j]).trim(); item[c] = val; if (val) tem = true; });
        if (tem) { fca.push(item); daAba++; }
      }
    }
  }
  const rateio = (vivo.frota || []).filter(f => f && f.placa && typeof f.valorCombustivel === "number").map(f => ({ placa: f.placa, valorTotal: f.valorCombustivel }));
  console.log("   FCA: " + (fca.length - daAba) + " do painel + " + daAba + " da aba FCA");
  return [JSON.stringify(fca), JSON.stringify(rateio), JSON.stringify(vivo.bancoHoras || []), JSON.stringify(vivo.turnover || [])];
}

// ---------------------------------------------------------------------
// NOVO (19/11/2026): ARQUIVO DAS FOLHAS FECHADAS (pendências de ponto)
// Toda pendência com data ANTES do início da folha atual vai para o
// historico_pend_bh.json e fica guardada ali para sempre — o painel não
// depende mais de a base (planilha) continuar com as folhas fechadas.
// Fonte de cada pendência fechada: o arquivo novo, se ainda tiver; senão o
// que estava no ar (última versão antes de sair da base); senão o histórico.
// ---------------------------------------------------------------------
function arquivarPendencias(novoTxt, atualTxt, histTxt) {
  const ler = t => { try { return JSON.parse(t || "null"); } catch (e) { return null; } };
  const novo = ler(novoTxt), atual = ler(atualTxt), hist = ler(histTxt);
  const itensDe = o => !o ? [] : (Array.isArray(o) ? o : (o.itens || []));
  const ini = novo && novo.folhaAtualIni ? String(novo.folhaAtualIni).slice(0, 10) : "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ini)) return null;
  const chave = i => String(i.m || "").replace(/\D/g, "").padStart(6, "0").slice(-6) + "|" + String(i.d || "").slice(0, 10) + "|" + String(i.tipo || "").trim().toUpperCase();
  const fechada = i => i && i.d && String(i.d).slice(0, 10) < ini;
  const mapa = new Map();
  let antes = 0;
  itensDe(hist).forEach(i => { if (i && i.d) { mapa.set(chave(i), i); antes++; } });
  // resposta (just) mais completa vence: não perde justificativa que já estava guardada
  const temResp = i => !!(i && i.just && (i.just.trat || i.just.status));
  const por = (i, origem) => {
    const k = chave(i), velho = mapa.get(k);
    const x = Object.assign({}, i); delete x.hist;
    if (velho && temResp(velho) && !temResp(x)) x.just = velho.just;
    x.arquivadoDe = origem;
    mapa.set(k, x);
  };
  itensDe(atual).filter(fechada).forEach(i => por(i, "painel"));
  itensDe(novo).filter(fechada).forEach(i => por(i, "base"));
  const itens = Array.from(mapa.values()).sort((a, b) => String(b.d).localeCompare(String(a.d)) || String(a.n || "").localeCompare(String(b.n || "")));
  const base = (hist && !Array.isArray(hist)) ? hist : {};
  return { txt: JSON.stringify(Object.assign({}, base, {
    geradoEm: new Date().toISOString(), fonte: "folhas fechadas guardadas pelo gerador (antes de " + ini + ") + TRATATIVA_GESTORES",
    total: itens.length, itens: itens })), antes: antes, depois: itens.length };
}

// ---------------------------------------------------------------------
// Conferências antes de publicar (não sobe arquivo vazio/quebrado)
// ---------------------------------------------------------------------
function conferir(t, txt) {
  let o;
  try { o = JSON.parse(txt); } catch (e) { return "não é um JSON válido"; }
  if (!o || typeof o !== "object") return "resultado vazio";
  if (o.erro) return "o script devolveu erro: " + o.erro;
  if (txt.length < 200) return "resultado pequeno demais (" + txt.length + " caracteres)";
  if (t.saida === "dados.json" && (!Array.isArray(o.meses) || o.meses.length === 0)) return "sem meses (aba Historico vazia?)";
  if (t.saida === "historico_mes.json" && !(o.total > 0)) return "nenhum incidente no mês";
  if (t.saida === "dados_afaz.json" && (!Array.isArray(o.itens) || o.itens.length === 0)) return "nenhum AFAZ aberto";
  if (t.saida === "dados_bh.json" && !(o.totalFuncionarios > 0)) return "nenhum funcionário";
  return "";
}

// resumo para comparar com o que está no ar (só contagens — o log é público)
function resumo(txt) {
  try {
    const o = JSON.parse(txt);
    return Object.keys(o).map(k => {
      const v = o[k];
      if (Array.isArray(v)) return k + "=" + v.length;
      if (v && typeof v === "object") return k + "{" + Object.keys(v).length + "}";
      return null;
    }).filter(Boolean).join(" ");
  } catch (e) { return "(inválido)"; }
}

async function main() {
  const entrada = argumento("entrada", "bases");
  const saida = argumento("saida", "..");
  const modo = argumento("modo", "teste");
  const lista = argumento("arquivos", "todos").split(/[,;\s]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
  const todos = lista.length === 0 || lista.indexOf("todos") >= 0;
  const pastaTeste = path.join(saida, "_teste");
  if (modo !== "producao") fs.mkdirSync(pastaTeste, { recursive: true });

  console.log("Modo: " + modo + " · planilhas: " + (todos ? "todas" : lista.join(", ")));
  const abertas = {};
  const relatorio = [];
  let falhas = 0;

  for (const t of TAREFAS) {
    const gatilhos = [t.planilha].concat((t.extras || []).map(e => e.planilha));
    if (!todos && !gatilhos.some(g => lista.indexOf(g) >= 0)) continue;
    const arqXlsx = path.join(entrada, t.planilha);
    if (!fs.existsSync(arqXlsx)) { relatorio.push("· " + t.saida + ": planilha " + t.planilha + " não foi enviada (pulei)"); continue; }
    if (!fs.existsSync(path.join(__dirname, "scripts", t.script))) { relatorio.push("· " + t.saida + ": falta o script " + t.script + " (pulei)"); continue; }
    const ini = Date.now();
    try {
      console.log("→ " + t.saida + " (" + t.planilha + ")");
      // cada script recebe a planilha "limpa" (o que um script grava não afeta o outro)
      const wb = await abrirPasta(fs.readFileSync(arqXlsx), t.planilha);
      // abas que vêm de outra planilha (ex.: Organograma) substituem as de mesmo nome
      for (const e of (t.extras || [])) {
        const arqExtra = path.join(entrada, e.planilha);
        if (!fs.existsSync(arqExtra)) continue;
        const wbE = await abrirPasta(fs.readFileSync(arqExtra), e.planilha);
        e.abas.forEach((nome, i) => {
          let aba = wbE.p.abas.find(a => a.nome.trim().toLowerCase() === nome.toLowerCase());
          if (e.cabecalho) {
            // aba com a coluna do cabeçalho (ex.: "Placa"); a tabela passa a começar em A1
            const alvo = sem(e.cabecalho);
            const pref = e.preferida ? wbE.p.abas.find(a => sem(a.nome) === sem(e.preferida)) : null;
            if (!aba && pref) aba = pref;
            const ordem = (aba ? [aba] : []).concat(wbE.p.abas.filter(a => a !== aba && a.visivel === "visible"), wbE.p.abas.filter(a => a.visivel !== "visible"));
            let achou = null;
            for (const a of ordem) {
              for (let r = 0; r < Math.min(a.linhas.length, 40) && !achou; r++) {
                const l = a.linhas[r] || [];
                if (l.some(v => sem(v) === alvo)) achou = { a, r };
              }
              if (achou) break;
            }
            if (!achou) { console.log("   " + e.planilha + ": não achei coluna " + e.cabecalho + " (pulei)"); return; }
            aba = achou.a;
            const lin = aba.linhas, hr = achou.r;
            const c0 = (lin[hr] || []).findIndex(v => v !== undefined && v !== null && String(v).trim() !== "");
            aba.linhas = []; aba.horas = new Set(); aba._usado = null; aba._usadoV = null;
            for (let r = hr; r < lin.length; r++) {
              const l = lin[r] || [];
              for (let c = Math.max(c0, 0); c < l.length; c++) if (l[c] !== undefined && l[c] !== null && l[c] !== "") aba.set(r - hr, c - Math.max(c0, 0), l[c]);
            }
            console.log("   " + e.planilha + ": cabeçalho na linha " + (hr + 1) + " da aba " + achou.a.nome);
          }
          if (!aba && i === 0 && !e.semPrimeira) aba = wbE.p.abas[0];   // sem aba com esse nome: usa a primeira
          if (!aba) return;
          wb.p.abas = wb.p.abas.filter(a => a.nome.trim().toLowerCase() !== nome.toLowerCase());
          aba.nome = nome;
          wb.p.abas.push(aba);
          console.log("   aba " + nome + " veio de " + e.planilha);
        });
      }
      abertas[t.planilha] = true;
      const main = carregarScript(t.script);
      const atual = lerJsonAtual(saida, t.saida);
      let txt;
      if (t.tipo === "dados") txt = main(wb, ...argsDados(wb, atual));
      else if (t.anterior) txt = main(wb, atual);
      else txt = main(wb);
      txt = String(txt);
      const problema = conferir(t, txt);
      if (problema) { falhas++; relatorio.push("✗ " + t.saida + ": NÃO PUBLICADO — " + problema); continue; }
      const destino = modo === "producao" ? path.join(saida, t.saida) : path.join(pastaTeste, t.saida);
      fs.writeFileSync(destino, txt);
      if (t.saida === "dados_pend_bh.json") {
        const ARQ_HIST = "historico_pend_bh.json";
        const r = arquivarPendencias(txt, atual, lerJsonAtual(saida, ARQ_HIST));
        if (r) {
          fs.writeFileSync(modo === "producao" ? path.join(saida, ARQ_HIST) : path.join(pastaTeste, ARQ_HIST), r.txt);
          relatorio.push("✓ " + ARQ_HIST + ": " + r.antes + " → " + r.depois + " pendências guardadas (folhas fechadas)");
        }
      }
      const seg = ((Date.now() - ini) / 1000).toFixed(1);
      relatorio.push("✓ " + t.saida + " (" + Math.round(txt.length / 1024) + " KB, " + seg + " s)" +
        (modo !== "producao" ? "\n     novo:  " + resumo(txt) + "\n     no ar: " + resumo(atual) : ""));
    } catch (e) {
      falhas++;
      relatorio.push("✗ " + t.saida + ": ERRO — " + (e && e.message ? e.message : String(e)));
      console.error(e && e.stack ? e.stack.split("\n").slice(0, 4).join("\n") : e);
    }
  }
  console.log("\n===== RESULTADO =====\n" + relatorio.join("\n"));
  if (falhas) process.exitCode = 1;
}

main();
