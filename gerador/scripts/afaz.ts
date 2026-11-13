// Office Script — gera o dados_afaz.json a partir das abas Afaz, Prioridades e Programação
// Cruza cada AFAZ em aberto com a aba "Prioridades" pelo Site (campos ps e fl no JSON).
// NOVO: lê a coluna "PRIORIDADE INFRAESTRUTURA" da aba Afaz (campo "pi" no JSON = Tipo Site)
//       e gera o resumo "porTipoSite" (AFAZ abertos por Tipo Site x PRIORIDADE do site).
// Uso no Power Automate: "Executar script" -> o retorno (result) é o JSON pronto para o PUT no GitHub.

function main(workbook: ExcelScript.Workbook): string {
  // ---------- utilitários ----------
  const ler = (nome: string): (string | number | boolean)[][] => {
    const ws = workbook.getWorksheet(nome);
    if (!ws) throw new Error(`Aba "${nome}" não encontrada`);
    const rg = ws.getUsedRange();
    return rg ? rg.getValues() : [];
  };

  // remove acentos, espaços extras e caixa — "Pendência Claro" == "pendencia claro"
  const norm = (v: string | number | boolean): string =>
    (v === null || v === undefined) ? "" :
      String(v).normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .trim().toLowerCase().replace(/\s+/g, " ");

  const idx = (cab: (string | number | boolean)[], nome: string): number =>
    cab.findIndex(c => norm(c) === norm(nome));

  const TRATADOS = ["finalizado", "resolvido", "pendencia claro"];
  const tratado = (v: string | number | boolean): boolean => TRATADOS.indexOf(norm(v)) >= 0;

  // Excel guarda datas como número serial; também aceita texto dd/mm/aaaa
  const paraData = (v: string | number | boolean): Date | null => {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400000));
    const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    const d = new Date(String(v));
    return isNaN(d.getTime()) ? null : d;
  };

  const iso = (d: Date | null): string | null => d ? d.toISOString().slice(0, 10) : null;
  const txt = (v: string | number | boolean): string => (v === null || v === undefined) ? "" : String(v).trim();

  // ---------- Afaz ----------
  const afaz = ler("Afaz");
  const c = afaz[0];
  const col = {
    id: idx(c, "id"), pr: idx(c, "Prioridade"), s: idx(c, "Site"), st: idx(c, "Status"),
    ti: idx(c, "Titulo"), ab: idx(c, "Abertura"), ar: idx(c, "Área/Sistema"),
    resp: idx(c, "Responsável"), prev: idx(c, "Previsão"), tst: idx(c, "Tempo do Status(horas)"),
    tp: idx(c, "Tipo"), ent: idx(c, "Entrou Em"), dst: idx(c, "Dias no Status"),
    pi: idx(c, "PRIORIDADE INFRAESTRUTURA"),  // PRIORIDADE 0 a 4 (P0 ... P4)
    ts: idx(c, "Tipo Site")                    // Tipo Site (Móvel Concentrador, Hub, Móvel Ponta ...)
  };
  if (col.pi < 0) throw new Error('Coluna "PRIORIDADE INFRAESTRUTURA" não encontrada na aba Afaz');

  // ---------- Prioridades (por Site) ----------
  // aba "Prioridades": id | Site | siglaClear | concentrador | estrategico | Isola Movel | Isola Gpon | gpon | PRIORIDADE
  type InfoSite = { p: number | null, cc: string, es: string, im: string, ig: string, gp: string };
  const prSite: { [site: string]: InfoSite } = {};
  const wsPr = workbook.getWorksheet("Prioridades");
  if (wsPr) {
    const pr = wsPr.getUsedRange().getValues();
    const cp = pr[0];
    const pc = {
      s: idx(cp, "Site"), sg: idx(cp, "siglaClear"), p: idx(cp, "PRIORIDADE"),
      cc: idx(cp, "concentrador"), es: idx(cp, "estrategico"),
      im: idx(cp, "Isola Movel"), ig: idx(cp, "Isola Gpon"), gp: idx(cp, "gpon")
    };
    const val = (r: (string | number | boolean)[], i: number): string => i >= 0 ? txt(r[i]) : "";
    for (let i = 1; i < pr.length; i++) {
      const r = pr[i];
      const n = Number(r[pc.p]);
      const info: InfoSite = {
        p: (r[pc.p] === "" || isNaN(n)) ? null : n,
        cc: val(r, pc.cc), es: val(r, pc.es), im: val(r, pc.im), ig: val(r, pc.ig), gp: val(r, pc.gp)
      };
      const s = norm(r[pc.s]);
      if (s && !prSite[s]) prSite[s] = info;
      const sg = pc.sg >= 0 ? norm(r[pc.sg]) : "";
      if (sg && !prSite[sg]) prSite[sg] = info; // reserva: casa também pela sigla
    }
  }

  const itens: object[] = [];
  const tratados: { [tipo: string]: number } = {};
  // tratados separados por status atual e tipo (gráfico "Por status e carteira"): { "Resolvido": { Backlog: 265, Entrante: 85 } }
  const tratadosPorStatus: { [st: string]: { [tipo: string]: number } } = {};

  // NOVO: AFAZ abertos por Tipo Site (coluna PRIORIDADE INFRAESTRUTURA)
  // { "Concentrador": { total: 505, sites: 167, atrasados: 144, porPrioridade: { "0": 7, "1": 498 } }, ... }
  type ResumoTipo = { total: number, atrasados: number, sites: number, porPrioridade: { [p: string]: number } };
  const porTipoSite: { [tipo: string]: ResumoTipo } = {};
  const sitesPorTipo: { [tipo: string]: { [s: string]: boolean } } = {};
  const hoje = new Date().toISOString().slice(0, 10);

  for (let i = 1; i < afaz.length; i++) {
    const r = afaz[i];
    if (r[col.id] === "" || r[col.id] === null) continue;

    const st = txt(r[col.st]);
    const tp = txt(r[col.tp]) || "Sem tipo";
    if (tratado(st)) {
      tratados[tp] = (tratados[tp] || 0) + 1;
      if (!tratadosPorStatus[st]) tratadosPorStatus[st] = {};
      tratadosPorStatus[st][tp] = (tratadosPorStatus[st][tp] || 0) + 1;
      continue;
    }

    let dias = Number(r[col.dst]);
    if (r[col.dst] === "" || r[col.dst] === null || isNaN(dias)) {
      const h = Number(r[col.tst]);
      dias = isNaN(h) ? 0 : Math.round(h / 24);
    }
    const ps = prSite[norm(r[col.s])];
    const pi = txt(r[col.pi]) || "Sem tipo site";
    const prev = iso(paraData(r[col.prev]));

    // resumo por Tipo Site
    if (!porTipoSite[pi]) { porTipoSite[pi] = { total: 0, atrasados: 0, sites: 0, porPrioridade: {} }; sitesPorTipo[pi] = {}; }
    const rt = porTipoSite[pi];
    rt.total++;
    if (prev && prev < hoje) rt.atrasados++;
    const sk = norm(r[col.s]);
    if (sk && !sitesPorTipo[pi][sk]) { sitesPorTipo[pi][sk] = true; rt.sites++; }
    const pk = ps && ps.p !== null ? String(ps.p) : "fora";
    rt.porPrioridade[pk] = (rt.porPrioridade[pk] || 0) + 1;

    itens.push({
      id: r[col.id], pr: txt(r[col.pr]), s: txt(r[col.s]), st: st, ti: txt(r[col.ti]),
      ab: iso(paraData(r[col.ab])), ar: txt(r[col.ar]), resp: txt(r[col.resp]),
      prev: prev, tp: tp, dst: dias,
      pi: pi,                        // PRIORIDADE INFRAESTRUTURA (P0 ... P4)
      tsi: col.ts >= 0 ? (txt(r[col.ts]) || "Sem tipo site") : "",   // coluna Tipo Site
      // colunas da aba Prioridades (vazias quando o site não está na lista)
      ps: ps ? ps.p : null,          // PRIORIDADE (0 = mais crítico)
      cc: ps ? ps.cc : "",           // concentrador
      es: ps ? ps.es : "",           // estrategico
      im: ps ? ps.im : "",           // Isola Movel
      ig: ps ? ps.ig : "",           // Isola Gpon
      gp: ps ? ps.gp : ""            // gpon
    });
  }

  const agora = new Date();

  // O operador ... espalha as propriedades (programado, execBacklog, execEntrante) diretamente no JSON de saída
  return JSON.stringify({
    geradoEm: agora.toISOString(),
    itens: itens,
    porTipoSite: porTipoSite,
    tratados: tratados,
    tratadosPorStatus: tratadosPorStatus,
    ...calcularExecucao(workbook)
  });
}

// ---- Programado x Executado (Backlog / Entrante) ----
function calcularExecucao(workbook: ExcelScript.Workbook) {
  const FECHADO = ["RESOLVIDO", "FINALIZADO", "PENDÊNCIA CLARO"];
  const mesDe = (v: string | number | boolean): string => {
    if (typeof v === "number" && v > 0) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0");
    }
    const m = String(v || "").match(/(\d{2})\/(\d{2})\/(\d{4})/);
    return m ? m[3] + "-" + m[2] : "";
  };

  // ---------- Executado ----------
  const execBacklog: { [k: string]: number } = {};
  const execEntrante: { [k: string]: number } = {};
  // executado separado por status fechado (para o filtro de Status do painel):
  // { "Resolvido": { b: { "2026-08": 88 }, e: { "2026-08": 22 } }, ... }
  const execPorStatus: { [st: string]: { b: { [k: string]: number }, e: { [k: string]: number } } } = {};
  const afaz = workbook.getWorksheet("Afaz").getUsedRange().getValues();
  const cab = afaz[0].map(c => String(c).trim());
  const cSt = cab.indexOf("Status"), cTp = cab.indexOf("Tipo"), cDt = cab.indexOf("Ultima Atualização"),
    cAb = cab.indexOf("Abertura");

  // ---------- Visão semanal (o painel agrupa por semana ISO, segunda a domingo) ----------
  // bTotal = total de AFAZ Backlog · bSemData = Backlog fechado sem data de atualização
  // bx = Backlog executado por dia · ex = Entrante executado por dia · ea = Entrante aberto por dia (Abertura)
  // ba = Backlog aberto por dia (Abertura) — o painel soma o Backlog aberto a partir da data de corte junto com os Entrantes
  const diaDe = (v: string | number | boolean): string => {
    if (typeof v === "number" && v > 0) return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
    const m = String(v || "").match(/(\d{2})\/(\d{2})\/(\d{4})/);
    return m ? m[3] + "-" + m[2] + "-" + m[1] : "";
  };
  const semanal = { bTotal: 0, bSemData: 0, bx: {} as { [d: string]: number }, ex: {} as { [d: string]: number }, ea: {} as { [d: string]: number }, ba: {} as { [d: string]: number } };
  const soma1 = (o: { [d: string]: number }, d: string) => { if (d) o[d] = (o[d] || 0) + 1; };
  for (let i = 1; i < afaz.length; i++) {
    const tp = String(afaz[i][cTp] || "").trim();
    const fechado = FECHADO.includes(String(afaz[i][cSt] || "").trim().toUpperCase());
    const dUp = diaDe(afaz[i][cDt]);
    if (tp === "Backlog") {
      semanal.bTotal++;
      if (cAb >= 0) soma1(semanal.ba, diaDe(afaz[i][cAb]));   // Backlog aberto por dia (Abertura)
      if (fechado) { if (dUp) soma1(semanal.bx, dUp); else semanal.bSemData++; }
    } else if (tp === "Entrante") {
      if (cAb >= 0) soma1(semanal.ea, diaDe(afaz[i][cAb]));
      if (fechado) soma1(semanal.ex, dUp);
    }
  }

  // ---------- Gráfico "Status por mês" (Abertura x Devolvido x Finalizado x Resolvido) ----------
  // logsMes = { "2026-08": { "Abertura": { Backlog: 500, Entrante: 239 }, "Resolvido": { Backlog: 150, Entrante: 61 }, ... } }
  // Abertura = mês da coluna Abertura da aba Afaz · demais = cada "Alterado Status ... para X" da aba Logs, no mês da Data
  const logsMes: { [mes: string]: { [st: string]: { [tp: string]: number } } } = {};
  const somaLog = (mes: string, st: string, tp: string) => {
    if (!mes || !st) return;
    if (!logsMes[mes]) logsMes[mes] = {};
    if (!logsMes[mes][st]) logsMes[mes][st] = {};
    logsMes[mes][st][tp] = (logsMes[mes][st][tp] || 0) + 1;
  };
  if (cAb >= 0) {
    for (let i = 1; i < afaz.length; i++) {
      const tp = String(afaz[i][cTp] || "").trim() || "Sem tipo";
      somaLog(mesDe(afaz[i][cAb]), "Abertura", tp);
    }
  }

  const contar = (ehEntrante: boolean, nomeSt: string, mes: string) => {
    const alvo = ehEntrante ? execEntrante : execBacklog;
    alvo[mes] = (alvo[mes] || 0) + 1;
    if (!execPorStatus[nomeSt]) execPorStatus[nomeSt] = { b: {}, e: {} };
    const porSt = ehEntrante ? execPorStatus[nomeSt].e : execPorStatus[nomeSt].b;
    porSt[mes] = (porSt[mes] || 0) + 1;
  };

  // ---------- Executado do gráfico Meta x Executado: pela aba Logs ----------
  // Cada AFAZ conta em CADA status de fechamento por onde passou, uma vez por mês:
  //   Resolvido em agosto e de novo em setembro -> conta em agosto e em setembro
  //   Resolvido e depois Finalizado            -> conta em Resolvido e em Finalizado
  //   Resolvido duas vezes no mesmo mês        -> conta 1
  // O Tipo (Backlog/Entrante) vem da aba Afaz pelo id; AFAZ que não está na aba Afaz fica de fora.
  // Sem a aba Logs, usa o jeito antigo (status atual da Afaz + Ultima Atualização).
  // O "Total tratado" do painel NÃO muda: continua vindo do status atual na aba Afaz.
  let origemExec = "Afaz (status atual)";
  const wsLogs = workbook.getWorksheet("Logs");
  if (wsLogs) {
    origemExec = "Logs (cada status de fechamento, uma vez por mês)";
    const tipoPorId: { [id: string]: string } = {};
    const cId = cab.indexOf("id");
    for (let i = 1; i < afaz.length; i++) tipoPorId[String(afaz[i][cId]).trim()] = String(afaz[i][cTp] || "").trim();

    const lg = wsLogs.getUsedRange().getValues();
    const lc = lg[0].map(c => String(c).trim());
    const lA = lc.indexOf("AFAZ"), lR = lc.indexOf("Registro"), lAc = lc.indexOf("Ação"), lD = lc.indexOf("Data");
    const NOMES: { [k: string]: string } = { "finalizado": "Finalizado", "resolvido": "Resolvido", "pendência claro": "Pendência Claro", "pendencia claro": "Pendência Claro" };
    // gráfico "Status por mês": conta CADA mudança de status registrada (sem juntar repetidos)
    const NOME_LOG: { [k: string]: string } = { "devolvido": "Devolvido", "finalizado": "Finalizado", "resolvido": "Resolvido", "pendencia claro": "Pendência Claro" };
    const semAcento = (s: string) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toLowerCase();
    for (let i = 1; i < lg.length; i++) {
      if (lAc >= 0 && String(lg[i][lAc]).trim() !== "Alterado Status") continue;
      const reg = String(lg[i][lR] || "");
      const p = reg.indexOf(" para ");
      if (p < 0) continue;
      const resto = reg.slice(p + 6), e = resto.indexOf(" em ");
      const st = NOME_LOG[semAcento(e < 0 ? resto : resto.slice(0, e))];
      if (!st) continue;
      const tp = tipoPorId[String(lg[i][lA]).trim()] || "Sem tipo";
      somaLog(mesDe(lg[i][lD]), st, tp);
    }

    const vistos: { [k: string]: boolean } = {};
    for (let i = 1; i < lg.length; i++) {
      if (lAc >= 0 && String(lg[i][lAc]).trim() !== "Alterado Status") continue;
      const reg = String(lg[i][lR] || "");
      const p = reg.indexOf(" para ");
      if (p < 0) continue;
      const resto = reg.slice(p + 6), e = resto.indexOf(" em ");
      if (e < 0) continue;
      const nomeSt = NOMES[resto.slice(0, e).trim().toLowerCase()];
      if (!nomeSt) continue;                       // não é status de fechamento
      const id = String(lg[i][lA]).trim();
      const tp = tipoPorId[id];
      if (tp !== "Backlog" && tp !== "Entrante") continue;   // AFAZ fora da aba Afaz
      const mes = mesDe(lg[i][lD]);
      if (!mes) continue;
      const chave = id + "|" + nomeSt + "|" + mes;
      if (vistos[chave]) continue;                 // mesmo status no mesmo mês: conta 1
      vistos[chave] = true;
      contar(tp === "Entrante", nomeSt, mes);
    }
  } else {
    for (let i = 1; i < afaz.length; i++) {
      const st = String(afaz[i][cSt] || "").trim().toUpperCase();
      if (!FECHADO.includes(st)) continue;
      const mes = mesDe(afaz[i][cDt]);
      if (!mes) continue;
      contar(String(afaz[i][cTp]).trim() === "Entrante", String(afaz[i][cSt]).trim(), mes);
    }
  }

  // ---------- Meta (aba Programação) ----------
  const programado: { [k: string]: number } = {};
  const MES: { [k: string]: string } = {
    JAN: "01", FEV: "02", MAR: "03", ABR: "04", MAI: "05", JUN: "06",
    JUL: "07", AGO: "08", SET: "09", OUT: "10", NOV: "11", DEZ: "12"
  };

  // Converte o rótulo do mês em "AAAA-MM". Aceita: JUL_2026, jul/26, Julho 2026, 07/2026, 2026-07 ou data do Excel
  const chaveMes = (v: string | number | boolean): string => {
    if (typeof v === "number" && v > 20000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0");
    }
    const t = String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
    let m = t.match(/^([A-Z]{3})[A-Z]*[\s_\/\-.]*(\d{4}|\d{2})$/);
    if (m && MES[m[1]]) return (m[2].length === 2 ? "20" + m[2] : m[2]) + "-" + MES[m[1]];
    m = t.match(/^(\d{1,2})[\/\-](\d{4})$/);
    if (m) return m[2] + "-" + m[1].padStart(2, "0");
    m = t.match(/^(\d{4})-(\d{2})/);
    if (m) return m[1] + "-" + m[2];
    return "";
  };

  const prog = workbook.getWorksheet("Programação").getUsedRange().getValues();

  // Layout 1: meses no cabeçalho (linha 1) e meta na linha 2
  for (let c = 0; c < prog[0].length; c++) {
    const k = chaveMes(prog[0][c]);
    if (k && prog.length > 1) programado[k] = Number(prog[1][c]) || 0;
  }

  // Layout 2 (se o 1 não achou nada): mês na coluna A e meta na coluna B
  if (Object.keys(programado).length === 0) {
    for (let r = 0; r < prog.length; r++) {
      const k = chaveMes(prog[r][0]);
      if (k) programado[k] = Number(prog[r][1]) || 0;
    }
  }

  return { programado, execBacklog, execEntrante, execPorStatus, semanal, origemExec, logsMes };
}