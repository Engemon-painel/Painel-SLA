/**
 * Office Script — Combustível (aba "Transações")  ·  formato v3
 *
 * Gera a análise de combustível para QUALQUER contrato (grupo), mais o total "TODOS".
 *  - Valor: VALOR EMISSAO · Mês: DATA TRANSACAO · separação por TIPO FROTA.
 *  - Top 10 técnicos (NOME MOTORISTA) só TIPO FROTA = FROTA; Top 10 gerador só GMG.
 *  - Por técnico (todos): valor de FROTA e de GMG + supervisor (coluna com "SUPERVISOR"
 *    no nome, se existir). O painel monta o ranking por supervisor e os 5 técnicos
 *    de cada um; sem a coluna, ele usa o supervisor do Organograma.
 *  - Por placa (só FROTA, todos os contratos): valor, litros, VL/litro, último
 *    hodômetro, KM rodados e KM/litro. KM > 3000 num abastecimento é ignorado.
 *
 * Grupos: o nome do contrato da planilha de combustível é ligado ao nome do
 * contrato do cadastro de frota (Operação). Regras, nesta ordem:
 *   1) MAPA_CONTRATOS abaixo (nomes que não batem automaticamente)
 *   2) o nome começa com o nome de um contrato da frota (CONTRATOS_FROTA)
 *   3) senão, o próprio nome do contrato vira um grupo
 */
const CONTRATOS_FROTA = [
  "CLARO INFRA SP", "CLARO INFRA RJ", "CLARO ARCON", "CIRION INFRA", "CIX - POUPATEMPO PR",
  "VOLANTE", "FEMSA NORTE", "MOTOROLA", "SISTEMAS/GESTAO", "CARE PLUS", "CMOC",
  "BRADESCO SPOT", "BRESCO VIRA COPOS", "EQUINIX SP4", "INDUSTRIAL", "SIEMENS"
];
const MAPA_CONTRATOS: Record<string, string> = {
  "CLARO - CONTRATO INFRA SPC": "CLARO INFRA SP",
  "CIRION -CONTRATO DE MANUT INFRA 1011233": "CIRION INFRA",
  "FEMSA MANUT D.NORTE 1009294 FILIAL": "FEMSA NORTE",
  "SISTEMAS": "SISTEMAS/GESTAO",
  "BRESCO - MANUTENCAO PREDIAL VIRA COPOS 1011509": "BRESCO VIRA COPOS"
};

function main(workbook: ExcelScript.Workbook): string {
  let ws = workbook.getWorksheet("Transações");
  if (!ws) {
    let melhor: ExcelScript.Worksheet | undefined = undefined;
    let maxLin = 0;
    for (const s of workbook.getWorksheets()) {
      const r = s.getUsedRange();
      if (!r) continue;
      const cab = (r.getRow(0).getValues()[0] as (string | number | boolean)[]).map(v => norm(v));
      if (cab.indexOf("TIPO FROTA") >= 0 && cab.indexOf("VALOR EMISSAO") >= 0 && r.getRowCount() > maxLin) {
        melhor = s; maxLin = r.getRowCount();
      }
    }
    ws = melhor;
  }
  if (!ws) return JSON.stringify({ erro: "Aba 'Transações' não encontrada" });
  const usado = ws.getUsedRange();
  if (!usado) return JSON.stringify({ erro: "Aba vazia" });
  return JSON.stringify(processar(usado.getValues() as (string | number | boolean)[][]));
}

// ---------- processamento ----------
function processar(valores: (string | number | boolean)[][]) {
  const TOP_N = 10, TOP_SUP = 5, KM_MAX = 3000;
  const cab = valores[0].map(v => norm(v));
  const col = (nome: string) => cab.indexOf(nome);
  const iData = col("DATA TRANSACAO"), iTipo = col("TIPO FROTA"), iMot = col("NOME MOTORISTA"),
    iContr = col("CONTRATO"), iValor = col("VALOR EMISSAO"), iPlaca = col("PLACA"), iLitros = col("LITROS"),
    iHod = col("HODOMETRO OU HORIMETRO"), iKm = col("KM RODADOS OU HORAS TRABALHADAS"), iSup = cab.findIndex(c => c.indexOf("SUPERV") >= 0);

  const faltando = [["DATA TRANSACAO", iData], ["TIPO FROTA", iTipo], ["NOME MOTORISTA", iMot],
  ["Contrato", iContr], ["VALOR EMISSAO", iValor]].filter(p => (p[1] as number) < 0).map(p => p[0]);
  if (faltando.length) return { erro: "Colunas não encontradas: " + faltando.join(", ") };

  const mapa: Record<string, string> = {};
  Object.keys(MAPA_CONTRATOS).forEach(k => mapa[norm(k)] = norm(MAPA_CONTRATOS[k]));
  const frotaPorTamanho = CONTRATOS_FROTA.map(c => norm(c)).sort((a, b) => b.length - a.length);
  const cacheGrupo: Record<string, string> = {};
  const grupoDe = (contrato: string) => {
    if (cacheGrupo[contrato] !== undefined) return cacheGrupo[contrato];
    let g = mapa[contrato];
    if (!g) g = frotaPorTamanho.find(f => contrato === f || contrato.indexOf(f + " ") === 0 || contrato.indexOf(f + "-") === 0) || "";
    if (!g) g = contrato || "SEM CONTRATO";
    cacheGrupo[contrato] = g;
    return g;
  };

  type Rank = Record<string, { valor: number; qtd: number; litros: number; placas: Record<string, boolean> }>;
  type Acum = {
    total: number; qtd: number; litros: number; contratos: Record<string, boolean>;
    porTipo: Record<string, { valor: number; qtd: number; litros: number }>; mot: Rank; motGmg: Rank; sup: Rank;
    pessoas: Record<string, { frota: number; gmg: number; qf: number; qg: number; lf: number; lg: number; sups: Record<string, number> }>
  };
  const novo = (): Acum => ({ total: 0, qtd: 0, litros: 0, contratos: {}, porTipo: {}, mot: {}, motGmg: {}, sup: {}, pessoas: {} });
  // grupos[g].meses[m] e grupos[g].geral
  const grupos: Record<string, { meses: Record<string, Acum>; geral: Acum }> = {};

  type Det = { valor: number; litros: number; km: number; litrosKm: number; hod: number; ordHod: number };
  const placasMes: Record<string, Record<string, Det>> = {};
  const placasGeral: Record<string, Det> = {};
  const somaDet = (m: Record<string, Det>, placa: string, valor: number, litros: number, km: number, hod: number, ord: number) => {
    if (!m[placa]) m[placa] = { valor: 0, litros: 0, km: 0, litrosKm: 0, hod: 0, ordHod: -1 };
    const d = m[placa];
    d.valor += valor; d.litros += litros;
    if (km > 0 && km <= KM_MAX) { d.km += km; d.litrosKm += litros; }
    if (hod > 0 && hod < 999999 && ord >= d.ordHod) { d.hod = hod; d.ordHod = ord; }
  };
  const somaRank = (r: Rank, nome: string, valor: number, litros: number, placa: string) => {
    if (!r[nome]) r[nome] = { valor: 0, qtd: 0, litros: 0, placas: {} };
    r[nome].valor += valor; r[nome].qtd++; r[nome].litros += litros;
    if (placa) r[nome].placas[placa] = true;
  };

  const todosMeses: Record<string, boolean> = {};
  let ignoradasData = 0, temSupervisor = false;

  for (let i = 1; i < valores.length; i++) {
    const l = valores[i];
    const mes = mesDe(l[iData]);
    if (!mes) { ignoradasData++; continue; }
    todosMeses[mes] = true;

    const valor = num(l[iValor]);
    const litros = iLitros >= 0 ? num(l[iLitros]) : 0;
    const tipo = norm(l[iTipo]) || "SEM TIPO";
    const mot = String(l[iMot] ?? "").trim() || "SEM MOTORISTA";
    const sup = iSup >= 0 ? String(l[iSup] ?? "").trim() : "";
    if (sup) temSupervisor = true;
    const placa = iPlaca >= 0 ? String(l[iPlaca] ?? "").trim().toUpperCase() : "";
    const contrato = norm(l[iContr]);
    const g = grupoDe(contrato);

    if (placa && tipo === "FROTA") {
      const km = iKm >= 0 ? num(l[iKm]) : 0, hod = iHod >= 0 ? num(l[iHod]) : 0;
      const ord = typeof l[iData] === "number" ? (l[iData] as number) : i;
      if (!placasMes[mes]) placasMes[mes] = {};
      somaDet(placasMes[mes], placa, valor, litros, km, hod, ord);
      somaDet(placasGeral, placa, valor, litros, km, hod, ord);
    }

    for (const chave of [g, "TODOS"]) {
      if (!grupos[chave]) grupos[chave] = { meses: {}, geral: novo() };
      if (!grupos[chave].meses[mes]) grupos[chave].meses[mes] = novo();
      for (const a of [grupos[chave].meses[mes], grupos[chave].geral]) {
        a.total += valor; a.qtd++; a.litros += litros; a.contratos[contrato] = true;
        if (!a.porTipo[tipo]) a.porTipo[tipo] = { valor: 0, qtd: 0, litros: 0 };
        a.porTipo[tipo].valor += valor; a.porTipo[tipo].qtd++; a.porTipo[tipo].litros += litros;
        if (tipo === "FROTA") {
          somaRank(a.mot, mot, valor, litros, placa);
          if (sup) somaRank(a.sup, sup, valor, litros, placa);
        }
        if (tipo === "GMG") somaRank(a.motGmg, mot, valor, litros, placa);
        if (tipo === "FROTA" || tipo === "GMG") {
          if (!a.pessoas[mot]) a.pessoas[mot] = { frota: 0, gmg: 0, qf: 0, qg: 0, lf: 0, lg: 0, sups: {} };
          const pe = a.pessoas[mot];
          if (tipo === "FROTA") { pe.frota += valor; pe.qf++; pe.lf += litros; } else { pe.gmg += valor; pe.qg++; pe.lg += litros; }
          if (sup) pe.sups[sup] = (pe.sups[sup] || 0) + 1;
        }
      }
    }
  }

  const r2 = (n: number) => Math.round(n * 100) / 100;
  const listaMeses = Object.keys(todosMeses).sort();
  const topDe = (r: Rank, n: number) => Object.keys(r).sort((x, y) => r[y].valor - r[x].valor).slice(0, n)
    .map(k => ({ nome: k, motorista: k, valor: r2(r[k].valor), qtd: r[k].qtd, litros: r2(r[k].litros), placas: Object.keys(r[k].placas) }));
  const fechar = (a: Acum) => ({
    total: r2(a.total), qtdAbastecimentos: a.qtd, litros: r2(a.litros),
    contratosOrigem: Object.keys(a.contratos),
    porTipoFrota: Object.keys(a.porTipo).sort((x, y) => a.porTipo[y].valor - a.porTipo[x].valor)
      .map(t => ({ tipo: t, valor: r2(a.porTipo[t].valor), qtd: a.porTipo[t].qtd, litros: r2(a.porTipo[t].litros) })),
    topMotoristasFrota: topDe(a.mot, TOP_N),
    topMotoristasGmg: topDe(a.motGmg, TOP_N),
    topSupervisores: topDe(a.sup, TOP_SUP),
    // todos os técnicos com uso de frota ou gerador (o painel agrupa por supervisor)
    tecnicos: Object.keys(a.pessoas).map(n => {
      const pe = a.pessoas[n];
      const supMaisUsado = Object.keys(pe.sups).sort((x, y) => pe.sups[y] - pe.sups[x])[0] || "";
      return {
        nome: n, frota: r2(pe.frota), gmg: r2(pe.gmg), qtdFrota: pe.qf, qtdGmg: pe.qg,
        litrosFrota: r2(pe.lf), litrosGmg: r2(pe.lg), supervisor: supMaisUsado
      };
    }).sort((x, y) => (y.frota + y.gmg) - (x.frota + x.gmg))
  });

  const saidaGrupos: Record<string, object> = {};
  for (const g of Object.keys(grupos)) {
    const G = grupos[g];
    const porMes: Record<string, object> = {};
    for (const m of Object.keys(G.meses).sort()) porMes[m] = fechar(G.meses[m]);
    const geral = fechar(G.geral);
    // no "geral", cada ranking traz o valor de cada mês (gráfico "Todos" mês a mês)
    const comMeses = (lista: ReturnType<typeof topDe>, chave: "mot" | "motGmg" | "sup") => lista.map(t => {
      const pm: Record<string, number> = {};
      for (const m of listaMeses) { const a = G.meses[m]; pm[m] = r2(a && a[chave][t.nome] ? a[chave][t.nome].valor : 0); }
      return { ...t, porMes: pm };
    });
    geral.topMotoristasFrota = comMeses(geral.topMotoristasFrota, "mot");
    geral.topMotoristasGmg = comMeses(geral.topMotoristasGmg, "motGmg");
    geral.topSupervisores = comMeses(geral.topSupervisores, "sup");
    saidaGrupos[g] = { porMes, geral };
  }

  const listaPlacas = (m: Record<string, Det>) => Object.keys(m || {}).sort((x, y) => m[y].valor - m[x].valor).map(p => {
    const d = m[p];
    return {
      Placa: p, "Valor Total": r2(d.valor), tipoFrota: "FROTA", litros: r2(d.litros),
      vlLitro: d.litros > 0 ? Math.round(d.valor / d.litros * 1000) / 1000 : 0,
      hodometro: d.hod, kmRodados: r2(d.km), kmLitro: d.litrosKm > 0 ? r2(d.km / d.litrosKm) : 0
    };
  });
  const placasPorMes: Record<string, object> = {};
  for (const m of listaMeses) placasPorMes[m] = listaPlacas(placasMes[m]);

  return {
    versao: 3,
    atualizadoEm: new Date().toISOString(),
    meses: listaMeses,
    temSupervisor: temSupervisor,
    colunaSupervisor: iSup >= 0 ? cab[iSup] : "",
    grupos: saidaGrupos,
    placas: { porMes: placasPorMes, geral: listaPlacas(placasGeral) },
    linhasSemData: ignoradasData
  };
}

// ---------- utilitários ----------
function norm(v: string | number | boolean | undefined | null): string {
  return String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
}
function num(v: string | number | boolean | undefined | null): number {
  if (typeof v === "number") return v;
  let s = String(v ?? "").replace(/[R$\s]/g, "");
  if (!s) return 0;
  if (s.indexOf(",") >= 0) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function mesDe(v: string | number | boolean | undefined | null): string {
  if (typeof v === "number" && v > 0) {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0");
  }
  const s = String(v ?? "").trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + "-" + m[2].padStart(2, "0");
  m = s.match(/^(\d{4})-(\d{2})/);
  if (m) return m[1] + "-" + m[2];
  return "";
}