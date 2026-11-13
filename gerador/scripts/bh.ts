/**
 * dados_bh  (Office Script — planilha "Controle de Ponto - HE - BH")
 * ---------------------------------------------------------------------------
 * Gera o dados_bh.json da Visão Geral / Funcionários por Faixa / Evolução,
 * no MESMO formato que o script original gerava, a partir da aba
 * "Controle_de_Ponto - HE - BH". Publique o resultado como dados_bh.json.
 * (As pendências ficam em outro script, que gera o dados_pend_bh.json.)
 */
const ABA_BH = "Controle_de_Ponto - HE - BH";
const COLS_IDADE = ["BH_0_30_h", "BH_31_60_h", "BH_61_90_h", "BH_91_120_h", "BH_121_150_h", "BH_151_180_h", "BH_Maior_180_h", "BH_PreBase_h"];

interface Colab {
  n: string; m: string; sup: string; co: string; ge: string; projeto: string; centroCusto: string; st: string; status: string;
  h: number; saldoTotal: string; tp: string; fx: string; faixaMaisCritica: string; faixaNegativa: string; cr: string;
  criticidadeNegativa: string; bh: number[];
}

function main(workbook: ExcelScript.Workbook): string {
  const ws = workbook.getWorksheet(ABA_BH);
  if (!ws) throw new Error("Aba '" + ABA_BH + "' não encontrada.");
  const tabs = ws.getTables();
  const v = tabs.length ? tabs[0].getRange().getValues() : ws.getUsedRange().getValues();
  const cab = v[0].map(c => String(c).trim());
  const ix = (nome: string) => cab.indexOf(nome);
  const txt = (row: (string | number | boolean)[], nome: string) => { const k = ix(nome); return k < 0 ? "" : String(row[k] == null ? "" : row[k]).trim(); };
  const num = (row: (string | number | boolean)[], nome: string) => { const k = ix(nome); return k < 0 ? 0 : (Number(row[k]) || 0); };
  const arred = (x: number) => Math.round(x * 100) / 100;

  const itens: Colab[] = [];
  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    const n = txt(row, "Nome");
    if (!n) continue;
    const st = txt(row, "Status");
    itens.push({
      n, m: txt(row, "Matricula_6"), sup: txt(row, "Supervisor"), co: txt(row, "Coordenador"), ge: txt(row, "Gerente"),
      projeto: txt(row, "Projeto"), centroCusto: txt(row, "Centro_Custo"), st, status: st,
      h: arred(num(row, "Saldo_Atual_h")), saldoTotal: txt(row, "Saldo_Total"), tp: txt(row, "Tipo_Saldo"),
      fx: txt(row, "Faixa_Final"),
      faixaMaisCritica: txt(row, "Faixa_Mais_Critica"), faixaNegativa: txt(row, "Faixa_Negativa"),
      cr: txt(row, "Criticidade_Final"), criticidadeNegativa: txt(row, "Criticidade_Negativa"),
      bh: COLS_IDADE.map(c => arred(num(row, c)))
    });
  }

  const resumoStatus: { [st: string]: { quantidade: number; saldoHoras: number } } = {};
  itens.forEach(i => {
    const k = i.st || "Não informado";
    if (!resumoStatus[k]) resumoStatus[k] = { quantidade: 0, saldoHoras: 0 };
    resumoStatus[k].quantidade++;
    resumoStatus[k].saldoHoras = Math.round((resumoStatus[k].saldoHoras + i.h) * 100) / 100;
  });

  return JSON.stringify({ geradoEm: new Date().toISOString(), fonte: ABA_BH, totalFuncionarios: itens.length, resumoStatus, itens });
}
