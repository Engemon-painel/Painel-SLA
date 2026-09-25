// Office Script — Exportar Base NR para o Painel do ESER
// Lê as abas BASE, Base_Supervisores e Preenchimento da Base_NR.xlsx e
// devolve um JSON (texto) com SÓ as colunas que o painel usa.
// O CPF NÃO é exportado.
// No Power Automate: "Executar script" → o resultado vai para base_nr.json no GitHub.
function main(workbook: ExcelScript.Workbook): string {
  const norm = (s: string | number | boolean): string =>
    String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/\s+/g, " ").trim();

  const CURSOS: string[] = ["NR6", "NR12", "NR17", "NR20", "NR18", "NR26", "SEP", "NR10", "NR35", "DIREÇÃO", "FICHA DE EPI", "INTEGRAÇÃO", "O.S"];
  const COLUNAS: { [aba: string]: string[] } = {
    "BASE": ["Matricula", "Nome", "Data Admis.", "Desc. Depto", "Status", "Área Técnica", "Supervisor", "Coordenador"]
      .concat(CURSOS).concat(["OBSERVAÇÃO", "Situação do Nr"]),
    "Base_Supervisores": ["NOME", "GESTÃO", "FUNÇÃO"].concat(CURSOS).concat(["OBSERVAÇÃO"]),
    "Preenchimento": ["Colaborador", "Supervisor", "Vencimento do Aso", "Data do Exame", "Status", "Observações"]
  };

  const abas: { [aba: string]: (string | number | boolean)[][] } = {};
  const planilhas = workbook.getWorksheets();

  for (const nomeAba of Object.keys(COLUNAS)) {
    const ws = planilhas.find(w => norm(w.getName()) === norm(nomeAba));
    if (!ws) throw new Error("Aba não encontrada: " + nomeAba);
    const usado = ws.getUsedRange(true);
    const linhas: (string | number | boolean)[][] = [COLUNAS[nomeAba]];
    if (usado) {
      const valores = usado.getValues();
      const cab = valores[0].map(h => String(h).replace(/\s+/g, " ").trim());
      const idx = COLUNAS[nomeAba].map(c => cab.findIndex(h => norm(h) === norm(c)));
      for (let i = 1; i < valores.length; i++) {
        const l = valores[i];
        if (l.every(v => v === "" || v === null)) continue;
        linhas.push(idx.map(j => (j >= 0 ? l[j] : "")));
      }
    }
    abas[nomeAba] = linhas;
  }

  const hoje = new Date();
  const dois = (n: number): string => (n < 10 ? "0" : "") + n;
  const atualizado_em = hoje.getFullYear() + "-" + dois(hoje.getMonth() + 1) + "-" + dois(hoje.getDate());
  return JSON.stringify({ atualizado_em: atualizado_em, abas: abas });
}
