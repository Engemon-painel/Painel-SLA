// Office Script — Histórico do mês (botão "Exportar base do mês" do painel, tela ESER)
// NOVO (25/11/2026): lê a aba "Historico" e devolve o historico_mes.json com TODAS as
// colunas da aba, só com os incidentes do MÊS ATUAL. O mês do incidente vem da
// CLEAR_DATE (se vazia, da RESOLVED_DATE) — igual à coluna "Mes" do Power Query.
// Se o mês de hoje ainda não tiver incidentes (dia 1º), usa o último mês da base.
// Telefone e e-mail do fornecedor NÃO são exportados (o repositório é público).
function main(workbook: ExcelScript.Workbook): string {
    const NOME_ABA = "Historico";
    const FORA: string[] = ["TELEFONE_FORNECEDOR", "EMAIL_FORNECEDOR"];

    const semAcento = (s: string): string => s.normalize("NFD").replace(/[̀-ͯ]/g, "");
    const lim = (v: string | number | boolean): string => String(v ?? "").trim();
    const norm = (v: string | number | boolean): string => semAcento(lim(v)).toUpperCase().replace(/\s+/g, "_");
    const dois = (n: number): string => (n < 10 ? "0" : "") + n;

    // abas: "Historico" (sem acento / maiúsculas tanto faz)
    const abas = workbook.getWorksheets();
    let ws: ExcelScript.Worksheet | undefined = undefined;
    for (let k = 0; k < abas.length; k++) {
        if (norm(abas[k].getName()) === norm(NOME_ABA)) { ws = abas[k]; break; }
    }
    if (!ws) return JSON.stringify({ erro: "Aba não encontrada: " + NOME_ABA });
    const usado = ws.getUsedRange(true);
    if (!usado) return JSON.stringify({ erro: "Aba Historico vazia" });
    const v = usado.getValues();

    // cabeçalho = linha com INCIDENT_NUMBER (nas primeiras 20)
    let lh = -1;
    for (let r = 0; r < Math.min(20, v.length); r++) {
        if (v[r].some(c => norm(c) === "INCIDENT_NUMBER")) { lh = r; break; }
    }
    if (lh < 0) return JSON.stringify({ erro: "Não achei a coluna INCIDENT_NUMBER na aba Historico" });
    const cab = v[lh].map(c => lim(c));
    const iInc = cab.findIndex(c => norm(c) === "INCIDENT_NUMBER");
    const iClear = cab.findIndex(c => norm(c) === "CLEAR_DATE");
    const iResol = cab.findIndex(c => norm(c) === "RESOLVED_DATE");
    // colunas que vão para o arquivo (todas com título, menos telefone/e-mail)
    const cols: number[] = [];
    cab.forEach((c, i) => { if (c && FORA.indexOf(norm(c)) < 0) cols.push(i); });

    // número do Excel → data. Colunas de data ficam como "dd/mm/aaaa hh:mm:ss"
    const ehData = (c: string): boolean => /DATE|^DT_|DATA/.test(norm(c));
    const dataTxt = (x: number): string => {
        const d = new Date(Date.UTC(1899, 11, 30) + Math.round(x * 86400000));
        return dois(d.getUTCDate()) + "/" + dois(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear() + " " +
            dois(d.getUTCHours()) + ":" + dois(d.getUTCMinutes()) + ":" + dois(d.getUTCSeconds());
    };
    const mesDe = (x: string | number | boolean): string => {
        if (typeof x === "number" && x > 30000 && x < 80000) {
            const d = new Date(Date.UTC(1899, 11, 30) + Math.round(x * 86400000));
            return d.getUTCFullYear() + "-" + dois(d.getUTCMonth() + 1);
        }
        const m = lim(x).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (m) return m[3] + "-" + dois(Number(m[2]));
        const m2 = lim(x).match(/^(\d{4})-(\d{2})/);
        return m2 ? m2[1] + "-" + m2[2] : "";
    };

    // mês de cada linha + o mês a exportar
    const linhasMes: { r: number, mes: string }[] = [];
    const meses: { [m: string]: boolean } = {};
    for (let r = lh + 1; r < v.length; r++) {
        if (!lim(v[r][iInc])) continue;
        let mes = iClear >= 0 ? mesDe(v[r][iClear]) : "";
        if (!mes && iResol >= 0) mes = mesDe(v[r][iResol]);
        if (!mes) continue;
        meses[mes] = true;
        linhasMes.push({ r: r, mes: mes });
    }
    const hoje = new Date();
    let alvo = hoje.getFullYear() + "-" + dois(hoje.getMonth() + 1);
    if (!meses[alvo]) alvo = Object.keys(meses).sort().pop() || alvo;

    // uma linha por incidente (igual ao "Remover duplicatas" da consulta)
    const vistos: { [inc: string]: boolean } = {};
    const linhas: (string | number | boolean)[][] = [];
    linhasMes.forEach(x => {
        if (x.mes !== alvo) return;
        const l = v[x.r];
        const inc = lim(l[iInc]).toUpperCase();
        if (vistos[inc]) return;
        vistos[inc] = true;
        linhas.push(cols.map(i => {
            const val = l[i];
            if (typeof val === "number" && ehData(cab[i]) && val > 30000 && val < 80000) return dataTxt(val);
            return typeof val === "string" ? val.trim() : val;
        }));
    });

    return JSON.stringify({
        geradoEm: hoje.getFullYear() + "-" + dois(hoje.getMonth() + 1) + "-" + dois(hoje.getDate()),
        fonte: "aba Historico (base_historico.xlsx)",
        mes: alvo,
        total: linhas.length,
        colunas: cols.map(i => cab[i]),
        linhas: linhas
    });
}
