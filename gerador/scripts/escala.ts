// Office Script — Exportar Escala para o Painel
// Lê a aba "ESCALA PAULO" e devolve o conteúdo do dados_escala.json.
// Telefone (Contato) e Residência NÃO são exportados.
// No Power Automate: "Executar script" → o resultado vai para dados_escala.json no GitHub.
function main(workbook: ExcelScript.Workbook): string {
    const NOME_ABA = "ESCALA PAULO";

    const semAcento = (s: string): string => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const lim = (v: string | number | boolean): string => String(v ?? "").replace(/\s+/g, " ").trim();
    const norm = (v: string | number | boolean): string => semAcento(lim(v)).toUpperCase();
    const dois = (n: number): string => (n < 10 ? "0" : "") + n;

    const PART: string[] = ["DE", "DA", "DO", "DOS", "DAS", "E"];
    const titulo = (v: string | number | boolean): string => {
        const s = lim(v);
        if (!s) return "";
        return s.toUpperCase().split(" ").map((w, i) =>
            (i > 0 && PART.indexOf(w) >= 0) ? w.toLowerCase() : w.charAt(0) + w.slice(1).toLowerCase()
        ).join(" ");
    };
    const horario = (v: string | number | boolean): string => {
        const m = String(v ?? "").match(/(\d{1,2}):(\d{2})/g);
        if (!m || m.length < 2) return lim(v);
        const f = (x: string): string => { const p = x.split(":"); return dois(Number(p[0])) + ":" + p[1]; };
        return f(m[0]) + "–" + f(m[1]);
    };
    // data do cabeçalho: número (data do Excel) ou texto dd/mm/aaaa
    const dataIso = (v: string | number | boolean): string => {
        if (typeof v === "number" && v > 30000 && v < 80000) {
            const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
            return d.getUTCFullYear() + "-" + dois(d.getUTCMonth() + 1) + "-" + dois(d.getUTCDate());
        }
        const m = String(v ?? "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        return m ? m[3] + "-" + m[2] + "-" + m[1] : "";
    };

    // acha a aba (laço simples: o Office Script no Power Automate não aceita .find com chamada ao Excel dentro)
    const planilhas = workbook.getWorksheets();
    let ws: ExcelScript.Worksheet | undefined = undefined;
    for (let k = 0; k < planilhas.length; k++) {
        if (norm(planilhas[k].getName()) === norm(NOME_ABA)) { ws = planilhas[k]; break; }
    }
    if (!ws) throw new Error("Aba não encontrada: " + NOME_ABA);

    const valores = ws.getUsedRange(true).getValues();

    // linha do cabeçalho = a que tem "Nome" (procura nas primeiras 20 linhas)
    let lh = -1;
    for (let r = 0; r < Math.min(20, valores.length); r++) {
        if (valores[r].some(c => norm(c) === "NOME")) { lh = r; break; }
    }
    if (lh < 0) throw new Error("Não achei a linha de cabeçalho (coluna Nome).");
    const cab = valores[lh];

    const col = (nome: string): number => cab.findIndex(c => norm(c) === semAcento(nome).toUpperCase());
    const iNome = col("Nome"), iSup = col("Supervisão"), iCoord = col("Coordenação"), iArea = col("Área"),
        iTurno = col("Turno"), iEscala = col("Escala"), iHor = col("Horário"), iAtu = col("Atuação");

    // colunas de data, em ordem cronológica
    const colsData: { iso: string, i: number }[] = [];
    cab.forEach((c, i) => { const iso = dataIso(c); if (iso) colsData.push({ iso: iso, i: i }); });
    colsData.sort((a, b) => (a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : 0));

    const pega = (linha: (string | number | boolean)[], i: number): string | number | boolean => (i >= 0 ? linha[i] : "");

    const tecnicos: {
        nome: string, duplado: boolean, supervisao: string, coordenacao: string, area: string,
        turno: string, escala: string, horario: string, atuacao: string, codigos: string[]
    }[] = [];

    for (let r = lh + 1; r < valores.length; r++) {
        const l = valores[r];
        let nome = lim(pega(l, iNome));
        if (!nome) continue;
        const duplado = /DUPLADO/i.test(nome);
        nome = nome.replace(/\s*[-(]\s*DUPLADO\s*\)?/i, "").trim();
        tecnicos.push({
            nome: nome,
            duplado: duplado,
            supervisao: titulo(pega(l, iSup)),
            coordenacao: titulo(pega(l, iCoord)),
            area: lim(pega(l, iArea)).toUpperCase(),
            turno: titulo(pega(l, iTurno)),
            escala: lim(pega(l, iEscala)).toUpperCase().replace("ESPANHOLA", "Espanhola"),
            horario: horario(pega(l, iHor)),
            atuacao: lim(pega(l, iAtu)),
            codigos: colsData.map(c => lim(l[c.i]).toUpperCase())
        });
    }

    const hoje = new Date();
    const atualizado_em = hoje.getFullYear() + "-" + dois(hoje.getMonth() + 1) + "-" + dois(hoje.getDate());
    return JSON.stringify({
        atualizado_em: atualizado_em,
        fonte: NOME_ABA,
        datas: colsData.map(c => c.iso),
        tecnicos: tecnicos
    });
}
