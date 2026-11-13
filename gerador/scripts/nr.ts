// Office Script — Base_NR (Gestão de NR / ASO)
// ALTERADO (07/10/2026): a aba "BASE" sai SEM duplicados — uma linha por
// Matrícula (sem matrícula, pelo Nome). Quando a mesma pessoa aparece mais de
// uma vez, fica a linha ATIVA; empatando, a que tem mais cursos preenchidos.
// Células vazias da linha mantida são completadas com as da linha descartada.
// As abas "Base_Supervisores" e "Preenchimento" não mudam.

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
        let ws: ExcelScript.Worksheet | undefined = undefined;
        for (let k = 0; k < planilhas.length; k++) {
            if (norm(planilhas[k].getName()) === norm(nomeAba)) { ws = planilhas[k]; break; }
        }
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

    // NOVO: remove duplicados da aba BASE (uma linha por matrícula).
    abas["BASE"] = deduplicarBase(abas["BASE"], COLUNAS["BASE"], CURSOS, norm);

    const hoje = new Date();
    const dois = (n: number): string => (n < 10 ? "0" : "") + n;
    const atualizado_em = hoje.getFullYear() + "-" + dois(hoje.getMonth() + 1) + "-" + dois(hoje.getDate());
    return JSON.stringify({ atualizado_em: atualizado_em, abas: abas });
}

function deduplicarBase(
    linhas: (string | number | boolean)[][],
    colunas: string[],
    cursos: string[],
    norm: (s: string | number | boolean) => string
): (string | number | boolean)[][] {
    const iMat = colunas.indexOf("Matricula");
    const iNome = colunas.indexOf("Nome");
    const iStatus = colunas.indexOf("Status");
    const iAdmis = colunas.indexOf("Data Admis.");
    const iCursos = cursos.map(c => colunas.indexOf(c));
    const vazio = (v: string | number | boolean) => v === "" || v === null || v === undefined;

    const chave = (l: (string | number | boolean)[]): string => {
        const mat = String(l[iMat] ?? "").replace(/\.0+$/, "").replace(/^0+/, "").trim();
        if (mat !== "") return "M:" + norm(mat);
        const nome = norm(l[iNome]);
        return nome !== "" ? "N:" + nome : "";
    };
    const ativo = (l: (string | number | boolean)[]) => norm(l[iStatus]) === "ATIVO";
    const cursosPreenchidos = (l: (string | number | boolean)[]) => iCursos.filter(i => !vazio(l[i])).length;
    const admissao = (l: (string | number | boolean)[]) => typeof l[iAdmis] === "number" ? (l[iAdmis] as number) : 0;

    // devolve a linha que deve ficar entre duas duplicadas
    const melhor = (a: (string | number | boolean)[], b: (string | number | boolean)[]) => {
        if (ativo(a) !== ativo(b)) return ativo(a) ? a : b;
        if (cursosPreenchidos(a) !== cursosPreenchidos(b)) return cursosPreenchidos(a) > cursosPreenchidos(b) ? a : b;
        return admissao(b) > admissao(a) ? b : a;
    };

    const saida: (string | number | boolean)[][] = [linhas[0]];
    const posicao: { [k: string]: number } = {};
    for (let i = 1; i < linhas.length; i++) {
        const l = linhas[i];
        const k = chave(l);
        if (k === "" || posicao[k] === undefined) {
            if (k !== "") posicao[k] = saida.length;
            saida.push(l.slice());
            continue;
        }
        const atual = saida[posicao[k]];
        const fica = melhor(atual, l).slice();
        const sai = fica[iMat] === atual[iMat] && melhor(atual, l) === atual ? l : atual;
        for (let c = 0; c < fica.length; c++) {
            if (vazio(fica[c]) && !vazio(sai[c])) fica[c] = sai[c];
        }
        saida[posicao[k]] = fica;
    }
    return saida;
}