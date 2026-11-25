// Office Script — Escala extra (CONTROLE DE FALTAS E ATESTADOS DOS TÉCNICOS)
// Lê as abas RESIDENTE e VOLANTE e devolve o dados_escala_extra.json
// (mesmo formato do dados_escala.json). Telefone (Contato) NÃO é exportado.
// Códigos: TRB → X (trabalho), DSR → FO (folga), FBH, FRD → Feriado,
//          FÉRIAS → Férias, ATS → Atestado; outros códigos vão como estão.
function main(workbook: ExcelScript.Workbook): string {
    const ABAS: { aba: string, supervisao: string }[] = [
        { aba: "RESIDENTE", supervisao: "Rodolfo De La Hermosa Alonso" },
        { aba: "VOLANTE", supervisao: "Thiago Patricio Da Silva" }
    ];
    const CODIGOS: { [k: string]: string } = { "TRB": "X", "DSR": "FO", "FBH": "FBH", "FRD": "Feriado", "FERIAS": "Férias", "ATS": "Atestado" };

    const semAcento = (s: string): string => s.normalize("NFD").replace(/[̀-ͯ]/g, "");
    const lim = (v: string | number | boolean): string => String(v ?? "").trim();
    const norm = (v: string | number | boolean): string => semAcento(lim(v).replace(/\s+/g, " ")).toUpperCase();
    const dois = (n: number): string => (n < 10 ? "0" : "") + n;
    const dataIso = (v: string | number | boolean): string => {
        if (typeof v === "number" && v > 30000 && v < 80000) {
            const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
            return d.getUTCFullYear() + "-" + dois(d.getUTCMonth() + 1) + "-" + dois(d.getUTCDate());
        }
        const t = String(v ?? "").trim();
        const m = t.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (m) return m[3] + "-" + m[2] + "-" + m[1];
        // "25/nov" (texto, sem ano): usa o ano atual
        const mm = norm(t).match(/^(\d{1,2})\/(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\.?$/);
        if (mm) { const mes = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"].indexOf(mm[2]) + 1; return new Date().getFullYear() + "-" + dois(mes) + "-" + dois(Number(mm[1])); }
        return "";
    };
    const codigo = (v: string | number | boolean): string => {
        const k = norm(v);
        if (!k) return "";
        return CODIGOS[k] !== undefined ? CODIGOS[k] : lim(v).toUpperCase();
    };

    interface Linha { nome: string, area: string, supervisao: string, turno: string, escala: string, horario: string, porData: { [iso: string]: string } }
    const linhas: Linha[] = [];
    const todasDatas: { [iso: string]: boolean } = {};

    const ehNomeCab = (c: string | number | boolean): boolean => ["COLABORADOR", "COLABORADORES", "TECNICO", "TECNICOS"].indexOf(norm(c)) >= 0;
    const planilhas = workbook.getWorksheets();
    for (const cfg of ABAS) {
        let ws: ExcelScript.Worksheet | undefined = undefined;
        for (let k = 0; k < planilhas.length; k++) {
            if (norm(planilhas[k].getName()) === cfg.aba) { ws = planilhas[k]; break; }
        }
        if (!ws) continue;
        const usado = ws.getUsedRange(true);
        if (!usado) continue;
        const v = usado.getValues();
        // A aba pode ter vários blocos (um por site), cada um com a sua linha
        // "Colaborador | ... | datas". Cada bloco usa as datas do seu cabeçalho.
        let iNome = 0, iTurno = -1, iHor = -1, iEscala = -1, iCont = -1;
        let colsData: { iso: string, i: number }[] = [];
        let temCab = false;
        for (let r = 0; r < v.length; r++) {
            const l = v[r];
            // ALTERADO (25/11/2026): a aba RESIDENTE usa "Técnicos" no lugar de "Colaborador" no cabeçalho de cada bloco
            if (l.some(c => ehNomeCab(c)) && (l.some(c => norm(c) === "TURNO") || l.some(c => norm(c) === "COLABORADOR"))) {
                const cab = l;
                const col = (nome: string): number => cab.findIndex(c => norm(c) === nome);
                iNome = cab.findIndex(c => ehNomeCab(c)); iTurno = col("TURNO"); iHor = col("HORARIO"); iCont = col("CONTATO");
                const novas: { iso: string, i: number }[] = [];
                cab.forEach((c, i) => { const iso = dataIso(c); if (iso) novas.push({ iso: iso, i: i }); });
                // bloco sem as datas no cabeçalho (ex.: Santo André): usa as colunas de data do bloco de cima
                if (novas.length) colsData = novas;
                // VOLANTE: a escala (5x2, 12x36) fica numa coluna sem título
                iEscala = -1;
                if (cfg.aba === "VOLANTE") {
                    for (let c = 0; c < cab.length && iEscala < 0; c++) {
                        if (lim(cab[c]) !== "") continue;
                        for (let rr = r + 1; rr < Math.min(v.length, r + 30); rr++) {
                            if (/^\d+\s*x\s*\d+$/i.test(lim(v[rr][c]))) { iEscala = c; break; }
                        }
                    }
                }
                temCab = true;
                continue;
            }
            if (!temCab) continue;
            const nome = lim(l[iNome]);
            if (!nome || /^(EM PROCESSO|VAGA ABERTA|VAGA)$/.test(norm(nome)) || ehNomeCab(nome)) continue;
            const turno = iTurno >= 0 ? lim(l[iTurno]) : "";
            const horario = iHor >= 0 ? lim(l[iHor]) : "";
            // só linhas de técnico: tem horário (06:00 as 18:00) ou turno de verdade (não é título de site/assinatura)
            // (ou tem telefone em Contato — técnico sem turno/horário preenchido, ex.: BRUNO HENRIQUE)
            const temFone = iCont >= 0 && /\d{4}\s*-?\s*\d{4}/.test(lim(l[iCont]));
            if (!/\d{1,2}:\d{2}/.test(horario) && !(turno && CODIGOS[norm(turno)] === undefined && norm(turno) !== "FERIAS") && !temFone) continue;
            const porData: { [iso: string]: string } = {};
            colsData.forEach(c => { todasDatas[c.iso] = true; const cod = codigo(l[c.i]); if (cod && porData[c.iso] === undefined) porData[c.iso] = cod; });
            linhas.push({
                nome: nome, area: cfg.aba, supervisao: cfg.supervisao, turno: turno,
                escala: iEscala >= 0 ? (lim(l[iEscala]) || turno) : turno,
                horario: horario, porData: porData
            });
        }
    }

    const datas = Object.keys(todasDatas).sort();
    const tecnicos = linhas.map(t => ({
        nome: t.nome, duplado: false, supervisao: t.supervisao, coordenacao: "", area: t.area,
        turno: t.turno, escala: t.escala, horario: t.horario, atuacao: "",
        codigos: datas.map(d => t.porData[d] || "")
    }));
    const hoje = new Date();
    return JSON.stringify({
        atualizado_em: hoje.getFullYear() + "-" + dois(hoje.getMonth() + 1) + "-" + dois(hoje.getDate()),
        fonte: "CONTROLE DE FALTAS E ATESTADOS DOS TÉCNICOS (abas RESIDENTE e VOLANTE)",
        datas: datas,
        tecnicos: tecnicos
    });
}
