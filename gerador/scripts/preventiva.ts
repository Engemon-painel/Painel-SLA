// ======================================================================
// Office Script — Exportar Preventiva (Painel do ESER)
// Lê a aba "Atualização_Preventiva" e a aba "Validação" (técnico por WO)
// e devolve o texto do dados_preventiva.json.
// Uso: Power Automate → "Executar script" → publicar o resultado no GitHub.
//
// NOVO — HISTÓRICO POR MÊS ("trava" o mês anterior):
// - O MÊS da base vem da própria base: o ponto do MEIO entre
//   "clr_sdr_data_inicio" e "clr_sdr_data_expiracao" de cada WO (o mês que
//   mais aparece). Ex.: setembro = 01/09 a 30/09 (meio 15/09 -> setembro);
//   outubro = 22/09 a 10/11 (meio 16/10 -> outubro). Dá para forçar outro mês
//   em MES_FORCADO (ex.: "2026-10"), mas normalmente fica vazio.
// - Cada execução guarda uma cópia dos registros desse mês na aba
//   "Historico_Preventiva" (criada sozinha, fica oculta). Enquanto a base for
//   de outubro, a cópia de outubro é atualizada; quando colar a de novembro,
//   outubro fica TRAVADO como estava na última execução.
// - Limpeza automática: se algum mês guardado tiver as MESMAS WOs da base
//   atual (a mesma base gravada com o mês errado), ele é apagado do histórico.
// - O JSON continua trazendo "mes", "gerado_em" e "registros" do mês atual
//   (o painel atual segue funcionando igual) e ganha "meses": a lista de
//   todos os meses guardados, para o filtro de mês do painel.
// - Parâmetro opcional "jsonAnterior": se o Power Automate passar o
//   dados_preventiva.json que já está publicado, os meses dele também são
//   aproveitados (segurança extra caso a aba de histórico se perca).
// ======================================================================
// Deixe "" para usar o mês de hoje. Preencha (ex.: "2026-09") só se precisar
// gravar a base atual em outro mês.
const MES_FORCADO = "";

function main(workbook: ExcelScript.Workbook, jsonAnterior?: string): string {
    const ABA = "Atualização_Preventiva";
    const ABA_VALIDACAO = "Validação";
    const ABA_HISTORICO = "Historico_Preventiva";

    const aba = workbook.getWorksheet(ABA);
    if (!aba) throw new Error(`Aba "${ABA}" não encontrada.`);
    const valores = aba.getUsedRange().getValues() as (string | number | boolean)[][];

    const abaVal = workbook.getWorksheet(ABA_VALIDACAO);
    const valoresVal = abaVal ? (abaVal.getUsedRange().getValues() as (string | number | boolean)[][]) : null;

    const atual = processar(valores, valoresVal);

    // ------------------------------------------------------------------
    // Histórico por mês
    // ------------------------------------------------------------------
    const meses: Record<string, MesPreventiva> = {};

    // 1) meses que já estavam publicados (opcional, vindo do Power Automate)
    try {
        const ant = JSON.parse(jsonAnterior || "{}") as { meses?: MesPreventiva[]; mes?: string; gerado_em?: string; registros?: object[] };
        if (Array.isArray(ant.meses)) {
            ant.meses.forEach(m => { if (m && m.mes) meses[m.mes] = m; });
        } else if (Array.isArray(ant.registros)) {
            // JSON do script antigo (sem "meses"): o "mes" dele podia vir errado,
            // então o mês vem da data em que aquele arquivo foi gerado.
            const mesAnt = /^\d{4}-\d{2}/.test(ant.gerado_em || "") ? (ant.gerado_em as string).slice(0, 7) : ant.mes;
            if (mesAnt) meses[mesAnt] = { mes: mesAnt, gerado_em: ant.gerado_em || "", registros: ant.registros };
        }
    } catch (e) {
        // jsonAnterior vazio ou inválido: segue só com a aba de histórico.
    }

    // 2) meses guardados na aba de histórico (tem prioridade sobre o item 1)
    let abaHist = workbook.getWorksheet(ABA_HISTORICO);
    if (!abaHist) {
        abaHist = workbook.addWorksheet(ABA_HISTORICO);
        abaHist.setVisibility(ExcelScript.SheetVisibility.hidden);
    }
    lerHistorico(abaHist).forEach(m => { meses[m.mes] = m; });

    // 3) limpeza: apaga meses guardados que são, na verdade, a MESMA base atual
    //    gravada com outro mês (mais da metade das WOs iguais às da base atual)
    const wosAtuais = new Set((atual.registros as { w?: string }[]).map(r => String(r.w || "")).filter(w => w !== ""));
    Object.keys(meses).forEach(k => {
        if (k === atual.mes) return;
        const wosMes = (meses[k].registros as { w?: string }[]).map(r => String(r.w || "")).filter(w => w !== "");
        if (!wosMes.length) return;
        const iguais = wosMes.filter(w => wosAtuais.has(w)).length;
        if (iguais / wosMes.length > 0.5) delete meses[k];
    });

    // 4) mês da base atual: sempre substitui a cópia desse mês
    //    (os outros meses ficam como estão = travados)
    meses[atual.mes] = atual;

    gravarHistorico(abaHist, Object.keys(meses).sort().map(k => meses[k]));

    const listaMeses = Object.keys(meses).sort().map(k => meses[k]);

    return JSON.stringify({
        mes: atual.mes,
        gerado_em: atual.gerado_em,
        registros: atual.registros,
        meses: listaMeses
    });
}

interface MesPreventiva {
    mes: string;
    gerado_em: string;
    registros: object[];
}

// Uma célula do Excel aceita no máximo 32.767 caracteres, então o JSON de
// cada mês é gravado em pedaços (uma linha por pedaço).
const TAMANHO_PEDACO = 30000;

function lerHistorico(abaHist: ExcelScript.Worksheet): MesPreventiva[] {
    const usado = abaHist.getUsedRange();
    if (!usado) return [];
    const linhas = usado.getValues() as (string | number | boolean)[][];
    if (linhas.length < 2) return [];

    const partes: Record<string, { gerado_em: string; pedacos: { n: number; t: string }[] }> = {};
    for (let r = 1; r < linhas.length; r++) {
        const mes = String(linhas[r][0] ?? "").trim();
        if (!/^\d{4}-\d{2}$/.test(mes)) continue;
        const n = Number(linhas[r][1]) || 0;
        const geradoEm = String(linhas[r][2] ?? "");
        const texto = String(linhas[r][3] ?? "");
        if (!partes[mes]) partes[mes] = { gerado_em: geradoEm, pedacos: [] };
        partes[mes].pedacos.push({ n: n, t: texto });
    }

    const saida: MesPreventiva[] = [];
    Object.keys(partes).forEach(mes => {
        const p = partes[mes];
        const json = p.pedacos.sort((a, b) => a.n - b.n).map(x => x.t).join("");
        try {
            const registros = JSON.parse(json) as object[];
            if (Array.isArray(registros)) saida.push({ mes: mes, gerado_em: p.gerado_em, registros: registros });
        } catch (e) {
            // pedaço corrompido: ignora esse mês em vez de quebrar o script
        }
    });
    return saida;
}

function gravarHistorico(abaHist: ExcelScript.Worksheet, meses: MesPreventiva[]): void {
    const linhas: (string | number)[][] = [["Mes", "Parte", "Gerado_em", "Registros_JSON"]];
    meses.forEach(m => {
        const json = JSON.stringify(m.registros);
        for (let i = 0, parte = 1; i < json.length; i += TAMANHO_PEDACO, parte++) {
            linhas.push([m.mes, parte, m.gerado_em, json.slice(i, i + TAMANHO_PEDACO)]);
        }
    });

    const usado = abaHist.getUsedRange();
    if (usado) usado.clear(ExcelScript.ClearApplyTo.contents);

    // Mês, data e JSON gravados como TEXTO, senão o Excel converte "2026-09" em data
    abaHist.getRangeByIndexes(0, 0, linhas.length, 1).setNumberFormat("@");
    abaHist.getRangeByIndexes(0, 2, linhas.length, 2).setNumberFormat("@");
    abaHist.getRangeByIndexes(0, 0, linhas.length, 4).setValues(linhas);
}

function processar(
    valores: (string | number | boolean)[][],
    valoresVal: (string | number | boolean)[][] | null
): MesPreventiva {
    const norm = (s: string) => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    const cab = valores[0].map(c => norm(String(c)));
    const idx = (...nomes: string[]): number => {
        for (const n of nomes) { const i = cab.indexOf(norm(n)); if (i >= 0) return i; }
        return -1;
    };

    const C = {
        wo: idx("WO"),
        chave: idx("Chave"),
        site: idx("clr_sdr_site/HELIX"),
        siteInfra: idx("programado/INFRATEL"),
        municipio: idx("MUNICIPIO"),
        tipologia: idx("TIPOLOGIA"),
        tipo: idx("Tipo de Preventiva"),
        equipe: idx("Equipe Responável", "Equipe Responsável"),
        equipeEng: idx("Equipe Responsável - Engemon"),
        status: idx("Status Compilado - Final"),
        dataFim: idx("Data Fim - Consolidado"),
        dataIni: idx("clr_sdr_data_inicio"),
        dataExp: idx("clr_sdr_data_expiracao"),
        expurgo: idx("Expurgo Situação"),
        infratel: idx("Infratel OK"),
        cronograma: idx("% Cronograma Infratel"),
        inventario: idx("% Inventário"),
        tipoInfra: idx("Tipo de Infra"),
        chamado: idx("Número do chamado / Acesso"),
        relatorio: idx("Entrega Relatórios"),
    };
    const faltando = Object.keys(C).filter(k => (C as Record<string, number>)[k] < 0 && k !== "chave");
    if (faltando.length) throw new Error("Colunas não encontradas: " + faltando.join(", "));

    const val = (l: (string | number | boolean)[], i: number) => (i < 0 ? "" : l[i]);
    const txt = (v: string | number | boolean) => String(v ?? "").replace(/\u00a0/g, " ").trim();
    const pad = (n: number) => String(n).padStart(2, "0");

    // Padroniza nome próprio: "JOAO DA SILVA" / "joao da silva" -> "João da Silva"
    const PREPOSICOES = new Set(["de", "da", "do", "das", "dos", "e"]);
    const nomeProprio = (v: string | number | boolean): string => {
        const s = txt(v);
        if (!s) return "";
        return s.toLowerCase().split(/\s+/).map((w, i) => {
            if (i > 0 && PREPOSICOES.has(w)) return w;
            return w.charAt(0).toUpperCase() + w.slice(1);
        }).join(" ");
    };

    // datas do Excel chegam como número serial (ex.: 46266) ou como texto dd/mm/aaaa
    const data = (v: string | number | boolean): string | null => {
        if (typeof v === "number" && v > 20000) {
            const d = new Date(Math.round((v - 25569) * 86400000));
            return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
        }
        const m = txt(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        return m ? `${m[3]}-${pad(+m[2])}-${pad(+m[1])}` : null;
    };
    const num = (v: string | number | boolean): number | null => {
        if (typeof v === "number") return Math.round(v * 10000) / 10000;
        const s = txt(v).replace("%", "").replace(",", ".");
        if (s === "") return null;
        const n = Number(s);
        return isNaN(n) ? null : Math.round(n * 10000) / 10000;
    };

    // --------------------------------------------------------------------
    // Aba "Validação": várias linhas por WO (uma por item de checklist).
    // Usada só para descobrir, por WO, o Técnico (Executor) e a Data de Início.
    // --------------------------------------------------------------------
    const mapaTec: Record<string, string> = {};
    const mapaIni: Record<string, string> = {};
    if (valoresVal && valoresVal.length > 1) {
        const cabVal = valoresVal[0].map(c => norm(String(c)));
        const idxVal = (...nomes: string[]): number => {
            for (const n of nomes) { const i = cabVal.indexOf(norm(n)); if (i >= 0) return i; }
            return -1;
        };
        const V = {
            wo: idxVal("WO.WO", "WO"),
            tec: idxVal("Executor"),
            inicio: idxVal("Início"),
        };
        if (V.wo >= 0) {
            for (let r = 1; r < valoresVal.length; r++) {
                const l = valoresVal[r];
                const wo = txt(val(l, V.wo));
                if (!wo) continue;
                if (V.tec >= 0 && !mapaTec[wo]) {
                    const tecNome = nomeProprio(val(l, V.tec));
                    if (tecNome) mapaTec[wo] = tecNome;
                }
                if (V.inicio >= 0 && !mapaIni[wo]) {
                    const ini = data(val(l, V.inicio));
                    if (ini) mapaIni[wo] = ini;
                }
            }
        }
    }

    // --------------------------------------------------------------------
    // "registros" (Atualização_Preventiva): UMA linha por WO.
    // --------------------------------------------------------------------
    const registros: object[] = [];
    const contMesIni: Record<string, number> = {};
    for (let r = 1; r < valores.length; r++) {
        const l = valores[r];
        if (!txt(val(l, C.wo)) && !txt(val(l, C.chave))) continue;
        const wo = txt(val(l, C.wo));
        const dFim = data(val(l, C.dataFim));
        const dIni = data(val(l, C.dataIni));
        const dExp = data(val(l, C.dataExp));
        if (dIni) {
            // meio do período da WO (início + expiração ÷ 2); sem expiração, o próprio início
            let ref = dIni;
            if (dExp) {
                const t = (Date.parse(dIni + "T00:00:00Z") + Date.parse(dExp + "T00:00:00Z")) / 2;
                const dm = new Date(t);
                ref = `${dm.getUTCFullYear()}-${pad(dm.getUTCMonth() + 1)}`;
            }
            const k = ref.slice(0, 7);
            contMesIni[k] = (contMesIni[k] || 0) + 1;
        }
        registros.push({
            w: wo,
            s: txt(val(l, C.site)),
            si: txt(val(l, C.siteInfra)),
            m: txt(val(l, C.municipio)),
            tp: txt(val(l, C.tipologia)),
            t: txt(val(l, C.tipo)).replace("Preventiva infra - ", ""),
            e: txt(val(l, C.equipe)) || "Sem equipe",
            x: txt(val(l, C.equipeEng)) || "Sem executor",
            st: txt(val(l, C.status)) || "Sem status",
            d: dFim,
            xp: txt(val(l, C.expurgo)),
            io: txt(val(l, C.infratel)),
            pc: num(val(l, C.cronograma)),
            pi: num(val(l, C.inventario)),
            ti: txt(val(l, C.tipoInfra)),
            ch: txt(val(l, C.chamado)),
            er: txt(val(l, C.relatorio)),
            tec: mapaTec[wo] || "",
            ini: mapaIni[wo] || null,
        });
    }

    // horário de Brasília (UTC-3)
    const agora = new Date(Date.now() - 3 * 3600000);
    const geradoEm = `${agora.getUTCFullYear()}-${pad(agora.getUTCMonth() + 1)}-${pad(agora.getUTCDate())} ${pad(agora.getUTCHours())}:${pad(agora.getUTCMinutes())}`;

    // Mês da base = mês que mais aparece no meio do período das WOs
    // (ou MES_FORCADO, se preenchido; sem datas, o mês em que o script roda).
    const ksIni = Object.keys(contMesIni).sort((a, b) => contMesIni[b] - contMesIni[a] || b.localeCompare(a));
    let mes = ksIni.length ? ksIni[0] : geradoEm.slice(0, 7);
    if (/^\d{4}-\d{2}$/.test(MES_FORCADO)) mes = MES_FORCADO;

    return { mes: mes, gerado_em: geradoEm, registros: registros };
}