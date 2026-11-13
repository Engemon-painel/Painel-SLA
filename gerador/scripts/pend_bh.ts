/**
 * SCRIPT DE PENDÊNCIAS DO BANCO DE HORAS  (Office Script — planilha "Controle de Ponto - HE - BH")
 * ---------------------------------------------------------------------------
 * As pendências vêm da BASE de pendências e TODA a tratativa é feita pelo painel.
 * O que o gestor responde fica guardado na tabela "tbJustificativasPainel"
 * (aba JUSTIFICATIVAS_PAINEL), criada sozinha. Ela é o histórico: guarda a
 * pendência completa + a resposta do gestor + as colunas do RH.
 *
 * Na PRIMEIRA execução, se a tabela estiver vazia, o script copia para ela tudo o
 * que já existe na TRATATIVA_GESTORES (tbTratativaGestoresClaro) — assim o que os
 * gestores já preencheram não se perde. Depois disso a TRATATIVA_GESTORES não é
 * mais usada.
 *
 * 1) SEM o campo "dados" (fluxo diário): junta a base de pendências com o histórico
 *    e devolve o dados_pend_bh.json → publique no GitHub.
 * 2) COM o campo "dados" (fluxo "BH - Gravar justificativa", chamado pelo painel):
 *    grava/atualiza a resposta do gestor na tbJustificativasPainel.
 *
 * Chave de cada pendência (igual em todas as bases): Matrícula (6 dígitos) | Data (aaaa-mm-dd) | Pendência
 */

// base de pendências (usa a primeira tabela que existir)
const TABELAS_BASE = ["tbBasePendenciasClaro", "Pendencias"];
const TAB_TRATATIVA_ANTIGA = "tbTratativaGestoresClaro";   // só para a carga inicial do histórico
const ABA_HIST = "JUSTIFICATIVAS_PAINEL";
const TAB_HIST = "tbJustificativasPainel";
const COLS_HIST: string[] = ["Chave", "ID_Pendencia", "Pendencia", "Nome", "Matricula_6", "Centro de Custo", "Turno", "Data", "Dia",
    "Detalhe / O que verificar", "Supervisor", "Coordenador", "Gerente", "Área Técnica", "Status_Colaborador", "BH_ATUAL_H",
    "Horas_Pendencia_h", "Pagamento_HE_Permitido", "Tratativa_Gestor", "Tarefa_HE", "Chamado", "Local_Atendimento",
    "Horario_Ajuste_Marcacoes", "Justificativa_Gestor", "Data_Resposta_Gestor", "Status_RH", "Observacao_RH", "Data_Ajuste_RH",
    "Status_Automatico", "Status Pendência", "Origem", "Justificativa_Padronizada"];
const COLS_GESTOR = ["Tratativa_Gestor", "Tarefa_HE", "Chamado", "Local_Atendimento", "Horario_Ajuste_Marcacoes", "Justificativa_Gestor", "Data_Resposta_Gestor"];
const COLS_RH = ["Status_RH", "Observacao_RH", "Data_Ajuste_RH"];
// texto que o script grava sozinho no Observacao_RH quando a pendência justificada sai da base
const OBS_RH_AUTO = "Corrigido no ponto (automático: saiu da base de pendências)";
const COLS_DATA = ["Data", "Data_Resposta_Gestor", "Data_Ajuste_RH"];

type Cel = string | number | boolean;
interface Just {
    trat: string; tarefa: string; chamado: string; local: string; horario: string; texto: string;
    dataResp: string; stRH: string; obsRH: string; dataRH: string; auto: string; status: string;
}
interface Item {
    id: string; idOrig: string; tipo: string; n: string; m: string; cc: string; turno: string; d: string; dia: string; det: string;
    sup: string; co: string; ge: string; area: string; st: string; saldo: number; tpSaldo: string; horas: number;
    hePerm: string; hist: boolean; just: Just;
    heDia: string;     // HE do dia da pendência (hh:mm): batidas x turno, como a base calcula (ex.: entrou 08:00, saiu 19:00, turno até 18:00 → 01:00)
    heFolha: string;   // HE da folha atual (hh:mm), da coluna da folha na base; vazio se a base não tiver essa coluna
    folha: string;     // período da folha da batida (dia 15 a 14), ex.: "15/09/2026 a 14/10/2026"
    folhaIni: string;  // início do período (aaaa-mm-dd), para ordenar e filtrar
}
interface Entrada { [campo: string]: string; }
interface Corpo { itens?: Entrada[]; ID?: string; acao?: string; }
interface Resultado { ID: string; ok: boolean; erro?: string; Status_Automatico?: string; "Status Pendência"?: string; }

function main(workbook: ExcelScript.Workbook, dados?: string): string {
    const hist = garantirHistorico(workbook);
    // ALTERADO (31/10/2026): VT/VR de trabalho em feriado ou fora da escala (supervisor aprova, RH lança)
    if (dados && /"acao"\s*:\s*"VTVR/i.test(String(dados))) return gravarVT(workbook, String(dados));
    if (dados && String(dados).trim()) return gravar(hist, String(dados));
    return JSON.stringify(exportar(workbook, hist));
}

// =====================================================================
// 1) EXPORTAR → dados_pend_bh.json
// =====================================================================
function exportar(workbook: ExcelScript.Workbook, hist: ExcelScript.Table): { geradoEm: string; folhaAtual: string; folhaAtualIni: string; total: number; itens: Item[]; absenteismo: Abs[]; espelho: EspSaida; vtvr: { [k: string]: VtVr } } {
    // ---- histórico (respostas dos gestores)
    const hv = hist.getRange().getValues();
    const hc = hv[0].map(c => String(c).trim());
    const hl = (r: number) => leitor(hc, hv[r]);
    const porChave: { [k: string]: number } = {};
    for (let r = 1; r < hv.length; r++) { const k = hl(r).txt(["Chave"]); if (k) porChave[k] = r; }

    // ---- base de pendências
    let base: ExcelScript.Table | undefined;
    for (const n of TABELAS_BASE) { base = workbook.getTable(n); if (base) break; }
    const bv = base ? base.getRange().getValues() : [[]];
    const bc = bv[0].map(c => String(c).trim());

    const itens: Item[] = [];
    const usados: { [k: string]: boolean } = {};
    const montar = (L: Leitor, hr: number | undefined, ehHist: boolean): Item | null => {
        const tipo = L.txt(["Pendencia"]).toUpperCase(), m = L.txt(["Matricula_6", "Matricula"]), d = L.data(["Data"]);
        if (!tipo || !m || !d) return null;
        const id = chave(m, d, tipo);
        const H = hr != null ? hl(hr) : null;
        const saldo = L.num(["BH_ATUAL_H", "BH_Atual_h", "Saldo_Atual_h"]);
        const j: Just = {
            trat: H ? H.txt(["Tratativa_Gestor"]) : "", tarefa: H ? H.txt(["Tarefa_HE"]) : "", chamado: H ? H.txt(["Chamado"]) : "",
            local: H ? H.txt(["Local_Atendimento"]) : "", horario: H ? H.txt(["Horario_Ajuste_Marcacoes"]) : "",
            texto: H ? H.txt(["Justificativa_Gestor"]) : "", dataResp: H ? H.data(["Data_Resposta_Gestor"]) : "",
            stRH: H ? H.txt(["Status_RH"]) : "", obsRH: H ? H.txt(["Observacao_RH"]) : "", dataRH: H ? H.data(["Data_Ajuste_RH"]) : "",
            auto: "", status: ""
        };
        // Pagamento_HE_Permitido / Status_RH / Observacao_RH / Data_Ajuste_RH: o RH preenche direto na JUSTIFICATIVAS_PAINEL
        const hePermRH = H ? H.txt(["Pagamento_HE_Permitido"]) : "";
        const hePerm = hePermRH || L.txt(["Pagamento_HE_Permitido"]);
        const s = calcularStatus(j, hePerm, tipo); j.auto = s[0]; j.status = s[1];
        return {
            id, idOrig: L.txt(["ID_Pendencia"]), tipo, n: L.txt(["Nome", "Colaborador"]), m: matPadrao(m), cc: L.txt(["Centro de Custo"]), turno: L.txt(["Turno"]),
            d, dia: L.txt(["Dia"]), det: L.txt(["Detalhe / O que verificar"]), sup: L.txt(["Supervisor"]), co: L.txt(["Coordenador"]),
            ge: L.txt(["Gerente"]), area: L.txt(["Área Técnica"]), st: L.txt(["Status_Colaborador", "Status"]),
            saldo: Math.round(saldo * 100) / 100, tpSaldo: saldo < 0 ? "NEGATIVO" : "POSITIVO", horas: L.horas(), hePerm, hist: ehHist, just: j,
            heDia: tipo === "HORA EXTRA" ? heDoDia(L.txt(["Detalhe / O que verificar"]), L.horas()) : "",
            heFolha: tipo === "HORA EXTRA" ? hhmm(L.heFolha()) : "",
            folha: periodoFolha(d)[1], folhaIni: periodoFolha(d)[0]
        };
    };
    for (let r = 1; r < bv.length; r++) {
        const it = montar(leitor(bc, bv[r]), undefined, false);
        if (!it) continue;
        const hr = porChave[it.id];
        const comResp = hr != null ? montar(leitor(bc, bv[r]), hr, false) : it;
        if (comResp) { itens.push(comResp); usados[it.id] = true; }
    }
    // ALTERADO (21/10/2026): ATRASOS — entrada (e1) mais de 10 min (era 5) depois do início do turno do dia
    // ALTERADO (22/10/2026): + FALTAS — observacao "Ausente" no Espelho de ponto
    const ABS: MapaAbs = {};
    const ESP: Esp = { mats: {}, ini: "", fim: "" };
    for (const linhaAtraso of lerAtrasos(workbook, ABS, ESP)) {
        const it = montar(linhaAtraso, undefined, false);
        if (!it || usados[it.id]) continue;
        const hr = porChave[it.id];
        const comResp = hr != null ? montar(linhaAtraso, hr, false) : it;
        if (comResp) { itens.push(comResp); usados[it.id] = true; }
    }

    // pendências que estão só no histórico (já saíram da base)
    for (let r = 1; r < hv.length; r++) {
        const k = hl(r).txt(["Chave"]);
        if (!k || usados[k]) continue;
        const it = montar(hl(r), r, true);
        // atraso antigo dentro da tolerância (ex.: 6 a 10 min, da regra anterior de 5 min) não volta para o painel
        if (it && it.tipo === "ATRASO" && Math.round((+it.horas || 0) * 60) <= TOLERANCIA_ATRASO_MIN) continue;
        if (it) itens.push(it);
    }

    // ALTERADO (27/10/2026): pendência JUSTIFICADA que saiu da base na atualização (dentro do período que a base
    // cobre: do início da folha da pendência mais antiga até a mais recente) foi corrigida no ponto →
    // preenche sozinho a resposta do RH (só se o RH ainda não respondeu): AJUSTADO + observação + data de hoje.
    const datasBase = itens.filter(i => !i.hist).map(i => i.d).sort();
    const fimBase = datasBase.length ? datasBase[datasBase.length - 1] : "";
    const iniBase = datasBase.length ? periodoFolha(datasBase[0])[0] : "";
    const hojeRH = hojeBrasiliaIso();
    const autoRH: { [k: string]: boolean } = {};
    if (fimBase && iniBase) {
        itens.forEach(i => {
            if (!i.hist || i.just.status !== "Justificado" || String(i.just.stRH || "").trim()) return;
            if (i.d < iniBase || i.d > fimBase) return;
            i.just.stRH = "AJUSTADO";
            i.just.obsRH = OBS_RH_AUTO + " em " + hojeRH.split("-").reverse().join("/");
            i.just.dataRH = hojeRH;
            const s2 = calcularStatus(i.just, i.hePerm, i.tipo); i.just.auto = s2[0]; i.just.status = s2[1];
            autoRH[i.id] = true;
        });
    }

    // grava os status recalculados no histórico (o RH pode ter mudado o Status_RH)
    if (hv.length > 1) {
        // resposta automática do RH (pendências que saíram da base)
        const iRH = hc.indexOf("Status_RH"), iObs = hc.indexOf("Observacao_RH"), iDt = hc.indexOf("Data_Ajuste_RH");
        if (Object.keys(autoRH).length && iRH >= 0 && iObs >= 0 && iDt >= 0) {
            const cRH: Cel[][] = [], cObs: Cel[][] = [], cDt: Cel[][] = [];
            for (let r = 1; r < hv.length; r++) {
                const k = hl(r).txt(["Chave"]), novo = autoRH[k] && !String(hv[r][iRH] || "").trim();
                cRH.push([novo ? "AJUSTADO" : hv[r][iRH]]);
                cObs.push([novo ? OBS_RH_AUTO + " em " + hojeRH.split("-").reverse().join("/") : hv[r][iObs]]);
                cDt.push([novo ? hojeBrasiliaSerial() : hv[r][iDt]]);
            }
            const corpoRH = hist.getRangeBetweenHeaderAndTotal();
            corpoRH.getColumn(iRH).setValues(cRH); corpoRH.getColumn(iObs).setValues(cObs); corpoRH.getColumn(iDt).setValues(cDt);
            formatarDatas(hist);
        }
        const iA = hc.indexOf("Status_Automatico"), iS = hc.indexOf("Status Pendência");
        const statusPor: { [k: string]: string[] } = {};
        itens.forEach(i => statusPor[i.id] = [i.just.auto, i.just.status]);
        const colA: Cel[][] = [], colS: Cel[][] = [];
        for (let r = 1; r < hv.length; r++) {
            const k = hl(r).txt(["Chave"]), s = statusPor[k];
            colA.push([s ? s[0] : hv[r][iA]]); colS.push([s ? s[1] : hv[r][iS]]);
        }
        const corpo = hist.getRangeBetweenHeaderAndTotal();
        if (iA >= 0) corpo.getColumn(iA).setValues(colA);
        if (iS >= 0) corpo.getColumn(iS).setValues(colS);
        // padroniza todas as justificativas do histórico
        const iJ = hc.indexOf("Justificativa_Gestor"), iP = hc.indexOf("Justificativa_Padronizada");
        if (iJ >= 0 && iP >= 0) {
            const colP: Cel[][] = [];
            for (let r = 1; r < hv.length; r++) colP.push([padronizarJust(hv[r][iJ])]);
            corpo.getColumn(iP).setValues(colP);
        }
    }
    // ALTERADO (28/10/2026): só PONTO ABERTO espera 2 dias úteis da data da batida
    // (seg a sex; ex.: batida sexta 16/10 → úteis 19 e 20 → aparece a partir de quarta 21/10).
    // As outras pendências aparecem na hora. Se o gestor já respondeu, a pendência continua aparecendo.
    const hojeIso = hojeBrasiliaIso();
    const visiveis = itens.filter(i => {
        if (i.tipo !== "PONTO ABERTO" && i.tipo !== "FALTA") return true;   // ALTERADO (03/11/2026): FALTA também espera 2 dias úteis
        if (i.just.trat) return true;
        return hojeIso > somarDiasUteis(i.d, 2);
    });
    const folhaAtual = periodoFolha(hojeIso);
    const absenteismo: Abs[] = Object.keys(ABS).map(k => {
        const a = ABS[k];
        return {
            f: a.f, m: matPadrao(a.m), n: a.n, sup: a.sup, co: a.co, cc: a.cc, esc: a.esc, dPrev: a.dPrev, dAus: a.dAus, dPass: a.dPass,
            hPrev: Math.round(a.hPrev * 100) / 100, hAus: Math.round(a.hAus * 100) / 100, dEscP: a.dEscP, dBat: a.dBat, dSem: a.dSem, ultBat: a.ultBat, dCal: a.dCal, dAusReg: a.dAusReg, hAusReg: Math.round(a.hAusReg * 100) / 100
        };
    });
    // ALTERADO (28/10/2026): matrículas que aparecem no Espelho de ponto (o painel avisa quem do BH não aparece)
    const espelho: EspSaida = { ini: ESP.ini, fim: ESP.fim, mats: Object.keys(ESP.mats).sort() };
    return { geradoEm: new Date().toISOString(), folhaAtual: folhaAtual[1], folhaAtualIni: folhaAtual[0], total: visiveis.length, itens: visiveis, absenteismo, espelho, vtvr: lerVT(workbook) };
}

// =====================================================================
// 1b) ATRASOS — lê a aba que tem as colunas e1 (entrada 1) e turno (escala)
// =====================================================================
// ALTERADO (30/10/2026): tolerância passou de 5 para 10 min — atraso só a partir de 11 min (conta o atraso inteiro)
const TOLERANCIA_ATRASO_MIN = 10;
const ABAS_ATRASO = ["Espelho de ponto", "Espelho de Ponto", "Atrasos"];
const DIA_SEM: { [k: string]: number } = { DOM: 0, SEG: 1, TER: 2, QUA: 3, QUI: 4, SEX: 5, SAB: 6 };
const NOME_DIA = ["Domingo", "Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado"];

function normCab(v: Cel): string {
    return String(v == null ? "" : v).normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase().replace(/[\s_]+/g, " ");
}
function colDe(cab: string[], nomes: string[]): number {
    for (const n of nomes) { const k = cab.indexOf(normCab(n)); if (k >= 0) return k; }
    return -1;
}
// hora → minutos do dia: aceita hora do Excel (0,35), data+hora do Excel (46316,35) ou texto "08:20"
function minutosDe(v: Cel): number {
    if (v === "" || v == null) return NaN;
    if (typeof v === "number") return Math.round((v - Math.floor(v)) * 1440) % 1440;
    const m = String(v).match(/(\d{1,2}):(\d{2})/);
    return m ? Number(m[1]) * 60 + Number(m[2]) : NaN;
}
function hhmmMin(min: number): string {
    return String(Math.floor(min / 60)).padStart(2, "0") + ":" + String(min % 60).padStart(2, "0");
}
// início do turno no dia da semana (aceita "SEG A SEX - 08:00 AS 17:48", "SEG A QUI - 08:00 AS 18:00 | SEX 08:00 AS 17:00", "12X36 - 06:00 AS 18:00")
function inicioTurno(turno: string, w: number): number {
    const t = String(turno || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
    const re = /(?:(DOM|SEG|TER|QUA|QUI|SEX|SAB)(?:\s*A\s*(DOM|SEG|TER|QUA|QUI|SEX|SAB))?|12X36|\d+X\d+)[^0-9]*(\d{1,2}:\d{2})\s*(?:AS|-|A)\s*(\d{1,2}:\d{2})/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(t)) !== null) {
        const a = m[1] ? DIA_SEM[m[1]] : 0, b = m[1] ? (m[2] ? DIA_SEM[m[2]] : a) : 6;
        if (w >= a && w <= b) return minutosDe(m[3]);
    }
    return NaN;
}

// fim do turno no dia da semana (mesma leitura do inicioTurno)
function fimTurno(turno: string, w: number): number {
    const t = String(turno || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
    const re = /(?:(DOM|SEG|TER|QUA|QUI|SEX|SAB)(?:\s*A\s*(DOM|SEG|TER|QUA|QUI|SEX|SAB))?|12X36|\d+X\d+)[^0-9]*(\d{1,2}:\d{2})\s*(?:AS|-|A)\s*(\d{1,2}:\d{2})/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(t)) !== null) {
        const a = m[1] ? DIA_SEM[m[1]] : 0, b = m[1] ? (m[2] ? DIA_SEM[m[2]] : a) : 6;
        if (w >= a && w <= b) return minutosDe(m[4]);
    }
    return NaN;
}

// absenteísmo por colaborador e folha (preenchido pelo lerAtrasos)
interface Abs { f: string; m: string; n: string; sup: string; co: string; cc: string; esc: string; dPrev: number; dAus: number; hPrev: number; hAus: number; dPass: number; dEscP: number; dBat: number; dSem: number; ultBat: string; dCal: number; dAusReg: number; hAusReg: number; }
interface MapaAbs { [chave: string]: Abs; }
interface DiaAbs { d: string; w: number; turno: string; estado: string; jor: number; n: string; sup: string; co: string; cc: string; esc: string; np: boolean; ausReg: boolean; antesAdm: boolean; }

// soma os dias previstos da folha inteira por colaborador:
//   "prev"/"aus" = dia que aconteceu (trabalhou / faltou) · "off" = não previsto · "proj" = ainda não aconteceu → pela escala
//   escala por dia da semana: dia da semana da escala (respeita "FOLGA QUI"); 12x36: dia sim, dia não a partir do último dia trabalhado/faltado
function projetarAbs(DIAS: { [m: string]: DiaAbs[] }, ABS: MapaAbs, fimEsp: string) {
    const hojeIso = hojeBrasiliaIso();
    // ALTERADO (29/10/2026): absenteísmo "por dias corridos" (como o relatório do RH):
    //   dCal = dias do período no espelho (até o último dia com batida) · dAusReg = dias com "Ausente" registrados
    //   hAusReg = horas da jornada desses dias (pela escala)
    for (const m of Object.keys(DIAS)) {
        for (const x of DIAS[m]) {
            if (!fimEsp || x.d > fimEsp || x.antesAdm) continue;   // só a partir da admissão
            const fo = periodoFolha(x.d)[0], ch = fo + "|" + m;
            if (!ABS[ch]) ABS[ch] = { f: fo, m: m, n: x.n, sup: x.sup, co: x.co, cc: x.cc, esc: x.esc, dPrev: 0, dAus: 0, hPrev: 0, hAus: 0, dPass: 0, dEscP: 0, dBat: 0, dSem: 0, ultBat: "", dCal: 0, dAusReg: 0, hAusReg: 0 };
            const a = ABS[ch];
            a.dCal++;
            if (x.ausReg) { a.dAusReg++; a.hAusReg += x.jor; }
        }
    }
    const diaNum = (iso: string) => Math.round(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 86400000);
    for (const m of Object.keys(DIAS)) {
        const dias = DIAS[m].slice().sort((a, b) => a.d < b.d ? -1 : 1);
        const reais = dias.filter(x => x.estado === "prev" || x.estado === "aus");
        const ref = reais.length ? diaNum(reais[reais.length - 1].d) : null;
        for (const x of dias) {
            let previsto = x.estado === "prev" || x.estado === "aus";
            if (x.estado === "proj") {
                const t = String(x.turno || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
                if (/\d+\s*X\s*\d+/.test(t)) {
                    previsto = ref != null ? Math.abs(diaNum(x.d) - ref) % 2 === 0 : (diaNum(x.d) - diaNum(periodoFolha(x.d)[0])) % 2 === 0;
                } else {
                    const nomes = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"];
                    const folga = new RegExp("FOLGA\\s+" + nomes[x.w] + "\\b").test(t);
                    previsto = !folga && !isNaN(inicioTurno(x.turno, x.w));
                }
            }
            if (!previsto) continue;
            const fo = periodoFolha(x.d)[0], ch = fo + "|" + m;
            if (!ABS[ch]) ABS[ch] = { f: fo, m: m, n: x.n, sup: x.sup, co: x.co, cc: x.cc, esc: x.esc, dPrev: 0, dAus: 0, hPrev: 0, hAus: 0, dPass: 0, dEscP: 0, dBat: 0, dSem: 0, ultBat: "", dCal: 0, dAusReg: 0, hAusReg: 0 };
            const a = ABS[ch];
            a.dPrev++; a.hPrev += x.jor;
            if (x.estado !== "proj") a.dPass++;
            if (x.estado === "aus") { a.dAus++; a.hAus += x.jor; }
            // ALTERADO (29/10/2026): dias da escala que JÁ PASSARAM (até ontem) com e sem batida
            //   "prev" = tem batida · "aus" (Ausente) ou "proj" já passado = dia da escala sem batida
            if (x.d < hojeIso && !x.np) {   // D.S.R., compensado, férias, afastamento… não contam
                a.dEscP++;
                if (x.estado === "prev") { a.dBat++; if (x.d > a.ultBat) a.ultBat = x.d; } else a.dSem++;
            }
        }
    }
}

// ATRASOS e FALTAS (lidos juntos do Espelho de ponto)
interface Esp { mats: { [m: string]: boolean }; ini: string; fim: string; }
interface EspSaida { ini: string; fim: string; mats: string[]; }
function lerAtrasos(workbook: ExcelScript.Workbook, ABS: MapaAbs, ESP: Esp): Leitor[] {
    // acha a aba: primeiro pelos nomes conhecidos, depois qualquer aba com as colunas e1 e turno
    const temCols = (ws: ExcelScript.Worksheet): (string | number | boolean)[][] | null => {
        const r = ws.getUsedRange();
        if (!r) return null;
        const v = r.getValues();
        if (v.length < 2) return null;
        const c = v[0].map(x => normCab(x));
        return colDe(c, ["e1", "entrada 1", "entrada1", "ent1"]) >= 0 && colDe(c, ["turno", "escala"]) >= 0 ? v : null;
    };
    let v: (string | number | boolean)[][] | null = null;
    for (const n of ABAS_ATRASO) { const ws = workbook.getWorksheet(n); if (ws) { v = temCols(ws); if (v) break; } }
    if (!v) for (const ws of workbook.getWorksheets()) { v = temCols(ws); if (v) break; }
    if (!v) return [];

    const c = v[0].map(x => normCab(x));
    const k = {
        n: colDe(c, ["nome", "colaborador", "funcionario"]), m: colDe(c, ["matricula 6", "matricula", "chapa", "re"]),
        d: colDe(c, ["data", "dia data", "data batida", "data marcacao"]), t: colDe(c, ["turno dia", "turno", "escala"]),
        sit: colDe(c, ["situacao"]), ab: colDe(c, ["abono"]), aus: colDe(c, ["ausencia"]), obs: colDe(c, ["observacao", "observação"]),
        e1: colDe(c, ["e1", "entrada 1", "entrada1", "ent1"]), s1: colDe(c, ["s1", "saida 1", "saida1", "sai1"]),
        sup: colDe(c, ["supervisor"]), co: colDe(c, ["coordenador"]), ge: colDe(c, ["gerente"]),
        ar: colDe(c, ["area tecnica", "area"]), st: colDe(c, ["status colaborador", "status"]), cc: colDe(c, ["centro de custo", "centro custo", "cc"]),
        te: colDe(c, ["turno", "escala"]),   // escala do colaborador (a coluna "turno", não o "turno dia")
        adm: colDe(c, ["admissao", "data admissao"])
    };
    if (k.m < 0 || k.d < 0) return [];

    // supervisor, coordenador, gerente e saldo pela matrícula (aba Controle_de_Ponto - HE - BH), quando a aba de atrasos não tiver
    const info: { [m: string]: Cel[] } = {};
    const wsBh = workbook.getWorksheet("Controle_de_Ponto - HE - BH");
    let ib: { [c: string]: number } = {};
    if (wsBh && wsBh.getUsedRange()) {
        const bv = wsBh.getUsedRange().getValues(), bc = bv[0].map(x => normCab(x));
        ib = {
            m: colDe(bc, ["matricula 6", "matricula"]), sup: colDe(bc, ["supervisor"]), co: colDe(bc, ["coordenador"]),
            ge: colDe(bc, ["gerente"]), ar: colDe(bc, ["area tecnica", "projeto"]), st: colDe(bc, ["status"]), sal: colDe(bc, ["saldo atual h"])
        };
        if (ib.m >= 0) for (let r = 1; r < bv.length; r++) info[mat6(bv[r][ib.m])] = bv[r];
    }
    const temInfo = Object.keys(info).length > 0;
    const de = (row: Cel[], i: number) => i >= 0 ? String(row[i] == null ? "" : row[i]).trim() : "";
    // ALTERADO (03/11/2026): dia ABONADO não é falta nem atraso — vale se a coluna "abono" tiver horas (maior que zero)
    // ou se qualquer coluna da linha disser "ABONADO" / "ABONO" (ex.: observação ou situação abonada).
    const abonado = (row: Cel[]): boolean => {
        if (k.ab >= 0) {
            const a = row[k.ab];
            if (typeof a === "number" ? a > 0 : (() => { const t = String(a == null ? "" : a).trim(); if (!t) return false; const mm = minutosDe(t); return isNaN(mm) ? !/^0+([,.]0+)?$/.test(t) : mm > 0; })()) return true;
        }
        return row.some(x => /ABONAD|ABONO/i.test(String(x == null ? "" : x)));
    };

    const CAB = ["Pendencia", "Matricula_6", "Data", "Nome", "Turno", "Dia", "Detalhe / O que verificar", "Supervisor", "Coordenador",
        "Gerente", "Área Técnica", "Status_Colaborador", "Centro de Custo", "BH_ATUAL_H", "Horas_Pendencia_h", "ID_Pendencia"];
    const saida: Leitor[] = [];
    const DIAS: { [m: string]: DiaAbs[] } = {};
    for (let r = 1; r < v.length; r++) {
        const row = v[r];
        const m = mat6(row[k.m]), d = dataIso(row[k.d]);
        if (!m || !d) continue;
        ESP.mats[m] = true;
        if (!ESP.ini || d < ESP.ini) ESP.ini = d;
        if (!isNaN(minutosDe(k.e1 >= 0 ? row[k.e1] : "")) && d > ESP.fim) ESP.fim = d;   // último dia com batida
        // ALTERADO (03/11/2026): entra quem está na aba Controle_de_Ponto - HE - BH OU quem é do contrato Claro/Cirion pelo
        // centro de custo do espelho (ex.: admitido recente que ainda não aparece no relatório de banco de horas)
        if (temInfo && !info[m] && !/CLARO|CIRION|8005400|3100001\d\d/i.test(de(row, k.cc))) continue;
        const I = info[m];
        const pegaI = (i: number) => I && i != null && i >= 0 ? String(I[i] == null ? "" : I[i]).trim() : "";
        const turno = de(row, k.t);
        const w = new Date(d + "T12:00:00Z").getUTCDay();
        const matTxt = matPadrao(I && ib.m >= 0 ? I[ib.m] : m);
        const comum = (tipo: string, det: string, horas: number): Cel[] => [tipo, matTxt, d, de(row, k.n), turno, NOME_DIA[w], det,
            de(row, k.sup) || pegaI(ib.sup), de(row, k.co) || pegaI(ib.co), de(row, k.ge) || pegaI(ib.ge),
            de(row, k.ar) || pegaI(ib.ar), de(row, k.st) || pegaI(ib.st), de(row, k.cc),
            I && ib.sal >= 0 ? (Number(I[ib.sal]) || 0) : 0, Math.round(horas * 100) / 100, m + "|" + d.replace(/-/g, "") + "|" + tipo];

        // ALTERADO (22/10/2026): ABSENTEÍSMO sobre a FOLHA INTEIRA (15 a 14)
        //   dias que já passaram (2 dias úteis): conta o que aconteceu — trabalhou ou "Ausente" = previsto; D.S.R., férias... = não previsto
        //   dias que ainda não aconteceram (em branco no espelho): projeta pela escala (ver projetarAbs)
        {
            const obsN = normCab(de(row, k.obs));
            const naoPrevisto = /^(d\.?s\.?r\.?|compensado|nao trabalhado|ferias|isento de marcacao|demissao)$/.test(obsN) ||
                /afastamento|ferias/.test(obsN) || /AFAST|FERIAS/i.test(de(row, k.sit));
            // ALTERADO (03/11/2026): "Ausente" em dia que ainda não chegou (hoje ou depois) não conta — só até ontem
            const ausente = obsN === "ausente" && !abonado(row) && d < hojeBrasiliaIso();
            const trabalhou = !isNaN(minutosDe(k.e1 >= 0 ? row[k.e1] : ""));
            const liberado = hojeBrasiliaIso() > somarDiasUteis(d, 2);
            const estado = liberado && naoPrevisto ? "off" : liberado && ausente ? "aus" : (trabalhou || (liberado && ausente)) ? "prev" : "proj";
            const fimT = fimTurno(turno, w), iniT = inicioTurno(turno, w);
            let jor = !isNaN(fimT) && !isNaN(iniT) ? fimT - iniT : 0; if (jor < 0) jor += 1440; if (jor > 360) jor -= 60;
            if (!DIAS[m]) DIAS[m] = [];
            DIAS[m].push({
                d: d, w: w, turno: turno, estado: estado, jor: jor / 60, n: de(row, k.n),
                sup: de(row, k.sup) || pegaI(ib.sup), co: de(row, k.co) || pegaI(ib.co),
                cc: de(row, k.cc), esc: (de(row, k.te) || turno).replace(/^\s*\d+\s*-\s*/, ""), np: naoPrevisto, ausReg: ausente, antesAdm: k.adm >= 0 && !!dataIso(row[k.adm]) && d < dataIso(row[k.adm])
            });
        }

        // ALTERADO (22/10/2026): FALTAS — coluna "observacao" = "Ausente" (sem abono)
        // ALTERADO (03/11/2026): só faltas até ontem (o espelho já traz "Ausente" nos dias futuros da folha)
        if (normCab(de(row, k.obs)) === "ausente" && !abonado(row)) {
            if (d >= hojeBrasiliaIso()) continue;
            const aus = k.aus >= 0 ? row[k.aus] : "";
            let horasF = typeof aus === "number" ? aus * 24 : (isNaN(minutosDe(aus)) ? 0 : minutosDe(aus) / 60);
            if (!horasF) {                               // sem horas de ausência: usa a jornada do turno do dia
                const fim = fimTurno(turno, w), ini0 = inicioTurno(turno, w);
                if (!isNaN(fim) && !isNaN(ini0)) { let j = fim - ini0; if (j < 0) j += 1440; horasF = (j > 360 ? j - 60 : j) / 60; }
            }
            const detF = "Falta (Ausente) | Turno: " + turno + (horasF ? " | Horas ausentes: " + hhmmMin(Math.round(horasF * 60)) : "");
            saida.push(leitor(CAB, comum("FALTA", detF, horasF)));
            continue;
        }

        const e1 = minutosDe(k.e1 >= 0 ? row[k.e1] : "");
        if (isNaN(e1) || /FOLGA|FERIAS|FÉRIAS|AFAST/i.test(turno)) continue;
        if (/FOLGA|FERIAS|FÉRIAS|AFAST|ATESTADO|LICEN|FERIADO|DSR|COMPENS/i.test(de(row, k.sit))) continue;
        if (abonado(row)) continue;                     // dia com abono: não é atraso a justificar
        const ini = inicioTurno(turno, w);
        if (isNaN(ini)) continue;
        let atraso = e1 - ini;
        if (atraso < -720) atraso += 1440; else if (atraso > 720) atraso -= 1440;   // virada da meia-noite
        if (atraso <= TOLERANCIA_ATRASO_MIN) continue;
        const s1 = minutosDe(k.s1 >= 0 ? row[k.s1] : "");
        if (isNaN(s1)) continue;                        // só uma batida = falta marcação (já vira PONTO ABERTO), não atraso
        const det = "Atraso: " + hhmmMin(atraso) + " | Turno: " + turno + " | Início do turno: " + hhmmMin(ini) +
            " | Entrada: " + hhmmMin(e1) + " | Saída: " + hhmmMin(s1);
        saida.push(leitor(CAB, comum("ATRASO", det, atraso / 60)));
    }
    projetarAbs(DIAS, ABS, ESP.fim);
    return saida;
}
// ALTERADO (23/10/2026): matrícula no padrão da base: "02 - 006841" (aceita 6841, "006841", "02 - 6841"...)
function matPadrao(v: Cel): string {
    const t = String(v == null ? "" : v).trim();
    const p = t.match(/^(\d{1,2})\s*-\s*(\d+)$/);
    if (p) return ("00" + p[1]).slice(-2) + " - " + ("000000" + p[2]).slice(-6);
    const m = mat6(v);
    return m ? "02 - " + m : t;
}
function mat6(v: Cel): string {
    // no Espelho de ponto a matrícula "02 - 3050" virou data (01/02/3050): o ano é a matrícula
    if (typeof v === "number" && v > 73051) {          // depois de 01/01/2100
        const dt = new Date(Math.round((v - 25569) * 86400000));
        if (dt.getUTCDate() === 1) return ("000000" + dt.getUTCFullYear()).slice(-6);
    }
    const dig = String(v == null ? "" : v).replace(/\D/g, "");
    return dig ? ("000000" + dig).slice(-6) : "";
}

// =====================================================================
// 2) GRAVAR (chamado pelo painel via Power Automate)
// =====================================================================
// ALTERADO (29/10/2026): Justificativa_Padronizada = texto do gestor agrupado em frases padrão
// (ex.: "HE para abatimento do banco de horas negativo", "Descontar do banco de horas", "Debitar do banco..."
// viram todas "Lançar no banco de horas"). O texto original continua na Justificativa_Gestor. Sem regra → mantém o original.
function padronizarJust(txt: Cel): string {
    const t0 = String(txt == null ? "" : txt).replace(/\s+/g, " ").trim();
    if (!t0) return "";
    const t = t0.normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
    const tem = (re: RegExp) => re.test(t);
    const HE = /\b(HE|H\.E\.?|HORA EXTRA|HORAS EXTRAS?)\b/;
    const BANCO = /BANCO|BANDO|\bBH\b|SALDO/;
    if (tem(/TIROU (BANCO|FOLGA)|COBRIR O TECNICO/)) return "Cobertura de plantão / NOC";
    if (tem(/NAO PROCEDE/) && tem(/ESCALA|FOLGA|TRABALHA (NORMALMENTE|EM ESCALA)/)) return "Dia de folga ou da escala – não é hora extra";
    if ((tem(/NAO AUTORIZAD|NAO PROCEDE/) && tem(HE)) || tem(/^NAO AUTORIZADA/)) return "HE não autorizada – aplicada escala pedagógica";
    if ((tem(HE) || tem(/^HORAS? PARA|BANCO DE HORAS PARA ABAD|COMPENSA/)) && tem(/ABAT|ABAD|DESCONTO DE BANCO|COMPENS|EQUILIBRIO/) && tem(BANCO))
        return "Lançar no banco de horas";
    if (tem(/COMPENSA/) && tem(BANCO)) return "Lançar no banco de horas";
    if (tem(/DESLOC|DESLOMENTO/)) return tem(/PEDAG/) ? "Horas de deslocamento – aplicada escala pedagógica" : "Deslocamento para residência após as atividades";
    if (tem(/DESCONT|DEBIT/) && tem(/FOLHA|PAGAMENTO|SALARI/)) return "Descontar em folha de pagamento";
    if (tem(/ENVIAR PARA|DESCONT|DEBIT|ABAT|RETIRAR|DEIXAR|COLOCAR COMO|SALDO NEGATIVO|PARA ABATER/) && tem(BANCO)) return "Lançar no banco de horas";
    if (tem(/^(BANCOS? DE HORAS|FOLGA DE BANCO DE HORAS|DESCONTAR DAS HORAS)\.?$/)) return "Lançar no banco de horas";
    if (tem(/SERA COMPENSADO EM FOLGA/)) return "Será compensado em folga";
    if (tem(/ATESTADO/)) return "Atestado médico";
    if (tem(/TREINAMENTO/) && tem(/NR|ABONAR|INTEGRACAO/)) return "Treinamento (NR/integração) – abonar";
    if (tem(/(COLABORADOR|TECNICO|FUNCIONARIO) DESLIGAD|DEMISSAO|DEMITID/)) return "Colaborador desligado";
    if (tem(/DEMANDAS? INTERNAS? DA CLARO|DEMANDAS CLARO|TREI\w* TECNICOS DO P\w*JETO CLARO/)) return "Demandas da Claro";
    if (tem(/CARRO (DO COLABORADOR )?QUEBR|GUINCHO/)) return "Problema com veículo – aguardando guincho";
    if (tem(/EXAME PERIODICO/)) return "Exame periódico";
    if (tem(/PROBLEMAS PARTICULARES/)) return "Ausência por motivo particular";
    if (tem(/FALTA CONFIRMADA|APLICADO A FALTA|NAO HA BATIDA NO SISTEMA|CONSIDERAR FALTA|FALTA SEM JUSTIFICATIVA|GENTILEZA CONSIDERAR FALTA/)) return "Falta confirmada";
    if (tem(/FALTA JUSTIFICADA|ABONAR/)) return "Falta justificada – abonar";
    if (tem(/ALMOCO/)) return "Sistema não registrou o horário de almoço";
    if (tem(/SAIDA DO COLABORADOR NAO CONFERE/)) return "Saída do colaborador não confere com a marcação informada";
    if (tem(/PEDAG/)) return "Marcação indevida – aplicada escala pedagógica";
    if (tem(/ESCALA ESPANHOLA|FOLGA FIXA|FOLGA ESCALA|ESCALA DE TRABALHO DE FINAL|FINAL DE SEMANA DE ESCALA|DIA DE ESCALA|DOMINGO DE ESCALA|NAO (DEVE|DEVERA) SER CONSIDERAD|NAO E HORA EXTRA|IMPROCEDENTE|ESTAVA DE FOLGA|^FOLGA\.?$|DIA DE FOLGA|NAO TEM HORA ADICIONAL|NAO TRABALHOU NESSE DIA, POIS ERA SUA FOLGA|EXPEDIENTE NORMAL/))
        return "Dia de folga ou da escala – não é hora extra";
    if (tem(/SINCRONI|SIMCRONI|DIMCRONI|APP\b|APLICATIVO|CLOCKIN|CELULAR|CORPORATIVO|NAO (FOI )?REGISTR|NAO CONSEGUI|NAO TER CONSEGUIDO|LOG DO APARELHO|PROBLEMAS? TECNIC|ERRO DE INTERNET|SISTEMA (NAO|ESTA)|FALHA NA MARCACAO|FALHA NO APP/))
        return "Falha no registro do ponto (app/sincronização) – ajustar marcação";
    if (tem(/ESQUEC|ESQUECIMENTO|ERROU|BATEU ERRAD|BATEU INCORRET|PONTO ERRAD|BATIDA ERRAD|ERRO NA BATIDA|INCORRETA|ERRONEAMENTE|ATRASAD|COM ATRASO|INDEVID|ORIENTAD|ADVERTID|ORIETADO|ATRASAR/))
        return "Esquecimento ou erro de marcação – colaborador orientado";
    if (tem(/CONSIDER|HORARIO (PADRAO|AJUSTADO|COMERCIAL|CORRETO|NORMAL|REAL)|AJUSTAR (O )?(PONTO|HORARIO|ESCALA)|AJUSTE|ENTRADA \d|INCLUIR (BATIDA|MARCACAO|PONTO)|SUBIR MARCACAO|INICIA E ENCERRA|INICIO E ENCERROU|AJUSTADO|ALTERAR O HORARIO|ALTERACAO DA ESCALA|MUDA(R)? TURNO|TROCA DE TURNO|ESCALA (DA COLABORADORA )?ERRADA|TRABALHANDO EM ESCALA|ESCALA PRATICA|VERIFICAR PONTO|CORRECAO|CORRIGIR|NAO (MARCOU|FEZ A MARCACAO)|INCLUIR HORARIO|INSERIR MARCACAO|SUBIR HORARIO|INICIOU SUA JORNADA|NAO CHEGOU AO SITE|CHEGOU AO LOCAL DE TRABALHO SOMENTE/))
        return "Ajuste de marcação – considerar horário da escala";
    if (tem(/TRABALHOU NORMALMENTE|TRABALHOU NO HORARIO|BATEU (O PONTO )?NOS? HORARIOS? CORRET|BATEU O PONTO NO HORARIO|ACRESCENTA(DO|NDO) HE/)) return "Colaborador trabalhou normalmente – ajustar marcação";
    if (tem(/VANDALI/)) return "Recuperação de vandalismo";
    if (tem(/PREVENTIVA|CRONOGRAMA|\bWO\b|CICLO/)) return "Execução de preventiva";
    if (tem(/RELATORIO|VALIDACAO|INSERCAO/)) return "Fechamento de atividades e relatórios";
    if (tem(/TREINAMENTO|TREINAMETNO|TREIANMENTO|DUPLADO|NOVATO|CURSO/)) return "Treinamento";
    if (tem(/PLANTAO|NOC\b|COBERTURA|COBRIR|COBRE O HORARIO|AUSENCIA DE TECNICO|FALTA DE TECNICO|UNICO TECNICO|RENDIMENTO DO PLANTAO|CONDICOES CLIMATICAS/)) return "Cobertura de plantão / NOC";
    if (tem(BANCO) && tem(/ENVIAR|MANDAR|LANCAR/)) return "Lançar no banco de horas";
    if (tem(/CRQ|SOLICITACAO DE CHAVES|\bSC EMERGENCIAL/)) return "Abertura de CRQ / SC / chaves";
    if (tem(/ABASTEC/)) return "Abastecimento de GMG";
    if (tem(/CLIMATIZ|MAQUINA|ARCON|AR CONDICIONADO|VOLUTA|COMPRESSOR|ALTA TEMPERATURA|TEMPERATURA ALTA|MOTORES QUEIMAD/)) return "Corretiva de climatização (ar-condicionado)";
    if (tem(/SEM ENERGIA|FALTA DE ENERGIA|FALTANDO ENERGIA|FALHA DE ENERGIA|QUEDA DE ENERGIA|AUSENCIA (DE ENERGIA|DE CC|DA CC)|ENERGIA (DA |PELA )?CC|ENERGIA DA CONCESSIONARIA|CONCESSIONARIA|FALHA (DE|EM) AC|FALHA DE ASC|RETORNO DE AC|NORMALIZACAO D[OA] (AC|ALIMENTACAO)|GMG (PORTATIL )?CONECTADO|GGP CONECTADO|CONECT\w* (O )?(GMG|GGP|GERADOR)|ACOPLAD|FUSIVEL|FALTA DE FASE|DESLIGAMENTO PROGRAMADO|SEM ALIMENTACAO|ATE A RENDICAO|RENDIMENTO/))
        return "Falha de energia – GMG conectado no site";
    if (tem(/GMG|GGP|GERADOR|ENERG\b|BOMBA INJETORA|DIESEL/)) return "Manutenção corretiva de GMG";
    if (tem(/CONTROLADORA|PLC|FCC|BASEBAND|FANS|PLACAS|ALARME|RETIFICADOR|\bUR\b|\bDJ\b|INFRA/)) return "Troca / configuração de equipamentos no site";
    if (tem(/EMERGENC|EMERGECIAL|INCIDENTE/)) return "Atendimento emergencial";
    if (tem(/4X4|LOGISTIC|DEVOLUCAO DO VEICULO|ENTREGAR A ESCADA|ENTREGA DE CHAVES/)) return "Logística (veículo 4x4 / chaves / materiais)";
    if (tem(/ATENDIMENTO|ATENDOMENTO|ATENDIMENYO|ACIONAD|ACIONAMENTO|CHAMADO|SITE|QUEDA|CRISE|GPON|CORRETIV|RECUPERACAO|INTERMITENTE/))
        return "Atendimento a chamado / site";
    if (tem(/FECHAMENTO|FINALIZ/)) return "Fechamento de atividades e relatórios";
    if (tem(/DEMANDAS? INTERNAS? DA CLARO|DEMANDAS CLARO|PEDIDO.*CLARO|CLARO/)) return "Demandas da Claro";
    if (tem(/PENDENCI|DEMANDA|TRATATIVA|EXTENSAO DE JORNADA|ESTENDEU|ESTENDER|PASSOU DO HORARIO|ANTECIPA|ENTROU MAIS CEDO|ENTRADA ANTECIPADA|MATERIAL|ALMOX|ACOMPANHAMENTO|VISTORIA|APOIO|AUXILIO/))
        return "Extensão de jornada por demanda operacional";
    return t0;
}

// ALTERADO (30/10/2026): resposta do RH pelo painel (aba "Resposta do RH").
// Recebe { acao: "RH", itens: [{ ID, Status_RH, Observacao_RH }] } e grava Status_RH, Observacao_RH e
// Data_Ajuste_RH (hoje) na linha da pendência (pela Chave). Recalcula Status_Automatico e Status Pendência.
function gravarRH(hist: ExcelScript.Table, lista: Entrada[]): string {
    const cab = hist.getHeaderRowRange().getValues()[0].map(x => String(x).trim());
    const corpoTab = hist.getRangeBetweenHeaderAndTotal();
    const v = corpoTab.getValues();
    const ix = (c: string) => cab.indexOf(c);
    const iK = ix("Chave");
    const porChave: { [k: string]: number } = {};
    v.forEach((row, r) => { const k = String(row[iK] == null ? "" : row[iK]).trim(); if (k) porChave[k] = r; });
    const hoje = hojeBrasiliaSerial();
    const resultados: Resultado[] = [];
    const tx = (row: Cel[], c: string) => ix(c) >= 0 ? String(row[ix(c)] == null ? "" : row[ix(c)]).trim() : "";
    for (const x of lista) {
        const k = String(x.ID || "").trim();
        const r = porChave[k];
        if (!k || r == null) { resultados.push({ ID: k, ok: false, erro: "Pendência não encontrada na JUSTIFICATIVAS_PAINEL." }); continue; }
        let st = String(x.Status_RH || "").trim().toUpperCase();
        if (/DEVOLV/.test(st)) st = "NÃO REALIZADO - DEVOLVIDO";
        else if (/^OK\b|SEM NECESSIDADE/.test(st)) st = "OK - SEM NECESSIDADE DE AJUSTE";   // ALTERADO (03/11/2026)
        else if (/AJUST/.test(st)) st = "AJUSTADO";
        else { resultados.push({ ID: k, ok: false, erro: "Status_RH inválido (use AJUSTADO, OK - SEM NECESSIDADE DE AJUSTE ou DEVOLVIDO)." }); continue; }
        const row = v[r].slice();
        if (ix("Status_RH") >= 0) row[ix("Status_RH")] = st;
        if (ix("Observacao_RH") >= 0) row[ix("Observacao_RH")] = String(x.Observacao_RH || "").trim();
        if (ix("Data_Ajuste_RH") >= 0) row[ix("Data_Ajuste_RH")] = hoje;
        const s = calcularStatus({
            trat: tx(row, "Tratativa_Gestor"), tarefa: tx(row, "Tarefa_HE"), horario: tx(row, "Horario_Ajuste_Marcacoes"),
            texto: tx(row, "Justificativa_Gestor"), stRH: st
        }, tx(row, "Pagamento_HE_Permitido"), tx(row, "Pendencia"));
        if (ix("Status_Automatico") >= 0) row[ix("Status_Automatico")] = s[0];
        if (ix("Status Pendência") >= 0) row[ix("Status Pendência")] = s[1];
        corpoTab.getRow(r).setValues([row]);
        v[r] = row;
        resultados.push({ ID: k, ok: true, Status_Automatico: s[0], "Status Pendência": s[1] });
    }
    formatarDatas(hist);
    return JSON.stringify({ ok: resultados.every(x => x.ok), gravadas: resultados.filter(x => x.ok).length, resultados });
}

function gravar(hist: ExcelScript.Table, dados: string): string {
    let corpo: Corpo;
    try { corpo = JSON.parse(dados) as Corpo; }
    catch (e) { return JSON.stringify({ ok: false, erro: "Conteúdo recebido não é JSON." }); }
    const lista: Entrada[] = Array.isArray(corpo.itens) ? corpo.itens : (corpo.ID ? [corpo as Entrada] : []);
    if (!lista.length) return JSON.stringify({ ok: false, erro: "Nenhuma justificativa recebida." });
    if (String(corpo.acao || "").toUpperCase() === "RH") return gravarRH(hist, lista);

    const cab = hist.getHeaderRowRange().getValues()[0].map(x => String(x).trim());
    let corpoTab = hist.getRangeBetweenHeaderAndTotal();
    const v = corpoTab.getValues();
    const iK = cab.indexOf("Chave");
    const porChave: { [k: string]: number } = {};
    const vazias: number[] = [];
    v.forEach((row, r) => {
        const k = String(row[iK] == null ? "" : row[iK]).trim();
        if (k) porChave[k] = r; else if (row.every(x => String(x == null ? "" : x).trim() === "")) vazias.push(r);
    });
    const hoje = hojeBrasiliaSerial();
    const resultados: Resultado[] = [];

    for (const x of lista) {
        const k = String(x.ID || "").trim();
        if (!k) { resultados.push({ ID: "", ok: false, erro: "Pendência sem chave." }); continue; }
        let r: number | undefined = porChave[k];
        const antigo = r != null ? v[r] : null;
        const velho = (c: string) => antigo && cab.indexOf(c) >= 0 ? antigo[cab.indexOf(c)] : "";

        let stRH = String(velho("Status_RH") || "").trim();
        if (/DEVOLV/i.test(stRH)) stRH = "";          // gestor respondeu de novo → volta para o RH
        const s = calcularStatus({
            trat: x.Tratativa_Gestor || "", tarefa: x.Tarefa_HE || "", horario: x.Horario_Ajuste_Marcacoes || "",
            texto: x.Justificativa_Gestor || "", stRH
        }, String(x.Pagamento_HE_Permitido || ""), String(x.Pendencia || ""));

        const linha: Cel[] = cab.map(c => {
            if (c === "Chave") return k;
            if (c === "Matricula_6") return matPadrao(x.Matricula_6 || velho(c) || k.split("|")[0]);
            if (c === "Status_RH") return stRH;
            if (COLS_RH.indexOf(c) >= 0) return velho(c);
            if (c === "Data_Resposta_Gestor") return hoje;
            if (c === "Status_Automatico") return s[0];
            if (c === "Status Pendência") return s[1];
            if (c === "Origem") return "PAINEL";
            if (c === "Justificativa_Padronizada") return padronizarJust(x.Justificativa_Gestor || "");
            const novo = x[c];
            if (novo == null || novo === "") return COLS_GESTOR.indexOf(c) >= 0 ? "" : velho(c);   // dados da pendência: mantém se não veio
            if (COLS_DATA.indexOf(c) >= 0) return serialDe(String(novo));
            return novo;
        });

        if (r == null && vazias.length) r = vazias.shift();
        if (r == null) {
            hist.addRow(-1, linha);
            corpoTab = hist.getRangeBetweenHeaderAndTotal();
        } else {
            corpoTab.getRow(r).setValues([linha]);
        }
        resultados.push({ ID: k, ok: true, Status_Automatico: s[0], "Status Pendência": s[1] });
    }
    formatarDatas(hist);
    return JSON.stringify({ ok: resultados.every(x => x.ok), gravadas: resultados.filter(x => x.ok).length, resultados });
}

// ALTERADO (23/10/2026): Matricula_6 da tbJustificativasPainel sempre "02 - 006841" (texto)
function padronizarMatricula(t: ExcelScript.Table) {
    const col = t.getColumnByName("Matricula_6");
    if (!col) return;
    const r = col.getRangeBetweenHeaderAndTotal();
    if (!r) return;
    const v = r.getValues();
    let mudou = false;
    const novo = v.map(l => { const p = matPadrao(l[0]); if (p !== String(l[0] == null ? "" : l[0])) mudou = true; return [p]; });
    if (mudou) { r.setNumberFormat("@"); r.setValues(novo); }
}

// colunas de data da tbJustificativasPainel em dd/mm/aaaa (o valor gravado é a data do Excel)
function formatarDatas(t: ExcelScript.Table) {
    for (const c of COLS_DATA) {
        const col = t.getColumnByName(c);
        if (!col) continue;
        const r = col.getRangeBetweenHeaderAndTotal();
        if (r) r.setNumberFormat("dd/mm/yyyy");
    }
}

// =====================================================================
// ALTERADO (31/10/2026): VT / VR — trabalho em feriado ou fora da escala
// Tabela tbVTVRPainel (aba VT_VR_PAINEL), criada sozinha. Uma linha por colaborador e dia (Chave = matrícula|data).
//   Supervisor (acao "VTVR_SUP"): VT, VR, Observacao_Supervisor → Data_Aprovacao_Supervisor = hoje
//   RH (acao "VTVR_RH"): Status_RH (LANÇADO ou DEVOLVIDO), Observacao_RH → Data_RH = hoje
//   Se o supervisor responde de novo depois de devolvido, o Status_RH volta a ficar vazio (aguardando RH).
// =====================================================================
const ABA_VT = "VT_VR_PAINEL";
const TAB_VT = "tbVTVRPainel";
const COLS_VT: string[] = ["Chave", "Matricula_6", "Nome", "Data", "Dia", "Motivo", "Turno", "Supervisor", "Horas_HE", "VT", "VR",
    "Observacao_Supervisor", "Data_Aprovacao_Supervisor", "Status_RH", "Observacao_RH", "Data_RH", "Origem"];
const COLS_VT_DATA = ["Data", "Data_Aprovacao_Supervisor", "Data_RH"];
interface VtVr { vt: string; vr: string; obsSup: string; dataSup: string; stRH: string; obsRH: string; dataRH: string; motivo: string; }

function garantirVT(workbook: ExcelScript.Workbook): ExcelScript.Table {
    let t = workbook.getTable(TAB_VT);
    if (!t) {
        const ws = workbook.getWorksheet(ABA_VT) || workbook.addWorksheet(ABA_VT);
        const cab = ws.getRangeByIndexes(0, 0, 1, COLS_VT.length);
        cab.setValues([COLS_VT]);
        t = workbook.addTable(cab, true);
        t.setName(TAB_VT);
    }
    const tem = t.getHeaderRowRange().getValues()[0].map(x => String(x).trim());
    COLS_VT.forEach(nome => { if (tem.indexOf(nome) < 0) t.addColumn(-1, null, nome); });
    return t;
}
function lerVT(workbook: ExcelScript.Workbook): { [k: string]: VtVr } {
    const out: { [k: string]: VtVr } = {};
    const t = workbook.getTable(TAB_VT);
    if (!t) return out;
    const v = t.getRange().getValues();
    const cab = v[0].map(c => String(c).trim());
    for (let r = 1; r < v.length; r++) {
        const L = leitor(cab, v[r]);
        const k = L.txt(["Chave"]);
        if (!k) continue;
        out[k] = {
            vt: L.txt(["VT"]), vr: L.txt(["VR"]), obsSup: L.txt(["Observacao_Supervisor"]), dataSup: L.data(["Data_Aprovacao_Supervisor"]),
            stRH: L.txt(["Status_RH"]), obsRH: L.txt(["Observacao_RH"]), dataRH: L.data(["Data_RH"]), motivo: L.txt(["Motivo"])
        };
    }
    return out;
}
function gravarVT(workbook: ExcelScript.Workbook, dados: string): string {
    let corpo: Corpo;
    try { corpo = JSON.parse(dados) as Corpo; }
    catch (e) { return JSON.stringify({ ok: false, erro: "Conteúdo recebido não é JSON." }); }
    const acao = String(corpo.acao || "").toUpperCase();
    const lista: Entrada[] = Array.isArray(corpo.itens) ? corpo.itens : [];
    if (!lista.length) return JSON.stringify({ ok: false, erro: "Nenhum item recebido." });
    const t = garantirVT(workbook);
    const cab = t.getHeaderRowRange().getValues()[0].map(x => String(x).trim());
    let corpoTab = t.getRangeBetweenHeaderAndTotal();
    const v = corpoTab.getValues();
    const iK = cab.indexOf("Chave");
    const porChave: { [k: string]: number } = {};
    const vazias: number[] = [];
    v.forEach((row, r) => {
        const k = String(row[iK] == null ? "" : row[iK]).trim();
        if (k) porChave[k] = r; else if (row.every(x => String(x == null ? "" : x).trim() === "")) vazias.push(r);
    });
    const hoje = hojeBrasiliaSerial();
    const resultados: Resultado[] = [];
    for (const x of lista) {
        const k = String(x.Chave || x.ID || "").trim();
        if (!k) { resultados.push({ ID: "", ok: false, erro: "Item sem chave." }); continue; }
        let r: number | undefined = porChave[k];
        const antigo = r != null ? v[r] : null;
        const velho = (c: string): Cel => antigo && cab.indexOf(c) >= 0 ? antigo[cab.indexOf(c)] : "";
        if (acao === "VTVR_RH" && r == null) { resultados.push({ ID: k, ok: false, erro: "O supervisor ainda não aprovou este VT/VR." }); continue; }
        let st = "";
        if (acao === "VTVR_RH") {
            st = String(x.Status_RH || "").toUpperCase();
            st = /DEVOLV/.test(st) ? "DEVOLVIDO" : /LAN/.test(st) ? "LANÇADO" : "";
            if (!st) { resultados.push({ ID: k, ok: false, erro: "Status_RH inválido (use LANÇADO ou DEVOLVIDO)." }); continue; }
        }
        const linha: Cel[] = cab.map(c => {
            if (c === "Chave") return k;
            if (acao === "VTVR_RH") {
                if (c === "Status_RH") return st;
                if (c === "Observacao_RH") return String(x.Observacao_RH || "").trim();
                if (c === "Data_RH") return hoje;
                return velho(c);
            }
            // supervisor
            if (c === "Data_Aprovacao_Supervisor") return hoje;
            if (c === "Status_RH" || c === "Observacao_RH" || c === "Data_RH") return /DEVOLV/i.test(String(velho("Status_RH"))) ? "" : velho(c);
            if (c === "Origem") return "PAINEL";
            if (c === "Matricula_6") return matPadrao(x.Matricula_6 || velho(c));
            const novo = x[c];
            if (novo == null || novo === "") return velho(c);
            if (COLS_VT_DATA.indexOf(c) >= 0) return serialDe(String(novo));
            return novo;
        });
        if (r == null && vazias.length) r = vazias.shift();
        if (r == null) { t.addRow(-1, linha); corpoTab = t.getRangeBetweenHeaderAndTotal(); }
        else corpoTab.getRow(r).setValues([linha]);
        resultados.push({ ID: k, ok: true });
    }
    for (const c of COLS_VT_DATA) {
        const col = t.getColumnByName(c);
        const rg = col ? col.getRangeBetweenHeaderAndTotal() : null;
        if (rg) rg.setNumberFormat("dd/mm/yyyy");
    }
    return JSON.stringify({ ok: resultados.every(x => x.ok), gravadas: resultados.filter(x => x.ok).length, resultados });
}

// =====================================================================
// tabela do histórico (+ carga inicial a partir da TRATATIVA_GESTORES)
// =====================================================================
function garantirHistorico(workbook: ExcelScript.Workbook): ExcelScript.Table {
    let t = workbook.getTable(TAB_HIST);
    if (!t) {
        const ws = workbook.getWorksheet(ABA_HIST) || workbook.addWorksheet(ABA_HIST);
        const cab = ws.getRangeByIndexes(0, 0, 1, COLS_HIST.length);
        cab.setValues([COLS_HIST]);
        t = workbook.addTable(cab, true);
        t.setName(TAB_HIST);
    }
    const tem = t.getHeaderRowRange().getValues()[0].map(x => String(x).trim());
    COLS_HIST.forEach(nome => { if (tem.indexOf(nome) < 0) t.addColumn(-1, null, nome); });

    // carga inicial: só se o histórico estiver vazio
    const atual = t.getRangeBetweenHeaderAndTotal().getValues();
    const vazio = atual.every(row => row.every(x => String(x == null ? "" : x).trim() === ""));
    const antiga = workbook.getTable(TAB_TRATATIVA_ANTIGA);
    if (vazio && antiga) {
        const av = antiga.getRange().getValues();
        const ac = av[0].map(c => String(c).trim());
        const cab = t.getHeaderRowRange().getValues()[0].map(x => String(x).trim());
        const linhas: Cel[][] = [];
        for (let r = 1; r < av.length; r++) {
            const L = leitor(ac, av[r]);
            const tipo = L.txt(["Pendencia"]).toUpperCase(), m = L.txt(["Matricula_6"]), d = L.data(["Data"]);
            if (!tipo || !m || !d) continue;
            linhas.push(cab.map(c => {
                if (c === "Chave") return chave(m, d, tipo);
                if (c === "Origem") return "TRATATIVA_GESTORES";
                if (c === "Horas_Pendencia_h") return L.horas();
                const k = ac.indexOf(c === "Nome" ? "Nome" : c);
                const val = k >= 0 ? av[r][k] : "";
                return typeof val === "string" && val.charAt(0) === "#" ? "" : val;   // #N/A das fórmulas antigas
            }));
        }
        if (linhas.length) {
            const ws = t.getWorksheet();
            const ini = t.getHeaderRowRange().getRowIndex() + 1, col = t.getHeaderRowRange().getColumnIndex();
            t.resize(ws.getRangeByIndexes(ini - 1, col, linhas.length + 1, cab.length));
            t.getRangeBetweenHeaderAndTotal().setValues(linhas);
        }
    }
    formatarDatas(t);
    padronizarMatricula(t);
    return t;
}

// =====================================================================
// utilidades
// =====================================================================
interface Leitor { txt: (n: string[]) => string; num: (n: string[]) => number; data: (n: string[]) => string; horas: () => number; heFolha: () => number; }
function leitor(cab: string[], row: Cel[]): Leitor {
    const pega = (nomes: string[]): Cel => { for (const n of nomes) { const k = cab.indexOf(n); if (k >= 0 && row[k] != null && row[k] !== "") return row[k]; } return ""; };
    return {
        txt: n => { const v = pega(n); const s = String(v).trim(); return s.charAt(0) === "#" ? "" : s; },
        num: n => Number(pega(n)) || 0,
        data: n => dataIso(pega(n)),
        // horas da pendência: Horas_Pendencia_h já vem em horas; Tempo_Pendencia vem como hora do Excel (fração do dia)
        horas: () => {
            const h = pega(["Horas_Pendencia_h"]);
            if (h !== "") return Math.round((Number(h) || 0) * 100) / 100;
            const t = Number(pega(["Tempo_Pendencia"])) || 0;
            return Math.round(t * 24 * 100) / 100;
        },
        // HE da folha atual: coluna da folha na base (HE_Folha, Horas_Extras_Folha ...)
        heFolha: () => {
            const f = pega(["HE_Folha", "HE Folha", "Horas_Extras_Folha", "Horas Extras Folha", "HE_Folha_h"]);
            if (f !== "") {
                const n = Number(f);
                if (!isNaN(n)) return n > 0 && n < 1 && !/_h$/i.test(String(f)) ? Math.round(n * 24 * 100) / 100 : n;  // hora do Excel ou horas
                const m = String(f).match(/(\d+):(\d{2})/);
                if (m) return Number(m[1]) + Number(m[2]) / 60;
            }
            return 0;
        }
    };
}
function chave(m: string, dIso: string, tipo: string): string {
    const dig = String(m).replace(/\D/g, "");
    return ("000000" + dig).slice(-6) + "|" + dIso + "|" + String(tipo).trim().toUpperCase();
}

// MESMA regra da fórmula do Status_Automatico e da LISTAS (de/para do Status Pendência).
function calcularStatus(j: { trat: string; tarefa: string; horario: string; texto: string; stRH: string }, hePerm: string, tipo: string): string[] {
    const up = (v: string) => String(v || "").trim().toUpperCase();
    const trat = up(j.trat), rh = up(j.stRH), pend = up(tipo);
    const auto =
        !trat ? "AGUARDANDO GESTOR" :
            !String(j.texto || "").trim() ? "INCOMPLETO - FALTA JUSTIFICATIVA" :
                rh === "NÃO REALIZADO - DEVOLVIDO" ? "DEVOLVIDO AO GESTOR" :
                    trat === "PAGAMENTO HE" && pend !== "HORA EXTRA" ? "INVÁLIDO - PAGAMENTO HE SOMENTE PARA HORA EXTRA" :
                        trat === "BANCO DE HORAS" && pend !== "HORA EXTRA" && pend !== "ATRASO" && pend !== "FALTA" ? "INVÁLIDO - BANCO DE HORAS SOMENTE PARA HORA EXTRA" :
                            trat === "PAGAMENTO HE" && up(hePerm) === "NÃO" ? "INVÁLIDO - BH NEGATIVO NÃO PERMITE PAGAMENTO" :
                                trat === "PAGAMENTO HE" && !String(j.tarefa || "").trim() ? "INCOMPLETO - FALTA TAREFA HE" :
                                    trat === "AJUSTAR PONTO" && (pend === "PONTO ABERTO" || pend === "SEM MARCACAO") && !String(j.horario || "").trim() ? "INCOMPLETO - FALTA HORÁRIO/MARCAÇÃO" :
                                        /^OK\b/.test(rh) ? "CONCLUÍDO" :   // ALTERADO (03/11/2026): RH "OK - Sem necessidade de ajuste"
                                            ["PAGAMENTO HE", "AJUSTAR PONTO", "AUSENTE", "FALTA INDEVIDA"].indexOf(trat) >= 0 && rh === "AJUSTADO" ? "CONCLUÍDO" :
                                                trat === "PAGAMENTO HE" ? "PRONTO PARA RH - HE" :
                                                    trat === "AJUSTAR PONTO" ? "AGUARDANDO RH - AJUSTE DE PONTO" :
                                                        trat === "AUSENTE" ? "AGUARDANDO RH - AUSÊNCIA" :
                                                            trat === "FALTA INDEVIDA" ? "AGUARDANDO RH - FALTA INDEVIDA" :
                                                                trat === "BANCO DE HORAS" ? "TRATADO - BANCO DE HORAS" :
                                                                    trat === "APENAS JUSTIFICAR" ? "TRATADO - JUSTIFICADO" : "CONCLUÍDO";
    let status = "Justificado";
    if (auto === "AGUARDANDO GESTOR") status = "Pendente Operação";
    else if (/^(INCOMPLETO|INVÁLIDO|DEVOLVIDO)/.test(auto)) status = "Erro na Justificativa";
    return [auto, status];
}

// HE do dia: o "HE: hh:mm" do Detalhe (calculado pela base com batidas x turno); se não tiver, as horas da pendência
function heDoDia(det: string, horas: number): string {
    const m = String(det || "").match(/HE:\s*(\d{1,3}:\d{2})/i);
    if (m) return m[1].padStart(5, "0");
    return hhmm(horas);
}
// período da folha: sempre do dia 15 ao dia 14 do mês seguinte
//   batida 20/10 → 15/10 a 14/11 · batida 10/10 → 15/09 a 14/10
function periodoFolha(iso: string): string[] {
    const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return ["", ""];
    let a = +m[1], mes = +m[2];
    if (+m[3] < 15) { mes--; if (mes < 1) { mes = 12; a--; } }
    const ini = new Date(Date.UTC(a, mes - 1, 15)), fim = new Date(Date.UTC(a, mes, 14));
    const br = (d: Date) => String(d.getUTCDate()).padStart(2, "0") + "/" + String(d.getUTCMonth() + 1).padStart(2, "0") + "/" + d.getUTCFullYear();
    return [ini.toISOString().slice(0, 10), br(ini) + " a " + br(fim)];
}
function hhmm(h: number): string {
    if (!h) return "";
    const min = Math.round(h * 60);
    return String(Math.floor(min / 60)).padStart(2, "0") + ":" + String(min % 60).padStart(2, "0");
}
function hojeBrasiliaIso(): string {
    return new Date(Date.now() - 3 * 3600000).toISOString().slice(0, 10);
}
// soma n dias úteis (seg a sex) a uma data aaaa-mm-dd
// ALTERADO (03/11/2026): feriados não contam como dia útil (nacionais + estado e cidade de SP — complete aqui)
const FERIADOS_UTEIS: string[] = ["2025-11-02", "2025-11-15", "2025-11-20", "2025-12-25", "2026-01-01", "2026-01-25", "2026-02-16", "2026-02-17", "2026-04-03", "2026-04-21", "2026-05-01", "2026-06-04", "2026-07-09", "2026-09-07", "2026-10-12", "2026-11-02", "2026-11-15", "2026-11-20", "2026-12-25", "2027-01-01", "2027-01-25"];
function somarDiasUteis(iso: string, n: number): string {
    const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return iso;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    let falta = n;
    while (falta > 0) {
        d.setUTCDate(d.getUTCDate() + 1);
        const w = d.getUTCDay();
        if (w !== 0 && w !== 6 && FERIADOS_UTEIS.indexOf(d.toISOString().slice(0, 10)) < 0) falta--;
    }
    return d.toISOString().slice(0, 10);
}
function hojeBrasiliaSerial(): number {
    const d = new Date(Date.now() - 3 * 3600000);
    return Math.floor(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) / 86400000) + 25569;
}
function serialDe(iso: string): Cel {
    const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? Math.floor(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000) + 25569 : iso;
}
function dataIso(v: Cel): string {
    if (typeof v === "number") { if (!v) return ""; return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10); }
    const s = String(v || "").trim();
    const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (br) return br[3] + "-" + ("0" + br[2]).slice(-2) + "-" + ("0" + br[1]).slice(-2);
    return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : "";
}