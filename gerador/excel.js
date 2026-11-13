// =====================================================================
// Leitor de .xlsx + imitação do "ExcelScript" (Office Scripts) em Node.
// Permite rodar os MESMOS scripts do Excel (Automatizar) fora do Excel.
// Lê só valores (o que getValues() devolve no Excel): números (datas = número
// serial do Excel), textos, verdadeiro/falso. Escritas ficam só na memória.
// =====================================================================
"use strict";
const JSZip = require("jszip");

// ---------- utilidades de XML ----------
function decodeXml(s) {
  if (s.indexOf("&") >= 0) {
    s = s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
      if (e === "amp") return "&"; if (e === "lt") return "<"; if (e === "gt") return ">";
      if (e === "quot") return '"'; if (e === "apos") return "'";
      if (e[1] === "x" || e[1] === "X") return String.fromCodePoint(parseInt(e.slice(2), 16));
      return String.fromCodePoint(parseInt(e.slice(1), 10));
    });
  }
  if (s.indexOf("_x") >= 0) s = s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  return s;
}
function attr(tag, nome) {
  const m = new RegExp("\\s" + nome + "\\s*=\\s*\"([^\"]*)\"").exec(tag);
  return m ? decodeXml(m[1]) : null;
}
function textosT(xml) {   // junta todos os <t>...</t> (texto rico)
  let out = "";
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
  let m;
  while ((m = re.exec(xml))) if (m[1] !== undefined) out += m[1];
  return decodeXml(out);
}
function colNum(letras) {
  let n = 0;
  for (let i = 0; i < letras.length; i++) n = n * 26 + (letras.charCodeAt(i) - 64);
  return n - 1;
}
function colLetras(n) {
  let s = ""; n++;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function parseRef(ref) {   // "B3" -> {r:2,c:1}; "B" -> {c:1}; "3" -> {r:2}
  const m = /^\$?([A-Z]*)\$?(\d*)$/.exec(ref.toUpperCase());
  if (!m) throw new Error("Endereço inválido: " + ref);
  return { c: m[1] ? colNum(m[1]) : null, r: m[2] ? parseInt(m[2], 10) - 1 : null };
}
function caminhoRel(base, alvo) {
  if (alvo.startsWith("/")) return alvo.slice(1);
  const partes = base.split("/"); partes.pop();
  for (const p of alvo.split("/")) { if (p === "..") partes.pop(); else if (p !== ".") partes.push(p); }
  return partes.join("/");
}

// formatos de número de hora/duração (para getTexts devolver "hh:mm")
const FMT_HORA_EMBUTIDO = new Set([18, 19, 20, 21, 22, 45, 46, 47]);
function ehFormatoHora(fmt) {
  const f = String(fmt || "").replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, m => (/^\[h+\]|\[m+\]|\[s+\]$/i.test(m) ? "h" : ""));
  return /h|s/i.test(f) || /:/.test(f);
}
function textoHora(v) {
  const seg = Math.round(v * 86400);
  const h = Math.floor(seg / 3600), m = Math.floor((seg % 3600) / 60), s = seg % 60;
  return h + ":" + String(m).padStart(2, "0") + ":" + String(s).padStart(2, "0");
}

// ---------- leitura da planilha de uma aba (varre o XML em bytes, sem carregar tudo como texto) ----------
function lerAba(buf, shared, estiloHora) {
  const linhas = [];          // linhas[r][c] = valor
  const horas = new Set();    // "r,c" de células com formato de hora
  let maxR = -1, maxC = -1, maxRV = -1, maxCV = -1, minR = Infinity, minC = Infinity, minRV = Infinity, minCV = Infinity;
  const ABRE_C = Buffer.from("<c");
  let pos = buf.indexOf("<sheetData");
  if (pos < 0) return { linhas, horas, usado: null, usadoValores: null };
  const fim = buf.lastIndexOf("</sheetData>");
  let linhaAtual = -1, colAtual = -1;
  const ABRE_ROW = Buffer.from("<row");
  // percorre <row ...> e <c ...> em ordem
  let pR = buf.indexOf(ABRE_ROW, pos);
  while (pos >= 0 && pos < fim) {
    if (pR >= 0 && pR < pos) pR = buf.indexOf(ABRE_ROW, pos);
    if (pR > fim) pR = -1;
    let pC = buf.indexOf(ABRE_C, pos);
    if (pC > fim) pC = -1;
    if (pC < 0 && pR < 0) break;
    if (pR >= 0 && (pC < 0 || pR < pC)) {
      const fecha = buf.indexOf(62, pR);   // '>'
      const tag = buf.toString("utf8", pR, fecha + 1);
      const rAttr = attr(tag, "r");
      linhaAtual = rAttr ? parseInt(rAttr, 10) - 1 : linhaAtual + 1;
      colAtual = -1;
      pos = fecha + 1;
      continue;
    }
    // garante que é "<c " ou "<c>" e não "<col" etc.
    const prox = buf[pC + 2];
    if (prox !== 32 && prox !== 62 && prox !== 47 && prox !== 9 && prox !== 10 && prox !== 13) { pos = pC + 2; continue; }
    const fechaTag = buf.indexOf(62, pC);
    const autoFecha = buf[fechaTag - 1] === 47;
    const tag = buf.toString("utf8", pC, fechaTag + 1);
    let corpo = "";
    let depois = fechaTag + 1;
    if (!autoFecha) {
      const fimC = buf.indexOf("</c>", fechaTag);
      corpo = buf.toString("utf8", fechaTag + 1, fimC);
      depois = fimC + 4;
    }
    pos = depois;
    const ref = attr(tag, "r");
    let r = linhaAtual, c;
    if (ref) { const p = parseRef(ref); c = p.c; if (p.r !== null) r = p.r; } else c = colAtual + 1;
    colAtual = c;
    if (r > maxR) maxR = r; if (c > maxC) maxC = c; if (r < minR) minR = r; if (c < minC) minC = c;
    const t = attr(tag, "t");
    let valor = null;
    if (t === "inlineStr") {
      valor = textosT(corpo);
    } else {
      const mv = /<v>([\s\S]*?)<\/v>/.exec(corpo);
      if (mv) {
        const raw = mv[1];
        if (t === "s") valor = shared[parseInt(raw, 10)] ?? "";
        else if (t === "str") valor = decodeXml(raw);
        else if (t === "b") valor = raw === "1" || raw === "true";
        else if (t === "e") valor = decodeXml(raw);
        else if (t === "d") valor = decodeXml(raw);
        else { const n = Number(raw); valor = isNaN(n) ? decodeXml(raw) : n; }
      } else if (t === "str") valor = "";
    }
    if (valor === null) continue;
    if (!linhas[r]) linhas[r] = [];
    linhas[r][c] = valor;
    if (valor !== "") {
      if (r > maxRV) maxRV = r; if (c > maxCV) maxCV = c; if (r < minRV) minRV = r; if (c < minCV) minCV = c;
    }
    if (typeof valor === "number") {
      const s = attr(tag, "s");
      if (s && estiloHora[parseInt(s, 10)]) horas.add(r + "," + c);
    }
  }
  return {
    linhas, horas,
    usado: maxR >= 0 ? { r0: minR, c0: minC, r1: maxR, c1: maxC } : null,
    usadoValores: maxRV >= 0 ? { r0: minRV, c0: minCV, r1: maxRV, c1: maxCV } : null
  };
}

async function lerXlsx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const texto = async nome => { const f = zip.file(nome); return f ? await f.async("string") : null; };

  // textos compartilhados
  const shared = [];
  const ss = await texto("xl/sharedStrings.xml");
  if (ss) {
    const re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
    let m;
    while ((m = re.exec(ss))) shared.push(m[1] === undefined ? "" : textosT(m[1]));
  }

  // estilos: quais índices de estilo são formato de hora
  const estiloHora = [];
  const st = await texto("xl/styles.xml");
  if (st) {
    const fmts = {};
    const reF = /<numFmt\s[^>]*>/g; let m;
    while ((m = reF.exec(st))) fmts[attr(m[0], "numFmtId")] = attr(m[0], "formatCode");
    const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(st);
    if (xfs) {
      const reX = /<xf\s[^>]*?\/?>/g; let i = 0;
      while ((m = reX.exec(xfs[1]))) {
        const id = parseInt(attr(m[0], "numFmtId") || "0", 10);
        estiloHora[i++] = FMT_HORA_EMBUTIDO.has(id) || (fmts[id] !== undefined && ehFormatoHora(fmts[id]));
      }
    }
  }

  // abas e tabelas
  const wbXml = await texto("xl/workbook.xml");
  const rels = await texto("xl/_rels/workbook.xml.rels") || "";
  const alvoRel = {};
  { const re = /<Relationship\s[^>]*>/g; let m; while ((m = re.exec(rels))) alvoRel[attr(m[0], "Id")] = attr(m[0], "Target"); }
  const abas = [];
  const tabelas = [];
  const reS = /<sheet\s[^>]*>/g; let m;
  while ((m = reS.exec(wbXml))) {
    const nome = attr(m[0], "name");
    const rid = attr(m[0], "r:id");
    const estado = attr(m[0], "state");
    const arq = caminhoRel("xl/workbook.xml", alvoRel[rid] || "");
    const f = zip.file(arq);
    if (!f) { abas.push(new Aba(nome, { linhas: [], horas: new Set(), usado: null, usadoValores: null }, estado)); continue; }
    const buf = await f.async("nodebuffer");
    abas.push(new Aba(nome, lerAba(buf, shared, estiloHora), estado));
    // tabelas desta aba
    const relsAba = arq.replace(/([^/]+)$/, "_rels/$1.rels");
    const ra = await texto(relsAba);
    if (ra) {
      const re = /<Relationship\s[^>]*>/g; let r;
      while ((r = re.exec(ra))) {
        if (!/\/table"?$/.test(attr(r[0], "Type") || "")) continue;
        const tx = await texto(caminhoRel(arq, attr(r[0], "Target")));
        if (!tx) continue;
        const tag = /<table\s[^>]*>/.exec(tx)[0];
        const ref = attr(tag, "ref").split(":");
        const a = parseRef(ref[0]), b = parseRef(ref[1] || ref[0]);
        tabelas.push({ nome: attr(tag, "displayName") || attr(tag, "name"), aba: nome, r0: a.r, c0: a.c, r1: b.r, c1: b.c,
          totais: parseInt(attr(tag, "totalsRowCount") || "0", 10), cabecalho: attr(tag, "headerRowCount") !== "0" });
      }
    }
  }
  return new Pasta(abas, tabelas);
}

// =====================================================================
// Imitação do ExcelScript
// =====================================================================
class Aba {
  constructor(nome, dados, estado) {
    this.nome = nome; this.linhas = dados.linhas; this.horas = dados.horas;
    this._usado = dados.usado; this._usadoV = dados.usadoValores; this.visivel = estado || "visible";
  }
  val(r, c) { const l = this.linhas[r]; if (!l) return ""; const v = l[c]; return v === undefined || v === null ? "" : v; }
  set(r, c, v) {
    if (!this.linhas[r]) this.linhas[r] = [];
    this.linhas[r][c] = v;
    const u = this._usado;
    if (!u) this._usado = { r0: r, c0: c, r1: r, c1: c };
    else { u.r0 = Math.min(u.r0, r); u.c0 = Math.min(u.c0, c); u.r1 = Math.max(u.r1, r); u.c1 = Math.max(u.c1, c); }
    if (v !== "" && v !== null && v !== undefined) {
      const w = this._usadoV;
      if (!w) this._usadoV = { r0: r, c0: c, r1: r, c1: c };
      else { w.r0 = Math.min(w.r0, r); w.c0 = Math.min(w.c0, c); w.r1 = Math.max(w.r1, r); w.c1 = Math.max(w.c1, c); }
    }
  }
}

class Range {
  constructor(pasta, aba, r0, c0, nr, nc) { this.p = pasta; this.a = aba; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; }
  getValues() {
    const out = new Array(this.nr);
    for (let i = 0; i < this.nr; i++) {
      const l = this.a.linhas[this.r0 + i];
      const row = new Array(this.nc);
      for (let j = 0; j < this.nc; j++) { const v = l ? l[this.c0 + j] : undefined; row[j] = v === undefined || v === null ? "" : v; }
      out[i] = row;
    }
    return out;
  }
  getTexts() {
    return this.getValues().map((row, i) => row.map((v, j) => {
      if (typeof v === "number" && this.a.horas.has((this.r0 + i) + "," + (this.c0 + j))) return textoHora(v);
      if (typeof v === "boolean") return v ? "VERDADEIRO" : "FALSO";
      return String(v);
    }));
  }
  getValue() { return this.a.val(this.r0, this.c0); }
  getText() { return this.getTexts()[0][0]; }
  getFormulas() { return this.getValues(); }
  setValues(v) {
    if (v.length !== this.nr || (v[0] && v[0].length !== this.nc)) {
      throw new Error("setValues: tamanho " + v.length + "x" + (v[0] ? v[0].length : 0) + " diferente do intervalo " + this.nr + "x" + this.nc);
    }
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.a.set(this.r0 + i, this.c0 + j, v[i][j]);
  }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.a.set(this.r0 + i, this.c0 + j, v); }
  setFormulas(v) { this.setValues(v); }
  getRowCount() { return this.nr; }
  getColumnCount() { return this.nc; }
  getRowIndex() { return this.r0; }
  getColumnIndex() { return this.c0; }
  getRow(i) { return new Range(this.p, this.a, this.r0 + i, this.c0, 1, this.nc); }
  getColumn(j) { return new Range(this.p, this.a, this.r0, this.c0 + j, this.nr, 1); }
  getCell(i, j) { return new Range(this.p, this.a, this.r0 + i, this.c0 + j, 1, 1); }
  getLastRow() { return this.getRow(this.nr - 1); }
  getLastColumn() { return this.getColumn(this.nc - 1); }
  getOffsetRange(dr, dc) { return new Range(this.p, this.a, this.r0 + dr, this.c0 + dc, this.nr, this.nc); }
  getResizedRange(dr, dc) { return new Range(this.p, this.a, this.r0, this.c0, this.nr + dr, this.nc + dc); }
  getEntireColumn() { return new Range(this.p, this.a, 0, this.c0, 1048576, this.nc); }
  getEntireRow() { return new Range(this.p, this.a, this.r0, 0, this.nr, 16384); }
  getWorksheet() { return new Worksheet(this.p, this.a); }
  getAddress() {
    return this.a.nome + "!" + colLetras(this.c0) + (this.r0 + 1) + (this.nr * this.nc > 1 ? ":" + colLetras(this.c0 + this.nc - 1) + (this.r0 + this.nr) : "");
  }
  clear() { for (let i = 0; i < this.nr; i++) { const l = this.a.linhas[this.r0 + i]; if (l) for (let j = 0; j < this.nc; j++) l[this.c0 + j] = undefined; } }
  delete() { this.clear(); }
  setNumberFormat() {} setNumberFormats() {} setNumberFormatLocal() {}
  getNumberFormats() { return this.getValues().map(r => r.map(() => "General")); }
  getFormat() { return FORMATO; }
  select() {}
}
const nada = () => {};
const FONTE = { setBold: nada, setColor: nada, setSize: nada, setItalic: nada, setName: nada };
const FORMATO = {
  getFont: () => FONTE, getFill: () => ({ setColor: nada, clear: nada }), setColumnWidth: nada, setRowHeight: nada,
  autofitColumns: nada, autofitRows: nada, setHorizontalAlignment: nada, setVerticalAlignment: nada, setWrapText: nada,
  getBorders: () => [], getRangeBorder: () => ({ setStyle: nada, setColor: nada })
};

class Worksheet {
  constructor(pasta, aba) { this.p = pasta; this.a = aba; }
  getName() { return this.a.nome; }
  setName(n) { this.a.nome = n; }
  getUsedRange(somenteValores) {
    const u = somenteValores ? this.a._usadoV : this.a._usado;
    if (!u) return undefined;
    return new Range(this.p, this.a, u.r0, u.c0, u.r1 - u.r0 + 1, u.c1 - u.c0 + 1);
  }
  getRangeByIndexes(r, c, nr, nc) { return new Range(this.p, this.a, r, c, nr, nc); }
  getRange(endereco) {
    if (!endereco) return new Range(this.p, this.a, 0, 0, 1048576, 16384);
    let e = String(endereco); if (e.indexOf("!") >= 0) e = e.split("!").pop();
    const [x, y] = e.replace(/'/g, "").split(":");
    const a = parseRef(x), b = parseRef(y || x);
    const r0 = a.r === null ? 0 : a.r, c0 = a.c === null ? 0 : a.c;
    const r1 = b.r === null ? 1048575 : b.r, c1 = b.c === null ? 16383 : b.c;
    return new Range(this.p, this.a, r0, c0, r1 - r0 + 1, c1 - c0 + 1);
  }
  getCell(r, c) { return new Range(this.p, this.a, r, c, 1, 1); }
  getTables() { return this.p.tabelas.filter(t => t.aba === this.a).map(t => new Table(this.p, t)); }
  getTable(n) { const t = this.getTables().filter(x => x.getName().toLowerCase() === String(n).toLowerCase())[0]; return t; }
  addTable(range, temCabecalho) { return this.p.wb.addTable(typeof range === "string" ? this.getRange(range) : range, temCabecalho); }
  setVisibility(v) { this.a.visivel = v; }
  getVisibility() { return this.a.visivel; }
  delete() { this.p.abas = this.p.abas.filter(x => x !== this.a); this.p.tabelas = this.p.tabelas.filter(t => t.aba !== this.a); }
  activate() {}
  getPosition() { return this.p.abas.indexOf(this.a); }
  getAutoFilter() { return { remove: nada, apply: nada }; }
  getFreezePanes() { return { freezeRows: nada, freezeColumns: nada, unfreeze: nada }; }
}

class TableColumn {
  constructor(t, j) { this.t = t; this.j = j; }
  getName() { return String(this.t.getHeaderRowRange().getValues()[0][this.j]); }
  getIndex() { return this.j; }
  getRange() { const d = this.t.d; return new Range(this.t.p, d.aba, d.r0, d.c0 + this.j, d.r1 - d.r0 + 1, 1); }
  getRangeBetweenHeaderAndTotal() { const d = this.t.d; return new Range(this.t.p, d.aba, d.r0 + 1, d.c0 + this.j, d.r1 - d.totais - d.r0, 1); }
  getHeaderRowRange() { const d = this.t.d; return new Range(this.t.p, d.aba, d.r0, d.c0 + this.j, 1, 1); }
}

class Table {
  constructor(pasta, d) { this.p = pasta; this.d = d; }
  getName() { return this.d.nome; }
  setName(n) { this.d.nome = n; }
  getWorksheet() { return new Worksheet(this.p, this.d.aba); }
  getRange() { const d = this.d; return new Range(this.p, d.aba, d.r0, d.c0, d.r1 - d.r0 + 1, d.c1 - d.c0 + 1); }
  getHeaderRowRange() { const d = this.d; return new Range(this.p, d.aba, d.r0, d.c0, 1, d.c1 - d.c0 + 1); }
  getRangeBetweenHeaderAndTotal() { const d = this.d; return new Range(this.p, d.aba, d.r0 + 1, d.c0, d.r1 - d.totais - d.r0, d.c1 - d.c0 + 1); }
  getRowCount() { return this.d.r1 - this.d.totais - this.d.r0; }
  getColumns() { const n = this.d.c1 - this.d.c0 + 1; const out = []; for (let j = 0; j < n; j++) out.push(new TableColumn(this, j)); return out; }
  getColumnByName(nome) {
    const cab = this.getHeaderRowRange().getValues()[0].map(x => String(x));
    const j = cab.indexOf(String(nome));
    return j < 0 ? undefined : new TableColumn(this, j);
  }
  getColumn(j) { return new TableColumn(this, j); }
  addRow(indice, valores) { this.addRows(indice, valores ? [valores] : [[]]); }
  addRows(indice, linhas) {
    const d = this.d, nc = d.c1 - d.c0 + 1;
    // Excel: uma tabela recém-criada tem 1 linha de dados vazia, que é usada pela primeira linha nova
    const corpo = this.getRangeBetweenHeaderAndTotal();
    const vazia = d.criadaVazia && corpo.getRowCount() === 1 && corpo.getValues()[0].every(x => x === "");
    let alvo = vazia ? d.r0 + 1 : d.r1 + 1;
    if (indice !== undefined && indice !== null && indice >= 0 && !vazia) {
      // insere no meio: desloca as linhas de baixo
      const inicio = d.r0 + 1 + indice;
      for (let r = d.r1; r >= inicio; r--) for (let j = 0; j < nc; j++) d.aba.set(r + linhas.length, d.c0 + j, d.aba.val(r, d.c0 + j));
      alvo = inicio;
    }
    linhas.forEach((l, i) => { for (let j = 0; j < nc; j++) d.aba.set(alvo + i, d.c0 + j, l && l[j] !== undefined ? l[j] : ""); });
    d.r1 = vazia ? d.r0 + linhas.length : d.r1 + linhas.length;
    d.criadaVazia = false;
  }
  addColumn(indice, valores, nome) {
    const d = this.d;
    d.c1++;
    d.aba.set(d.r0, d.c1, nome || ("Coluna" + (d.c1 - d.c0 + 1)));
    if (valores) valores.forEach((v, i) => { if (i > 0) d.aba.set(d.r0 + i, d.c1, Array.isArray(v) ? v[0] : v); });
  }
  resize(range) { const d = this.d; d.r0 = range.r0; d.c0 = range.c0; d.r1 = range.r0 + range.nr - 1; d.c1 = range.c0 + range.nc - 1; d.criadaVazia = false; }
  delete() { this.p.tabelas = this.p.tabelas.filter(t => t !== this.d); }
  getAutoFilter() { return { clearCriteria: nada, apply: nada }; }
  getSort() { return { apply: nada, clear: nada }; }
  setShowTotals() {} setPredefinedTableStyle() {} setShowBandedRows() {}
}

class Pasta {
  constructor(abas, tabelas) {
    this.abas = abas;
    this.tabelas = tabelas.map(t => Object.assign({}, t, { aba: abas.find(a => a.nome === t.aba) })).filter(t => t.aba);
    this.wb = new Workbook(this);
  }
}

class Workbook {
  constructor(p) { this.p = p; }
  getWorksheets() { return this.p.abas.map(a => new Worksheet(this.p, a)); }
  getWorksheet(nome) {
    const n = String(nome).toLowerCase();
    const a = this.p.abas.find(x => x.nome.toLowerCase() === n);
    return a ? new Worksheet(this.p, a) : undefined;
  }
  addWorksheet(nome) {
    let n = nome || ("Planilha" + (this.p.abas.length + 1));
    if (this.getWorksheet(n)) throw new Error("Já existe uma aba chamada " + n);
    const a = new Aba(n, { linhas: [], horas: new Set(), usado: null, usadoValores: null });
    this.p.abas.push(a);
    return new Worksheet(this.p, a);
  }
  getTables() { return this.p.tabelas.map(t => new Table(this.p, t)); }
  getTable(nome) {
    const n = String(nome).toLowerCase();
    const t = this.p.tabelas.find(x => x.nome.toLowerCase() === n);
    return t ? new Table(this.p, t) : undefined;
  }
  addTable(range, temCabecalho) {
    const r = typeof range === "string" ? this.getActiveWorksheet().getRange(range) : range;
    const d = { nome: "Tabela" + (this.p.tabelas.length + 1), aba: r.a, r0: r.r0, c0: r.c0, r1: r.r0 + r.nr - 1, c1: r.c0 + r.nc - 1, totais: 0, cabecalho: true };
    if (r.nr === 1) { d.r1 = r.r0 + 1; d.criadaVazia = true; }   // cabeçalho + 1 linha vazia, como no Excel
    this.p.tabelas.push(d);
    return new Table(this.p, d);
  }
  getActiveWorksheet() { return new Worksheet(this.p, this.p.abas[0]); }
  getName() { return this.p.nome || "pasta.xlsx"; }
  getNamedItem() { return undefined; }
  refreshAllDataConnections() {}
  getApplication() { return { calculate: nada, getCalculationMode: () => "Automatic" }; }
}

// enums usados pelos scripts
const ExcelScript = {
  ClearApplyTo: { all: "All", formats: "Formats", contents: "Contents", hyperlinks: "Hyperlinks", removeHyperlinks: "RemoveHyperlinks" },
  SheetVisibility: { visible: "Visible", hidden: "Hidden", veryHidden: "VeryHidden" },
  DeleteShiftDirection: { up: "Up", left: "Left" },
  InsertShiftDirection: { down: "Down", right: "Right" },
  HorizontalAlignment: {}, VerticalAlignment: {}, BorderIndex: {}, BorderLineStyle: {}, SortOn: {}, SortOrientation: {}
};

async function abrirPasta(buffer, nome) {
  const p = await lerXlsx(buffer);
  p.nome = nome;
  return p.wb;
}

module.exports = { abrirPasta, ExcelScript };
