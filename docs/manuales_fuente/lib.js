// Librería común para maquetar el manual de la Consola Registro Info SAP (docx-js), con el mismo
// estilo que los de Control de Pedidos y el Organizador Princess (de donde se copió).
// Idea y desarrollo de la aplicación: Víctor Alejandro Martín Andrade.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, HeadingLevel, LevelFormat, PageBreak, Footer, Header, PageNumber,
  TableOfContents, VerticalAlign, TabStopType, TabStopPosition,
} = require("docx");

const CAP = process.env.CAPTURAS || path.join(__dirname, "cap");
const NAVY = "0F1F3D", GOLD = "B8913A", NARANJA = "E8590C", GRIS = "5B6475", CLARO = "F4F6FA";
const FONT = "Arial";
const ANCHO = 9638; // A4 con márgenes de 2 cm (11906 − 2×1134)

const sinBorde = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: sinBorde, bottom: sinBorde, left: sinBorde, right: sinBorde };

// Índice para «Consultar el manual» de la app: cada h1/h2 abre un apartado y el texto de los
// párrafos, viñetas, pasos, leyendas, recuadros y tablas se le añade. generar() lo guarda junto al
// .docx (manual_indice.json) y paginar_indice.py le pone la página de cada apartado.
const INDICE = [];
let _capitulo = "";
function _texto(t) { return String(t).replace(/\*\*/g, ""); }
function _anota(t) { if (INDICE.length) INDICE[INDICE.length - 1].texto += (INDICE[INDICE.length - 1].texto ? "\n" : "") + _texto(t); }

function run(t, o = {}) { return new TextRun({ text: t, font: FONT, ...o }); }

// Texto con **negritas** sencillas
function rico(texto, o = {}) {
  const partes = String(texto).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return partes.map(p => p.startsWith("**") ? run(p.slice(2, -2), { ...o, bold: true }) : run(p, o));
}

function p(texto, o = {}) {
  _anota(texto);
  return new Paragraph({ children: rico(texto, { size: 21, color: o.color }), spacing: { after: 120, line: 300 },
    alignment: o.align, keepNext: o.keepNext });
}

function vineta(texto, nivel = 0) {
  _anota(texto);
  return new Paragraph({ numbering: { reference: "vinetas", level: nivel }, children: rico(texto, { size: 21 }),
    spacing: { after: 60, line: 290 } });
}

function pasos(lista) {
  pasos._n += 1; const ref = "pasos" + pasos._n;
  lista.forEach((t, i) => _anota(`${i + 1}. ${t}`));
  return lista.map((t, i) => new Paragraph({ numbering: { reference: ref, level: 0 },
    children: rico(t, { size: 21 }), spacing: { after: 80, line: 290 } }));
}
pasos._n = 0;
function nuevaLista() { pasos._n += 1; return pasos._n; }

// Párrafo vacío tras una tabla o un recuadro. Si cae justo al final de una página, pasaría solo a la
// siguiente y, con el salto del capítulo, dejaría una página en blanco: generar() lo quita delante de un h1.
function _separador(after) { const x = new Paragraph({ spacing: { after }, children: [] }); x._separador = true; return x; }
function h1(t) { _capitulo = t; INDICE.push({ capitulo: t, titulo: t, texto: "" }); const x = new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [run(t)] }); x._h1 = true; return x; }
function h2(t) { INDICE.push({ capitulo: _capitulo, titulo: t, texto: "" }); return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(t)] }); }

function imagen(nombre, anchoCm = 17, pie = "") {
  const f = path.join(CAP, nombre + ".png");
  const buf = fs.readFileSync(f);
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const anchoPx = Math.round(anchoCm / 2.54 * 96);
  let altoPx = Math.round(anchoPx * h / w);
  const maxAlto = Math.round(21 / 2.54 * 96);
  let ancho = anchoPx;
  if (altoPx > maxAlto) { ancho = Math.round(anchoPx * maxAlto / altoPx); altoPx = maxAlto; }
  const out = [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: pie ? 40 : 160 }, keepNext: !!pie,
    children: [new ImageRun({ type: "png", data: buf, transformation: { width: ancho, height: altoPx },
      altText: { title: pie || nombre, description: pie || nombre, name: nombre } })] })];
  if (pie) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 },
    children: [run(pie, { size: 17, italics: true, color: GRIS })] }));
  return out;
}

// Leyenda numerada: [[n, "texto"], ...]
function leyenda(items) {
  items.forEach(([n, t]) => _anota(`(${n}) ${t}`));
  const filas = items.map(([n, t]) => new TableRow({ cantSplit: true, children: [
    new TableCell({ width: { size: 560, type: WidthType.DXA }, borders: sinBordes, verticalAlign: VerticalAlign.TOP,
      margins: { top: 50, bottom: 50, left: 40, right: 40 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: "png",
        data: fs.readFileSync(path.join(__dirname, "badges", n + ".png")), transformation: { width: 19, height: 19 },
        altText: { title: "Marca " + n, description: "Número " + n, name: "marca" + n } })] })] }),
    new TableCell({ width: { size: ANCHO - 560, type: WidthType.DXA }, borders: sinBordes,
      margins: { top: 50, bottom: 50, left: 80, right: 80 },
      children: [new Paragraph({ children: rico(t, { size: 20 }), spacing: { line: 280 } })] }),
  ] }));
  return [new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [560, ANCHO - 560], rows: filas }),
    _separador(120)];
}

// Recuadro de aviso: tipo = consejo | importante | nota
function recuadro(tipo, titulo, lineas) {
  _anota(titulo + ". " + [].concat(lineas).join(" "));
  const cfg = { consejo: ["💡", "E8F5EE", "2E8B57"], importante: ["⚠️", "FFF4E5", "D9822B"], nota: ["ℹ️", "EAF1FB", "2F6FB5"] }[tipo];
  const hijos = [new Paragraph({ spacing: { after: 60 }, children: [run(`${cfg[0]}  ${titulo}`, { bold: true, size: 21, color: cfg[2] })] })];
  for (const l of [].concat(lineas)) hijos.push(new Paragraph({ spacing: { after: 50, line: 280 }, children: rico(l, { size: 20 }) }));
  const borde = { style: BorderStyle.SINGLE, size: 24, color: cfg[2] };
  return [new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [ANCHO], rows: [new TableRow({ cantSplit: true, children: [
    new TableCell({ width: { size: ANCHO, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: cfg[1], color: "auto" },
      borders: { top: sinBorde, bottom: sinBorde, right: sinBorde, left: borde },
      margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: hijos })] })] }),
    _separador(160)];
}

// Tabla sencilla con cabecera
function tabla(cabecera, filas, anchos) {
  filas.forEach(f => _anota(f.join(" — ") + "."));
  const total = anchos.reduce((a, b) => a + b, 0);
  const celda = (t, i, cab, fill) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: cab ? NAVY : fill, color: "auto" },
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "D5DAE3" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5DAE3" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "D5DAE3" }, right: { style: BorderStyle.SINGLE, size: 4, color: "D5DAE3" } },
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    children: [new Paragraph({ spacing: { line: 270 }, children: rico(t, { size: cab ? 18 : 19, bold: cab, color: cab ? "FFFFFF" : undefined }) })] });
  const rows = [new TableRow({ tableHeader: true, children: cabecera.map((t, i) => celda(t, i, true)) })]
    .concat(filas.map((f, k) => new TableRow({ cantSplit: true, children: f.map((t, i) => celda(t, i, false, k % 2 ? "FFFFFF" : CLARO)) })));
  return [new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: anchos, rows }),
    _separador(160)];
}

// Flujo en cajas: [["Texto", "quién"], ...]
function flujo(cajas) {
  const n = cajas.length, w = Math.floor(ANCHO / n);
  const anchos = Array(n).fill(w); anchos[n - 1] = ANCHO - w * (n - 1);
  const row = new TableRow({ cantSplit: true, children: cajas.map(([t, quien, color], i) => new TableCell({
    width: { size: anchos[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: color || "E3EEFF", color: "auto" },
    borders: { top: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, bottom: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" },
      left: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, right: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" } },
    margins: { top: 120, bottom: 120, left: 80, right: 80 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(`${i + 1}`, { bold: true, size: 26, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 250 }, children: [run(t, { bold: true, size: 17, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 }, children: [run(quien, { size: 15, color: GRIS, italics: true })] })] })) });
  return [new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: anchos, rows: [row] }), _separador(160)];
}

function portada({ rol, subtitulo, version, fecha, publico }) {
  const logo = fs.readFileSync(path.join(__dirname, "..", "..", "assets", "logo.png"));
  const cel = (hijos, fill, h) => new TableCell({ width: { size: ANCHO, type: WidthType.DXA }, borders: sinBordes,
    shading: { type: ShadingType.CLEAR, fill, color: "auto" }, margins: { top: 300, bottom: 300, left: 400, right: 400 }, children: hijos });
  const bloque = new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [ANCHO], rows: [
    new TableRow({ children: [cel([
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 300 }, children: [new ImageRun({ type: "png", data: logo,
        transformation: { width: 170, height: Math.round(170 * logo.readUInt32BE(20) / logo.readUInt32BE(16)) }, altText: { title: "Princess Hotels & Resorts", description: "Logotipo", name: "logo" } })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [run("CONSOLA REGISTRO INFO SAP", { bold: true, size: 40, color: "FFFFFF", characterSpacing: 50 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, children: [run("Manual de usuario", { size: 30, color: "D9C48F" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 500, after: 120 }, children: [run(rol, { bold: true, size: 56, color: "FFFFFF" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 500 }, children: [run(subtitulo, { size: 24, color: "C9D2E3" })] }),
    ], NAVY)] }),
    new TableRow({ children: [cel([
      new Paragraph({ alignment: AlignmentType.CENTER, children: [run(`Versión de la aplicación ${version}  ·  ${fecha}`, { size: 20, bold: true, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [run(publico, { size: 19, color: GRIS })] }),
    ], "F3ECDD")] }),
  ] });
  return [bloque,
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 500 }, children: [run("CENTRAL DE COMPRAS · PRINCESS HOTELS & RESORTS CANARIAS", { size: 18, bold: true, color: NAVY, characterSpacing: 30 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [run("Idea y desarrollo de la aplicación: Víctor Alejandro Martín Andrade, para el Departamento de Central de Compras de Princess Hotels en Canarias.", { size: 16, italics: true, color: GRIS })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [run("Las capturas de este manual usan datos de ejemplo inventados.", { size: 16, italics: true, color: GRIS })] }),
  ];
}

function indice(capitulos) {
  const out = [new Paragraph({ pageBreakBefore: true, spacing: { after: 240 }, children: [run("Contenido", { bold: true, size: 36, color: NAVY })] })];
  for (const [t, sub] of capitulos) {
    out.push(new Paragraph({ spacing: { before: 140, after: 20 }, border: { bottom: { style: BorderStyle.DOTTED, size: 4, color: "C9CED8", space: 4 } },
      children: [run(t, { bold: true, size: 23, color: NAVY })] }));
    if (sub) out.push(new Paragraph({ spacing: { after: 40 }, indent: { left: 360 }, children: [run(sub, { size: 18, color: GRIS })] }));
  }
  return out;
}

async function generar(salida, titulo, rolCorto, secciones) {
  secciones = secciones.filter((x, i) => !(x._separador && (!secciones[i + 1] || secciones[i + 1]._h1)));
  const listas = [{ reference: "vinetas", levels: [
    { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } },
    { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 900, hanging: 260 } } } }] }];
  for (let i = 1; i <= Math.max(pasos._n, 1); i++) listas.push({ reference: "pasos" + i, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
    alignment: AlignmentType.LEFT, style: { run: { bold: true, color: GOLD }, paragraph: { indent: { left: 500, hanging: 320 } } } }] });
  const doc = new Document({
    creator: "Víctor Alejandro Martín Andrade", title: titulo, description: "Manual de la Consola Registro Info SAP",
    features: { updateFields: true },
    styles: {
      default: { document: { run: { font: FONT, size: 21 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 36, bold: true, color: NAVY },
          paragraph: { spacing: { before: 0, after: 220 }, keepNext: true, outlineLevel: 0,
            border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: GOLD, space: 6 } } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 26, bold: true, color: GOLD },
          paragraph: { spacing: { before: 300, after: 120 }, keepNext: true, outlineLevel: 1 } },
      ],
    },
    numbering: { config: listas },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } }, titlePage: true },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: [run(`Consola Registro Info SAP · Manual ${rolCorto}`, { size: 16, color: GRIS })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }],
        children: [run("Central de Compras · Princess Hotels & Resorts Canarias", { size: 16, color: GRIS }),
          run("\tPágina ", { size: 16, color: GRIS }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GRIS })] })] }) },
      children: secciones.flat(),
    }],
  });
  fs.writeFileSync(salida, await Packer.toBuffer(doc));
  console.log("Generado", salida);
  const rutaIndice = path.join(path.dirname(salida), "manual_indice.json");
  fs.writeFileSync(rutaIndice, JSON.stringify({ titulo, secciones: INDICE.filter(s => s.texto.trim()) }, null, 1));
  console.log("Índice para la consulta:", rutaIndice, INDICE.length, "apartados");
}

module.exports = { p, vineta, pasos, nuevaLista, h1, h2, imagen, leyenda, recuadro, tabla, flujo, portada, indice, generar, run, rico };
