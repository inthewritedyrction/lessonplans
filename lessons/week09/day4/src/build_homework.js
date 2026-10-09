// Builds Week9_Day4_Guided_Reading_Homework.docx (editable Word). Run: node build_homework.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, HeightRule, Footer, PageNumber, TableLayoutType,
} = require("docx");

const FONT = "Century Gothic";
const INK = "1E2A32", MUTED = "4F5F69", LINE = "9AA7AF";
const TEAL = "00A39B", TEAL_D = "006E68", TEAL_L = "D4F3EE";
const YEL_L = "FFF1C9", YEL_D = "9A6C00";
const PUR = "8B6BE8", PUR_D = "5F3FC0", PUR_L = "ECE5FF";
const COR_D = "CF3A20";
const PAGE_W = 12240, MARGIN = 900, CONTENT_W = PAGE_W - 2 * MARGIN; // 0.625in margins

const run = (text, o = {}) => new TextRun({ text, font: FONT, size: (o.size || 12) * 2, bold: o.bold, italics: o.italic, color: o.color || INK, characterSpacing: o.spacing });
const para = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before || 0, after: o.after ?? 80, line: o.line || 276 }, alignment: o.align, keepNext: o.keepNext });

const border = (c = LINE, sz = 12) => ({ style: BorderStyle.SINGLE, size: sz, color: c });
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });

// A full-width one-cell table: used for colored callouts and for open answer boxes.
function box(children, { fill, line, height, sz = 12 } = {}) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    layout: TableLayoutType.FIXED,
    rows: [new TableRow({
      height: height ? { value: Math.round(height * 1440), rule: HeightRule.ATLEAST } : undefined,
      cantSplit: true,
      children: [new TableCell({
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined,
        borders: allBorders(line ? border(line, sz) : noBorder),
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: children.length ? children : [new Paragraph({ children: [] })],
      })],
    })],
  });
}
const answer = (h) => box([], { line: LINE, height: h });
const gap = (after = 120) => new Paragraph({ children: [], spacing: { after, before: 0 } });

function question(n, text, hint) {
  const runs = [run(`${n}.  `, { bold: true, color: PUR_D, size: 13 }), run(text, { size: 12.5, bold: true })];
  const out = [para(runs, { before: 120, after: hint ? 20 : 80, keepNext: true })];
  if (hint) out.push(para(run(hint, { size: 11, italic: true, color: MUTED }), { after: 80, keepNext: true }));
  return out;
}

const children = [
  para(run("WEEK 9 · DAY 4 · GUIDED READING HOMEWORK · ENCOUNTERING EVIL · GRADE 8 ELA · MS. BRIGGS", { size: 8, color: MUTED, spacing: 20 }), { after: 160 }),
  para([run("Night", { size: 22, bold: true, italic: true }), run(" — Opening Reading Check  |  pp. 3–22", { size: 22, bold: true })], { after: 120 }),
  // name / date / class
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [5640, 2400, 2400],
    layout: TableLayoutType.FIXED,
    rows: [new TableRow({
      height: { value: 600, rule: HeightRule.ATLEAST },
      children: [["NAME", 5640], ["DATE", 2400], ["CLASS", 2400]].map(([t, w]) => new TableCell({
        width: { size: w, type: WidthType.DXA },
        borders: allBorders(border()),
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [para(run(t, { size: 8, color: MUTED, spacing: 20 }))],
      })),
    })],
  }),
  gap(100),
  box([
    para([run("READ: ", { bold: true, color: TEAL_D, size: 13 }), run("printed book pp. 3–22  /  PDF pages 18–36", { bold: true, size: 13 })], { after: 60 }),
    para(run("Bring your printed reading AND this completed sheet to class tomorrow.", { size: 12 }), { after: 0 }),
  ], { fill: TEAL_L }),
  gap(80),
  box([
    para(run("This is your FIRST READ.", { bold: true, size: 12, color: YEL_D }), { after: 40 }),
    para(run("Your job tonight is to understand what happens. Tomorrow, we will reread important sections together and analyze why Wiesel wrote them this way.", { size: 12 }), { after: 0 }),
  ], { fill: YEL_L }),

  ...question(1, "Who is Moishe the Beadle, and why is he important to young Elie at the beginning of the memoir?", "Give one specific detail."),
  answer(1.0),
  ...question(2, "What happens to Moishe after the foreign Jews are expelled from Sighet? What does he report when he returns?", "Give a brief summary and the page number."),
  answer(1.15),
  ...question(3, "How do the people of Sighet respond to Moishe’s warning?", "Name at least one specific reaction and give the page number."),
  answer(1.0),
  ...question(4, "Why do you think the people are able to dismiss what Moishe tells them?", "Explain in 1–2 sentences. This is your interpretation. There is no single correct answer."),
  answer(1.0),
  ...question(5, "Put these events in order. Write 1, 2, 3, 4 in the boxes.", "1 = first, 4 = last."),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [900, CONTENT_W - 900],
    layout: TableLayoutType.FIXED,
    rows: [
      "German soldiers enter Sighet",
      "Jews are forced into ghettos",
      "Moishe returns to Sighet with his warning",
      "Jewish community leaders are arrested / restrictions escalate",
    ].map((t) => new TableRow({
      height: { value: 470, rule: HeightRule.ATLEAST },
      cantSplit: true,
      children: [
        new TableCell({ width: { size: 900, type: WidthType.DXA }, borders: allBorders(border(LINE, 12)), children: [new Paragraph({ children: [] })] }),
        new TableCell({ width: { size: CONTENT_W - 900, type: WidthType.DXA }, borders: { top: noBorder, bottom: noBorder, right: noBorder, left: border(LINE, 12) }, margins: { left: 200, top: 80 }, children: [para(run(t, { size: 12 }), { after: 0 })] }),
      ],
    })),
  }),
  ...question(6, "Find one moment when the people convince themselves that everything will still be okay, even though the danger is increasing.", "Briefly describe the moment and include the page number."),
  answer(1.25),
  ...question(7, "Wiesel repeatedly uses the word “night.” While reading pp. 11–22, underline or mark at least TWO places where the word appears in your printed text."),
  para([run("Write ONE of those page numbers here:  ", { size: 12.5 }), run("p. ________", { size: 12.5, bold: true })], { after: 60 }),
  para(run("You do not need to explain it yet. Just notice it. We will figure out what it means together tomorrow.", { size: 11, italic: true, color: MUTED }), { after: 60 }),
  ...question(8, "Choose ONE and finish the sentence."),
  para([run("☐  A.  ", { bold: true, color: PUR_D }), run("One moment that confused me was ______ because ______.")], { after: 40 }),
  para([run("☐  B.  ", { bold: true, color: PUR_D }), run("One moment I think will matter later is ______ because ______.")], { after: 80 }),
  answer(1.3),
  gap(100),
  box([
    para(run("HOMEWORK CHECK", { bold: true, color: COR_D, size: 11, spacing: 30 }), { after: 60 }),
    ...["I read all assigned pages.", "I marked at least two uses of “night.”", "I answered every question.", "I am bringing my printed reading back tomorrow."]
      .map((t) => para([run("☐  ", { size: 13, bold: true }), run(t, { size: 12 })], { after: 40 })),
  ], { line: "FF6B4E", sz: 16 }),
];

const doc = new Document({
  creator: "Ms. Briggs",
  title: "Night — Opening Reading Check (pp. 3–22)",
  styles: { default: { document: { run: { font: FONT, size: 24, color: INK } } } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 15840 }, margin: { top: 720, bottom: 720, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "NIGHT · OPENING READING CHECK · PAGE ", font: FONT, size: 14, color: MUTED }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 14, color: MUTED }),
    ] })] }) },
    children,
  }],
});

const out = path.join(__dirname, "..", "Week9_Day4_Guided_Reading_Homework.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("homework ok:", out); });
