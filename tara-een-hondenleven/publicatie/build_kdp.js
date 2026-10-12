// Bouwt het KDP-binnenwerk (6 x 9 inch) van "Zij koos mij" uit de hoofdstukbestanden.
// Gebruik: node build_kdp.js <map-met-hoofdstukken> <uitvoer.docx>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
  TableOfContents, Footer, PageNumber, NumberFormat, BorderStyle,
} = require("docx");

const [, , srcDir, outFile] = process.argv;

const TITLE = "Zij koos mij";
const SUBTITLE = "Het geheime onderzoeksrapport van een Labrador";
const AUTHOR = "Patrik Verstraelen";
const YEAR = "2026";
const FONT = "Georgia";
const BODY_SIZE = 22; // 11 pt

const KDP_PAGE = {
  size: { width: 8640, height: 12960 },
  margin: { top: 1080, bottom: 1080, left: 720, right: 720, gutter: 720 },
};

// Zet **vet** en *cursief* om in TextRuns.
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0;
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: BODY_SIZE, ...base }));
    const tok = m[0];
    if (tok.startsWith("**")) out.push(new TextRun({ text: tok.slice(2, -2), bold: true, font: FONT, size: BODY_SIZE, ...base }));
    else out.push(new TextRun({ text: tok.slice(1, -1), italics: !base.italics, font: FONT, size: BODY_SIZE, ...base, ...(base.italics ? { italics: false } : {}) }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: BODY_SIZE, ...base }));
  return out;
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 2000, after: 600 },
    children: [new TextRun({ text, bold: true, size: 30, font: FONT, color: "000000" })],
  });
}

function partTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 4000, after: 600 },
    children: [new TextRun({ text, bold: true, size: 34, font: FONT, color: "000000" })],
  });
}

function subheading(text) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    keepNext: true,
    children: [new TextRun({ text, bold: true, smallCaps: true, size: BODY_SIZE, font: FONT })],
  });
}

function bodyPara(text, first) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 0, line: 300 },
    indent: first ? undefined : { firstLine: 360 },
    children: runs(text),
  });
}

function boxPara(text, isFirst, isLast) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "808080", space: 6 };
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 540, right: 540 },
    spacing: { before: isFirst ? 360 : 0, after: isLast ? 360 : 60, line: 276 },
    border: { left: border },
    keepLines: true,
    keepNext: !isLast,
    children: runs(text, { size: 20 }),
  });
}

function bulletPara(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 540, hanging: 270 },
    spacing: { after: 80, line: 276 },
    children: [new TextRun({ text: "•\t", font: FONT, size: BODY_SIZE }), ...runs(text)],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Zet één markdownbestand om in alinea's.
function convert(md) {
  const out = [];
  const blocks = md.replace(/\r/g, "").split(/\n\s*\n/);
  let firstAfterHeading = true;
  for (const raw of blocks) {
    const block = raw.trim();
    if (!block) continue;
    const lines = block.split("\n");
    if (lines[0].startsWith("# ")) {
      for (const l of lines) {
        const t = l.replace(/^#\s+/, "").trim();
        if (/^Deel /.test(t)) { out.push(partTitle(t)); out.push(pageBreak()); }
        else out.push(heading1(t));
      }
      firstAfterHeading = true;
      continue;
    }
    if (lines[0].startsWith("### ")) {
      out.push(subheading(lines[0].replace(/^###\s+/, "")));
      firstAfterHeading = true;
      continue;
    }
    if (lines[0].startsWith("##")) {
      out.push(subheading(lines[0].replace(/^#+\s+/, "")));
      firstAfterHeading = true;
      continue;
    }
    if (lines.every((l) => l.startsWith(">"))) {
      const content = lines.map((l) => l.replace(/^>\s?/, "")).filter((l) => l.trim() !== "");
      content.forEach((l, i) => out.push(boxPara(l.trim(), i === 0, i === content.length - 1)));
      firstAfterHeading = true;
      continue;
    }
    if (lines.every((l) => /^\s*[-*] /.test(l))) {
      lines.forEach((l) => out.push(bulletPara(l.replace(/^\s*[-*]\s+/, "").trim())));
      firstAfterHeading = true;
      continue;
    }
    out.push(bodyPara(lines.join(" ").trim(), firstAfterHeading));
    firstAfterHeading = false;
  }
  return out;
}

const files = fs.readdirSync(srcDir).filter((f) => /^\d\d[a-z]?_.*\.md$/.test(f)).sort();
const motto = files.find((f) => f.startsWith("00a_"));
const chapters = files.filter((f) => f !== motto);

// Voorwerk: titelpagina, colofon, motto.
const center = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: opts.after ?? 200, before: opts.before ?? 0 },
  children: [new TextRun({ text, font: FONT, size: opts.size ?? BODY_SIZE, bold: opts.bold, italics: opts.italics })],
});
const small = (text, after = 160) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { after },
  children: [new TextRun({ text, font: FONT, size: 18 })],
});

const front = [
  center(TITLE.toUpperCase(), { size: 44, bold: true, before: 3200, after: 400 }),
  center(SUBTITLE, { size: 24, italics: true, after: 1600 }),
  center(AUTHOR, { size: 28 }),
  pageBreak(),
  small(`${TITLE}. ${SUBTITLE}`, 80),
  small(`Een roman van ${AUTHOR}`, 400),
  small(`© ${YEAR} ${AUTHOR}`, 400),
  small("Alle rechten voorbehouden. Niets uit deze uitgave mag worden verveelvoudigd, opgeslagen in een geautomatiseerd gegevensbestand en/of openbaar gemaakt in enige vorm of op enige wijze, hetzij elektronisch, mechanisch, door fotokopieën, opnamen of op enige andere manier, zonder voorafgaande schriftelijke toestemming van de auteur."),
  small("Dit is een werk van fictie. Namen, personages, plaatsen en gebeurtenissen zijn ofwel het product van de verbeelding van de auteur, ofwel op fictieve wijze gebruikt. Heultje, Westerlo en de Kempen bestaan echt; de huizen, straten en mensen in dit boek zijn verzonnen. Elke gelijkenis met bestaande personen, levend of overleden, berust op toeval."),
  small("Het motto is een eigen vertaling uit Franz Kafka, Forschungen eines Hundes (1922).", 400),
  small(`Eerste druk, ${YEAR}`, 80),
  small("ISBN: [in te vullen]", 80),
  small("Omslagontwerp: [in te vullen]", 80),
  small("Uitgegeven in eigen beheer."),
  pageBreak(),
];
if (motto) {
  const lines = fs.readFileSync(path.join(srcDir, motto), "utf8").split("\n")
    .map((l) => l.replace(/^>\s?/, "").trim()).filter(Boolean);
  front.push(new Paragraph({ spacing: { before: 3600 }, children: [] }));
  lines.forEach((l, i) => front.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    indent: { left: 1440 },
    spacing: { after: i === 0 ? 240 : 80 },
    children: runs(l, { size: i === 0 ? 22 : 18 }),
  })));
}

const body = [];
chapters.forEach((f, i) => {
  if (i > 0) body.push(pageBreak());
  body.push(...convert(fs.readFileSync(path.join(srcDir, f), "utf8")));
});

function footerWithPageNumber() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })],
    })],
  });
}

const doc = new Document({
  features: { updateFields: true },
  title: `${TITLE}. ${SUBTITLE}`,
  creator: AUTHOR,
  styles: { default: { document: { run: { font: FONT, size: BODY_SIZE } } } },
  sections: [
    {
      properties: { page: KDP_PAGE, titlePage: true },
      footers: {
        default: new Footer({ children: [new Paragraph({ children: [] })] }),
        first: new Footer({ children: [new Paragraph({ children: [] })] }),
      },
      children: front,
    },
    {
      properties: { page: { ...KDP_PAGE, pageNumbers: { start: 4, formatType: NumberFormat.LOWER_ROMAN } } },
      footers: { default: footerWithPageNumber() },
      children: [
        new TableOfContents("Inhoud", { hyperlink: true, headingStyleRange: "1-1" }),
        pageBreak(),
      ],
    },
    {
      properties: { page: { ...KDP_PAGE, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } } },
      footers: { default: footerWithPageNumber() },
      children: body,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outFile, buf);
  console.log(`Geschreven: ${outFile} (${chapters.length} bestanden)`);
});
