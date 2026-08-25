const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType,
  BorderStyle, Table, TableRow, TableCell, WidthType, VerticalAlign,
  PageBorderDisplay, PageBorderOffsetFrom, PageBorderZOrder,
} = require("docx");

const MAROON = "6B1D2A";
const GOLD = "A98A44";
const CHARCOAL = "2B2B2B";
const SERIF = "Georgia";

const img = fs.readFileSync(__dirname + "/model-a.jpg");

// image is 728x546 (4:3). Frame it at 5.6" wide.
const IMG_W = 5.6 * 96;
const IMG_H = IMG_W * 546 / 728;

const spacer = (pts) => new Paragraph({ spacing: { after: pts * 20 }, children: [] });

const rule = (widthIndent, color, size) => new Paragraph({
  alignment: AlignmentType.CENTER,
  indent: { left: widthIndent, right: widthIndent },
  spacing: { before: 0, after: 0 },
  border: { bottom: { style: BorderStyle.SINGLE, size: size, color: color, space: 1 } },
  children: [],
});

const blankLine = (label, labelSize, lineWidthTwips) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 340, after: 0 },
  children: [
    new TextRun({ text: label + "  ", font: SERIF, size: labelSize, color: CHARCOAL, bold: true, allCaps: true, characterSpacing: 20 }),
    new TextRun({ text: "_".repeat(lineWidthTwips), font: SERIF, size: labelSize, color: CHARCOAL }),
  ],
});

const doc = new Document({
  styles: {
    default: { document: { run: { font: SERIF, color: CHARCOAL } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, bottom: 1080, left: 1152, right: 1152 },
        borders: {
          pageBorders: {
            display: PageBorderDisplay.ALL_PAGES,
            offsetFrom: PageBorderOffsetFrom.PAGE,
            zOrder: PageBorderZOrder.FRONT,
          },
          pageBorderTop:    { style: BorderStyle.DOUBLE, size: 32, color: MAROON, space: 24 },
          pageBorderBottom: { style: BorderStyle.DOUBLE, size: 32, color: MAROON, space: 24 },
          pageBorderLeft:   { style: BorderStyle.DOUBLE, size: 32, color: MAROON, space: 24 },
          pageBorderRight:  { style: BorderStyle.DOUBLE, size: 32, color: MAROON, space: 24 },
        },
      },
    },
    children: [
      // Presenter line
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 80 },
        children: [new TextRun({
          text: "LIGHTHOUSE FELLOWSHIP PRESENTS", font: SERIF, size: 22,
          color: GOLD, characterSpacing: 60, bold: true,
        })],
      }),
      // Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 100 },
        children: [new TextRun({
          text: "Certificate of Entry", font: SERIF, size: 76, color: MAROON,
          smallCaps: true,
        })],
      }),
      rule(3600, GOLD, 12),
      // Event name
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 220, after: 60 },
        children: [new TextRun({
          text: "1st Annual Dry Dock Freedom", font: SERIF, size: 44,
          color: CHARCOAL, bold: true, smallCaps: true, characterSpacing: 20,
        })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({
          text: "CAR SHOW", font: SERIF, size: 30, color: GOLD,
          bold: true, characterSpacing: 90,
        })],
      }),
      // Framed photo
      new Table({
        alignment: AlignmentType.CENTER,
        columnWidths: [Math.round(IMG_W * 15) + 240],
        borders: undefined,
        rows: [new TableRow({
          children: [new TableCell({
            width: { size: Math.round(IMG_W * 15) + 240, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 100, bottom: 60, left: 100, right: 100 },
            borders: {
              top:    { style: BorderStyle.SINGLE, size: 12, color: MAROON },
              bottom: { style: BorderStyle.SINGLE, size: 12, color: MAROON },
              left:   { style: BorderStyle.SINGLE, size: 12, color: MAROON },
              right:  { style: BorderStyle.SINGLE, size: 12, color: MAROON },
            },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              children: [new ImageRun({
                type: "jpg", data: img,
                transformation: { width: IMG_W, height: IMG_H },
              })],
            })],
          })],
        })],
      }),
      // Date
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 80 },
        children: [new TextRun({
          text: "Saturday  •  August 29th, 2026", font: SERIF, size: 30,
          color: MAROON, bold: true, smallCaps: true, characterSpacing: 30,
        })],
      }),
      rule(3600, GOLD, 12),
      // Entry number — the big one
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 420, after: 0 },
        children: [
          new TextRun({ text: "ENTRY NO.  ", font: SERIF, size: 40, color: CHARCOAL, bold: true, characterSpacing: 40 }),
          new TextRun({ text: "_".repeat(14), font: SERIF, size: 40, color: CHARCOAL }),
        ],
      }),
      // Owner / vehicle lines
      blankLine("Owner", 24, 52),
      blankLine("Vehicle (Year / Make / Model)", 24, 34),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(__dirname + "/Certificate_of_Entry_Car_Show.docx", buf);
  console.log("written");
});
