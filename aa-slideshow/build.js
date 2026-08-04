const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

const NAVY = "1B2A41";
const MUTED = "6B7A8C";
const GREEN = "2C5F2D";
const GREEN_TINT = "EAF1E6";
const BLUE = "065A82";
const BLUE_TINT = "E2EDF4";
const MAGENTA = "9E1B60";
const MAGENTA_SOFT = "F7D9E8";
const HEAD = "Cambria";
const BODY = "Calibri";

const slide = pres.addSlide();
slide.background = { color: "FFFFFF" };

// ---- Title ----
slide.addText("The Big Book", {
  x: 0.6, y: 0.35, w: 7.9, h: 0.62,
  fontFace: HEAD, fontSize: 40, bold: true, color: NAVY, margin: 0, valign: "middle",
});
slide.addText("Problem   ·   Solution   ·   Action", {
  x: 0.6, y: 1.0, w: 7.9, h: 0.38,
  fontFace: BODY, fontSize: 17, color: MUTED, margin: 0, valign: "middle", charSpacing: 2,
});

// ---- "MUSTS" callout ----
slide.addShape(pres.ShapeType.roundRect, {
  x: 8.9, y: 0.35, w: 3.83, h: 1.05, fill: { color: MAGENTA }, rectRadius: 0.1,
});
slide.addText(
  [
    { text: "The “MUSTS”", options: { fontFace: HEAD, fontSize: 19, bold: true, color: "FFFFFF", breakLine: true } },
    { text: "Start on page 142", options: { fontFace: BODY, fontSize: 14, color: MAGENTA_SOFT } },
  ],
  { x: 8.9, y: 0.35, w: 3.83, h: 1.05, align: "center", valign: "middle", margin: 0 }
);

// ---- Group cards ----
const groups = [
  {
    y: 1.55, h: 1.2, dark: GREEN, tint: GREEN_TINT,
    step: "STEP 1", theme: "Problem  ·  Powerless",
    rows: [
      { n: "", title: "The Doctor’s Opinion", pg: "xxi" },
      { n: "1", title: "Bill’s Story", pg: "1" },
    ],
  },
  {
    y: 2.95, h: 1.5, dark: BLUE, tint: BLUE_TINT,
    step: "STEP 2", theme: "Solution  ·  Power",
    rows: [
      { n: "2", title: "There Is a Solution", pg: "17" },
      { n: "3", title: "More About Alcoholism", pg: "30" },
      { n: "4", title: "We Agnostics", pg: "44" },
    ],
  },
  {
    y: 4.65, h: 1.5, dark: GREEN, tint: GREEN_TINT,
    step: "STEPS 3–12", theme: "Action Necessary for Recovery",
    rows: [
      { n: "5", title: "How It Works", pg: "58" },
      { n: "6", title: "Into Action", pg: "72" },
      { n: "7", title: "Working With Others", pg: "89" },
    ],
  },
];

groups.forEach((g) => {
  slide.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: g.y, w: 12.13, h: g.h, fill: { color: g.tint }, rectRadius: 0.08,
  });
  slide.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: g.y, w: 3.3, h: g.h, fill: { color: g.dark }, rectRadius: 0.08,
  });
  slide.addText(
    [
      { text: g.step, options: { fontFace: HEAD, fontSize: 21, bold: true, color: "FFFFFF", breakLine: true } },
      { text: g.theme, options: { fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" } },
    ],
    { x: 0.85, y: g.y, w: 2.8, h: g.h, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
  );

  const rowH = 0.38;
  const startY = g.y + (g.h - (g.rows.length * rowH + (g.rows.length - 1) * 0.04)) / 2;
  g.rows.forEach((r, i) => {
    const ry = startY + i * (rowH + 0.04);
    slide.addText(r.n, {
      x: 4.15, y: ry, w: 0.5, h: rowH,
      fontFace: BODY, fontSize: 15, bold: true, color: g.dark, align: "right", valign: "middle", margin: 0,
    });
    slide.addText(r.title, {
      x: 4.78, y: ry, w: 6.2, h: rowH,
      fontFace: BODY, fontSize: 17, color: NAVY, valign: "middle", margin: 0,
    });
    slide.addText("p. " + r.pg, {
      x: 11.15, y: ry, w: 1.4, h: rowH,
      fontFace: BODY, fontSize: 14, color: MUTED, align: "right", valign: "middle", margin: 0,
    });
  });
});

// ---- Chapter 8 / 9 rewrites ----
slide.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 6.4, w: 12.13, h: 0.6, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
slide.addText(
  [
    { text: "8   ", options: { bold: true, color: MUTED } },
    { text: "To Wives", options: { strike: true, color: MUTED } },
    { text: "  →  ", options: { color: MUTED } },
    { text: "To Al-Anons", options: { bold: true, color: NAVY } },
    { text: "   p. 104", options: { color: MUTED } },
  ],
  { x: 0.95, y: 6.4, w: 5.5, h: 0.6, fontFace: BODY, fontSize: 14, valign: "middle", margin: 0 }
);
slide.addText(
  [
    { text: "9   ", options: { bold: true, color: MUTED } },
    { text: "The Family Afterward", options: { strike: true, color: MUTED } },
    { text: "  →  ", options: { color: MUTED } },
    { text: "The Alateens Afterward", options: { bold: true, color: NAVY } },
    { text: "   p. 122", options: { color: MUTED } },
  ],
  { x: 6.9, y: 6.4, w: 5.5, h: 0.6, fontFace: BODY, fontSize: 14, valign: "middle", margin: 0 }
);

slide.addNotes(
  "Roadmap slide. The Big Book is built in three movements.\n\n" +
  "Doctor's Opinion + Bill's Story = Step 1. The Problem. We are powerless.\n" +
  "Chapters 2-4 = Step 2. The Solution. There is a Power.\n" +
  "Chapters 5-7 = Steps 3 through 12. The action necessary for recovery.\n\n" +
  "Chapters 8 and 9 speak to the family - read today as Al-Anon and Alateen.\n\n" +
  "The \"MUSTS\" exercise begins on page 142."
);

pres.writeFile({ fileName: "/tmp/claude-0/-home-user/d255f53f-ba4d-5eef-a926-701b8fb15d3a/scratchpad/aa-deck/aa-big-book-study.pptx" })
  .then((f) => console.log("wrote " + f));
