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
const YELLOW_TINT = "FFF7D9";
const ON_NAVY = "A9B8CC";
const HEAD = "Cambria";
const BODY = "Calibri";

// Required by A.A.W.S. whenever this literature is screen-shared.
const DISCLAIMER =
  "“These materials are copyright © by Alcoholics Anonymous World Services, Inc. (\"A.A.W.S.\"). " +
  "All rights reserved. Individual printing or photocopying of a single copy is permitted. " +
  "Intended for personal use only, and not to be reproduced further or distributed for resale.”";

const FOOTER =
  "Copyright © Alcoholics Anonymous World Services, Inc. All rights reserved. " +
  "Intended for personal use only, and not to be reproduced further or distributed for resale.";

function addFooter(s) {
  s.addText(FOOTER, {
    x: 0.6, y: 6.95, w: 12.13, h: 0.32,
    fontFace: BODY, fontSize: 9.5, color: MUTED, valign: "middle", margin: 0,
  });
}

// =====================================================================
// SLIDE 1 — title + required copyright notice
// =====================================================================
const s1 = pres.addSlide();
s1.background = { color: NAVY };

s1.addText("The Big Book Study", {
  x: 0.9, y: 1.0, w: 11.53, h: 0.85,
  fontFace: HEAD, fontSize: 44, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});
s1.addText("Alcoholics Anonymous   ·   Fourth Edition", {
  x: 0.9, y: 1.9, w: 11.53, h: 0.4,
  fontFace: BODY, fontSize: 18, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});

s1.addShape(pres.ShapeType.roundRect, {
  x: 0.9, y: 2.9, w: 11.53, h: 2.2, fill: { color: "FFFFFF" }, rectRadius: 0.1,
});
s1.addText("COPYRIGHT NOTICE", {
  x: 1.22, y: 3.05, w: 10.9, h: 0.32,
  fontFace: BODY, fontSize: 12, bold: true, color: MAGENTA, margin: 0, valign: "middle", charSpacing: 2,
});
s1.addText(DISCLAIMER, {
  x: 1.22, y: 3.42, w: 10.9, h: 1.5,
  fontFace: BODY, fontSize: 17, italic: true, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
});

s1.addText(
  "Displayed at the request of A.A. World Services, Inc. A.A.W.S. has advised that it has no objection to " +
  "screen-sharing A.A. literature from aa.org or from authorized e-books during an A.A. meeting, provided " +
  "this notice is also displayed.",
  {
    x: 0.9, y: 5.35, w: 11.53, h: 0.8,
    fontFace: BODY, fontSize: 13, color: ON_NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  }
);

s1.addNotes(
  "Leave this slide up while people are settling in.\n\n" +
  "A.A.W.S. granted permission by email (Drew Deetz, Intellectual Property Administrator,\n" +
  "General Service Office) on the condition that this copyright notice is displayed when\n" +
  "A.A. literature is screen-shared during a meeting."
);

// =====================================================================
// SLIDE 2 — contents roadmap (Problem / Solution / Action)
// =====================================================================
const s2 = pres.addSlide();
s2.background = { color: "FFFFFF" };

s2.addText("The Big Book", {
  x: 0.6, y: 0.35, w: 7.9, h: 0.62,
  fontFace: HEAD, fontSize: 40, bold: true, color: NAVY, margin: 0, valign: "middle",
});
s2.addText("Problem   ·   Solution   ·   Action", {
  x: 0.6, y: 1.0, w: 7.9, h: 0.38,
  fontFace: BODY, fontSize: 17, color: MUTED, margin: 0, valign: "middle", charSpacing: 2,
});

s2.addShape(pres.ShapeType.roundRect, {
  x: 8.9, y: 0.35, w: 3.83, h: 1.05, fill: { color: MAGENTA }, rectRadius: 0.1,
});
s2.addText(
  [
    { text: "The “MUSTS”", options: { fontFace: HEAD, fontSize: 19, bold: true, color: "FFFFFF", breakLine: true } },
    { text: "Start on page 142", options: { fontFace: BODY, fontSize: 14, color: MAGENTA_SOFT } },
  ],
  { x: 8.9, y: 0.35, w: 3.83, h: 1.05, align: "center", valign: "middle", margin: 0 }
);

const groups = [
  {
    y: 1.5, h: 1.15, dark: GREEN, tint: GREEN_TINT,
    step: "STEP 1", theme: "Problem  ·  Powerless",
    rows: [
      { n: "", title: "The Doctor’s Opinion", pg: "xxi" },
      { n: "1", title: "Bill’s Story", pg: "1" },
    ],
  },
  {
    y: 2.8, h: 1.45, dark: BLUE, tint: BLUE_TINT,
    step: "STEP 2", theme: "Solution  ·  Power",
    rows: [
      { n: "2", title: "There Is a Solution", pg: "17" },
      { n: "3", title: "More About Alcoholism", pg: "30" },
      { n: "4", title: "We Agnostics", pg: "44" },
    ],
  },
  {
    y: 4.4, h: 1.45, dark: GREEN, tint: GREEN_TINT,
    step: "STEPS 3–12", theme: "Action Necessary for Recovery",
    rows: [
      { n: "5", title: "How It Works", pg: "58" },
      { n: "6", title: "Into Action", pg: "72" },
      { n: "7", title: "Working With Others", pg: "89" },
    ],
  },
];

groups.forEach((g) => {
  s2.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: g.y, w: 12.13, h: g.h, fill: { color: g.tint }, rectRadius: 0.08,
  });
  s2.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: g.y, w: 3.3, h: g.h, fill: { color: g.dark }, rectRadius: 0.08,
  });
  s2.addText(
    [
      { text: g.step, options: { fontFace: HEAD, fontSize: 21, bold: true, color: "FFFFFF", breakLine: true } },
      { text: g.theme, options: { fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" } },
    ],
    { x: 0.85, y: g.y, w: 2.8, h: g.h, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
  );

  const rowH = 0.36;
  const startY = g.y + (g.h - (g.rows.length * rowH + (g.rows.length - 1) * 0.04)) / 2;
  g.rows.forEach((r, i) => {
    const ry = startY + i * (rowH + 0.04);
    s2.addText(r.n, {
      x: 4.15, y: ry, w: 0.5, h: rowH,
      fontFace: BODY, fontSize: 15, bold: true, color: g.dark, align: "right", valign: "middle", margin: 0,
    });
    s2.addText(r.title, {
      x: 4.78, y: ry, w: 6.2, h: rowH,
      fontFace: BODY, fontSize: 17, color: NAVY, valign: "middle", margin: 0,
    });
    s2.addText("p. " + r.pg, {
      x: 11.15, y: ry, w: 1.4, h: rowH,
      fontFace: BODY, fontSize: 14, color: MUTED, align: "right", valign: "middle", margin: 0,
    });
  });
});

s2.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 6.05, w: 12.13, h: 0.55, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
s2.addText(
  [
    { text: "8   ", options: { bold: true, color: MUTED } },
    { text: "To Wives", options: { strike: true, color: MUTED } },
    { text: "   →   ", options: { color: MUTED } },
    { text: "To Al-Anons", options: { bold: true, color: NAVY } },
    { text: "     p. 104", options: { color: MUTED } },
  ],
  { x: 0.95, y: 6.05, w: 5.5, h: 0.55, fontFace: BODY, fontSize: 14, valign: "middle", margin: 0 }
);
s2.addText(
  [
    { text: "9   ", options: { bold: true, color: MUTED } },
    { text: "The Family Afterward", options: { strike: true, color: MUTED } },
    { text: "   →   ", options: { color: MUTED } },
    { text: "The Alateens Afterward", options: { bold: true, color: NAVY } },
    { text: "     p. 122", options: { color: MUTED } },
  ],
  { x: 6.9, y: 6.05, w: 5.5, h: 0.55, fontFace: BODY, fontSize: 14, valign: "middle", margin: 0 }
);
addFooter(s2);

s2.addNotes(
  "Roadmap slide. The Big Book is built in three movements.\n\n" +
  "Doctor's Opinion + Bill's Story = Step 1. The Problem. We are powerless.\n" +
  "Chapters 2-4 = Step 2. The Solution. There is a Power.\n" +
  "Chapters 5-7 = Steps 3 through 12. The action necessary for recovery.\n\n" +
  "Chapters 8 and 9 speak to the family - read today as Al-Anon and Alateen.\n\n" +
  "The \"MUSTS\" exercise begins on page 142."
);

// =====================================================================
// SLIDE 3 — page 142, "Check out the newcomer" (the MUSTS)
// =====================================================================
const s3 = pres.addSlide();
s3.background = { color: "FFFFFF" };

s3.addText("Check Out the Newcomer", {
  x: 0.6, y: 0.35, w: 8.6, h: 0.62,
  fontFace: HEAD, fontSize: 40, bold: true, color: NAVY, margin: 0, valign: "middle",
});
s3.addText("“Start here with the newcomer.”", {
  x: 0.6, y: 1.0, w: 8.6, h: 0.38,
  fontFace: BODY, fontSize: 17, italic: true, color: MUTED, margin: 0, valign: "middle",
});

s3.addShape(pres.ShapeType.roundRect, {
  x: 9.6, y: 0.35, w: 3.13, h: 1.05, fill: { color: MAGENTA }, rectRadius: 0.1,
});
s3.addText(
  [
    { text: "The “MUSTS”", options: { fontFace: HEAD, fontSize: 19, bold: true, color: "FFFFFF", breakLine: true } },
    { text: "Page 142", options: { fontFace: BODY, fontSize: 14, color: MAGENTA_SOFT } },
  ],
  { x: 9.6, y: 0.35, w: 3.13, h: 1.05, align: "center", valign: "middle", margin: 0 }
);

const quotes = [
  {
    y: 1.6, h: 1.0,
    text: "Will he take every necessary step, submit to anything to get well, to stop drinking forever?",
  },
  {
    y: 2.8, h: 1.7,
    text: "…does he think he is fooling you, and that after rest and treatment he will be able to get " +
          "away with a few drinks now and then? We believe a man should be thoroughly probed on these " +
          "points. Be satisfied he is not deceiving himself or you.",
  },
  {
    y: 4.7, h: 1.3,
    text: "Either you are dealing with a man who can and will get well or you are not. If not, why waste " +
          "time with him? This may seem severe, but it is usually the best course.",
  },
];

quotes.forEach((q, i) => {
  s3.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: q.y, w: 12.13, h: q.h, fill: { color: YELLOW_TINT }, rectRadius: 0.08,
  });
  s3.addShape(pres.ShapeType.ellipse, {
    x: 0.9, y: q.y + q.h / 2 - 0.24, w: 0.48, h: 0.48, fill: { color: MAGENTA },
  });
  s3.addText(String(i + 1), {
    x: 0.9, y: q.y + q.h / 2 - 0.24, w: 0.48, h: 0.48,
    fontFace: BODY, fontSize: 16, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });
  s3.addText(q.text, {
    x: 1.6, y: q.y, w: 10.95, h: q.h,
    fontFace: BODY, fontSize: 20, color: NAVY, valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
});

s3.addText(
  [
    { text: "Alcoholics Anonymous", options: { italic: true } },
    { text: ", 4th ed., “To Employers,” p. 142." },
  ],
  { x: 0.6, y: 6.25, w: 12.13, h: 0.32, fontFace: BODY, fontSize: 12, color: MUTED, valign: "middle", margin: 0 }
);
addFooter(s3);

s3.addNotes(
  "Page 142. Start here with the newcomer - this is where we check him out.\n\n" +
  "Three questions, and they are not rhetorical:\n" +
  "1. Will he take every necessary step? Submit to anything? Stop drinking forever?\n" +
  "2. Or does he think he can get away with a few drinks later? Probe this one thoroughly.\n" +
  "3. He either can and will get well, or he will not. If not, why waste the time?\n\n" +
  "Note for the room: this passage sits in the chapter To Employers, but the test applies\n" +
  "to anyone carrying the message to a newcomer."
);

// =====================================================================
// SLIDE 4 — Preface, p. vii: the basic text
// =====================================================================
const s4 = pres.addSlide();
s4.background = { color: "FFFFFF" };

s4.addText("The Basic Text", {
  x: 0.6, y: 0.35, w: 8.6, h: 0.62,
  fontFace: HEAD, fontSize: 40, bold: true, color: NAVY, margin: 0, valign: "middle",
});
s4.addText("“…the basic text for our Society.”", {
  x: 0.6, y: 1.0, w: 8.6, h: 0.38,
  fontFace: BODY, fontSize: 17, italic: true, color: MUTED, margin: 0, valign: "middle",
});

s4.addShape(pres.ShapeType.roundRect, {
  x: 9.6, y: 0.35, w: 3.13, h: 1.05, fill: { color: BLUE }, rectRadius: 0.1,
});
s4.addText(
  [
    { text: "PREFACE", options: { fontFace: HEAD, fontSize: 19, bold: true, color: "FFFFFF", breakLine: true, charSpacing: 2 } },
    { text: "Page vii", options: { fontFace: BODY, fontSize: 14, color: "C7DEEC" } },
  ],
  { x: 9.6, y: 0.35, w: 3.13, h: 1.05, align: "center", valign: "middle", margin: 0 }
);

s4.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 1.9, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s4.addText(
  [
    { text: "Because this book has become the " },
    { text: "basic text", options: { bold: true } },
    { text: " for our Society and has helped such large numbers of alcoholic men and women to " +
            "recovery, there exists strong sentiment against any radical changes being made in it. " +
            "Therefore, the first portion of this volume, describing the A.A. recovery program, has " +
            "been left largely untouched in the course of revisions made for the second, third, and " +
            "fourth editions." },
  ],
  { x: 1.0, y: 1.6, w: 11.33, h: 1.9, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s4.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.7, w: 12.13, h: 1.2, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s4.addText(
  "The section called “The Doctor’s Opinion” has been kept intact, just as it was originally " +
  "written in 1939 by the late Dr. William D. Silkworth, our Society’s great medical benefactor.",
  { x: 1.0, y: 3.7, w: 11.33, h: 1.2, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s4.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 5.15, w: 12.13, h: 1.0, fill: { color: NAVY }, rectRadius: 0.08,
});
s4.addText("MARGIN NOTE", {
  x: 1.0, y: 5.28, w: 11.33, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});
s4.addText("A book for transferring knowledge.", {
  x: 1.0, y: 5.56, w: 11.33, h: 0.45,
  fontFace: HEAD, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});

s4.addText(
  [
    { text: "Alcoholics Anonymous", options: { italic: true } },
    { text: ", 4th ed., Preface, p. vii." },
  ],
  { x: 0.6, y: 6.3, w: 12.13, h: 0.32, fontFace: BODY, fontSize: 12, color: MUTED, valign: "middle", margin: 0 }
);
addFooter(s4);

s4.addNotes(
  "Preface, page vii. Why this book is the one we work from.\n\n" +
  "It is the basic text - that word is circled for a reason. The first portion, the part\n" +
  "that describes the recovery program, has been left largely untouched through three\n" +
  "revisions. The Doctor's Opinion is exactly as Dr. Silkworth wrote it in 1939.\n\n" +
  "So: this is a book for transferring knowledge. Not a memoir, not history. The directions\n" +
  "have not changed because they did not need to."
);

// =====================================================================
// SLIDE 5 — Foreword to the First Edition, p. ix
// =====================================================================
const s5 = pres.addSlide();
s5.background = { color: "FFFFFF" };

s5.addText("Precisely How We Have Recovered", {
  x: 0.6, y: 0.35, w: 8.9, h: 0.62,
  fontFace: HEAD, fontSize: 38, bold: true, color: NAVY, margin: 0, valign: "middle",
});
s5.addText("The main purpose of this book.", {
  x: 0.6, y: 1.0, w: 8.9, h: 0.38,
  fontFace: BODY, fontSize: 17, italic: true, color: MUTED, margin: 0, valign: "middle",
});

s5.addShape(pres.ShapeType.roundRect, {
  x: 9.7, y: 0.35, w: 3.03, h: 1.05, fill: { color: GREEN }, rectRadius: 0.1,
});
s5.addText(
  [
    { text: "FOREWORD", options: { fontFace: HEAD, fontSize: 17, bold: true, color: "FFFFFF", breakLine: true, charSpacing: 1 } },
    { text: "First Edition  ·  p. ix", options: { fontFace: BODY, fontSize: 13, color: "D5E3CE" } },
  ],
  { x: 9.7, y: 0.35, w: 3.03, h: 1.05, align: "center", valign: "middle", margin: 0 }
);

s5.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 1.5, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s5.addText(
  [
    { text: "We, of Alcoholics Anonymous, are more than " },
    { text: "one hundred men and women", options: { bold: true, color: GREEN } },
    { text: " who have recovered from a seemingly hopeless state of mind and body. To show other " +
            "alcoholics " },
    { text: "precisely how we have recovered", options: { bold: true, italic: true, underline: { style: "sng" } } },
    { text: " is the main purpose of this book. For…" },
  ],
  { x: 1.0, y: 1.55, w: 11.33, h: 1.5, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s5.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.3, w: 12.13, h: 1.5, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s5.addText(
  "…them, we hope these pages will prove so convincing that no further authentication will be " +
  "necessary. We think this account of our experiences will help everyone to better understand " +
  "the alcoholic.",
  { x: 1.0, y: 3.3, w: 11.33, h: 1.5, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s5.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 5.05, w: 12.13, h: 1.15, fill: { color: NAVY }, rectRadius: 0.08,
});
s5.addText("MARGIN NOTE", {
  x: 1.0, y: 5.18, w: 11.33, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});
s5.addText("Authored this book — originally 87, raised to one hundred at printing time.", {
  x: 1.0, y: 5.48, w: 11.33, h: 0.6,
  fontFace: HEAD, fontSize: 22, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});

s5.addText(
  [
    { text: "Alcoholics Anonymous", options: { italic: true } },
    { text: ", 4th ed., Foreword to the First Edition, p. ix." },
  ],
  { x: 0.6, y: 6.4, w: 12.13, h: 0.32, fontFace: BODY, fontSize: 12, color: MUTED, valign: "middle", margin: 0 }
);
addFooter(s5);

s5.addNotes(
  "Foreword to the First Edition - the Foreword as it appeared in 1939.\n\n" +
  "Circled: one hundred men and women. That is who authored this book. The margin note\n" +
  "is fellowship history, not text from the book - the count was 87 and was raised to one\n" +
  "hundred by printing time.\n\n" +
  "Underlined: precisely how we have recovered. Precisely. Not roughly, not our impressions.\n" +
  "That is the main purpose of the book, stated on page ix before anything else.\n\n" +
  "And the promise that follows: no further authentication will be necessary."
);

// =====================================================================
// Shared helpers for the Foreword-to-Second-Edition run
// =====================================================================
const BLUE_SOFT = "C7DEEC";
const GREEN_SOFT = "D5E3CE";

function header(s, title, sub, box) {
  s.addText(title, {
    x: 0.6, y: 0.35, w: 8.9, h: 0.62,
    fontFace: HEAD, fontSize: 38, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s.addText(sub, {
    x: 0.6, y: 1.0, w: 8.9, h: 0.38,
    fontFace: BODY, fontSize: 17, italic: true, color: MUTED, margin: 0, valign: "middle",
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 9.7, y: 0.35, w: 3.03, h: 1.05, fill: { color: box.fill }, rectRadius: 0.1,
  });
  s.addText(
    [
      { text: box.title, options: { fontFace: HEAD, fontSize: 17, bold: true, color: "FFFFFF", breakLine: true, charSpacing: 1 } },
      { text: box.sub, options: { fontFace: BODY, fontSize: 13, color: box.subColor } },
    ],
    { x: 9.7, y: 0.35, w: 3.03, h: 1.05, align: "center", valign: "middle", margin: 0 }
  );
}

function marginNote(s, y, h, text) {
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: y, w: 12.13, h: h, fill: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText("MARGIN NOTE", {
    x: 1.0, y: y + 0.13, w: 11.33, h: 0.28,
    fontFace: BODY, fontSize: 11, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
  });
  s.addText(text, {
    x: 1.0, y: y + 0.43, w: 11.33, h: h - 0.56,
    fontFace: HEAD, fontSize: 22, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
  });
}

function cite(s, y, tail) {
  s.addText(
    [
      { text: "Alcoholics Anonymous", options: { italic: true } },
      { text: ", 4th ed., " + tail },
    ],
    { x: 0.6, y: y, w: 12.13, h: 0.32, fontFace: BODY, fontSize: 12, color: MUTED, valign: "middle", margin: 0 }
  );
}

// =====================================================================
// SLIDE 6 — The spark, Akron 1935 (pp. xi-xii)
// =====================================================================
const s6 = pres.addSlide();
s6.background = { color: "FFFFFF" };
header(s6, "The Spark", "Akron, Ohio — June 1935.",
  { fill: BLUE, title: "FOREWORD", sub: "Second Ed.  ·  p. xi", subColor: BLUE_SOFT });

s6.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 2.3, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s6.addText(
  "The spark that was to flare into the first A.A. group was struck at Akron, Ohio, in June 1935, " +
  "during a talk between a New York stockbroker and an Akron physician. Six months earlier, the " +
  "broker had been relieved of his drink obsession by a sudden spiritual experience, following a " +
  "meeting with an alcoholic friend who had been in contact with the Oxford Groups of that day.",
  { x: 1.0, y: 1.6, w: 11.33, h: 2.3, fontFace: BODY, fontSize: 21, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s6, 4.2, 1.35, "Tell the history from Roland.");
cite(s6, 5.75, "Foreword to the Second Edition, pp. xi–xii.");
addFooter(s6);

s6.addNotes(
  "Stop here and tell the Roland story before reading on.\n\n" +
  "Roland Hazard - the chain starts with him. He goes to Carl Jung, is told his case is\n" +
  "hopeless short of a vital spiritual experience, finds the Oxford Groups, and carries it to\n" +
  "Ebby. Ebby is the alcoholic friend in this paragraph. Ebby carries it to the broker.\n\n" +
  "The broker is Bill W. The Akron physician is Dr. Bob. June 1935."
);

// =====================================================================
// SLIDE 7 — What the broker had learned (p. xii)
// =====================================================================
const s7 = pres.addSlide();
s7.background = { color: "FFFFFF" };
header(s7, "The Grave Nature of Alcoholism", "What the broker had already learned.",
  { fill: BLUE, title: "FOREWORD", sub: "Second Ed.  ·  p. xii", subColor: BLUE_SOFT });

s7.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 1.95, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s7.addText(
  "He had also been greatly helped by the late Dr. William D. Silkworth, a New York specialist in " +
  "alcoholism who is now accounted no less than a medical saint by A.A. members, and whose story " +
  "of the early days of our Society appears in the next pages. From this doctor, the broker had " +
  "learned the grave nature of alcoholism.",
  { x: 1.0, y: 1.55, w: 11.33, h: 1.95, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s7.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.7, w: 12.13, h: 2.0, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s7.addText(
  [
    { text: "Though he could not accept all the tenets of the Oxford Groups, he was convinced of the need for " },
    { text: "moral inventory", options: { bold: true, color: BLUE } },
    { text: ", " },
    { text: "confession of personality defects", options: { bold: true, color: BLUE } },
    { text: ", " },
    { text: "restitution to those harmed", options: { bold: true, color: BLUE } },
    { text: ", " },
    { text: "helpfulness to others", options: { bold: true, color: BLUE } },
    { text: ", and the necessity of " },
    { text: "belief in and dependence upon God", options: { bold: true, color: BLUE } },
    { text: "." },
  ],
  { x: 1.0, y: 3.7, w: 11.33, h: 2.0, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s7, 5.95, "Foreword to the Second Edition, p. xii.");
addFooter(s7);

s7.addNotes(
  "Five things, and every one of them ends up in the Twelve Steps.\n\n" +
  "Moral inventory - Step 4. Confession of personality defects - Step 5. Restitution to those\n" +
  "harmed - Steps 8 and 9. Helpfulness to others - Step 12. Dependence upon God - Steps 3 and 11.\n\n" +
  "He could not accept all of the Oxford Groups. He kept these."
);

// =====================================================================
// SLIDE 8 — He must carry his message (p. xii)
// =====================================================================
const s8 = pres.addSlide();
s8.background = { color: "FFFFFF" };
header(s8, "He Must Carry His Message", "In order to save himself.",
  { fill: GREEN, title: "FOREWORD", sub: "Second Ed.  ·  p. xii", subColor: GREEN_SOFT });

s8.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 2.5, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s8.addText(
  [
    { text: "Prior to his journey to Akron, the broker had worked hard with many alcoholics on the theory " +
            "that only an alcoholic could help an alcoholic, but he had succeeded only in keeping sober " +
            "himself. The broker had gone to Akron on a business venture which had collapsed, leaving him " +
            "greatly in fear that he might start drinking again. He suddenly realized that in order to " +
            "save himself " },
    { text: "he must carry his message to another alcoholic", options: { bold: true, color: MAGENTA } },
    { text: ". That alcoholic turned out to be the Akron physician." },
  ],
  { x: 1.0, y: 1.55, w: 11.33, h: 2.5, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s8.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 4.3, w: 12.13, h: 1.35, fill: { color: MAGENTA }, rectRadius: 0.08,
});
s8.addText("A “MUST”", {
  x: 1.0, y: 4.43, w: 11.33, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: MAGENTA_SOFT, margin: 0, valign: "middle", charSpacing: 2,
});
s8.addText("…he must carry his message to another alcoholic.", {
  x: 1.0, y: 4.73, w: 11.33, h: 0.75,
  fontFace: HEAD, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});

cite(s8, 5.9, "Foreword to the Second Edition, p. xii.");
addFooter(s8);

s8.addNotes(
  "He had it backwards for six months. He worked hard on other alcoholics on the theory that\n" +
  "only an alcoholic could help an alcoholic - and all it did was keep him sober. Which,\n" +
  "it turns out, was the point.\n\n" +
  "Akron. Business venture collapsed. Afraid he is going to drink. And what he realizes is\n" +
  "not that he needs a meeting or a drink or a plan - he must carry the message to another\n" +
  "alcoholic in order to save himself.\n\n" +
  "That alcoholic turned out to be Dr. Bob."
);

// =====================================================================
// SLIDE 9 — Two things we learned (pp. xii-xiii)
// =====================================================================
const s9 = pres.addSlide();
s9.background = { color: "FFFFFF" };
header(s9, "Two Things We Learned", "From the first talk between the broker and the physician.",
  { fill: NAVY, title: "FOREWORD", sub: "Second Ed.  ·  pp. xii–xiii", subColor: ON_NAVY });

const lessons = [
  { y: 1.75, h: 1.7, dark: BLUE, tint: BLUE_TINT, n: "1",
    text: "This seemed to prove that one alcoholic could affect another as no nonalcoholic could." },
  { y: 3.7, h: 1.7, dark: GREEN, tint: GREEN_TINT, n: "2",
    text: "It also indicated that strenuous work, one alcoholic with another, was vital to permanent recovery." },
];

lessons.forEach((l) => {
  s9.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: l.y, w: 12.13, h: l.h, fill: { color: l.tint }, rectRadius: 0.08,
  });
  s9.addShape(pres.ShapeType.ellipse, {
    x: 1.0, y: l.y + l.h / 2 - 0.32, w: 0.64, h: 0.64, fill: { color: l.dark },
  });
  s9.addText(l.n, {
    x: 1.0, y: l.y + l.h / 2 - 0.32, w: 0.64, h: 0.64,
    fontFace: HEAD, fontSize: 22, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
  });
  s9.addText(l.text, {
    x: 1.95, y: l.y, w: 10.4, h: l.h,
    fontFace: BODY, fontSize: 22, color: NAVY, valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
});

cite(s9, 5.65, "Foreword to the Second Edition, pp. xii–xiii.");
addFooter(s9);

s9.addNotes(
  "These two are numbered in the book for a reason. Everything A.A. does rests on them.\n\n" +
  "One: an alcoholic reaches another alcoholic in a way that no nonalcoholic can. Not a\n" +
  "doctor, not a preacher, not a spouse. That is why we have sponsors.\n\n" +
  "Two: strenuous work, one alcoholic with another. Strenuous. Not casual, not when it is\n" +
  "convenient. Vital to permanent recovery - vital meaning you do not live without it."
);

// =====================================================================
// SLIDE 10 — A.A. Number Three (p. xiii)
// =====================================================================
const s10 = pres.addSlide();
s10.background = { color: "FFFFFF" };
header(s10, "A.A. Number Three", "Their very first case.",
  { fill: GREEN, title: "FOREWORD", sub: "Second Ed.  ·  p. xiii", subColor: GREEN_SOFT });

// Photo well - Will supplies the image on the desktop.
s10.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 5.35, h: 4.0,
  fill: { color: "F1F4F7" }, rectRadius: 0.08,
  line: { color: "AEBBC9", width: 1.5, dashType: "dash" },
});
s10.addText(
  [
    { text: "PHOTOGRAPH", options: { fontSize: 11, bold: true, color: MUTED, charSpacing: 2, breakLine: true } },
    { text: "\n", options: { fontSize: 8, breakLine: true } },
    { text: "Bill D. in the hospital bed", options: { fontSize: 17, bold: true, color: NAVY, breakLine: true } },
    { text: "Akron City Hospital, 1935", options: { fontSize: 15, color: MUTED, breakLine: true } },
    { text: "\n", options: { fontSize: 8, breakLine: true } },
    { text: "Drop the image into this frame in PowerPoint.", options: { fontSize: 13, italic: true, color: MUTED } },
  ],
  { x: 0.9, y: 1.6, w: 4.75, h: 4.0, fontFace: BODY, align: "center", valign: "middle", margin: 0 }
);

s10.addShape(pres.ShapeType.roundRect, {
  x: 6.3, y: 1.6, w: 6.43, h: 2.25, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s10.addText(
  [
    { text: "Their very first case, a desperate one, recovered immediately and became A.A. number three.",
      options: { bold: true, underline: { style: "sng" } } },
    { text: " He never had another drink." },
  ],
  { x: 6.7, y: 1.6, w: 5.63, h: 2.25, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s10.addShape(pres.ShapeType.roundRect, {
  x: 6.3, y: 4.05, w: 6.43, h: 1.55, fill: { color: NAVY }, rectRadius: 0.08,
});
s10.addText("MARGIN NOTE", {
  x: 6.7, y: 4.2, w: 5.63, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});
s10.addText("Bill D. — A.A. Number Three", {
  x: 6.7, y: 4.52, w: 5.63, h: 0.85,
  fontFace: HEAD, fontSize: 22, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});

cite(s10, 5.85, "Foreword to the Second Edition, p. xiii.");
addFooter(s10);

s10.addNotes(
  "The first man they worked on. A desperate case, in a bed at Akron City Hospital, summer 1935.\n" +
  "He recovered immediately and never had another drink.\n\n" +
  "Bill D. - A.A. number three. The man in the photograph.\n\n" +
  "Two men who could not stay sober alone stayed sober by working on a third."
);

// =====================================================================
// SLIDE 11 — Hang together or die separately (pp. xiv-xv)
// =====================================================================
const s11 = pres.addSlide();
s11.background = { color: "FFFFFF" };
header(s11, "Hang Together or Die Separately", "The adolescent period, and the test it faced.",
  { fill: BLUE, title: "FOREWORD", sub: "Second Ed.  ·  pp. xiv–xv", subColor: BLUE_SOFT });

s11.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 1.8, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s11.addText(
  "Our Society then entered a fearsome and exciting adolescent period. The test that it faced was " +
  "this: Could these large numbers of erstwhile erratic alcoholics successfully meet and work " +
  "together? Would there be quarrels over membership, leadership, and money? Would there be " +
  "strivings for power and prestige? Would there be schisms which would split A.A. apart?",
  { x: 1.0, y: 1.55, w: 11.33, h: 1.8, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s11.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.55, w: 12.13, h: 1.45, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s11.addText(
  [
    { text: "Soon A.A. was beset by these very problems on every side and in every group. But out of this " +
            "frightening and at first disrupting experience the conviction grew that " },
    { text: "A.A.’s had to hang together or die separately", options: { bold: true, color: BLUE } },
    { text: ". We had to unify our Fellowship or pass off the scene." },
  ],
  { x: 1.0, y: 3.55, w: 11.33, h: 1.45, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s11, 5.3, 1.0, "The Fellowship was named by the book — Alcoholics Anonymous.");
cite(s11, 6.45, "Foreword to the Second Edition, pp. xiv–xv.");
addFooter(s11);

s11.addNotes(
  "Everything in this paragraph is a question, and every one of them got answered the hard way.\n" +
  "Quarrels over membership, leadership, money. Strivings for power and prestige. Schisms.\n" +
  "A.A. had all of it.\n\n" +
  "What came out of it was the Twelve Traditions - and the conviction underneath them:\n" +
  "hang together or die separately.\n\n" +
  "Note the order of events. The book came first, in 1939. The Fellowship took its name\n" +
  "from the book, not the other way around."
);

// =====================================================================
// SLIDE 12 — Really tried (pp. xv-xvi)
// =====================================================================
const s12 = pres.addSlide();
s12.background = { color: "FFFFFF" };
header(s12, "Came to A.A. and Really Tried", "Why public acceptance grew.",
  { fill: GREEN, title: "FOREWORD", sub: "Second Ed.  ·  pp. xv–xvi", subColor: GREEN_SOFT });

s12.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 1.5, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s12.addText(
  "While the internal difficulties of our adolescent period were being ironed out, public " +
  "acceptance of A.A. grew by leaps and bounds. For this there were two principal reasons: the " +
  "large numbers of recoveries, and reunited homes. These made their impressions everywhere.",
  { x: 1.0, y: 1.55, w: 11.33, h: 1.5, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s12.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.25, w: 12.13, h: 2.2, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s12.addText(
  [
    { text: "Of alcoholics who came to A.A. and " },
    { text: "really tried", options: { bold: true, italic: true } },
    { text: ", " },
    { text: "50%", options: { bold: true, color: GREEN } },
    { text: " got sober at once and remained that way; " },
    { text: "25%", options: { bold: true, color: GREEN } },
    { text: " sobered up after some relapses, and among the remainder, those who stayed on with A.A. " +
            "showed improvement. Other thousands came to a few A.A. meetings and at first decided they " +
            "didn’t want the program. But great numbers of these—" },
    { text: "about two out of three", options: { bold: true, color: GREEN } },
    { text: "—began to return as time passed." },
  ],
  { x: 1.0, y: 3.25, w: 11.33, h: 2.2, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s12, 5.75, "Foreword to the Second Edition, pp. xv–xvi.");
addFooter(s12);

s12.addNotes(
  "Two principal reasons the public came around: recoveries, and reunited homes. Not\n" +
  "advertising. Not argument. Results people could see in their own neighborhoods.\n\n" +
  "Watch the qualifier - came to A.A. and really tried. The numbers are attached to that\n" +
  "phrase, not to attendance.\n\n" +
  "And the last line is the one to sit on: of the thousands who came, decided they did not\n" +
  "want it, and left - about two out of three came back."
);

// =====================================================================
// SLIDE 13 — Not a religious organization (p. xvi)
// =====================================================================
const s13 = pres.addSlide();
s13.background = { color: "FFFFFF" };
header(s13, "Not a Religious Organization", "What A.A. is not.",
  { fill: BLUE, title: "FOREWORD", sub: "Second Ed.  ·  p. xvi", subColor: BLUE_SOFT });

s13.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.9, w: 12.13, h: 2.4, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s13.addText(
  [
    { text: "Alcoholics Anonymous is not a religious organization.", options: { bold: true } },
    { text: " Neither does A.A. take any particular medical point of view, though we cooperate widely " +
            "with the men of medicine as well as with the men of religion." },
  ],
  { x: 1.0, y: 1.9, w: 11.33, h: 2.4, fontFace: BODY, fontSize: 26, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s13, 4.65, 1.0, "Newcomers.");
cite(s13, 6.05, "Foreword to the Second Edition, p. xvi.");
addFooter(s13);

s13.addNotes(
  "This is the paragraph for the person sitting in the back who has been told all his life\n" +
  "that A.A. is a religion, or a cult, or a hospital program.\n\n" +
  "Not a religious organization. No particular medical point of view. We cooperate with\n" +
  "both and we are neither.\n\n" +
  "Ninety words that have kept more newcomers in the room than any argument ever has."
);

// =====================================================================
// SLIDE 14 — December 1941 membership (margin note, p. xvii)
// =====================================================================
const s14 = pres.addSlide();
s14.background = { color: "FFFFFF" };
header(s14, "December 1941", "National Directory — members by city.",
  { fill: NAVY, title: "FOREWORD", sub: "Second Ed.  ·  p. xvii", subColor: ON_NAVY });

const cities = [
  { name: "New York City", n: "450", lead: false },
  { name: "Akron", n: "225", lead: false },
  { name: "Detroit", n: "175", lead: false },
  { name: "Cleveland", n: "1,100", lead: true },
];

cities.forEach((c, i) => {
  const y = 1.5 + i * 0.86;
  s14.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: y, w: 12.13, h: 0.76,
    fill: { color: c.lead ? GREEN : "F1F4F7" }, rectRadius: 0.08,
  });
  s14.addText(c.name, {
    x: 1.0, y: y, w: 7.0, h: 0.76,
    fontFace: BODY, fontSize: 22, bold: c.lead, color: c.lead ? "FFFFFF" : NAVY,
    valign: "middle", margin: 0,
  });
  s14.addText(c.n, {
    x: 8.2, y: y, w: 4.13, h: 0.76,
    fontFace: HEAD, fontSize: 28, bold: true, color: c.lead ? "FFFFFF" : NAVY,
    align: "right", valign: "middle", margin: 0,
  });
});

marginNote(s14, 5.25, 1.05, "Clarence Snyder was using the book — is why the #’s were higher.");

s14.addText("Handwritten margin note, Foreword to the Second Edition, p. xvii.", {
  x: 0.6, y: 6.45, w: 12.13, h: 0.3, fontFace: BODY, fontSize: 12, color: MUTED,
  valign: "middle", margin: 0,
});
addFooter(s14);

s14.addNotes(
  "These figures are the margin note, not text from the book. Read them as such.\n\n" +
  "Look at the shape of it. New York, where Bill was, 450. Akron, where it started, 225.\n" +
  "Detroit, 175. Cleveland - 1,100. More than the other three put together.\n\n" +
  "Clarence Snyder started the Cleveland group in 1939, the first to call itself Alcoholics\n" +
  "Anonymous. Cleveland took newcomers straight through the book and built sponsorship\n" +
  "around it. That is fellowship history rather than book text, but the numbers are hard\n" +
  "to argue with.\n\n" +
  "They used the book. The numbers were higher."
);

// =====================================================================
// SLIDE 15 — The Doctor's Opinion (p. xxi)
// =====================================================================
const MAGENTA_TINT = "F9E3EE";

const s15 = pres.addSlide();
s15.background = { color: "FFFFFF" };
header(s15, "The Doctor’s Opinion", "“Just an opinion.”",
  { fill: BLUE, title: "THE BIG BOOK", sub: "p. xxi", subColor: BLUE_SOFT });

const doctors = [
  {
    x: 0.6,
    name: "Dr. Benjamin Rush",
    role: "Surgeon General of the Continental Army",
    body: [
      { text: "In 1784 he published " },
      { text: "An Inquiry into the Effects of Ardent Spirits upon the Human Body and Mind",
        options: { italic: true } },
      { text: " — the first American physician to argue in print that alcoholism is a disease." },
    ],
  },
  {
    x: 6.78,
    name: "Dr. Thomas Trotter",
    role: "Physician to the British Fleet",
    body: [
      { text: "His " },
      { text: "Essay on Drunkenness, and its Effects on the Human Body",
        options: { italic: true } },
      { text: ", published in 1804, was the first book-length medical study of alcohol dependence." },
    ],
  },
];

doctors.forEach((d) => {
  s15.addShape(pres.ShapeType.roundRect, {
    x: d.x, y: 1.7, w: 5.95, h: 2.9, fill: { color: BLUE_TINT }, rectRadius: 0.08,
  });
  s15.addText(d.name, {
    x: d.x + 0.35, y: 2.0, w: 5.25, h: 0.45,
    fontFace: HEAD, fontSize: 24, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s15.addText(d.role, {
    x: d.x + 0.35, y: 2.45, w: 5.25, h: 0.3,
    fontFace: BODY, fontSize: 14, bold: true, color: BLUE, margin: 0, valign: "middle", charSpacing: 1,
  });
  s15.addText(d.body, {
    x: d.x + 0.35, y: 2.85, w: 5.25, h: 1.5,
    fontFace: BODY, fontSize: 17, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });
});

marginNote(s15, 4.9, 1.1,
  "Page 1 in every first edition. Moved to the front matter in the second.");
cite(s15, 6.2, "“The Doctor’s Opinion,” p. xxi.");
addFooter(s15);

s15.addNotes(
  "Note what is written under the title: just an opinion. The chapter is a doctor's letter,\n" +
  "not doctrine, and the book presents it that way.\n\n" +
  "But the disease idea did not start with A.A. Rush was writing about it in 1784, Trotter in\n" +
  "1804 - a hundred and fifty years before this book. Silkworth is standing on that.\n\n" +
  "Dates on this slide are the landmark publications. If you have a source for the 1776 and\n" +
  "1782 dates from the margin, use those instead.\n\n" +
  "And this chapter was page 1 in every first edition. It was moved to the front matter in the\n" +
  "second - which is how Bill's Story became page 1."
);

// =====================================================================
// SLIDE 16 — Two musts from The Doctor's Opinion (p. xxi)
// =====================================================================
const s16 = pres.addSlide();
s16.background = { color: "FFFFFF" };
header(s16, "Two Musts", "The Doctor’s Opinion.",
  { fill: MAGENTA, title: "THE “MUSTS”", sub: "p. xxi", subColor: MAGENTA_SOFT });

const xxiMusts = [
  { y: 1.7, h: 1.65,
    text: "Convincing testimony must surely come from medical men who have had experience with the " +
          "sufferings of our members and have witnessed our return to health." },
  { y: 3.6, h: 1.65,
    text: "As part of his rehabilitation he commenced to present his conceptions to other alcoholics, " +
          "impressing upon them that they must do likewise with still others." },
];

xxiMusts.forEach((q, i) => {
  s16.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: q.y, w: 12.13, h: q.h, fill: { color: MAGENTA_TINT }, rectRadius: 0.08,
  });
  s16.addShape(pres.ShapeType.ellipse, {
    x: 1.0, y: q.y + q.h / 2 - 0.3, w: 0.6, h: 0.6, fill: { color: MAGENTA },
  });
  s16.addText(String(i + 1), {
    x: 1.0, y: q.y + q.h / 2 - 0.3, w: 0.6, h: 0.6,
    fontFace: HEAD, fontSize: 21, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
  });
  s16.addText(q.text, {
    x: 1.9, y: q.y, w: 10.45, h: q.h,
    fontFace: BODY, fontSize: 21, color: NAVY, valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
});

cite(s16, 5.55, "“The Doctor’s Opinion,” p. xxi.");
addFooter(s16);

s16.addNotes(
  "Both of these are highlighted pink in the study copy - both are musts.\n\n" +
  "The first is about where the testimony has to come from. Not from us saying we feel better.\n" +
  "From medical men who watched us suffer and then watched us recover.\n\n" +
  "The second is Silkworth describing his patient. Rehabilitation was not finished when the man\n" +
  "stopped drinking. Part of it was carrying it to other alcoholics - and impressing on them\n" +
  "that they must do the same with still others.\n\n" +
  "That is the chain. It is in the doctor's letter before the book even starts."
);

// =====================================================================
// SLIDE 17 — The signature (p. xxii)
// =====================================================================
const s17 = pres.addSlide();
s17.background = { color: "FFFFFF" };
header(s17, "William D. Silkworth, M.D.", "The letter is signed.",
  { fill: BLUE, title: "THE BIG BOOK", sub: "p. xxii", subColor: BLUE_SOFT });

s17.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.8, w: 12.13, h: 2.3, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s17.addText("Very truly yours,", {
  x: 1.0, y: 2.2, w: 11.33, h: 0.5,
  fontFace: BODY, fontSize: 22, italic: true, color: NAVY, align: "right", valign: "middle", margin: 0,
});
s17.addText("William D. Silkworth, M.D.", {
  x: 1.0, y: 2.8, w: 11.33, h: 0.75,
  fontFace: HEAD, fontSize: 34, bold: true, color: BLUE, align: "right", valign: "middle", margin: 0,
});

marginNote(s17, 4.4, 1.1, "Signed “Anonymous” in the first editions.");
cite(s17, 5.7, "“The Doctor’s Opinion,” p. xxii.");
addFooter(s17);

s17.addNotes(
  "The doctor who wrote this letter is the same Silkworth from the Foreword - the one the\n" +
  "book calls a medical saint, the one who told the broker the grave nature of alcoholism.\n\n" +
  "His name is on it now. It was not always. In the first editions the letter was signed\n" +
  "Anonymous, and the name was added later.\n\n" +
  "A doctor put his professional reputation behind a book written by drunks, in 1939, when\n" +
  "there was nothing to gain by it."
);

// =====================================================================
// SLIDE 18 — As abnormal as his mind (p. xxii)
// =====================================================================
const s18 = pres.addSlide();
s18.background = { color: "FFFFFF" };
header(s18, "As Abnormal as His Mind", "What we who have suffered must believe.",
  { fill: MAGENTA, title: "THE “MUSTS”", sub: "p. xxii", subColor: MAGENTA_SOFT });

s18.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 2.3, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s18.addText(
  [
    { text: "The physician who, at our request, gave us this letter, has been kind enough to enlarge " +
            "upon his views in another statement which follows. In this statement he " },
    { text: "confirms what we who have suffered alcoholic torture must believe—",
      options: { bold: true, color: MAGENTA } },
    { text: "that the body of the alcoholic is quite as abnormal as his mind.",
      options: { bold: true, color: MAGENTA, underline: { style: "sng" } } },
  ],
  { x: 1.0, y: 1.6, w: 11.33, h: 2.3, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s18.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 4.2, w: 12.13, h: 1.35, fill: { color: MAGENTA }, rectRadius: 0.08,
});
s18.addText("A “MUST”", {
  x: 1.0, y: 4.33, w: 11.33, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: MAGENTA_SOFT, margin: 0, valign: "middle", charSpacing: 2,
});
s18.addText("…the body of the alcoholic is quite as abnormal as his mind.", {
  x: 1.0, y: 4.63, w: 11.33, h: 0.75,
  fontFace: HEAD, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});

cite(s18, 5.8, "“The Doctor’s Opinion,” p. xxii.");
addFooter(s18);

s18.addNotes(
  "Body and mind. Both. That is the whole claim, and it is a must - what we who have\n" +
  "suffered alcoholic torture must believe.\n\n" +
  "Not a weak mind housed in a healthy body. Not a moral failure a person could think his\n" +
  "way out of. The body is as abnormal as the mind.\n\n" +
  "This is Step One territory. If only the mind were the problem, willpower would work."
);

// =====================================================================
// SLIDE 19 — The physical factor (p. xxii)
// =====================================================================
const s19 = pres.addSlide();
s19.background = { color: "FFFFFF" };
header(s19, "The Physical Factor", "It did not satisfy us to be told…",
  { fill: GREEN, title: "THE BIG BOOK", sub: "p. xxii", subColor: GREEN_SOFT });

s19.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.85, w: 12.13, h: 3.05, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s19.addText(
  [
    { text: "It did not satisfy us to be told that we could not " },
    { text: "control", options: { bold: true, underline: { style: "dbl" } } },
    { text: " our drinking just because we were maladjusted to life, that we were in full flight from " +
            "reality, or were outright mental defectives. These things were true to some extent, in " +
            "fact, to a considerable extent with some of us. But we are sure that our bodies were " +
            "sickened as well. " },
    { text: "In our belief, any picture of the alcoholic which leaves out this physical factor is " +
            "incomplete.", options: { bold: true, color: GREEN, underline: { style: "sng" } } },
  ],
  { x: 1.0, y: 1.85, w: 11.33, h: 3.05, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s19, 5.2, "“The Doctor’s Opinion,” p. xxii.");
addFooter(s19);

s19.addNotes(
  "Control is double-underlined in the study copy. Sit on that word.\n\n" +
  "They are not arguing that the psychological picture is wrong. Read it again - these things\n" +
  "were true to some extent, in fact to a considerable extent with some of us. Maladjusted,\n" +
  "in flight from reality, all of it. They grant it.\n\n" +
  "What they will not grant is that it is the whole story. Our bodies were sickened as well.\n\n" +
  "Any picture that leaves out the physical factor is incomplete. Incomplete, not wrong -\n" +
  "and an incomplete picture is what keeps a man trying to solve a physical problem with\n" +
  "resolutions."
);

// =====================================================================
// SLIDE 20 — An allergy to alcohol (p. xxii)
// =====================================================================
const s20 = pres.addSlide();
s20.background = { color: "FFFFFF" };
header(s20, "An Allergy to Alcohol", "The doctor’s theory.",
  { fill: BLUE, title: "THE BIG BOOK", sub: "p. xxii", subColor: BLUE_SOFT });

s20.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.75, w: 12.13, h: 2.6, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s20.addText(
  [
    { text: "The doctor’s theory that we have an allergy to alcohol interests us.",
      options: { bold: true, color: BLUE, underline: { style: "sng" } } },
    { text: " As laymen, our opinion as to its soundness may, of course, mean little. But as " +
            "ex-problem drinkers, we can say that his explanation makes good sense. It explains " +
            "many things for which we cannot otherwise account." },
  ],
  { x: 1.0, y: 1.75, w: 11.33, h: 2.6, fontFace: BODY, fontSize: 21, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s20, 4.65, 1.05, "Allergy = abnormal reaction.");
cite(s20, 5.9, "“The Doctor’s Opinion,” p. xxii.");
addFooter(s20);

s20.addNotes(
  "Interests us. Not proves, not settles - interests. They are careful here.\n\n" +
  "As laymen our opinion means little. They say so themselves. But as ex-problem drinkers -\n" +
  "and that is the standing that counts in this room - the explanation makes good sense.\n\n" +
  "It explains many things for which we cannot otherwise account. That is the test they are\n" +
  "applying. Not is it proven, but does it account for what happened to me."
);

// =====================================================================
// SLIDE 21 — Of paramount importance (p. xxiii)
// =====================================================================
const s21 = pres.addSlide();
s21.background = { color: "FFFFFF" };
header(s21, "Of Paramount Importance", "The doctor writes:",
  { fill: GREEN, title: "THE BIG BOOK", sub: "p. xxiii", subColor: GREEN_SOFT });

s21.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 2.9, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s21.addText(
  [
    { text: "The subject presented in this book seems to me to be of paramount importance to those " +
            "afflicted with alcoholic addiction.",
      options: { breakLine: true, paraSpaceAfter: 10 } },
    { text: "I say this after many years’ experience as Medical Director of " },
    { text: "one of the oldest hospitals in the country treating alcoholic and drug addiction",
      options: { bold: true, color: GREEN } },
    { text: ".", options: { breakLine: true, paraSpaceAfter: 10 } },
    { text: "There was, therefore, a sense of real satisfaction when I was asked to contribute a few " +
            "words on a subject which is covered in such masterly detail in these pages." },
  ],
  { x: 1.0, y: 1.6, w: 11.33, h: 2.9, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s21, 4.8, 1.05, "Towns Hospital, New York City.");
cite(s21, 6.05, "“The Doctor’s Opinion,” p. xxiii.");
addFooter(s21);

s21.addNotes(
  "The green arrow in the study copy starts here. This is the doctor speaking in his own\n" +
  "voice, and the first thing he does is establish standing.\n\n" +
  "The hospital is the Charles B. Towns Hospital, 293 Central Park West in Manhattan.\n" +
  "Silkworth was its medical director. That is where he saw thousands of alcoholics come\n" +
  "through, and it is where he met Bill.\n\n" +
  "Note his word for the book - masterly. A physician calling a book written by drunks\n" +
  "masterly, in 1939."
);

// =====================================================================
// SLIDE 22 — Beyond our conception (p. xxiii)
// =====================================================================
const s22 = pres.addSlide();
s22.background = { color: "FFFFFF" };
header(s22, "Beyond Our Conception", "What we doctors have realized.",
  { fill: GREEN, title: "THE BIG BOOK", sub: "p. xxiii", subColor: GREEN_SOFT });

[
  { y: 1.8, h: 1.6,
    runs: [
      { text: "We doctors have realized for a long time that some form of " },
      { text: "moral psychology", options: { bold: true, color: GREEN } },
      { text: " was of urgent importance to alcoholics, but its application presented difficulties " +
              "beyond our conception." },
    ] },
  { y: 3.7, h: 1.6,
    runs: [
      { text: "What with our ultra-modern standards, our scientific approach to everything, we are " +
              "perhaps not well equipped to apply the powers of good that lie outside our synthetic " +
              "knowledge." },
    ] },
].forEach((c) => {
  s22.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: c.y, w: 12.13, h: c.h, fill: { color: GREEN_TINT }, rectRadius: 0.08,
  });
  s22.addText(c.runs, {
    x: 1.0, y: c.y, w: 11.33, h: c.h, fontFace: BODY, fontSize: 21, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
});

cite(s22, 5.6, "“The Doctor’s Opinion,” p. xxiii.");
addFooter(s22);

s22.addNotes(
  "A doctor admitting the limits of medicine. Read that second paragraph slowly.\n\n" +
  "They knew some form of moral psychology was urgently needed. They could not deliver it.\n" +
  "Its application presented difficulties beyond our conception.\n\n" +
  "And then he says why - our ultra-modern standards, our scientific approach to everything.\n" +
  "The training that makes a good physician is the same training that leaves him not well\n" +
  "equipped to apply powers of good outside his synthetic knowledge.\n\n" +
  "He is describing the gap that the Twelve Steps walked into."
);

// =====================================================================
// SLIDE 23 — One of the leading contributors (p. xxiii)
// =====================================================================
const s23 = pres.addSlide();
s23.background = { color: "FFFFFF" };
header(s23, "One of the Leading Contributors", "Came under our care in this hospital.",
  { fill: GREEN, title: "THE BIG BOOK", sub: "p. xxiii", subColor: GREEN_SOFT });

s23.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.7, w: 12.13, h: 1.9, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s23.addText(
  [
    { text: "Many years ago " },
    { text: "one of the leading contributors to this book", options: { bold: true, color: GREEN } },
    { text: " came under our care in " },
    { text: "this hospital", options: { bold: true, color: GREEN } },
    { text: " and while here " },
    { text: "he", options: { bold: true, color: GREEN } },
    { text: " acquired some ideas which he put into practical application at once." },
  ],
  { x: 1.0, y: 1.7, w: 11.33, h: 1.9, fontFace: BODY, fontSize: 22, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

[
  { x: 0.6, label: "THIS HOSPITAL", big: "Charles B. Towns Hospital",
    small: "293 Central Park West, New York City." },
  { x: 6.78, label: "HE", big: "Bill W.",
    small: "Admitted four times, 1933–34. Last drink 11 December 1934." },
].forEach((c) => {
  s23.addShape(pres.ShapeType.roundRect, {
    x: c.x, y: 3.9, w: 5.95, h: 1.7, fill: { color: NAVY }, rectRadius: 0.08,
  });
  s23.addText(c.label, {
    x: c.x + 0.4, y: 4.05, w: 5.15, h: 0.28,
    fontFace: BODY, fontSize: 11, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
  });
  s23.addText(c.big, {
    x: c.x + 0.4, y: 4.35, w: 5.15, h: 0.5,
    fontFace: HEAD, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
  });
  s23.addText(c.small, {
    x: c.x + 0.4, y: 4.9, w: 5.15, h: 0.55,
    fontFace: BODY, fontSize: 14, color: ON_NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });
});

cite(s23, 5.8, "“The Doctor’s Opinion,” p. xxiii.");
addFooter(s23);

s23.addNotes(
  "Silkworth is being discreet. He does not name him. But everyone reading this in 1939 who\n" +
  "knew anything knew who it was.\n\n" +
  "One of the leading contributors to this book is Bill W. This hospital is Towns. Bill was\n" +
  "admitted four times between 1933 and 1934 - his last drink was 11 December 1934, and\n" +
  "three days later, still in that hospital, he had the experience he describes in his story.\n\n" +
  "The ideas he acquired came from Silkworth: the allergy, and the obsession. And he put them\n" +
  "into practical application at once - meaning he went out and started talking to other\n" +
  "alcoholics.\n\n" +
  "The doctor who explained the illness to him is the same doctor writing this letter."
);

// =====================================================================
// SLIDE 24 — The phenomenon of craving (p. xxiv)
// =====================================================================
const s24 = pres.addSlide();
s24.background = { color: "FFFFFF" };
header(s24, "The Phenomenon of Craving", "A manifestation of an allergy.",
  { fill: GREEN, title: "THE BIG BOOK", sub: "p. xxiv", subColor: GREEN_SOFT });

s24.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 1.5, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s24.addText(
  [
    { text: "We believe, and so suggested a few years ago, that the action of alcohol on these chronic " +
            "alcoholics is a manifestation of an allergy; that the " },
    { text: "phenomenon of craving is limited to this class", options: { bold: true, color: GREEN } },
    { text: " and never occurs in the average temperate drinker." },
  ],
  { x: 1.0, y: 1.55, w: 11.33, h: 1.5, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s24.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.25, w: 12.13, h: 1.7, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s24.addText(
  [
    { text: "These allergic types can " },
    { text: "never safely use alcohol in any form at all", options: { bold: true, color: BLUE } },
    { text: "; and once having formed the habit and found they cannot break it, once having lost their " +
            "self-confidence, their reliance upon things human, their problems pile up on them and " +
            "become astonishingly difficult to solve." },
  ],
  { x: 1.0, y: 3.25, w: 11.33, h: 1.7, fontFace: BODY, fontSize: 19, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s24, 5.2, 0.95, "Crave = physical, in this book.");
cite(s24, 6.35, "“The Doctor’s Opinion,” p. xxiv.");
addFooter(s24);

s24.addNotes(
  "Craving in this book means the physical thing. Not wanting a drink - the body's reaction\n" +
  "once alcohol is in it.\n\n" +
  "And it is limited to this class. It never occurs in the average temperate drinker. That is\n" +
  "why the normal drinker cannot understand us and why his advice never helps.\n\n" +
  "Never safely use alcohol in any form at all. Not less. Not carefully. Not at all.\n\n" +
  "Then watch the order of the collapse: form the habit, cannot break it, lose self-confidence,\n" +
  "lose reliance upon things human. The problems pile up after that, not before."
);

// =====================================================================
// SLIDE 25 — Frothy emotional appeal (p. xxiv)
// =====================================================================
const s25 = pres.addSlide();
s25.background = { color: "FFFFFF" };
header(s25, "Frothy Emotional Appeal", "Seldom suffices.",
  { fill: MAGENTA, title: "THE “MUSTS”", sub: "p. xxiv", subColor: MAGENTA_SOFT });

s25.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 1.9, fill: { color: MAGENTA_TINT }, rectRadius: 0.08,
});
s25.addText(
  [
    { text: "Frothy emotional appeal seldom suffices. The message which can interest and hold these " +
            "alcoholic people " },
    { text: "must have depth and weight", options: { bold: true, color: MAGENTA } },
    { text: ". In nearly all cases, their ideals " },
    { text: "must be grounded in a power greater than themselves", options: { bold: true, color: MAGENTA } },
    { text: ", if they are to re-create their lives." },
  ],
  { x: 1.0, y: 1.6, w: 11.33, h: 1.9, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s25.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.8, w: 12.13, h: 1.5, fill: { color: MAGENTA }, rectRadius: 0.08,
});
s25.addText("A “MUST”", {
  x: 1.0, y: 3.95, w: 11.33, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: MAGENTA_SOFT, margin: 0, valign: "middle", charSpacing: 2,
});
s25.addText("…their ideals must be grounded in a power greater than themselves.", {
  x: 1.0, y: 4.25, w: 11.33, h: 0.85,
  fontFace: HEAD, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});

cite(s25, 5.55, "“The Doctor’s Opinion,” p. xxiv.");
addFooter(s25);

s25.addNotes(
  "Frothy emotional appeal seldom suffices. A doctor wrote that, about pep talks and\n" +
  "sentiment and getting people fired up. It does not hold.\n\n" +
  "Depth and weight. That is what he says the message needs, and it is why this book reads\n" +
  "the way it does rather than like a motivational pamphlet.\n\n" +
  "And there is Step Two, stated by a physician in the front matter: a power greater than\n" +
  "themselves. Re-create their lives - not improve, not manage. Re-create."
);

// =====================================================================
// SLIDE 26 — The effect produced by alcohol (pp. xxiv-xxv)
// =====================================================================
const s26 = pres.addSlide();
s26.background = { color: "FFFFFF" };
header(s26, "The Effect Produced by Alcohol", "Restless, irritable and discontented.",
  { fill: GREEN, title: "THE BIG BOOK", sub: "pp. xxiv–xxv", subColor: GREEN_SOFT });

s26.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 1.55, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s26.addText(
  [
    { text: "Men and women drink essentially because they " },
    { text: "like the effect produced by alcohol", options: { bold: true, color: GREEN } },
    { text: ". The sensation is so elusive that, while they admit it is injurious, they cannot after " +
            "a time differentiate the true from the false." },
  ],
  { x: 1.0, y: 1.6, w: 11.33, h: 1.55, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s26.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.35, w: 12.13, h: 1.95, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s26.addText(
  [
    { text: "To them, their alcoholic life seems the only normal one. They are " },
    { text: "restless, irritable and discontented", options: { bold: true, color: BLUE } },
    { text: ", unless they can again experience the sense of ease and comfort which comes at once by " +
            "taking a few drinks—drinks which they see others taking with impunity." },
  ],
  { x: 1.0, y: 3.35, w: 11.33, h: 1.95, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s26, 5.6, "“The Doctor’s Opinion,” pp. xxiv–xxv.");
addFooter(s26);

s26.addNotes(
  "Not because of the wife, the job, the childhood. Because they like the effect. Start there.\n\n" +
  "Cannot differentiate the true from the false - he is describing a man who no longer knows\n" +
  "which of his own thoughts to trust.\n\n" +
  "Restless, irritable and discontented. Anyone who has been around a while can say that line\n" +
  "from memory, and it is in a doctor's letter before page 1.\n\n" +
  "And the last clause is the whole trouble: drinks which they see others taking with impunity."
);

// =====================================================================
// SLIDE 27 — An entire psychic change (p. xxv)
// =====================================================================
const s27 = pres.addSlide();
s27.background = { color: "FFFFFF" };
header(s27, "An Entire Psychic Change", "Or very little hope of recovery.",
  { fill: BLUE, title: "THE BIG BOOK", sub: "p. xxv", subColor: BLUE_SOFT });

s27.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.7, w: 12.13, h: 2.25, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s27.addText(
  [
    { text: "After they have succumbed to the desire again, as so many do, and the phenomenon of craving " +
            "develops, they pass through the well-known stages of a spree, emerging remorseful, with a " +
            "firm resolution not to drink again. This is repeated over and over, and unless this person " +
            "can experience " },
    { text: "an entire psychic change", options: { bold: true, color: BLUE } },
    { text: " there is very little hope of his recovery." },
  ],
  { x: 1.0, y: 1.7, w: 11.33, h: 2.25, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s27, 4.25, 1.05, "Psychic = thinking.");
cite(s27, 5.5, "“The Doctor’s Opinion,” p. xxv.");
addFooter(s27);

s27.addNotes(
  "SHOW SHEET here - that is the star in the margin.\n\n" +
  "Psychic change means a change in thinking. Circled three times on this page in the study\n" +
  "copy, so it is worth saying out loud: not mystical, not spooky. Thinking.\n\n" +
  "The cycle first: succumb, craving develops, the spree runs its stages, remorse, firm\n" +
  "resolution. Repeated over and over. Everyone in this room has run that lap.\n\n" +
  "The resolution is not the way out. An entire change in thinking is. And on the next page\n" +
  "Silkworth says that once it happens, the man who seemed doomed finds he can follow a few\n" +
  "simple rules."
);

// =====================================================================
// SLIDE 28 — Doctor, I cannot go on like this (p. xxv)
// =====================================================================
const s28 = pres.addSlide();
s28.background = { color: "FFFFFF" };
header(s28, "Doctor, I Cannot Go On Like This", "Where medicine runs out.",
  { fill: MAGENTA, title: "THE “MUSTS”", sub: "p. xxv", subColor: MAGENTA_SOFT });

const xxvMusts = [
  { y: 1.6, h: 1.4, size: 19,
    runs: [
      { text: "Men have cried out to me in sincere and despairing appeal: “Doctor, I cannot go on like " +
              "this! I have everything to live for! " },
      { text: "I must stop, but I cannot!", options: { bold: true, color: MAGENTA } },
      { text: " You must help me!”" },
    ] },
  { y: 3.2, h: 1.2, size: 19,
    runs: [
      { text: "Faced with this problem, if a doctor is honest with himself, " },
      { text: "he must sometimes feel his own inadequacy", options: { bold: true, color: MAGENTA } },
      { text: "." },
    ] },
  { y: 4.6, h: 1.2, size: 19,
    runs: [
      { text: "…" },
      { text: "we physicians must admit we have made little impression upon the problem as a whole",
        options: { bold: true, color: MAGENTA } },
      { text: "." },
    ] },
];

xxvMusts.forEach((q, i) => {
  s28.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: q.y, w: 12.13, h: q.h, fill: { color: MAGENTA_TINT }, rectRadius: 0.08,
  });
  s28.addShape(pres.ShapeType.ellipse, {
    x: 1.0, y: q.y + q.h / 2 - 0.28, w: 0.56, h: 0.56, fill: { color: MAGENTA },
  });
  s28.addText(String(i + 1), {
    x: 1.0, y: q.y + q.h / 2 - 0.28, w: 0.56, h: 0.56,
    fontFace: HEAD, fontSize: 20, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
  });
  s28.addText(q.runs, {
    x: 1.85, y: q.y, w: 10.5, h: q.h,
    fontFace: BODY, fontSize: q.size, color: NAVY, valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
});

cite(s28, 5.95, "“The Doctor’s Opinion,” p. xxv.");
addFooter(s28);

s28.addNotes(
  "Three musts, and all three are the doctor admitting he is beaten.\n\n" +
  "First, the patient: I must stop, but I cannot. That sentence is Step One in a man's own\n" +
  "words, said to a doctor, years before there were Steps.\n\n" +
  "Second, the physician: if he is honest with himself, he must sometimes feel his own\n" +
  "inadequacy. He gives all that is in him and it often is not enough.\n\n" +
  "Third: we physicians must admit we have made little impression upon the problem as a whole.\n\n" +
  "That is the man who signed this letter, with a hospital and thousands of cases behind him,\n" +
  "saying medicine could not do it. Something more than human power is needed."
);

// =====================================================================
// SLIDE 29 — The five types (p. xxvi)
// =====================================================================
const s29 = pres.addSlide();
s29.background = { color: "FFFFFF" };
header(s29, "The Classification of Alcoholics", "…seems most difficult, and in much detail is outside the scope of this book.",
  { fill: NAVY, title: "THE BIG BOOK", sub: "p. xxvi", subColor: ON_NAVY });

const types = [
  { n: "TYPE I", y: 1.45, h: 1.0, dark: GREEN, tint: GREEN_TINT,
    text: "There are, of course, the psychopaths who are emotionally unstable. We are all familiar " +
          "with this type. They are always “going on the wagon for keeps.” They are over-remorseful " +
          "and make many resolutions, but never a decision." },
  { n: "TYPE II", y: 2.52, h: 0.86, dark: BLUE, tint: BLUE_TINT,
    text: "There is the type of man who is unwilling to admit that he cannot take a drink. He plans " +
          "various ways of drinking. He changes his brand or his environment." },
  { n: "TYPE III", y: 3.45, h: 0.86, dark: GREEN, tint: GREEN_TINT,
    text: "There is the type who always believes that after being entirely free from alcohol for a " +
          "period of time he can take a drink without danger." },
  { n: "TYPE IV", y: 4.38, h: 0.86, dark: BLUE, tint: BLUE_TINT,
    text: "There is the manic-depressive type, who is, perhaps, the least understood by his friends, " +
          "and about whom a whole chapter could be written." },
  { n: "TYPE V", y: 5.31, h: 0.86, dark: GREEN, tint: GREEN_TINT,
    text: "Then there are types entirely normal in every respect except in the effect alcohol has " +
          "upon them. They are often able, intelligent, friendly people." },
];

types.forEach((t) => {
  s29.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: t.y, w: 12.13, h: t.h, fill: { color: t.tint }, rectRadius: 0.08,
  });
  s29.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: t.y, w: 1.85, h: t.h, fill: { color: t.dark }, rectRadius: 0.08,
  });
  s29.addText(t.n, {
    x: 0.6, y: t.y, w: 1.85, h: t.h,
    fontFace: HEAD, fontSize: 16, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });
  s29.addText(t.text, {
    x: 2.68, y: t.y, w: 10.05, h: t.h,
    fontFace: BODY, fontSize: 16, color: NAVY, valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
});

cite(s29, 6.3, "“The Doctor’s Opinion,” p. xxvi.");
addFooter(s29);

s29.addNotes(
  "Five types, and the doctor is not being unkind about any of them. He is showing that the\n" +
  "illness does not care what sort of person you are.\n\n" +
  "Type I - over-remorseful, always going on the wagon for keeps. Many resolutions, but never\n" +
  "a decision. Sit on that difference.\n\n" +
  "Type II - unwilling to admit he cannot take a drink. Changes his brand, changes his\n" +
  "environment. Anything but the one admission.\n\n" +
  "Type III - believes that after enough time off he can drink safely again.\n\n" +
  "Type IV - the manic-depressive, least understood by his friends.\n\n" +
  "Type V - entirely normal in every respect except one. Able, intelligent, friendly people.\n" +
  "That one is on the list so nobody in the room gets to disqualify himself for looking fine."
);

// =====================================================================
// SLIDE 30 — One symptom in common (p. xxvi)
// =====================================================================
const s30 = pres.addSlide();
s30.background = { color: "FFFFFF" };
header(s30, "One Symptom in Common", "What every type shares.",
  { fill: BLUE, title: "THE BIG BOOK", sub: "p. xxvi", subColor: BLUE_SOFT });

s30.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.7, w: 12.13, h: 2.2, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s30.addText(
  [
    { text: "All these, and many others, have " },
    { text: "one symptom in common", options: { bold: true, color: BLUE } },
    { text: ": they cannot start drinking without developing the phenomenon of craving. This " +
            "phenomenon, as we have suggested, may be the manifestation of an allergy which " +
            "differentiates these people, and " },
    { text: "sets them apart as a distinct entity", options: { bold: true, color: BLUE } },
    { text: ". It has never been, by any treatment with which we are familiar, " },
    { text: "permanently eradicated", options: { bold: true, color: BLUE } },
    { text: "." },
  ],
  { x: 1.0, y: 1.7, w: 11.33, h: 2.2, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s30.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 4.2, w: 12.13, h: 1.4, fill: { color: BLUE }, rectRadius: 0.08,
});
s30.addText("IMPORTANT", {
  x: 1.0, y: 4.35, w: 11.33, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: BLUE_SOFT, margin: 0, valign: "middle", charSpacing: 2,
});
s30.addText(
  [
    { text: "The only relief we have to suggest is " },
    { text: "entire abstinence", options: { underline: { style: "dbl" } } },
    { text: "." },
  ],
  { x: 1.0, y: 4.67, w: 11.33, h: 0.75,
    fontFace: HEAD, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "middle" }
);

cite(s30, 5.85, "“The Doctor’s Opinion,” p. xxvi.");
addFooter(s30);

s30.addNotes(
  "This is the paragraph the margin marks IMPORTANT, and it is the one that ties the five\n" +
  "types together.\n\n" +
  "Whatever type a man is, one symptom is common to all of them: he cannot start drinking\n" +
  "without developing the phenomenon of craving. Cannot start.\n\n" +
  "Sets them apart as a distinct entity - not worse people, a different category.\n\n" +
  "Never permanently eradicated, by any treatment he knows of. That word is doing a lot of\n" +
  "work. Not cured. Arrested, at best.\n\n" +
  "So the only relief he has to suggest is entire abstinence - double underlined in the study\n" +
  "copy. Entire. Not moderation, not management. And relief, not cure."
);

// =====================================================================
// A card that carries Will's margin note in its right gutter. Used from
// Bill's Story on, where nearly every highlight is tagged with a "level".
// =====================================================================
function markedCard(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: o.y, w: 12.13, h: o.h, fill: { color: o.tint }, rectRadius: 0.08,
  });
  s.addText(o.runs, {
    x: 1.0, y: o.y, w: o.note ? 8.55 : 11.33, h: o.h,
    fontFace: BODY, fontSize: o.size || 17, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });
  if (o.note) {
    s.addText(o.note, {
      x: 9.75, y: o.y, w: 2.6, h: o.h,
      fontFace: BODY, fontSize: 12.5, bold: true, color: o.dark,
      align: "right", valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
  }
}

// =====================================================================
// SLIDE 31 — Close of The Doctor's Opinion
// =====================================================================
const s31 = pres.addSlide();
s31.background = { color: "FFFFFF" };
header(s31, "He May Remain to Pray", "How the doctor closes his letter.",
  { fill: BLUE, title: "THE BIG BOOK", sub: "The Doctor’s Opinion", subColor: BLUE_SOFT });

s31.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.8, w: 12.13, h: 2.1, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s31.addText(
  "I earnestly advise every alcoholic to read this book through, and though perhaps he came to " +
  "scoff, he may remain to pray.",
  { x: 1.0, y: 1.8, w: 11.33, h: 2.1, fontFace: BODY, fontSize: 26, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);
s31.addText("William D. Silkworth, M.D.", {
  x: 1.0, y: 4.05, w: 11.33, h: 0.45,
  fontFace: HEAD, fontSize: 20, bold: true, color: BLUE, align: "right", valign: "middle", margin: 0,
});

marginNote(s31, 4.75, 1.1,
  "Some people are alcoholics on their first drink. Others crossed the line.");
cite(s31, 6.05, "close of “The Doctor’s Opinion.”");
addFooter(s31);

s31.addNotes(
  "The last line of the letter, and he is still being careful. He does not promise. He advises.\n\n" +
  "Came to scoff, may remain to pray. He knows exactly who is picking up this book and in\n" +
  "what frame of mind.\n\n" +
  "NOTE: this slide was built from dictation, not from a photograph of the page. Check the\n" +
  "wording and add the page number before presenting."
);

// =====================================================================
// SLIDE 32 — Bill's Story (chapter opener)
// =====================================================================
const s32 = pres.addSlide();
s32.background = { color: NAVY };

s32.addText("Bill’s Story", {
  x: 0.9, y: 0.95, w: 11.53, h: 0.9,
  fontFace: HEAD, fontSize: 46, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});
s32.addText("Chapter 1   ·   Page 1", {
  x: 0.9, y: 1.85, w: 11.53, h: 0.4,
  fontFace: BODY, fontSize: 18, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});

s32.addShape(pres.ShapeType.roundRect, {
  x: 0.9, y: 2.75, w: 11.53, h: 1.15, fill: { color: "FFFFFF" }, rectRadius: 0.1,
});
s32.addText(
  [
    { text: "Identify — " },
    { text: "don’t", options: { underline: { style: "sng" } } },
    { text: " compare." },
  ],
  { x: 1.3, y: 2.75, w: 10.73, h: 1.15,
    fontFace: HEAD, fontSize: 30, bold: true, color: NAVY, valign: "middle", margin: 0 }
);

s32.addShape(pres.ShapeType.roundRect, {
  x: 0.9, y: 4.2, w: 11.53, h: 1.7, fill: { color: "2A3B54" }, rectRadius: 0.1,
});
s32.addText("MARGIN NOTE", {
  x: 1.3, y: 4.38, w: 10.73, h: 0.28,
  fontFace: BODY, fontSize: 11, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});
s32.addText(
  "In A.A. the word “bottom” is used a lot, but the only true bottom is death. What there is, " +
  "is different levels of alcoholism that we can get out at.",
  { x: 1.3, y: 4.7, w: 10.73, h: 1.05,
    fontFace: BODY, fontSize: 19, color: "FFFFFF", valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s32.addNotes(
  "Two instructions before we read a word of it.\n\n" +
  "Identify, do not compare. The moment you start measuring your drinking against Bill's,\n" +
  "you are looking for a reason this does not apply to you.\n\n" +
  "And the bottom. We say it constantly - he had to hit bottom. The only true bottom is\n" +
  "death. Everything above that is a level, and every level has a door out of it.\n\n" +
  "Watch for those levels as we read. They are marked all the way down this chapter."
);

// =====================================================================
// SLIDE 33 — Ominous warning (p. 1)
// =====================================================================
const s33 = pres.addSlide();
s33.background = { color: "FFFFFF" };
header(s33, "Ominous Warning", "Which I failed to heed.",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 1", subColor: BLUE_SOFT });

s33.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.55, w: 12.13, h: 2.15, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
s33.addText(
  [
    { text: "“Here lies a Hampshire Grenadier", options: { breakLine: true } },
    { text: "Who caught his death", options: { breakLine: true } },
    { text: "Drinking cold small beer.", options: { breakLine: true } },
    { text: "A good soldier is ne’er forgot", options: { breakLine: true } },
    { text: "Whether he dieth by musket", options: { breakLine: true } },
    { text: "Or by " },
    { text: "pot", options: { bold: true, color: BLUE } },
    { text: ".”" },
  ],
  { x: 1.0, y: 1.55, w: 8.0, h: 2.15, fontFace: HEAD, fontSize: 17, italic: true, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.1 }
);
s33.addText(
  [
    { text: "“POT”", options: { fontSize: 11, bold: true, color: MUTED, charSpacing: 2, breakLine: true } },
    { text: "A pail — a vessel of drink.", options: { fontSize: 15, color: NAVY } },
  ],
  { x: 9.3, y: 1.55, w: 3.05, h: 2.15, fontFace: BODY, align: "right", valign: "middle", margin: 0 }
);

s33.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.95, w: 12.13, h: 2.0, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s33.addText(
  [
    { text: "Ominous warning—which I failed to heed.", options: { bold: true, color: BLUE } },
    { text: " Twenty-two, and a veteran of foreign wars, I went home at last. I fancied myself a " +
            "leader, for had not the men of my battery given me a special token of appreciation? " +
            "My talent for leadership, I imagined, would place me at the head of vast enterprises " +
            "which I would manage with the utmost assurance." },
  ],
  { x: 1.0, y: 3.95, w: 11.33, h: 2.0, fontFace: BODY, fontSize: 18, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s33, 6.15, "“Bill’s Story,” p. 1.");
addFooter(s33);

s33.addNotes(
  "Read all of this one - the arrow in the margin says so.\n\n" +
  "The epitaph is real, from Winchester Cathedral. Pot means a drinking vessel, a pail.\n" +
  "A soldier who survived the muskets and died of the drink.\n\n" +
  "Bill reads his own obituary on a tombstone at twenty-two and calls it an ominous warning\n" +
  "he failed to heed. Then in the very next breath: leader, special token of appreciation,\n" +
  "vast enterprises, utmost assurance.\n\n" +
  "The warning and the ego, back to back, in one paragraph."
);

// =====================================================================
// SLIDE 34 — The drive for success (p. 2)
// =====================================================================
const s34 = pres.addSlide();
s34.background = { color: "FFFFFF" };
header(s34, "The Drive for Success", "Wall Street, and a law course.",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 2", subColor: BLUE_SOFT });

[
  { y: 1.6, h: 1.15, tint: BLUE_TINT, dark: BLUE, size: 19,
    runs: [{ text: "The drive for success was on. My work took me about Wall Street and little by " +
                   "little I became interested in the market." }] },
  { y: 2.95, h: 1.15, tint: GREEN_TINT, dark: GREEN, size: 19,
    runs: [{ text: "I nearly failed my law course. At one of the finals I was too drunk to think " +
                   "or write." }] },
  { y: 4.3, h: 1.55, tint: GREEN_TINT, dark: GREEN, size: 19,
    runs: [{ text: "Though my drinking was not yet continuous, it disturbed my wife. We had long " +
                   "talks when I would still her forebodings by telling her that " },
           { text: "men of genius conceived their best projects when drunk", options: { bold: true, color: GREEN } },
           { text: ";" }] },
].forEach((c) => markedCard(s34, c));

cite(s34, 6.15, "“Bill’s Story,” p. 2.");
addFooter(s34);

s34.addNotes(
  "Too drunk to think or write at a final - and he nearly failed the course he was paying for\n" +
  "with the job that drinking would eventually cost him.\n\n" +
  "And then the excuse, which is the part to sit on: men of genius conceived their best\n" +
  "projects when drunk. He is not lying to her. He believes it.\n\n" +
  "Everyone in this room has had a version of that sentence."
);

// =====================================================================
// SLIDE 35 — The weapon (p. 2)
// =====================================================================
const s35 = pres.addSlide();
s35.background = { color: "FFFFFF" };
header(s35, "Forging the Weapon", "Drink and speculation.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 2", subColor: GREEN_SOFT });

[
  { y: 1.55, h: 1.1, tint: BLUE_TINT, dark: BLUE, note: "Level —\nquitting projects",
    runs: [{ text: "By the time I had completed the course, I knew the law was not for me. The " +
                   "inviting maelstrom of Wall Street had me in its grip." }] },
  { y: 2.75, h: 1.35, tint: GREEN_TINT, dark: GREEN,
    runs: [{ text: "Out of this alloy of drink and speculation, I commenced to " },
           { text: "forge the weapon that one day would turn in its flight like a boomerang and all " +
                   "but cut me to ribbons", options: { bold: true, color: GREEN } },
           { text: "." }] },
  { y: 4.2, h: 0.9, tint: BLUE_TINT, dark: BLUE, note: "Some people’s\nyearly income",
    runs: [{ text: "I saved $1,000.", options: { bold: true } }] },
  { y: 5.2, h: 1.05, tint: BLUE_TINT, dark: BLUE, note: "First to write\na prospectus",
    runs: [{ text: "I failed to persuade my broker friends to send me out looking over factories " +
                   "and managements." }] },
].forEach((c) => markedCard(s35, c));

cite(s35, 6.4, "“Bill’s Story,” p. 2.");
addFooter(s35);

s35.addNotes(
  "Quitting the law after finishing the course - that is a level. Not a catastrophe yet.\n" +
  "Just a man walking away from the thing he trained for.\n\n" +
  "The boomerang is his own image for it, written years later. Drink and speculation, alloyed\n" +
  "into a weapon that comes back around.\n\n" +
  "A thousand dollars in the early twenties was a year's income for a lot of people. And the\n" +
  "idea nobody would fund - going out to look at the factories himself - is the one he was\n" +
  "right about. He was among the first to write what we would now call a prospectus."
);

// =====================================================================
// SLIDE 36 — I had arrived (p. 3)
// =====================================================================
const s36 = pres.addSlide();
s36.background = { color: "FFFFFF" };
header(s36, "I Had Arrived", "The levels, one after another.",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 3", subColor: BLUE_SOFT });

[
  { y: 1.5, h: 0.9, tint: BLUE_TINT, dark: BLUE, note: "9 people’s\nyearly income",
    runs: [{ text: "The exercise of an option brought in more money, leaving us with a profit of " +
                   "several thousand dollars for that year." }] },
  { y: 2.5, h: 0.8, tint: BLUE_TINT, dark: BLUE, note: "Egotistic", size: 24,
    runs: [{ text: "I had arrived.", options: { bold: true } }] },
  { y: 3.4, h: 0.85, tint: BLUE_TINT, dark: BLUE, note: "Another level",
    runs: [{ text: "My drinking assumed more serious proportions, continuing all day and almost " +
                   "every night." }] },
  { y: 4.35, h: 0.95, tint: BLUE_TINT, dark: BLUE, note: "Another level —\ncarousing",
    runs: [{ text: "There had been no real infidelity, for loyalty to my wife, helped at times by " +
                   "extreme drunkenness, kept me out of those scrapes." }] },
  { y: 5.4, h: 0.9, tint: BLUE_TINT, dark: BLUE, note: "Another level",
    runs: [{ text: "Liquor caught up with me much faster than I came up behind Walter. I began to " +
                   "be jittery in the morning." }] },
].forEach((c) => markedCard(s36, c));

cite(s36, 6.45, "“Bill’s Story,” p. 3.");
addFooter(s36);

s36.addNotes(
  "Three words: I had arrived. That is the top of the arc, and the book gives it its own\n" +
  "sentence for a reason.\n\n" +
  "Then read the levels coming down off it. All day and almost every night. Kept out of the\n" +
  "scrapes only by being too drunk for them - and notice he counts that as loyalty.\n\n" +
  "Jittery in the morning. That is the body speaking now, not the ego."
);

// =====================================================================
// SLIDE 37 — I would not jump (p. 4)
// =====================================================================
const s37 = pres.addSlide();
s37.background = { color: "FFFFFF" };
header(s37, "I Would Not Jump", "October 1929.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 4", subColor: GREEN_SOFT });

[
  { y: 1.7, h: 1.4, tint: GREEN_TINT, dark: GREEN, note: "Insanity", size: 19,
    runs: [{ text: "The papers reported men jumping to death from the towers of High Finance. " +
                   "That disgusted me. " },
           { text: "I would not jump. I went back to the bar.", options: { bold: true, color: GREEN } }] },
  { y: 3.25, h: 1.3, tint: BLUE_TINT, dark: BLUE, note: "Another level", size: 19,
    runs: [{ text: "Next morning I telephoned a friend in Montreal. He had plenty of money left " +
                   "and thought I had better go to Canada." }] },
].forEach((c) => markedCard(s37, c));

marginNote(s37, 4.85, 1.05, "Level — sponging off of family.");
cite(s37, 6.1, "“Bill’s Story,” p. 4.");
addFooter(s37);

s37.addNotes(
  "Men are going out of windows and his reaction is disgust - and then the bar. He notices\n" +
  "the insanity in other people and walks straight into his own.\n\n" +
  "Montreal is the next level: somebody else still has money, so go there.\n\n" +
  "And when that ends, they move in with his wife's parents. Sponging off family. Nobody\n" +
  "calls it that at the time. It is always temporary, always somebody being generous."
);

// =====================================================================
// SLIDE 38 — I still thought I could control it (p. 5)
// =====================================================================
const s38 = pres.addSlide();
s38.background = { color: "FFFFFF" };
header(s38, "I Still Thought I Could Control It", "Liquor ceased to be a luxury.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 5", subColor: GREEN_SOFT });

[
  { y: 1.7, h: 1.0, tint: GREEN_TINT, dark: GREEN, note: "Another level", size: 20,
    runs: [{ text: "Liquor ceased to be a luxury; it became a necessity.", options: { bold: true } }] },
  { y: 2.85, h: 1.3, tint: GREEN_TINT, dark: GREEN, note: "Another level", size: 19,
    runs: [{ text: "Nevertheless, I still thought I could " },
           { text: "control", options: { bold: true, underline: { style: "dbl" } } },
           { text: " the situation, and there were periods of sobriety which renewed my wife’s hope." }] },
  { y: 4.3, h: 1.0, tint: BLUE_TINT, dark: BLUE, note: "Blinded to life", size: 20,
    runs: [{ text: "Gradually things got worse.", options: { bold: true } }] },
].forEach((c) => markedCard(s38, c));

cite(s38, 5.55, "“Bill’s Story,” p. 5.");
addFooter(s38);

s38.addNotes(
  "Luxury to necessity. One clause, and a whole line has been crossed.\n\n" +
  "Control is double-underlined here, the same as on page xxii where the book says we could\n" +
  "not control our drinking. Same word, both places.\n\n" +
  "And notice what keeps the thing alive: periods of sobriety which renewed my wife's hope.\n" +
  "The dry stretches are not evidence of control. They are what makes everyone keep believing.\n\n" +
  "Gradually things got worse. Gradually. Nobody sees it happening, which is the note in the\n" +
  "margin - blinded to life."
);

// =====================================================================
// SLIDE 39 — Was I crazy? (p. 5)
// =====================================================================
const s39 = pres.addSlide();
s39.background = { color: "FFFFFF" };
header(s39, "Was I Crazy?", "Will power, and what it was worth.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 5", subColor: GREEN_SOFT });

[
  { y: 1.6, h: 1.75, tint: GREEN_TINT, dark: GREEN, note: "Will\nPower", size: 19,
    runs: [{ text: "I woke up. This had to be stopped. I saw I could not take so much as one drink. " },
           { text: "I was through forever.", options: { bold: true, color: GREEN } },
           { text: " Before then, I had written lots of sweet promises, but my wife happily observed " +
                   "that this time I meant business. And so I did." }] },
  { y: 3.55, h: 1.05, tint: BLUE_TINT, dark: BLUE, size: 20,
    runs: [{ text: "Shortly afterward I came home drunk. There had been no fight.", options: { bold: true } }] },
  { y: 4.8, h: 1.05, tint: GREEN_TINT, dark: GREEN, note: "Level", size: 26,
    runs: [{ text: "Was I crazy?", options: { bold: true, underline: { style: "sng" } } }] },
].forEach((c) => markedCard(s39, c));

cite(s39, 6.05, "“Bill’s Story,” p. 5.");
addFooter(s39);

s39.addNotes(
  "This is the will power level, and it is the honest one. He is not lying, he is not\n" +
  "half-hearted, he is not making a sweet promise. He means it, and his wife can tell the\n" +
  "difference. And so I did - he actually stopped.\n\n" +
  "Then: shortly afterward I came home drunk. There had been no fight.\n\n" +
  "No fight. No reason. No trigger to point at. That is the whole argument for the physical\n" +
  "factor, told as a story instead of a theory.\n\n" +
  "Was I crazy? He underlines it. He is asking the question the Doctor's Opinion already\n" +
  "answered - the body is as abnormal as the mind."
);

// A compact navy card that resolves a circled word on the page.
function idCard(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText(o.label, {
    x: o.x + 0.32, y: o.y + 0.18, w: o.w - 0.64, h: 0.26,
    fontFace: BODY, fontSize: 10.5, bold: true, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
  });
  s.addText(o.name, {
    x: o.x + 0.32, y: o.y + 0.48, w: o.w - 0.64, h: 0.5,
    fontFace: HEAD, fontSize: 20, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
  });
  if (o.sub) {
    s.addText(o.sub, {
      x: o.x + 0.32, y: o.y + 1.0, w: o.w - 0.64, h: 0.4,
      fontFace: BODY, fontSize: 13, color: ON_NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.1,
    });
  }
}

// =====================================================================
// SLIDE 40 — Gin would fix that (p. 6)
// =====================================================================
const s40 = pres.addSlide();
s40.background = { color: "FFFFFF" };
header(s40, "Gin Would Fix That", "Remorse, horror and hopelessness.",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 6", subColor: BLUE_SOFT });

[
  { y: 1.9, h: 1.3, tint: BLUE_TINT, dark: BLUE, note: "Sickest reason\nto drink", size: 24,
    runs: [{ text: "Gin would fix that. So two bottles, and—oblivion.", options: { bold: true } }] },
  { y: 3.5, h: 1.4, tint: BLUE_TINT, dark: BLUE, size: 20,
    runs: [{ text: "Sometimes I stole from my wife’s slender purse when the morning terror and " +
                   "madness were on me." }] },
].forEach((c) => markedCard(s40, c));

cite(s40, 5.2, "“Bill’s Story,” p. 6.");
addFooter(s40);

s40.addNotes(
  "Read what comes right before it - the remorse, the horror, the hopelessness of the next\n" +
  "morning. And the answer to all of that is gin.\n\n" +
  "The sickest reason to drink: not to feel good. To go out. Two bottles and oblivion.\n\n" +
  "Then the purse. A grown man taking coins from his wife's handbag because the terror is\n" +
  "on him. He does not dress it up."
);

// =====================================================================
// SLIDE 41 — Forty pounds under weight (p. 7)
// =====================================================================
const s41 = pres.addSlide();
s41.background = { color: "FFFFFF" };
header(s41, "Forty Pounds Under Weight", "Who is who on this page.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 7", subColor: GREEN_SOFT });

markedCard(s41, {
  y: 1.55, h: 1.0, tint: GREEN_TINT, dark: GREEN, note: "Level", size: 20,
  runs: [{ text: "I could eat little or nothing when drinking, and I was forty pounds under weight.",
           options: { bold: true } }],
});

idCard(s41, { x: 0.6, y: 2.85, w: 3.89, h: 1.6, label: "MY BROTHER-IN-LAW",
  name: "Dr. Leonard Strong", sub: "The physician who got him in." });
idCard(s41, { x: 4.72, y: 2.85, w: 3.89, h: 1.6, label: "A NATIONALLY-KNOWN HOSPITAL",
  name: "Towns", sub: "First visit." });
idCard(s41, { x: 8.84, y: 2.85, w: 3.89, h: 1.6, label: "A DOCTOR",
  name: "Dr. Silkworth", sub: "Who explained I had been seriously ill, bodily and mentally." });

s41.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 4.7, w: 12.13, h: 0.95, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
s41.addText(
  [
    { text: "BELLADONNA TREATMENT   ", options: { fontSize: 11, bold: true, color: MUTED, charSpacing: 2 } },
    { text: "The nightshade family — a hallucinogenic.", options: { fontSize: 17, color: NAVY } },
  ],
  { x: 1.0, y: 4.7, w: 11.33, h: 0.95, fontFace: BODY, valign: "middle", margin: 0 }
);

cite(s41, 5.9, "“Bill’s Story,” p. 7.");
addFooter(s41);

s41.addNotes(
  "Forty pounds under weight. That is not a figure of speech, it is a man starving.\n\n" +
  "Three people get circled on this page and none of them are named in the book.\n" +
  "The brother-in-law is Dr. Leonard Strong. The nationally-known hospital is Towns -\n" +
  "this is the first of the visits. The doctor is Silkworth, the same man who wrote\n" +
  "the letter we just read.\n\n" +
  "And the belladonna treatment - deadly nightshade. They were treating delirium with a\n" +
  "hallucinogen. That was the state of the art in 1934."
);

// =====================================================================
// SLIDE 42 — Surely this was the answer (p. 7)
// =====================================================================
const s42 = pres.addSlide();
s42.background = { color: "FFFFFF" };
header(s42, "Surely This Was the Answer", "Self-knowledge.",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 7", subColor: BLUE_SOFT });

s42.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 2.5, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s42.addText(
  [
    { text: "It relieved me somewhat to learn that in alcoholics the will is amazingly weakened when " +
            "it comes to combating liquor, though it often remains strong in other respects. My " +
            "incredible behavior in the face of a desperate desire to stop was explained. Understanding " +
            "myself now, I fared forth in high hope. For three or four months the goose hung high. I " +
            "went to town regularly and even made a little money. Surely this was the answer—" },
    { text: "self-knowledge", options: { bold: true, color: BLUE, underline: { style: "sng" } } },
    { text: "." },
  ],
  { x: 1.0, y: 1.6, w: 11.33, h: 2.5, fontFace: BODY, fontSize: 18, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

markedCard(s42, {
  y: 4.3, h: 1.15, tint: BLUE_TINT, dark: BLUE, note: "Second visit —\nTowns", size: 28,
  runs: [{ text: "But it was not.", options: { bold: true, underline: { style: "sng" } } }],
});

cite(s42, 5.7, "“Bill’s Story,” p. 7.");
addFooter(s42);

s42.addNotes(
  "Three or four months. The goose hung high. He is going to work, making money, and he\n" +
  "understands his own condition for the first time in his life.\n\n" +
  "Surely this was the answer - self-knowledge. Underlined, because it is the trap.\n\n" +
  "But it was not. Four words, underlined, on its own.\n\n" +
  "Knowing exactly what is wrong with you does not fix what is wrong with you. He goes back\n" +
  "to the same hospital a second time. This is the argument against 'if I just understood\n" +
  "myself better,' and Bill made it with his own life before the book ever made it in print."
);

// =====================================================================
// SLIDE 43 — Alcohol was my master (p. 8)
// =====================================================================
const s43 = pres.addSlide();
s43.background = { color: "FFFFFF" };
header(s43, "Alcohol Was My Master", "No words can tell.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 8", subColor: GREEN_SOFT });

s43.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.8, w: 12.13, h: 2.4, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s43.addText(
  [
    { text: "No words can tell of the loneliness and despair I found in that bitter morass of " +
            "self-pity. Quicksand stretched around me in all directions. I had met my match. I had " +
            "been overwhelmed. " },
    { text: "Alcohol was my master.", options: { bold: true, color: GREEN } },
  ],
  { x: 1.0, y: 1.8, w: 11.33, h: 2.4, fontFace: BODY, fontSize: 22, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s43, 4.5, 1.1, "The beginning of Step One — Bill’s.");
cite(s43, 5.85, "“Bill’s Story,” p. 8.");
addFooter(s43);

s43.addNotes(
  "This is where Step One starts for him, and notice it is not a decision. It is a report.\n\n" +
  "I had met my match. I had been overwhelmed. Alcohol was my master. Three flat statements\n" +
  "of fact, past tense.\n\n" +
  "Nobody talks him into this. He arrives at it in a bitter morass of self-pity with\n" +
  "quicksand in all directions - which is exactly where most people arrive at it."
);

// =====================================================================
// SLIDE 44 — Ebby Thatcher (p. 9)
// =====================================================================
const s44 = pres.addSlide();
s44.background = { color: "FFFFFF" };
header(s44, "He Was Sober", "An old school friend comes to the door.",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 9", subColor: BLUE_SOFT });

s44.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 12.13, h: 1.35, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
s44.addText(
  "Rumor had it that he had been committed for alcoholic insanity. I wondered how he had escaped.",
  { x: 1.0, y: 1.6, w: 11.33, h: 1.35, fontFace: BODY, fontSize: 20, color: NAVY,
    underline: { style: "sng" }, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

s44.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 3.15, w: 12.13, h: 1.35, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s44.addText("“I’ve got religion.”", {
  x: 1.0, y: 3.15, w: 11.33, h: 1.35,
  fontFace: HEAD, fontSize: 30, bold: true, color: BLUE, valign: "middle", margin: 0,
});

marginNote(s44, 4.8, 1.1, "Ebby Thatcher — Bill’s sponsor.");
cite(s44, 6.15, "“Bill’s Story,” p. 9.");
addFooter(s44);

s44.addNotes(
  "The man at the door is Ebby Thatcher - the alcoholic friend from the Foreword, the one\n" +
  "Roland carried it to. Bill's sponsor, before the word existed.\n\n" +
  "Bill had heard he was locked up for alcoholic insanity. And here he is at the door, sober,\n" +
  "fresh-skinned and glowing, refusing a drink.\n\n" +
  "I've got religion. Bill is aghast. Watch what he does with that in the next paragraph -\n" +
  "he pours himself a drink and lets the old boy rant."
);

// =====================================================================
// SLIDE 45 — Idea and action (p. 9)
// =====================================================================
const s45 = pres.addSlide();
s45.background = { color: "FFFFFF" };
header(s45, "An Idea and a Program", "What they had told him.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 9", subColor: GREEN_SOFT });

[
  { y: 1.9, h: 1.5, tint: GREEN_TINT, dark: GREEN, note: "STEP 2", size: 24,
    runs: [{ text: "They had told of a simple religious idea…", options: { bold: true } }] },
  { y: 3.6, h: 1.5, tint: BLUE_TINT, dark: BLUE, note: "STEPS 3 THRU 12", size: 24,
    runs: [{ text: "…and a practical program of action.", options: { bold: true } }] },
].forEach((c) => markedCard(s45, c));

cite(s45, 5.4, "“Bill’s Story,” p. 9.");
addFooter(s45);

s45.addNotes(
  "One sentence, and the whole program is in it.\n\n" +
  "A simple religious idea - that is Step Two. Not a doctrine, not a denomination. An idea.\n\n" +
  "And a practical program of action - Steps Three through Twelve. Practical. Action.\n\n" +
  "The idea without the action is a conversation. The action without the idea is willpower,\n" +
  "and we already watched that fail on page five."
);

// =====================================================================
// SLIDE 46 — I was hopeless (p. 10)
// =====================================================================
const s46 = pres.addSlide();
s46.background = { color: "FFFFFF" };
header(s46, "I Had to Be", "Certainly I was interested.",
  { fill: NAVY, title: "BILL’S STORY", sub: "p. 10", subColor: ON_NAVY });

s46.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 2.2, w: 12.13, h: 1.8, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
s46.addText(
  [
    { text: "Certainly I was interested. " },
    { text: "I had to be, for I was hopeless.", options: { bold: true, color: NAVY } },
  ],
  { x: 1.0, y: 2.2, w: 11.33, h: 1.8, fontFace: BODY, fontSize: 28, color: MUTED,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s46, 4.3, 1.1, "Bill completes Step One.");
cite(s46, 5.65, "“Bill’s Story,” p. 10.");
addFooter(s46);

s46.addNotes(
  "Circled in the study copy, and rightly. This is where Step One closes for him.\n\n" +
  "He is not interested because it sounds good. He is interested because he has run out of\n" +
  "everything else. I had to be, for I was hopeless.\n\n" +
  "Hopeless is not despair here. It is an accurate description of his position, and it is\n" +
  "the only thing that got him to listen to a man he thought was a crackpot."
);

// =====================================================================
// SLIDE 47 — The solution (p. 11)
// =====================================================================
const s47 = pres.addSlide();
s47.background = { color: "FFFFFF" };
header(s47, "The Solution", "My friend sat before me.",
  { fill: GREEN, title: "BILL’S STORY", sub: "p. 11", subColor: GREEN_SOFT });

s47.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.7, w: 12.13, h: 3.2, fill: { color: GREEN_TINT }, rectRadius: 0.08,
});
s47.addText(
  [
    { text: "But my friend sat before me, and he made the point-blank declaration that " },
    { text: "God had done for him what he could not do for himself", options: { bold: true, color: GREEN } },
    { text: ". His human will had failed. Doctors had pronounced him incurable. Society was about to " +
            "lock him up. Like myself, he had admitted complete defeat. Then he had, in effect, been " +
            "raised from the dead, suddenly taken from the scrap heap to a level of life better than " +
            "the best he had ever known!" },
  ],
  { x: 1.0, y: 1.7, w: 11.33, h: 3.2, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s47, 5.2, "“Bill’s Story,” p. 11.");
addFooter(s47);

s47.addNotes(
  "The margin marks this whole paragraph The Solution, and that is what it is.\n\n" +
  "Count what had already failed: his human will, the doctors, and society, which was about\n" +
  "to lock him up. Everything available had been tried and had lost.\n\n" +
  "Like myself, he had admitted complete defeat. That is the entry requirement, stated plainly.\n\n" +
  "And what follows defeat is not managing better. Raised from the dead. Scrap heap to a level\n" +
  "of life better than the best he had ever known. Better than his best, not back to average."
);

// =====================================================================
// SLIDE 48 — None at all (p. 11)
// =====================================================================
const s48 = pres.addSlide();
s48.background = { color: "FFFFFF" };
header(s48, "This Was None at All", "Had this power originated in him?",
  { fill: BLUE, title: "BILL’S STORY", sub: "p. 11", subColor: BLUE_SOFT });

s48.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 2.0, w: 12.13, h: 2.5, fill: { color: BLUE_TINT }, rectRadius: 0.08,
});
s48.addText(
  [
    { text: "Had this power originated in him? Obviously it had not. There had been no more power " +
            "in him than there was in me at that minute; and " },
    { text: "this was none at all", options: { bold: true, color: BLUE } },
    { text: "." },
  ],
  { x: 1.0, y: 2.0, w: 11.33, h: 2.5, fontFace: BODY, fontSize: 24, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

cite(s48, 4.8, "“Bill’s Story,” p. 11.");
addFooter(s48);

s48.addNotes(
  "Bill checks the obvious explanation first. Maybe Ebby just found something in himself.\n\n" +
  "Obviously it had not. He had known Ebby for years. There was no more power in that man\n" +
  "than there was in Bill sitting across the table - and Bill knew exactly how much that was.\n\n" +
  "None at all.\n\n" +
  "So whatever did it came from somewhere else. That is the reasoning, and he does it cold,\n" +
  "as a skeptic, at his own kitchen table."
);

// =====================================================================
// SLIDE 49 — A miracle across the kitchen table (p. 11)
// =====================================================================
const s49 = pres.addSlide();
s49.background = { color: "FFFFFF" };
header(s49, "Across the Kitchen Table", "That floored me.",
  { fill: NAVY, title: "BILL’S STORY", sub: "p. 11", subColor: ON_NAVY });

s49.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.7, w: 12.13, h: 2.6, fill: { color: "F1F4F7" }, rectRadius: 0.08,
});
s49.addText(
  [
    { text: "That floored me. It began to look as though religious people were right after all. " +
            "Here was something at work in a human heart which had done the impossible. My ideas " +
            "about miracles were drastically revised right then. Never mind the musty past; " },
    { text: "here sat a miracle directly across the kitchen table", options: { bold: true, color: NAVY } },
    { text: ". He shouted great tidings." },
  ],
  { x: 1.0, y: 1.7, w: 11.33, h: 2.6, fontFace: BODY, fontSize: 20, color: NAVY,
    valign: "middle", margin: 0, lineSpacingMultiple: 1.15 }
);

marginNote(s49, 4.6, 1.1, "Bill sees a miracle in front of him.");
cite(s49, 5.95, "“Bill’s Story,” p. 11.");
addFooter(s49);

s49.addNotes(
  "Nothing is highlighted in this paragraph, but the margin says why it is here: Bill sees a\n" +
  "miracle in front of him.\n\n" +
  "Never mind the musty past. He is not being asked to believe something that happened two\n" +
  "thousand years ago to somebody he never met. The evidence is sitting across the kitchen\n" +
  "table drinking coffee.\n\n" +
  "That is still how it works. Nobody gets argued into this. They watch somebody they knew\n" +
  "when he was drinking, and they cannot account for what they are looking at."
);

pres.writeFile({ fileName: __dirname + "/aa-big-book-study.pptx" })
  .then((f) => console.log("wrote " + f));
