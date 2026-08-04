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

pres.writeFile({ fileName: __dirname + "/aa-big-book-study.pptx" })
  .then((f) => console.log("wrote " + f));
