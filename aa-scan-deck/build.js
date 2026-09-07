const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const NAVY = "1B2A41";
const ON_NAVY = "A9B8CC";
const MUTED = "6B7A8C";
const FOLIO = "9AAABC";
const MAGENTA = "9E1B60";
const HEAD = "Cambria";
const BODY = "Calibri";

const DISCLAIMER =
  "“These materials are copyright © by Alcoholics Anonymous World Services, Inc. (\"A.A.W.S.\"). " +
  "All rights reserved. Individual printing or photocopying of a single copy is permitted. " +
  "Intended for personal use only, and not to be reproduced further or distributed for resale.”";

const FOOTER =
  "Copyright © Alcoholics Anonymous World Services, Inc. All rights reserved. " +
  "Intended for personal use only, and not to be reproduced further or distributed for resale.";


const SECTIONS = [
  { key: "1-front-matter", label: "Front Matter through Page 16",
    src: "5215354a-08182026_ALCOHOLIC_S_ANONYMOUS.pdf" },
  { key: "2-there-is-a-solution", label: "There Is a Solution — Pages 17 to 72",
    src: "43630d84-08182026_THERE_IS__SOLUTION_D_Njj__t7_i5uri_Cr__ft_VG_Lt_6TLi.pdf" },
  { key: "3-into-action", label: "Into Action — Pages 73 to 140",
    src: "231419ad-08182026___INTO_ACTION_73_invariably_they_got_drunk._Having_per.pdf" },
  { key: "4-to-employers", label: "To Employers — Pages 141 to the End",
    src: "6c79fd5d-08182026_TO_EMPLOYERS_141_normal_will_do_incredible_things._Aft.pdf" },
];

function pairUp(pages) {
  const pairs = [];
  for (let i = 0; i < pages.length; ) {
    const a = pages[i], b = pages[i + 1];
    if (a.side === "verso" && b && b.side === "recto") { pairs.push([a, b]); i += 2; }
    else { pairs.push([a]); i += 1; }
  }
  return pairs;
}


// Chris Zimmer's placement list, from his notebook. Each entry becomes a slide
// inserted straight after the spread carrying that page. Images are dropped in
// by hand - see README.
const PHOTOS = [
  { at: "xi",   title: "Rowland Hazard",            kind: "photo",   note: "Photograph, with a brief history." },
  { at: "xii",  title: "The Oxford Group",          kind: "photo",   note: "Photograph." },
  { at: "xiii", title: "A.A. Number Three",         kind: "photo",   note: "Bill D. \u2014 the man in the bed, Akron City Hospital, 1935." },
  { at: "xvii", title: "Clarence Snyder",           kind: "photo",   note: "Photograph." },
  { at: "xxii", title: "William D. Silkworth, M.D.",kind: "photo",   note: "Photograph." },
  { at: "xxiii",title: "Towns Hospital",            kind: "photo",   note: "293 Central Park West, New York City." },
  { at: "xxv",  title: "Alcohol Metabolism",        kind: "diagram", note: "Diagram." },
  { at: "1",    title: "Bill W.",                   kind: "photo",   note: "Photograph." },
  { at: "1",    title: "The Tombstone",             kind: "photo",   note: "Thomas Thetcher, Winchester Cathedral \u2014 \u201cOr by pot.\u201d" },
  { at: "7",    title: "Dr. Leonard Strong, M.D.",  kind: "photo",   note: "Bill\u2019s brother-in-law." },
  { at: "9",    title: "Ebby Thatcher",             kind: "photo",   note: "Bill\u2019s sponsor." },
  { at: "17",   title: "The Fellowship",            kind: "diagram", note: "Diagram." },
  { at: "26",   title: "Dr. Carl Jung",             kind: "photo",   note: "Photograph." },
  { at: "28",   title: "William James",             kind: "photo",   note: "Photograph." },
  { at: "63",   title: "Handout Sheet",             kind: "handout", note: "Handout." },
  { at: "64",   title: "Inventory Handouts",        kind: "handout", note: "Handout." },
  { at: "86",   title: "Eleventh Step Inventory",   kind: "handout", note: "Handout." },
  { at: "136",  title: "Hank Parkhurst",            kind: "photo",   note: "Chapter 10, To Employers." },
  { at: "165",  title: "Dr. Bob",                   kind: "photo",   note: "Photograph." },
];

function addPlaceholder(pres, s, item, ref) {
  s.background = { color: "FFFFFF" };
  s.addText(item.title, {
    x: 0.6, y: 0.42, w: 9.0, h: 0.7,
    fontFace: HEAD, fontSize: 38, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s.addText(ref, {
    x: 0.6, y: 1.08, w: 9.0, h: 0.4,
    fontFace: BODY, fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "middle",
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 3.05, y: 1.75, w: 7.2, h: 4.45,
    fill: { color: "F7F9FB" }, rectRadius: 0.08,
    line: { color: "AEBBC9", width: 1.5, dashType: "dash" },
  });
  s.addText(
    [
      { text: item.kind.toUpperCase(), options: { fontSize: 11, bold: true, color: MUTED, charSpacing: 2, breakLine: true } },
      { text: "\n", options: { fontSize: 8, breakLine: true } },
      { text: item.note, options: { fontSize: 17, color: NAVY, breakLine: true } },
      { text: "\n", options: { fontSize: 8, breakLine: true } },
      { text: "Drop the image into this frame in PowerPoint.", options: { fontSize: 13, italic: true, color: MUTED } },
    ],
    { x: 3.45, y: 1.75, w: 6.4, h: 4.45, fontFace: BODY, align: "center", valign: "middle", margin: 0 }
  );
  s.addText("Credit:", {
    x: 3.05, y: 6.35, w: 7.2, h: 0.35,
    fontFace: BODY, fontSize: 12, color: MUTED, margin: 0, valign: "middle",
  });
}

function buildDeck(pairs, outName, subtitle) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5

  // ---- title ----
  const s1 = pres.addSlide();
  s1.background = { color: NAVY };
  s1.addText("Big Book Study Group", {
    x: 0.9, y: 2.05, w: 11.53, h: 1.1,
    fontFace: HEAD, fontSize: 52, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
  });
  s1.addText("by District 23", {
    x: 0.9, y: 3.18, w: 11.53, h: 0.45,
    fontFace: BODY, fontSize: 20, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
  });
  if (subtitle) {
    s1.addText(subtitle, {
      x: 0.9, y: 3.68, w: 11.53, h: 0.45,
      fontFace: BODY, fontSize: 16, italic: true, color: "7E92AC", margin: 0, valign: "middle",
    });
  }
  s1.addShape(pres.ShapeType.roundRect, {
    x: 0.9, y: 4.45, w: 6.0, h: 1.15, fill: { color: "2A3B54" }, rectRadius: 0.1,
  });
  s1.addText(
    [
      { text: "903 Court Street, Port Huron", options: { fontSize: 20, bold: true, color: "FFFFFF", breakLine: true } },
      { text: "Wednesdays  \u00b7  7:00 PM", options: { fontSize: 15, color: ON_NAVY } },
    ],
    { x: 1.25, y: 4.45, w: 5.3, h: 1.15, fontFace: BODY, valign: "middle", margin: 0 }
  );

  // ---- copyright notice ----
  const s2 = pres.addSlide();
  s2.background = { color: "FFFFFF" };
  s2.addText("Copyright Notice", {
    x: 0.6, y: 0.9, w: 12.13, h: 0.7,
    fontFace: HEAD, fontSize: 36, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s2.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 1.9, w: 12.13, h: 2.2, fill: { color: "F1F4F7" }, rectRadius: 0.1,
  });
  s2.addText(DISCLAIMER, {
    x: 1.0, y: 1.9, w: 11.33, h: 2.2,
    fontFace: BODY, fontSize: 19, italic: true, color: NAVY, margin: 0, valign: "middle",
    lineSpacingMultiple: 1.2,
  });
  s2.addText(
    "Displayed at the request of A.A. World Services, Inc. Reproduced with permission for use " +
    "during this A.A. meeting.",
    { x: 0.6, y: 4.4, w: 12.13, h: 0.7, fontFace: BODY, fontSize: 14, color: MUTED,
      margin: 0, valign: "top", lineSpacingMultiple: 1.2 }
  );

  // ---- spreads: pages run the full height of the slide ----
  const TOP = 0.06, BOT = 7.44, BAND = BOT - TOP;
  const SW = 13.333;
  pairs.forEach((pair) => {
    const s = pres.addSlide();
    s.background = { color: "FFFFFF" };
    const widths = pair.map((p) => (p.w / p.h) * BAND);
    const total = widths.reduce((a, b) => a + b, 0);
    const scale = total > 11.4 ? 11.4 / total : 1;   // keep room for the margins
    const h = BAND * scale;
    const y = TOP + (BAND - h) / 2;
    let x = (SW - total * scale) / 2;
    const x0 = x;
    pair.forEach((p, i) => {
      const w = widths[i] * scale;
      s.addImage({ path: path.join(__dirname, process.env.PAGES_DIR || "pages", p.file), x, y, w, h });
      x += w;
    });
    const x1 = x;

    // Printed folios, out in the margins where a book sets them.
    const left = pair[0], right = pair.length > 1 ? pair[1] : null;
    if (left && left.label && left.side === "verso") {
      s.addText(left.label, {
        x: x0 - 1.05, y: y + h / 2 - 0.35, w: 0.85, h: 0.7,
        fontFace: HEAD, fontSize: 28, bold: true, color: FOLIO,
        align: "right", valign: "middle", margin: 0,
      });
    }
    const outer = right || (left && left.side === "recto" ? left : null);
    if (outer && outer.label) {
      s.addText(outer.label, {
        x: x1 + 0.2, y: y + h / 2 - 0.35, w: 0.85, h: 0.7,
        fontFace: HEAD, fontSize: 28, bold: true, color: FOLIO,
        align: "left", valign: "middle", margin: 0,
      });
    }

    // Copyright, set on its side up the left edge. The box is defined
    // horizontally and rotated about its centre, so x is deliberately
    // negative - after the 270 turn it lands in the left margin.
    const CW = 7.2, CH = 0.46, CX = 0.32, CY = 3.75;
    s.addText(FOOTER, {
      x: CX - CW / 2, y: CY - CH / 2, w: CW, h: CH,
      rotate: 270,
      fontFace: BODY, fontSize: 7, color: MUTED,
      align: "center", valign: "middle", margin: 0,
    });

    // Zimmer's photo slides, straight after the spread carrying that page.
    pair.forEach((p) => {
      PHOTOS.filter((it) => it.at === p.label).forEach((it) => {
        const ps = pres.addSlide();
        const roman = !/^[0-9]+$/.test(p.label);
        addPlaceholder(pres, ps, it, (roman ? "Page " : "Page ") + p.label);
      });
    });
  });

  return pres.writeFile({ fileName: path.join(__dirname, outName) })
    .then(() => console.log("  " + outName + "  (" + pairs.length + " spreads)"));
}

const pages = JSON.parse(fs.readFileSync(path.join(__dirname, "pages.json"), "utf8"));

(async () => {
  if (process.env.COMBINED) {
    await buildDeck(pairUp(pages), "aa-big-book-spreads.pptx", null);
  } else {
    for (const sec of SECTIONS) {
      const sub = pages.filter((p) => p.source === sec.src);
      if (!sub.length) { console.log("  (no pages for " + sec.key + ")"); continue; }
      await buildDeck(pairUp(sub), "big-book-" + sec.key + ".pptx", sec.label);
    }
  }
})();
