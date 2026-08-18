const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

const NAVY = "1B2A41";
const ON_NAVY = "A9B8CC";
const MUTED = "6B7A8C";
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

// ---------- Slide 1: title ----------
const s1 = pres.addSlide();
s1.background = { color: NAVY };
s1.addText("Big Book Study", {
  x: 0.9, y: 2.15, w: 11.53, h: 1.1,
  fontFace: HEAD, fontSize: 54, bold: true, color: "FFFFFF", margin: 0, valign: "middle",
});
s1.addText("Presented by A.A. District 23", {
  x: 0.9, y: 3.3, w: 11.53, h: 0.5,
  fontFace: BODY, fontSize: 22, color: ON_NAVY, margin: 0, valign: "middle", charSpacing: 2,
});
s1.addShape(pres.ShapeType.roundRect, {
  x: 0.9, y: 4.3, w: 5.2, h: 1.15, fill: { color: "2A3B54" }, rectRadius: 0.1,
});
s1.addText(
  [
    { text: "903 Court Street", options: { fontSize: 20, bold: true, color: "FFFFFF", breakLine: true } },
    { text: "Wednesdays  ·  7:00 PM", options: { fontSize: 15, color: ON_NAVY } },
  ],
  { x: 1.25, y: 4.3, w: 4.5, h: 1.15, fontFace: BODY, valign: "middle", margin: 0 }
);
s1.addNotes(
  "Leave this up while people settle in.\n\n" +
  "Big Book Study, District 23, 903 Court Street, Wednesdays at 7."
);

// ---------- Slide 2: required copyright notice ----------
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
s2.addNotes("Required by A.A.W.S. whenever this literature is displayed.");

// ---------- Spreads ----------
// Reads pages.json (from prep.py). Verso pages go left, recto right.
const MANIFEST = path.join(__dirname, "pages.json");
let spreads = 0;
if (fs.existsSync(MANIFEST)) {
  const pages = JSON.parse(fs.readFileSync(MANIFEST, "utf8"));

  // Pair them: a verso followed by a recto is one spread. Unpaired pages
  // stand alone, centered, rather than being silently dropped.
  const pairs = [];
  for (let i = 0; i < pages.length; ) {
    const a = pages[i], b = pages[i + 1];
    if (a && b && a.side === "verso" && b.side === "recto") { pairs.push([a, b]); i += 2; }
    else if (a && b && a.side !== "recto" && b.side !== "verso") { pairs.push([a, b]); i += 2; }
    else { pairs.push([a, null]); i += 1; }
  }

  const TOP = 0.35, BOT = 6.85;           // image band
  const BAND = BOT - TOP;

  pairs.forEach((pair) => {
    const s = pres.addSlide();
    s.background = { color: "FFFFFF" };
    const imgs = pair.filter(Boolean);
    // Scale both pages to a common height, butt them at the spine.
    const widths = imgs.map((p) => (p.w / p.h) * BAND);
    const total = widths.reduce((a, b) => a + b, 0);
    let scale = 1;
    if (total > 12.5) scale = 12.5 / total;
    const h = BAND * scale;
    const y = TOP + (BAND - h) / 2;
    let x = (13.33 - total * scale) / 2;
    imgs.forEach((p, i) => {
      const w = widths[i] * scale;
      s.addImage({ path: path.join(__dirname, "pages", p.file), x, y, w, h });
      x += w;
    });
    s.addText(FOOTER, {
      x: 0.6, y: 6.98, w: 12.13, h: 0.3,
      fontFace: BODY, fontSize: 9, color: MUTED, valign: "middle", margin: 0,
    });
    spreads++;
  });
}

pres.writeFile({ fileName: path.join(__dirname, "aa-big-book-spreads.pptx") })
  .then((f) => console.log("wrote " + f + "  (" + spreads + " spreads)"));
