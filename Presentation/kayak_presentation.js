const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "JEDHA — Bloc 1";
pres.title = "Projet Kayak";

// Color palette
const C = {
  navy:    "0A2342",
  teal:    "0D7377",
  ltTeal:  "14BDAC",
  lightBg: "EEF7F7",
  white:   "FFFFFF",
  ltGray:  "F4F6F8",
  dark:    "1A1A2E",
  orange:  "F4731C",
  gray:    "888888",
  midGray: "CCCCCC",
};

const makeShadow = () => ({
  type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.12
});

const makeLightShadow = () => ({
  type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.08
});

// ─────────────────────────────────────────────
// SLIDE 1 — Title
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Decorative circle top-right
  s.addShape(pres.shapes.OVAL, {
    x: 8.2, y: -0.9, w: 2.6, h: 2.6,
    fill: { color: C.ltTeal, transparency: 75 },
    line: { color: C.ltTeal, transparency: 75 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 8.8, y: -0.4, w: 1.6, h: 1.6,
    fill: { color: C.ltTeal, transparency: 55 },
    line: { color: C.ltTeal, transparency: 55 },
  });

  // Bottom decorative bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.7, w: 10, h: 0.925,
    fill: { color: C.teal, transparency: 65 },
    line: { color: C.teal, transparency: 65 },
  });

  // Left accent stripe
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.3, w: 0.07, h: 1.8,
    fill: { color: C.ltTeal },
    line: { color: C.ltTeal },
  });

  // Main title
  s.addText("Projet Kayak", {
    x: 0.7, y: 1.2, w: 8.5, h: 1.1,
    fontSize: 52, fontFace: "Georgia", bold: true, color: C.white,
    align: "left", valign: "middle", margin: 0,
  });

  // Subtitle
  s.addText("Recommandation de destinations basée sur la météo et les hôtels", {
    x: 0.7, y: 2.4, w: 8.0, h: 0.7,
    fontSize: 20, fontFace: "Calibri", color: C.ltTeal,
    align: "left", valign: "middle", margin: 0,
  });

  // Tag bottom-left (inside the bar)
  s.addText("JEDHA  —  Bloc 1  —  Data Collection & Management", {
    x: 0.5, y: 4.78, w: 7, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: C.white,
    transparency: 30, align: "left", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — Problématique
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Problématique", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  const cards = [
    { icon: "🎯", title: "Objectif",  text: "Identifier les meilleures destinations kayak en France selon la météo et les hôtels disponibles" },
    { icon: "📊", title: "Données",   text: "Combiner APIs météo, géocodage et scraping Booking en un pipeline automatisé" },
    { icon: "🗺️", title: "Livrable",  text: "Un dataset exploitable, une base SQL et des cartes interactives" },
  ];

  const cardW = 2.8, cardH = 3.0, gap = 0.25;
  const startX = (10 - 3 * cardW - 2 * gap) / 2;
  const cardY = 1.15;

  cards.forEach((card, i) => {
    const cx = startX + i * (cardW + gap);

    // Card background
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: C.white },
      line: { color: C.midGray, width: 0.5 },
      shadow: makeLightShadow(),
    });

    // Top teal banner
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.75,
      fill: { color: C.teal },
      line: { color: C.teal },
    });

    // Icon circle
    s.addShape(pres.shapes.OVAL, {
      x: cx + cardW / 2 - 0.32, y: cardY + 0.38, w: 0.64, h: 0.64,
      fill: { color: C.ltTeal },
      line: { color: C.ltTeal },
    });

    // Icon text
    s.addText(card.icon, {
      x: cx + cardW / 2 - 0.32, y: cardY + 0.38, w: 0.64, h: 0.64,
      fontSize: 18, align: "center", valign: "middle", margin: 0,
    });

    // Card title
    s.addText(card.title, {
      x: cx + 0.15, y: cardY + 0.85, w: cardW - 0.3, h: 0.45,
      fontSize: 18, fontFace: "Georgia", bold: true, color: C.navy,
      align: "center", valign: "middle", margin: 0,
    });

    // Card text
    s.addText(card.text, {
      x: cx + 0.2, y: cardY + 1.4, w: cardW - 0.4, h: 1.4,
      fontSize: 13, fontFace: "Calibri", color: C.dark,
      align: "center", valign: "top", wrap: true, margin: 0,
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — Architecture du pipeline
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightBg };

  s.addText("Architecture du pipeline", {
    x: 0.5, y: 0.25, w: 9, h: 0.65,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  const steps = [
    { name: "Nominatim",     sub: "35 villes" },
    { name: "OpenWeather",   sub: "5 jours" },
    { name: "Score météo",   sub: "Top 5" },
    { name: "Booking",       sub: "50 hôtels" },
    { name: "Neon SQL",      sub: "50 lignes" },
    { name: "Cartes",        sub: "2 cartes" },
  ];

  const boxW = 1.28, boxH = 0.72, arrowW = 0.22;
  const totalW = steps.length * boxW + (steps.length - 1) * arrowW;
  const startX = (10 - totalW) / 2;
  const boxY = 1.9;

  steps.forEach((step, i) => {
    const bx = startX + i * (boxW + arrowW);

    // Number badge
    s.addShape(pres.shapes.OVAL, {
      x: bx + boxW / 2 - 0.18, y: boxY - 0.5, w: 0.36, h: 0.36,
      fill: { color: C.ltTeal },
      line: { color: C.ltTeal },
    });
    s.addText(String(i + 1), {
      x: bx + boxW / 2 - 0.18, y: boxY - 0.5, w: 0.36, h: 0.36,
      fontSize: 11, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });

    // Main box
    s.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: boxY, w: boxW, h: boxH,
      fill: { color: C.teal },
      line: { color: C.teal },
      shadow: makeLightShadow(),
    });

    s.addText(step.name, {
      x: bx, y: boxY, w: boxW, h: boxH,
      fontSize: 12, fontFace: "Calibri", bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });

    // Sub-label
    s.addText(step.sub, {
      x: bx, y: boxY + boxH + 0.08, w: boxW, h: 0.28,
      fontSize: 10, fontFace: "Calibri", color: C.gray,
      align: "center", valign: "top", margin: 0,
    });

    // Arrow
    if (i < steps.length - 1) {
      const ax = bx + boxW;
      s.addShape(pres.shapes.LINE, {
        x: ax, y: boxY + boxH / 2, w: arrowW, h: 0,
        line: { color: C.teal, width: 2 },
      });
      // Arrowhead triangle approximation using text "▶"
      s.addText("▶", {
        x: ax + arrowW - 0.12, y: boxY + boxH / 2 - 0.13, w: 0.2, h: 0.26,
        fontSize: 10, color: C.teal, align: "left", valign: "middle", margin: 0,
      });
    }
  });

  // Description below
  s.addText("Chaque étape collecte, transforme ou stocke les données du projet.", {
    x: 1, y: 3.6, w: 8, h: 0.45,
    fontSize: 13, fontFace: "Calibri", color: C.gray, italic: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Flow summary box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.15, w: 8.8, h: 0.9,
    fill: { color: C.navy },
    line: { color: C.navy },
    shadow: makeLightShadow(),
  });
  s.addText("Collecte des coordonnées  →  Météo  →  Classement  →  Hôtels  →  Stockage  →  Visualisation", {
    x: 0.6, y: 4.15, w: 8.8, h: 0.9,
    fontSize: 13, fontFace: "Calibri", color: C.ltTeal,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — Géocodage — Nominatim
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Géocodage — Nominatim", {
    x: 0.5, y: 0.25, w: 9, h: 0.65,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  // Left column background
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.5, h: 3.9,
    fill: { color: C.lightBg },
    line: { color: C.lightBg },
  });

  // Left accent
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 0.07, h: 3.9,
    fill: { color: C.ltTeal },
    line: { color: C.ltTeal },
  });

  s.addText("À propos de l'API", {
    x: 0.65, y: 1.2, w: 4.1, h: 0.38,
    fontSize: 15, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  s.addText([
    { text: "API OpenStreetMap — gratuite, pas de clé requise", options: { bullet: true, breakLine: true } },
    { text: "35 villes françaises géocodées", options: { bullet: true, breakLine: true } },
    { text: "Résultat : latitude + longitude par ville", options: { bullet: true, breakLine: true } },
    { text: "Politique : 1 requête/seconde max (time.sleep)", options: { bullet: true } },
  ], {
    x: 0.65, y: 1.65, w: 4.1, h: 2.9,
    fontSize: 14, fontFace: "Calibri", color: C.dark,
    align: "left", valign: "top",
  });

  // Right column table
  const tableData = [
    [
      { text: "Ville",       options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Latitude",    options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Longitude",   options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
    ],
    [
      { text: "Avignon",    options: { fill: { color: C.white }, align: "center" } },
      { text: "43.95",      options: { fill: { color: C.white }, align: "center" } },
      { text: "4.81",       options: { fill: { color: C.white }, align: "center" } },
    ],
    [
      { text: "Marseille",  options: { fill: { color: C.lightBg }, align: "center" } },
      { text: "43.30",      options: { fill: { color: C.lightBg }, align: "center" } },
      { text: "5.38",       options: { fill: { color: C.lightBg }, align: "center" } },
    ],
    [
      { text: "Strasbourg", options: { fill: { color: C.white }, align: "center" } },
      { text: "48.58",      options: { fill: { color: C.white }, align: "center" } },
      { text: "7.75",       options: { fill: { color: C.white }, align: "center" } },
    ],
    [
      { text: "Biarritz",   options: { fill: { color: C.lightBg }, align: "center" } },
      { text: "43.48",      options: { fill: { color: C.lightBg }, align: "center" } },
      { text: "-1.56",      options: { fill: { color: C.lightBg }, align: "center" } },
    ],
    [
      { text: "Annecy",     options: { fill: { color: C.white }, align: "center" } },
      { text: "45.90",      options: { fill: { color: C.white }, align: "center" } },
      { text: "6.13",       options: { fill: { color: C.white }, align: "center" } },
    ],
  ];

  s.addTable(tableData, {
    x: 5.3, y: 1.1, w: 4.2, h: 2.8,
    fontSize: 13, fontFace: "Calibri",
    border: { pt: 0.5, color: C.midGray },
    colW: [1.7, 1.25, 1.25],
  });

  s.addText("Extrait des 35 villes géocodées", {
    x: 5.3, y: 4.0, w: 4.2, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: C.gray, italic: true,
    align: "center", margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — Score météo
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightBg };

  s.addText("Météo & Score", {
    x: 0.5, y: 0.2, w: 9, h: 0.65,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  s.addText("OpenWeather  /forecast  —  plan gratuit  —  prévisions 5 jours / 3h", {
    x: 0.5, y: 0.9, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.teal,
    align: "left", valign: "middle", margin: 0,
  });

  // Formula box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.45, w: 8.4, h: 0.78,
    fill: { color: C.navy },
    line: { color: C.navy },
    shadow: makeShadow(),
  });
  s.addText("Score = 100  −  |temp−24|×2  −  pop×30  −  pluie×1.5  −  vent×0.8", {
    x: 0.8, y: 1.45, w: 8.4, h: 0.78,
    fontSize: 14, fontFace: "Courier New", color: C.white,
    align: "center", valign: "middle", margin: 0,
  });

  // Metric cards
  const metrics = [
    { label: "Température",       detail: "−2 pts/°C\nd’écart à 24°C" },
    { label: "Probabilité pluie", detail: "−30 pts\nsi 100%" },
    { label: "Pluie totale",      detail: "−1.5 pt/mm" },
    { label: "Vent",              detail: "−0.8 pt/m·s⁻¹" },
  ];

  const mW = 2.0, mH = 1.7, mGap = 0.2;
  const mStart = (10 - 4 * mW - 3 * mGap) / 2;
  const mY = 2.55;

  metrics.forEach((m, i) => {
    const mx = mStart + i * (mW + mGap);

    s.addShape(pres.shapes.RECTANGLE, {
      x: mx, y: mY, w: mW, h: mH,
      fill: { color: C.white },
      line: { color: C.midGray, width: 0.5 },
      shadow: makeLightShadow(),
    });

    // Top teal accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: mx, y: mY, w: mW, h: 0.08,
      fill: { color: C.ltTeal },
      line: { color: C.ltTeal },
    });

    s.addText(m.label, {
      x: mx + 0.1, y: mY + 0.15, w: mW - 0.2, h: 0.45,
      fontSize: 13, fontFace: "Georgia", bold: true, color: C.navy,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText(m.detail, {
      x: mx + 0.1, y: mY + 0.65, w: mW - 0.2, h: 0.85,
      fontSize: 14, fontFace: "Calibri", color: C.teal, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — Top 5 destinations
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Top 5 destinations", {
    x: 0.5, y: 0.2, w: 9, h: 0.65,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  const rows = [
    ["Rang", "Ville", "Score", "Temp.", "Pluie", "Vent"],
    ["🥇  1", "Bayonne",         "89.37", "19.7°C", "3.2 mm", "3.1 m/s"],
    ["🥈  2", "Aix-en-Provence", "89.30", "20.0°C", "2.8 mm", "2.9 m/s"],
    ["🥉  3", "Montauban",       "88.60", "19.1°C", "3.0 mm", "3.3 m/s"],
    ["4",           "Nîmes",          "88.56", "19.6°C", "3.1 mm", "3.2 m/s"],
    ["5",           "Avignon",             "88.30", "19.5°C", "3.0 mm", "3.1 m/s"],
  ];

  const tableData = rows.map((row, ri) => {
    const isHeader = ri === 0;
    const isEven   = ri % 2 === 0 && !isHeader;
    const bg = isHeader ? C.navy : (isEven ? C.lightBg : C.white);
    const fg = isHeader ? C.white : C.dark;

    return row.map((cell, ci) => ({
      text: cell,
      options: {
        fill: { color: bg },
        color: fg,
        bold: isHeader,
        align: ci === 1 ? "left" : "center",
        fontSize: isHeader ? 13 : 14,
      },
    }));
  });

  s.addTable(tableData, {
    x: 0.6, y: 1.1, w: 8.8, h: 3.8,
    fontFace: "Calibri",
    border: { pt: 0.3, color: C.midGray },
    colW: [1.1, 2.1, 1.2, 1.4, 1.4, 1.6],
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — Scraping Booking
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightBg };

  s.addText("Scraping Booking — ScrapingBee", {
    x: 0.5, y: 0.2, w: 9, h: 0.65,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  // ── Left column ──
  const lx = 0.4, lw = 4.6, colY = 1.05;

  // Problem label
  s.addText("Problème", {
    x: lx, y: colY, w: lw, h: 0.3,
    fontSize: 13, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText("Booking bloque les requêtes Python directes", {
    x: lx, y: colY + 0.32, w: lw, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.dark, italic: true,
    align: "left", valign: "middle", margin: 0,
  });

  // Solution box
  s.addShape(pres.shapes.RECTANGLE, {
    x: lx, y: colY + 0.85, w: lw, h: 1.0,
    fill: { color: C.teal },
    line: { color: C.teal },
    shadow: makeLightShadow(),
  });
  s.addText("ScrapingBee agit comme un vrai navigateur avec proxy furtif + rendu JavaScript", {
    x: lx + 0.15, y: colY + 0.85, w: lw - 0.3, h: 1.0,
    fontSize: 13, fontFace: "Calibri", color: C.white,
    align: "left", valign: "middle", wrap: true,
  });

  // Strategy label
  s.addText("Stratégie", {
    x: lx, y: colY + 2.0, w: lw, h: 0.3,
    fontSize: 13, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });
  s.addText("4 variantes d'URL testées par ville — on garde la première valide", {
    x: lx, y: colY + 2.32, w: lw, h: 0.5,
    fontSize: 13, fontFace: "Calibri", color: C.dark,
    align: "left", valign: "top", wrap: true, margin: 0,
  });

  // ── Right column ──
  const rx = 5.4, rw = 4.2;

  // Big stat
  s.addText("50", {
    x: rx, y: colY, w: rw, h: 1.0,
    fontSize: 72, fontFace: "Calibri", bold: true, color: C.teal,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("hôtels extraits", {
    x: rx, y: colY + 0.95, w: rw, h: 0.35,
    fontSize: 16, fontFace: "Calibri", color: C.navy,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("10 par ville × 5 villes", {
    x: rx, y: colY + 1.33, w: rw, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.gray,
    align: "center", valign: "middle", margin: 0,
  });

  // Separator
  s.addShape(pres.shapes.LINE, {
    x: rx + 0.5, y: colY + 1.78, w: rw - 1, h: 0,
    line: { color: C.midGray, width: 1 },
  });

  // Code snippet
  s.addShape(pres.shapes.RECTANGLE, {
    x: rx + 0.15, y: colY + 1.98, w: rw - 0.3, h: 0.5,
    fill: { color: C.ltGray },
    line: { color: C.midGray, width: 0.5 },
  });
  s.addText("data-testid='property-card'", {
    x: rx + 0.25, y: colY + 1.98, w: rw - 0.5, h: 0.5,
    fontSize: 11, fontFace: "Courier New", color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — Stack technique
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("Stack technique", {
    x: 0.5, y: 0.2, w: 9, h: 0.65,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.navy,
    align: "left", valign: "middle", margin: 0,
  });

  const techs = [
    { name: "Nominatim",       desc: "Géocodage gratuit OpenStreetMap" },
    { name: "OpenWeather",     desc: "Prévisions météo 5 jours" },
    { name: "ScrapingBee",     desc: "Scraping JS anti-bot" },
    { name: "BeautifulSoup",   desc: "Parsing HTML" },
    { name: "PostgreSQL Neon", desc: "Base cloud serverless" },
    { name: "Folium",          desc: "Cartes interactives HTML" },
  ];

  const cols = 3, rows = 2;
  const cW = 2.85, cH = 1.55, cGapX = 0.18, cGapY = 0.22;
  const totalW = cols * cW + (cols - 1) * cGapX;
  const startX = (10 - totalW) / 2;
  const startY = 1.15;

  techs.forEach((t, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = startX + col * (cW + cGapX);
    const cy = startY + row * (cH + cGapY);

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cW, h: cH,
      fill: { color: C.white },
      line: { color: C.midGray, width: 0.5 },
      shadow: makeLightShadow(),
    });

    // Left teal border
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 0.07, h: cH,
      fill: { color: C.ltTeal },
      line: { color: C.ltTeal },
    });

    // Number badge
    s.addShape(pres.shapes.OVAL, {
      x: cx + 0.18, y: cy + cH / 2 - 0.22, w: 0.44, h: 0.44,
      fill: { color: C.teal },
      line: { color: C.teal },
    });
    s.addText(String(i + 1), {
      x: cx + 0.18, y: cy + cH / 2 - 0.22, w: 0.44, h: 0.44,
      fontSize: 13, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });

    // Tech name
    s.addText(t.name, {
      x: cx + 0.74, y: cy + 0.2, w: cW - 0.84, h: 0.42,
      fontSize: 16, fontFace: "Georgia", bold: true, color: C.navy,
      align: "left", valign: "middle", margin: 0,
    });

    // Description
    s.addText(t.desc, {
      x: cx + 0.74, y: cy + 0.65, w: cW - 0.84, h: 0.7,
      fontSize: 13, fontFace: "Calibri", color: C.gray,
      align: "left", valign: "top", margin: 0,
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 9 — Conclusion
// ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Decorative circles
  s.addShape(pres.shapes.OVAL, {
    x: -0.5, y: -0.5, w: 2.5, h: 2.5,
    fill: { color: C.teal, transparency: 80 },
    line: { color: C.teal, transparency: 80 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 8.5, y: 3.5, w: 2.2, h: 2.2,
    fill: { color: C.ltTeal, transparency: 85 },
    line: { color: C.ltTeal, transparency: 85 },
  });

  s.addText("Résultats & Bilan", {
    x: 0.5, y: 0.25, w: 9, h: 0.75,
    fontSize: 44, fontFace: "Georgia", bold: true, color: C.white,
    align: "left", valign: "middle", margin: 0,
  });

  // 4 stat cards
  const stats = [
    { num: "35", label: "villes\ngéocodées" },
    { num: "50", label: "hôtels\nscrapés" },
    { num: "2",  label: "cartes\ngénérées" },
    { num: "1",  label: "base SQL\nchargée" },
  ];

  const sW = 2.0, sH = 1.9, sGap = 0.2;
  const sStart = (10 - 4 * sW - 3 * sGap) / 2;
  const sY = 1.2;

  stats.forEach((st, i) => {
    const sx = sStart + i * (sW + sGap);

    s.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: sY, w: sW, h: sH,
      fill: { color: C.white, transparency: 92 },
      line: { color: C.ltTeal, width: 1 },
    });

    s.addText(st.num, {
      x: sx, y: sY + 0.1, w: sW, h: 0.95,
      fontSize: 60, fontFace: "Calibri", bold: true, color: C.ltTeal,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(st.label, {
      x: sx + 0.1, y: sY + 1.08, w: sW - 0.2, h: 0.72,
      fontSize: 14, fontFace: "Calibri", color: C.white,
      align: "center", valign: "top", margin: 0,
    });
  });

  // Pipeline summary line
  s.addText("Pipeline complet  Nominatim → OpenWeather → Booking → Neon → Folium", {
    x: 0.5, y: 3.35, w: 9, h: 0.55,
    fontSize: 16, fontFace: "Calibri", color: C.ltTeal,
    align: "center", valign: "middle", margin: 0,
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.88, w: 10, h: 0.745,
    fill: { color: C.teal, transparency: 50 },
    line: { color: C.teal, transparency: 50 },
  });
  s.addText("JEDHA  —  Bloc 1  —  Data Collection & Management", {
    x: 0.5, y: 4.9, w: 9, h: 0.42,
    fontSize: 13, fontFace: "Calibri", color: C.white,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────
// Write file
// ─────────────────────────────────────────────
pres.writeFile({
  fileName: "C:\\Users\\negi\\Desktop\\JEDHA EXAMEN FINAL\\Dossier FINAL JEDHA CDSD\\Bloc_1 Kayak\\Kayak_Presentation.pptx"
}).then(() => {
  console.log("✅ Kayak_Presentation.pptx created successfully.");
}).catch(err => {
  console.error("❌ Error:", err);
});
