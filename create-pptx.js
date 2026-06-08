const pptxgen = require("pptxgenjs");
const path = require("path");

const IMG = (name) => path.resolve(__dirname, "images", name);
const FONT = "Inter";
const ACCENT = "B43B2B";
const GOLD = "E8A838";
const DARK = "1A1A2E";
const WHITE = "FFFFFF";
const LGRAY = "F4F0EA";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Ministère de la Santé Publique du Tchad";
pres.title = "Amaan — Dossier Médical Biométrique";

// Helper: full-bleed image slide
function imgSlide(name, overlay) {
  const s = pres.addSlide();
  s.background = { path: IMG(name) };
  if (overlay) {
    s.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: 13.33, h: 7.5,
      fill: { color: DARK, transparency: overlay }
    });
  }
  return s;
}

// =============== SLIDE 1 — COVER ===============
{
  const s = imgSlide("hospital.jpg", 35);
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 7.5,
    fill: { color: DARK, transparency: 50 }
  });

  s.addText("MINISTÈRE DE LA SANTÉ PUBLIQUE DU TCHAD", {
    x: 0.8, y: 0.6, w: 6, h: 0.4,
    fontSize: 11, color: GOLD, fontFace: FONT, bold: true,
    letterSpacing: 2
  });

  s.addText("أَمَان", {
    x: 0.8, y: 1.3, w: 11, h: 1.4,
    fontSize: 96, color: GOLD, fontFace: FONT, bold: true
  });

  s.addText("Amaan", {
    x: 0.8, y: 2.6, w: 11, h: 0.8,
    fontSize: 40, color: WHITE, fontFace: FONT, bold: true
  });

  s.addText("DOSSIER MÉDICAL BIOMÉTRIQUE POUR CHAQUE TCHADIEN", {
    x: 0.8, y: 3.5, w: 10, h: 0.5,
    fontSize: 16, color: WHITE, fontFace: FONT
  });

  s.addText("صِحَّتَك فِي أَمَان — Ta santé en sécurité", {
    x: 0.8, y: 4.1, w: 10, h: 0.4,
    fontSize: 13, color: LGRAY, fontFace: FONT, italic: true
  });

  // Stats row
  const stats = [
    ["17,4M", "Habitants ciblés"],
    ["3%", "Médecins / 10k"],
    ["ODD 3", "Objectif aligné"],
    ["3 ans", "Déploiement"],
    ["3,5M$", "Budget requis"]
  ];
  stats.forEach(([v, l], i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 0.8 + i * 2.4, y: 5.2, w: 2.1, h: 1.6,
      fill: { color: DARK, transparency: 40 },
      line: { color: GOLD, width: 0.5, transparency: 50 },
      rectRadius: 6
    });
    s.addText(v, {
      x: 0.8 + i * 2.4, y: 5.3, w: 2.1, h: 0.7,
      fontSize: 28, color: GOLD, fontFace: FONT, bold: true, align: "center"
    });
    s.addText(l, {
      x: 0.8 + i * 2.4, y: 6.1, w: 2.1, h: 0.5,
      fontSize: 10, color: LGRAY, fontFace: FONT, align: "center"
    });
  });

  s.addText("Présentation aux Bailleurs de Fonds • 2025-2028", {
    x: 0.8, y: 7, w: 8, h: 0.35,
    fontSize: 10, color: LGRAY, fontFace: FONT
  });
}

// =============== SLIDE 2 — LE PROBLÈME ===============
{
  const s = imgSlide("africa.jpg", 25);
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 7.5,
    fill: { color: DARK, transparency: 45 }
  });

  s.addText("02", {
    x: 0.8, y: 0.4, w: 1, h: 0.5,
    fontSize: 14, color: GOLD, fontFace: FONT, bold: true
  });
  s.addText("LE PROBLÈME", {
    x: 0.8, y: 0.8, w: 8, h: 0.8,
    fontSize: 36, color: WHITE, fontFace: FONT, bold: true
  });
  s.addText("60% des carnets de santé papier sont perdus ou détruits", {
    x: 0.8, y: 1.5, w: 8, h: 0.45,
    fontSize: 16, color: GOLD, fontFace: FONT
  });

  // Problem list
  const problems = [
    "Impossibilité de vérifier le statut vaccinal",
    "Aucun suivi longitudinal pour patients nomades",
    "Risque de double traitement ou allergies non détectées",
    "Données épidémiologiques nationales incomplètes",
    "Aucune identité médicale fiable en zones rurales"
  ];
  problems.forEach((p, i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 0.8, y: 2.3 + i * 0.55, w: 6.5, h: 0.42,
      fill: { color: DARK, transparency: 50 },
      rectRadius: 4
    });
    s.addText(`  ✕  ${p}`, {
      x: 0.8, y: 2.3 + i * 0.55, w: 6.5, h: 0.42,
      fontSize: 12, color: WHITE, fontFace: FONT
    });
  });

  // Stats box right side
  s.addShape(pres.ShapeType.rect, {
    x: 8, y: 2, w: 4.5, h: 5,
    fill: { color: DARK, transparency: 30 },
    line: { color: GOLD, width: 1, transparency: 40 },
    rectRadius: 8
  });
  s.addText("INDICATEURS SANITAIRES", {
    x: 8.2, y: 2.2, w: 4.1, h: 0.4,
    fontSize: 10, color: GOLD, fontFace: FONT, bold: true, letterSpacing: 1
  });

  const indics = [
    ["117‰", "Mortalité infantile", "F87171"],
    ["860", "Mortalité maternelle /100k", "F87171"],
    ["45%", "Couverture vaccinale", "E8A838"],
    ["1/27k", "Médecins/habitant", "FB923C"],
    ["22%", "Alphabétisation femmes rurales", "FB923C"]
  ];
  indics.forEach(([v, l, c], i) => {
    s.addText(v, {
      x: 8.3, y: 2.8 + i * 0.7, w: 1.5, h: 0.5,
      fontSize: 22, color: c, fontFace: FONT, bold: true
    });
    s.addText(l, {
      x: 9.8, y: 2.8 + i * 0.7, w: 2.5, h: 0.5,
      fontSize: 11, color: LGRAY, fontFace: FONT
    });
  });
}

// =============== SLIDE 3 — LA SOLUTION ===============
{
  const s = pres.addSlide();
  s.background = { fill: DARK };

  s.addText("03", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: GOLD, fontFace: FONT, bold: true
  });
  s.addText("LA SOLUTION", {
    x: 0.8, y: 0.6, w: 8, h: 0.8,
    fontSize: 36, color: WHITE, fontFace: FONT, bold: true
  });
  s.addText("Dossier Médical Biométrique Numérique", {
    x: 0.8, y: 1.3, w: 8, h: 0.4,
    fontSize: 16, color: GOLD, fontFace: FONT
  });

  // Left: 3-step flow
  const steps = [
    ["1", "Enregistrement", "Carte à puce ou bracelet remis au patient. Son empreinte est numérisée et liée à un identifiant unique national."],
    ["2", "Identification", "L'agent de santé identifie le patient en posant son doigt sur un lecteur biométrique."],
    ["3", "Consultation", "Le dossier médical complet s'affiche instantanément — même hors-ligne."]
  ];
  steps.forEach(([n, t, d], i) => {
    const y = 2.1 + i * 1.3;
    s.addShape(pres.ShapeType.rect, {
      x: 0.8, y, w: 0.7, h: 0.7,
      fill: { color: ACCENT },
      rectRadius: 35
    });
    s.addText(n, {
      x: 0.8, y, w: 0.7, h: 0.7,
      fontSize: 24, color: WHITE, fontFace: FONT, bold: true, align: "center", valign: "middle"
    });
    s.addText(t, {
      x: 1.7, y, w: 4.8, h: 0.35,
      fontSize: 14, color: WHITE, fontFace: FONT, bold: true
    });
    s.addText(d, {
      x: 1.7, y: y + 0.3, w: 4.8, h: 0.5,
      fontSize: 11, color: LGRAY, fontFace: FONT
    });
    if (i < 2) {
      s.addShape(pres.ShapeType.line, {
        x: 1.15, y: y + 0.7, w: 0, h: 0.55,
        line: { color: ACCENT, width: 2 }
      });
    }
  });

  // Right: image + key features
  s.addShape(pres.ShapeType.rect, {
    x: 7.2, y: 1.8, w: 5.5, h: 5.2,
    fill: { path: IMG("fingerprint.jpg") },
    rectRadius: 8
  });
  s.addShape(pres.ShapeType.rect, {
    x: 7.2, y: 1.8, w: 5.5, h: 5.2,
    fill: { color: DARK, transparency: 30 },
    rectRadius: 8
  });

  const features = [
    ["Puce biométrique", "Carte ISO 7816 / bracelet NFC"],
    ["Lecteur d'empreinte", "FAR < 0.001% (Suprema)"],
    ["Mode 100% hors-ligne", "Sync automatique via 2G"],
    ["Chiffrement AES-256", "RGPD / OMS conforme"]
  ];
  features.forEach(([t, d], i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 7.5, y: 2.1 + i * 1.15, w: 5, h: 0.95,
      fill: { color: DARK, transparency: 40 },
      rectRadius: 6
    });
    s.addText(`▸ ${t}`, {
      x: 7.7, y: 2.15 + i * 1.15, w: 4.6, h: 0.4,
      fontSize: 13, color: GOLD, fontFace: FONT, bold: true
    });
    s.addText(d, {
      x: 7.7, y: 2.55 + i * 1.15, w: 4.6, h: 0.35,
      fontSize: 11, color: LGRAY, fontFace: FONT
    });
  });
}

// =============== SLIDE 4 — FONCTIONNEMENT OFFLINE ===============
{
  const s = pres.addSlide();
  s.background = { fill: WHITE };

  s.addText("04", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: ACCENT, fontFace: FONT, bold: true
  });
  s.addText("Fonctionnement 100% hors-ligne", {
    x: 0.8, y: 0.6, w: 10, h: 0.8,
    fontSize: 32, color: DARK, fontFace: FONT, bold: true
  });
  s.addText("Conçu pour les zones sans connexion internet, avec synchronisation automatique dès que le réseau est disponible", {
    x: 0.8, y: 1.3, w: 10, h: 0.4,
    fontSize: 12, color: "666666", fontFace: FONT
  });

  // 4 tech cards
  const cards = [
    ["📇", "Carte à puce / Bracelet", "Résistant à l'eau et à la chaleur. NFC ISO 7816.", ACCENT],
    ["🔍", "Lecteur biométrique", "Capteur optique DigitalPersona/Suprema. IP65.", ACCENT],
    ["📱", "Application Web PWA", "React.js, IndexedDB, sync différée.", DARK],
    ["🔒", "Base de données centrale", "PostgreSQL chiffré AES-256. Datacenter N'Djamena.", DARK]
  ];
  cards.forEach(([icon, title, desc, color], i) => {
    const x = 0.6 + i * 3.1;
    s.addShape(pres.ShapeType.rect, {
      x, y: 2, w: 2.9, h: 3.5,
      fill: { color: "FAF9F6" },
      line: { color: "E0D8CC", width: 1 },
      rectRadius: 8
    });
    s.addText(icon, {
      x, y: 2.2, w: 2.9, h: 0.6,
      fontSize: 28, align: "center"
    });
    s.addText(title, {
      x: x + 0.2, y: 2.9, w: 2.5, h: 0.5,
      fontSize: 13, color: color, fontFace: FONT, bold: true, align: "center"
    });
    s.addText(desc, {
      x: x + 0.2, y: 3.4, w: 2.5, h: 0.6,
      fontSize: 10, color: "666666", fontFace: FONT, align: "center"
    });
  });

  // Bottom: data content section
  s.addText("Contenu du dossier médical numérique", {
    x: 0.8, y: 5.8, w: 8, h: 0.4,
    fontSize: 15, color: DARK, fontFace: FONT, bold: true
  });

  const items = [
    "Identité et données biométriques",
    "Carnet de vaccination complet",
    "Suivi de grossesse",
    "Historique des consultations",
    "Maladies chroniques et traitements",
    "Résultats biologiques et imagerie"
  ];
  items.forEach((item, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    s.addShape(pres.ShapeType.rect, {
      x: 0.8 + col * 4, y: 6.25 + row * 0.5, w: 3.7, h: 0.38,
      fill: { color: LGRAY },
      rectRadius: 4
    });
    s.addText(`✓  ${item}`, {
      x: 0.9 + col * 4, y: 6.25 + row * 0.5, w: 3.5, h: 0.38,
      fontSize: 10, color: DARK, fontFace: FONT
    });
  });

  // Image right side
  s.addShape(pres.ShapeType.rect, {
    x: 9.5, y: 2, w: 3.3, h: 4.5,
    fill: { path: IMG("doc-tablet.jpg") },
    rectRadius: 8
  });
}

// =============== SLIDE 5 — CONTENU & TECHNOLOGIE ===============
{
  const s = pres.addSlide();
  s.background = { fill: DARK };

  s.addText("05", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: GOLD, fontFace: FONT, bold: true
  });
  s.addText("Architecture technique", {
    x: 0.8, y: 0.6, w: 8, h: 0.8,
    fontSize: 36, color: WHITE, fontFace: FONT, bold: true
  });
  s.addText("Sécurité, interopérabilité et résilience", {
    x: 0.8, y: 1.3, w: 8, h: 0.4,
    fontSize: 14, color: GOLD, fontFace: FONT
  });

  const arch = [
    ["Application Web PWA", "React.js, mode hors-ligne", "IndexedDB + file d'attente de sync via API REST sécurisée (HTTPS/TLS)"],
    ["Authentification", "Double facteur", "Badge professionnel + code PIN pour chaque agent de santé"],
    ["Protection des données", "Empreintes en gabarit chiffré", "Jamais en image brute. Conforme législation tchadienne."],
    ["Hébergement", "Datacenter N'Djamena", "Réplication cloud Africa Zone. RBAC strict."]
  ];
  arch.forEach(([t, sbt, desc], i) => {
    const y = 2 + i * 1.25;
    s.addShape(pres.ShapeType.rect, {
      x: 0.8, y, w: 6.5, h: 1.05,
      fill: { color: DARK, transparency: 40 },
      line: { color: GOLD, width: 0.5, transparency: 60 },
      rectRadius: 6
    });
    s.addText(t, {
      x: 1, y: y + 0.05, w: 4, h: 0.3,
      fontSize: 13, color: WHITE, fontFace: FONT, bold: true
    });
    s.addText(sbt, {
      x: 1, y: y + 0.35, w: 4, h: 0.25,
      fontSize: 11, color: GOLD, fontFace: FONT
    });
    s.addText(desc, {
      x: 1, y: y + 0.6, w: 6, h: 0.35,
      fontSize: 10, color: LGRAY, fontFace: FONT
    });
  });

  // Right image
  s.addShape(pres.ShapeType.rect, {
    x: 8, y: 2, w: 4.8, h: 5,
    fill: { path: IMG("doctor.jpg") },
    rectRadius: 8
  });
  s.addShape(pres.ShapeType.rect, {
    x: 8, y: 2, w: 4.8, h: 5,
    fill: { color: DARK, transparency: 20 },
    rectRadius: 8
  });

  s.addShape(pres.ShapeType.rect, {
    x: 8.3, y: 5.8, w: 4.2, h: 1,
    fill: { color: ACCENT },
    rectRadius: 6
  });
  s.addText("Authentification 2 facteurs", {
    x: 8.3, y: 5.85, w: 4.2, h: 0.35,
    fontSize: 11, color: WHITE, fontFace: FONT, bold: true, align: "center"
  });
  s.addText("Badge professionnel + Code PIN", {
    x: 8.3, y: 6.2, w: 4.2, h: 0.3,
    fontSize: 10, color: WHITE, fontFace: FONT, align: "center"
  });
}

// =============== SLIDE 6 — IMPACT ===============
{
  const s = imgSlide("mother-child.jpg", 30);

  s.addText("06", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: GOLD, fontFace: FONT, bold: true
  });
  s.addText("IMPACT ATTENDU", {
    x: 0.8, y: 0.6, w: 10, h: 0.8,
    fontSize: 36, color: WHITE, fontFace: FONT, bold: true
  });
  s.addText("Résultats cibles à 3 ans", {
    x: 0.8, y: 1.3, w: 8, h: 0.4,
    fontSize: 16, color: GOLD, fontFace: FONT
  });

  // Target metrics
  const targets = [
    ["500K", "Patients\nenregistrés"],
    ["150", "Centres de santé\néquipés"],
    ["80%", "Couverture\nvaccinale"],
    ["-40%", "Mortalité maternelle\névitable"]
  ];
  targets.forEach(([v, l], i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 0.8 + i * 3, y: 2, w: 2.7, h: 1.6,
      fill: { color: DARK, transparency: 35 },
      line: { color: GOLD, width: 1, transparency: 40 },
      rectRadius: 8
    });
    s.addText(v, {
      x: 0.8 + i * 3, y: 2.1, w: 2.7, h: 0.8,
      fontSize: 40, color: GOLD, fontFace: FONT, bold: true, align: "center"
    });
    s.addText(l, {
      x: 0.8 + i * 3, y: 2.9, w: 2.7, h: 0.6,
      fontSize: 11, color: LGRAY, fontFace: FONT, align: "center"
    });
  });

  // ODD
  s.addText("Alignement ODD", {
    x: 0.8, y: 4, w: 5, h: 0.4,
    fontSize: 14, color: WHITE, fontFace: FONT, bold: true
  });

  const odds = [
    ["ODD 3", "Bonne santé"],
    ["ODD 10", "Inégalités réduites"],
    ["ODD 9", "Innovation"],
    ["ODD 16", "Institutions"],
    ["ODD 17", "Partenariats"]
  ];
  odds.forEach(([n, t], i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 0.8 + i * 2.4, y: 4.5, w: 2.1, h: 1.2,
      fill: { color: ACCENT, transparency: 20 },
      line: { color: GOLD, width: 0.5 },
      rectRadius: 6
    });
    s.addText(n, {
      x: 0.8 + i * 2.4, y: 4.55, w: 2.1, h: 0.5,
      fontSize: 18, color: GOLD, fontFace: FONT, bold: true, align: "center"
    });
    s.addText(t, {
      x: 0.8 + i * 2.4, y: 5.0, w: 2.1, h: 0.5,
      fontSize: 11, color: WHITE, fontFace: FONT, align: "center"
    });
  });

  // Bottom bar
  s.addShape(pres.ShapeType.rect, {
    x: 0.8, y: 6, w: 11.7, h: 0.8,
    fill: { color: DARK, transparency: 30 },
    rectRadius: 6
  });
  s.addText("ODD 3 — Bonne santé et bien-être : Amélioration du suivi médical, réduction des erreurs, meilleure couverture vaccinale", {
    x: 1, y: 6, w: 11.5, h: 0.8,
    fontSize: 12, color: LGRAY, fontFace: FONT, valign: "middle"
  });
}

// =============== SLIDE 7 — PLAN DE DÉPLOIEMENT ===============
{
  const s = pres.addSlide();
  s.background = { fill: WHITE };

  s.addText("07", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: ACCENT, fontFace: FONT, bold: true
  });
  s.addText("Plan de déploiement", {
    x: 0.8, y: 0.6, w: 10, h: 0.8,
    fontSize: 36, color: DARK, fontFace: FONT, bold: true
  });
  s.addText("3 phases — 3 ans", {
    x: 0.8, y: 1.3, w: 8, h: 0.4,
    fontSize: 16, color: ACCENT, fontFace: FONT
  });

  // Timeline
  const phases = [
    {
      title: "Phase 1 — Pilote N'Djamena",
      period: "Mois 1-6",
      items: [
        "5 hôpitaux de référence",
        "10 000 patients enregistrés",
        "200 agents de santé formés",
        "Validation technique"
      ]
    },
    {
      title: "Phase 2 — Extension régionale",
      period: "Mois 7-18",
      items: [
        "23 chefs-lieux de province",
        "100 000 patients enregistrés",
        "Intégration DHIS2",
        "Équipes mobiles"
      ]
    },
    {
      title: "Phase 3 — Déploiement national",
      period: "Mois 19-36",
      items: [
        "500+ centres de santé ruraux",
        "500 000 patients",
        "Transfert de compétences",
        "Autonomie des équipes locales"
      ]
    }
  ];

  phases.forEach((ph, i) => {
    const x = 0.6 + i * 4.2;
    s.addShape(pres.ShapeType.rect, {
      x, y: 2, w: 3.9, h: 4.8,
      fill: { color: "FAF9F6" },
      line: { color: "E0D8CC", width: 1 },
      rectRadius: 8
    });

    // Phase header
    s.addShape(pres.ShapeType.rect, {
      x, y: 2, w: 3.9, h: 0.9,
      fill: { color: ACCENT },
      rectRadius: 8
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 2.7, w: 3.9, h: 0.2,
      fill: { color: ACCENT }
    });
    s.addText(ph.title, {
      x: x + 0.2, y: 2.05, w: 3.5, h: 0.5,
      fontSize: 13, color: WHITE, fontFace: FONT, bold: true, align: "center"
    });
    s.addText(ph.period, {
      x: x + 0.2, y: 2.5, w: 3.5, h: 0.3,
      fontSize: 11, color: WHITE, fontFace: FONT, align: "center"
    });

    // Items
    ph.items.forEach((item, j) => {
      s.addShape(pres.ShapeType.rect, {
        x: x + 0.2, y: 3.1 + j * 0.7, w: 3.5, h: 0.55,
        fill: { color: DARK, transparency: 95 },
        rectRadius: 4
      });
      s.addText(`→  ${item}`, {
        x: x + 0.3, y: 3.1 + j * 0.7, w: 3.3, h: 0.55,
        fontSize: 12, color: "444444", fontFace: FONT, valign: "middle"
      });
    });

    // Arrow between phases
    if (i < 2) {
      s.addText("▶", {
        x: x + 3.9, y: 3.6, w: 0.3, h: 0.4,
        fontSize: 18, color: ACCENT, fontFace: FONT, align: "center"
      });
    }
  });

  // Bottom: final goal
  s.addShape(pres.ShapeType.rect, {
    x: 0.8, y: 7, w: 11.7, h: 0.35,
    fill: { color: ACCENT },
    rectRadius: 4
  });
  s.addText("Objectif final : Doter chaque Tchadien d'un dossier médical numérique unique, sécurisé et accessible par empreinte digitale", {
    x: 0.8, y: 7, w: 11.7, h: 0.35,
    fontSize: 11, color: WHITE, fontFace: FONT, align: "center", valign: "middle"
  });
}

// =============== SLIDE 8 — BUDGET ===============
{
  const s = pres.addSlide();
  s.background = { fill: DARK };

  s.addText("08", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: GOLD, fontFace: FONT, bold: true
  });
  s.addText("Budget prévisionnel", {
    x: 0.8, y: 0.6, w: 10, h: 0.8,
    fontSize: 36, color: WHITE, fontFace: FONT, bold: true
  });
  s.addText("3 500 000 USD sur 3 ans", {
    x: 0.8, y: 1.3, w: 8, h: 0.5,
    fontSize: 24, color: GOLD, fontFace: FONT, bold: true
  });

  // Budget table
  const rows = [
    [{ text: "Poste budgétaire", options: { bold: true, color: WHITE, fill: { color: ACCENT } } },
     { text: "Année 1", options: { bold: true, color: WHITE, fill: { color: ACCENT }, align: "center" } },
     { text: "Année 2", options: { bold: true, color: WHITE, fill: { color: ACCENT }, align: "center" } },
     { text: "Année 3", options: { bold: true, color: WHITE, fill: { color: ACCENT }, align: "center" } },
     { text: "Total", options: { bold: true, color: WHITE, fill: { color: ACCENT }, align: "center" } }],
    [{ text: "Matériel biométrique" }, { text: "450 K", options: { align: "center" } }, { text: "180 K", options: { align: "center" } }, { text: "120 K", options: { align: "center" } }, { text: "750 K", options: { align: "center" } }],
    [{ text: "Développement logiciel" }, { text: "350 K", options: { align: "center" } }, { text: "120 K", options: { align: "center" } }, { text: "80 K", options: { align: "center" } }, { text: "550 K", options: { align: "center" } }],
    [{ text: "Infrastructure serveurs" }, { text: "180 K", options: { align: "center" } }, { text: "90 K", options: { align: "center" } }, { text: "90 K", options: { align: "center" } }, { text: "360 K", options: { align: "center" } }],
    [{ text: "Formation des agents" }, { text: "150 K", options: { align: "center" } }, { text: "180 K", options: { align: "center" } }, { text: "120 K", options: { align: "center" } }, { text: "450 K", options: { align: "center" } }],
    [{ text: "Déploiement terrain" }, { text: "200 K", options: { align: "center" } }, { text: "250 K", options: { align: "center" } }, { text: "150 K", options: { align: "center" } }, { text: "600 K", options: { align: "center" } }],
    [{ text: "Sécurité & conformité" }, { text: "80 K", options: { align: "center" } }, { text: "60 K", options: { align: "center" } }, { text: "60 K", options: { align: "center" } }, { text: "200 K", options: { align: "center" } }],
    [{ text: "Communication" }, { text: "50 K", options: { align: "center" } }, { text: "80 K", options: { align: "center" } }, { text: "60 K", options: { align: "center" } }, { text: "190 K", options: { align: "center" } }],
    [{ text: "Gestion de projet" }, { text: "100 K", options: { align: "center" } }, { text: "130 K", options: { align: "center" } }, { text: "170 K", options: { align: "center" } }, { text: "400 K", options: { align: "center" } }],
    [{ text: "TOTAL", options: { bold: true, color: GOLD } },
     { text: "1 560 K", options: { bold: true, color: GOLD, align: "center" } },
     { text: "1 090 K", options: { bold: true, color: GOLD, align: "center" } },
     { text: "850 K", options: { bold: true, color: GOLD, align: "center" } },
     { text: "3 500 K", options: { bold: true, color: GOLD, align: "center" } }]
  ];

  s.addTable(rows, {
    x: 0.6, y: 2, w: 7.8,
    fontSize: 10,
    fontFace: FONT,
    border: { color: "2D2D44", width: 0.5 },
    rowH: 0.38,
    color: "CCCCCC",
    fill: { color: DARK, transparency: 50 },
    colW: [2.5, 0.8, 0.8, 0.8, 0.8],
    margin: [4, 6, 4, 6]
  });

  // Funding sources
  s.addText("Sources de financement recherchées", {
    x: 8.8, y: 2, w: 4, h: 0.4,
    fontSize: 13, color: WHITE, fontFace: FONT, bold: true
  });

  const funds = [
    ["Banque Mondiale — REDISSE", "1 200 000 USD", 34],
    ["Union Européenne — NDICI", "800 000 USD", 23],
    ["USAID — Global Health", "700 000 USD", 20],
    ["Gouvernement du Tchad", "500 000 USD", 14],
    ["Fondations philanthropiques", "300 000 USD", 9]
  ];
  funds.forEach(([name, amt, pct], i) => {
    const y = 2.6 + i * 0.85;
    s.addShape(pres.ShapeType.rect, {
      x: 8.8, y, w: 3.8, h: 0.7,
      fill: { color: DARK, transparency: 40 },
      line: { color: GOLD, width: 0.5, transparency: 50 },
      rectRadius: 4
    });
    s.addShape(pres.ShapeType.rect, {
      x: 8.8, y: y + 0.55, w: 3.8 * pct / 100, h: 0.06,
      fill: { color: GOLD },
      rectRadius: 2
    });
    s.addText(name, {
      x: 9, y: y + 0.02, w: 2.4, h: 0.3,
      fontSize: 10, color: WHITE, fontFace: FONT
    });
    s.addText(amt, {
      x: 11.5, y: y + 0.02, w: 1, h: 0.3,
      fontSize: 10, color: GOLD, fontFace: FONT, bold: true, align: "right"
    });
  });
}

// =============== SLIDE 9 — GOUVERNANCE & RISQUES ===============
{
  const s = pres.addSlide();
  s.background = { fill: WHITE };

  s.addText("09", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: ACCENT, fontFace: FONT, bold: true
  });
  s.addText("Gouvernance, sécurité & durabilité", {
    x: 0.8, y: 0.6, w: 11, h: 0.8,
    fontSize: 32, color: DARK, fontFace: FONT, bold: true
  });

  // 3 governance cards
  const govCards = [
    {
      title: "Gouvernance",
      items: [
        "Comité de Pilotage National - Ministère Santé",
        "Unité de Gestion de Projet (UGP) dédiée",
        "Comité technique (IT, médecins, communautés)",
        "Audits indépendants semestriels"
      ]
    },
    {
      title: "Protection des données",
      items: [
        "Conformité législation tchadienne",
        "Empreintes en gabarit chiffré (jamais brut)",
        "Consentement éclairé obligatoire",
        "Droit d'accès, rectification et suppression"
      ]
    },
    {
      title: "Viabilité financière",
      items: [
        "Intégration budget national année 3",
        "Frais symboliques (300 FCFA)",
        "Partenariats télécom (Airtel, Moov)",
        "Maintenance par développeurs tchadiens"
      ]
    }
  ];

  govCards.forEach((card, i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.6, w: 3.8, h: 2.8,
      fill: { color: "FAF9F6" },
      line: { color: "E0D8CC", width: 1 },
      rectRadius: 8
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 1.6, w: 3.8, h: 0.55,
      fill: { color: ACCENT },
      rectRadius: 8
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: 2, w: 3.8, h: 0.15,
      fill: { color: ACCENT }
    });
    s.addText(card.title, {
      x: x + 0.2, y: 1.63, w: 3.4, h: 0.45,
      fontSize: 14, color: WHITE, fontFace: FONT, bold: true, align: "center"
    });
    card.items.forEach((item, j) => {
      s.addText(`✓  ${item}`, {
        x: x + 0.2, y: 2.35 + j * 0.45, w: 3.4, h: 0.4,
        fontSize: 10, color: "444444", fontFace: FONT
      });
    });
  });

  // Risks section
  s.addText("Analyse des risques", {
    x: 0.8, y: 4.7, w: 5, h: 0.4,
    fontSize: 15, color: DARK, fontFace: FONT, bold: true
  });

  const risks = [
    ["📡", "Faible connectivité", "Mode offline 2G"],
    ["💪", "Résistance au changement", "Formation intensive"],
    ["🌡️", "Défaillance matérielle", "Matériel IP65, stock 20%"],
    ["🔒", "Cyber-attaques", "AES-256, audits trimestriels"],
    ["🌍", "Instabilité politique", "Architecture décentralisée"],
    ["👥", "Faible adoption", "Sensibilisation locale"]
  ];

  risks.forEach(([icon, risk, mitigation], i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1;
    const y = 5.2 + row * 0.7;

    s.addShape(pres.ShapeType.rect, {
      x, y, w: 3.8, h: 0.55,
      fill: { color: "FFF5F0" },
      line: { color: "FED7AA", width: 1 },
      rectRadius: 6
    });
    s.addText(`${icon}  ${risk}`, {
      x: x + 0.15, y, w: 1.8, h: 0.55,
      fontSize: 10, color: "9A3412", fontFace: FONT, valign: "middle"
    });
    s.addText(mitigation, {
      x: x + 2, y, w: 1.7, h: 0.55,
      fontSize: 10, color: ACCENT, fontFace: FONT, bold: true, valign: "middle", align: "right"
    });
  });

  // Right image
  s.addShape(pres.ShapeType.rect, {
    x: 9.5, y: 1.6, w: 3.3, h: 2.8,
    fill: { path: IMG("team.jpg") },
    rectRadius: 8
  });
}

// =============== SLIDE 10 — APPEL À FINANCEMENT ===============
{
  const s = imgSlide("doc-tablet.jpg", 30);

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 7.5,
    fill: { color: DARK, transparency: 40 }
  });

  s.addText("10", {
    x: 0.8, y: 0.3, w: 1, h: 0.5,
    fontSize: 14, color: GOLD, fontFace: FONT, bold: true
  });
  s.addText("Appel à financement", {
    x: 0.8, y: 0.6, w: 10, h: 0.9,
    fontSize: 40, color: WHITE, fontFace: FONT, bold: true
  });

  s.addText("Ensemble, offrons à chaque Tchadien — du nomade du Sahel à la mère en zone rurale — le droit à un suivi médical digne et sécurisé.", {
    x: 0.8, y: 1.5, w: 8, h: 0.7,
    fontSize: 14, color: LGRAY, fontFace: FONT
  });

  // 4 impact items
  const items = [
    ["❤️", "Sauver des vies en améliorant le diagnostic"],
    ["💪", "Renforcer les systèmes de santé nationaux"],
    ["🌍", "Construire une infrastructure numérique durable"],
    ["🎯", "Atteindre les ODD d'ici 2030"]
  ];
  items.forEach(([icon, text], i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 0.8, y: 2.5 + i * 0.55, w: 6.5, h: 0.44,
      fill: { color: DARK, transparency: 50 },
      rectRadius: 4
    });
    s.addText(`${icon}  ${text}`, {
      x: 1, y: 2.5 + i * 0.55, w: 6.2, h: 0.44,
      fontSize: 13, color: WHITE, fontFace: FONT, valign: "middle"
    });
  });

  // Contact info box
  s.addShape(pres.ShapeType.rect, {
    x: 8.2, y: 2, w: 4.5, h: 3.5,
    fill: { color: DARK, transparency: 25 },
    line: { color: GOLD, width: 1 },
    rectRadius: 8
  });

  s.addText("CONTACT", {
    x: 8.4, y: 2.2, w: 4.1, h: 0.3,
    fontSize: 11, color: GOLD, fontFace: FONT, bold: true, letterSpacing: 2
  });
  s.addText("Unité de Gestion Amaan\nMinistère de la Santé Publique\nN'Djamena, Tchad", {
    x: 8.4, y: 2.6, w: 4.1, h: 0.9,
    fontSize: 12, color: WHITE, fontFace: FONT
  });
  s.addText("contact@amaan-tchad.gov.td", {
    x: 8.4, y: 3.5, w: 4.1, h: 0.3,
    fontSize: 13, color: GOLD, fontFace: FONT, bold: true
  });

  // Next steps
  s.addText("PROCHAINES ÉTAPES", {
    x: 8.4, y: 4, w: 4.1, h: 0.3,
    fontSize: 11, color: GOLD, fontFace: FONT, bold: true, letterSpacing: 2
  });

  const nextSteps = [
    "1. Réunion de présentation (30 min)",
    "2. Due diligence technique",
    "3. Signature accord de financement",
    "4. Lancement Phase Pilote (T1 2025)"
  ];
  nextSteps.forEach((step, i) => {
    s.addText(step, {
      x: 8.4, y: 4.4 + i * 0.35, w: 4.1, h: 0.3,
      fontSize: 11, color: LGRAY, fontFace: FONT
    });
  });

  // Bottom CTA
  s.addShape(pres.ShapeType.rect, {
    x: 0.8, y: 6.2, w: 11.7, h: 0.8,
    fill: { color: ACCENT },
    rectRadius: 6
  });
  s.addText("INVESTISSEZ DANS LA SANTÉ NUMÉRIQUE — Contactez-nous : contact@amaan-tchad.gov.td", {
    x: 0.8, y: 6.2, w: 11.7, h: 0.8,
    fontSize: 15, color: WHITE, fontFace: FONT, bold: true, align: "center", valign: "middle"
  });
}

// Save
const outPath = path.resolve(__dirname, "Amaan_Presentation_Bailleurs.pptx");
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("✅ PPTX créé :", outPath);
}).catch(err => {
  console.error("❌ Erreur :", err);
});
