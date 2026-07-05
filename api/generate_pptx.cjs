const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const {
  FaUser, FaUniversity, FaFlask, FaBuilding, FaLightbulb,
  FaChartLine, FaUsers, FaBullseye, FaRocket, FaHandshake,
  FaMoneyBillWave, FaTrophy, FaMapMarkedAlt, FaPhone, FaEnvelope,
  FaFacebook, FaInstagram, FaYoutube, FaTwitter, FaGlobe,
  FaGraduationCap, FaCogs, FaStore, FaBolt, FaChartBar
} = require("react-icons/fa");
const { MdBusiness, MdPeople, MdTrendingUp } = require("react-icons/md");

// ── Helpers ────────────────────────────────────────────────────────────────
async function iconB64(Icon, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const C = {
  navy:   "0A1172",
  navyM:  "1a237e",
  navyL:  "283593",
  gold:   "C9A84C",
  goldL:  "FFD54F",
  white:  "FFFFFF",
  light:  "E8EAF6",
  gray:   "9E9E9E",
  dark:   "1A1A2E",
  red:    "C62828",
  green:  "2E7D32",
  teal:   "00695C",
};

function makeShadow() {
  return { type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.18 };
}

// Common header bar (navy background strip at top)
function addHeader(slide, title, pres) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.85,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  slide.addText(title, {
    x: 0.4, y: 0, w: 8.5, h: 0.85,
    fontSize: 22, bold: true, color: C.white,
    fontFace: "Calibri", valign: "middle", margin: 0
  });
}

// Logo text badge top-right
function addLogoBadge(slide, pres, logoText = "StartupAI") {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 8.8, y: 0.05, w: 1.1, h: 0.75,
    fill: { color: C.gold }, line: { color: C.gold }, rectRadius: 0.08
  });
  slide.addText(logoText, {
    x: 8.8, y: 0.05, w: 1.1, h: 0.75,
    fontSize: 9, bold: true, color: C.navy,
    fontFace: "Calibri", align: "center", valign: "middle", margin: 0
  });
}

// Footer line
function addFooter(slide, pres, universityName, incubatorName) {
  slide.addShape(pres.shapes.LINE, {
    x: 0.3, y: 5.4, w: 9.4, h: 0,
    line: { color: C.gold, width: 1.5 }
  });
  slide.addText(`${universityName}  |  ${incubatorName}`, {
    x: 0.3, y: 5.45, w: 9.4, h: 0.18,
    fontSize: 7, color: C.gray, fontFace: "Calibri", align: "center", margin: 0
  });
}

// Numbered circle badge
function addCircleBadge(slide, pres, num, x, y, size = 0.55) {
  slide.addShape(pres.shapes.OVAL, {
    x, y, w: size, h: size,
    fill: { color: C.gold }, line: { color: C.gold }
  });
  slide.addText(String(num).padStart(2, "0"), {
    x, y, w: size, h: size,
    fontSize: 14, bold: true, color: C.navy,
    fontFace: "Calibri", align: "center", valign: "middle", margin: 0
  });
}

// Card shape
function addCard(slide, pres, x, y, w, h, fillColor = "F5F5FF") {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: fillColor },
    line: { color: "D0D0E8", width: 0.5 },
    rectRadius: 0.1,
    shadow: makeShadow()
  });
}

// ── MAIN ──────────────────────────────────────────────────────────────────
async function generatePPTX(data, outputPath) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.projectTitle;
  pres.author = "StartupAI — Diplôme/Startup 1275";

  const uni  = data.universityName  || "Université ...";
  const inc  = data.incubatorName   || "Incubateur ...";
  const fac  = data.facultyName     || "Faculté ...";
  const dept = data.departmentName  || "Département ...";
  const spec = data.speciality      || "Spécialité ...";
  const year = data.academicYear    || "2025/2026";
  const logo = (data.projectTitle || "Startup").substring(0, 8);
  const members = data.teamMembers  || [];
  const slides  = data.slides       || {};

  // Icons (pre-render)
  const icoUser    = await iconB64(FaUser,          "#" + C.navy, 256);
  const icoUni     = await iconB64(FaUniversity,    "#" + C.gold, 256);
  const icoLight   = await iconB64(FaLightbulb,     "#" + C.gold, 256);
  const icoChart   = await iconB64(FaChartLine,     "#" + C.gold, 256);
  const icoUsers   = await iconB64(FaUsers,         "#" + C.white, 256);
  const icoTarget  = await iconB64(FaBullseye,      "#" + C.gold, 256);
  const icoRocket  = await iconB64(FaRocket,        "#" + C.gold, 256);
  const icoHand    = await iconB64(FaHandshake,     "#" + C.gold, 256);
  const icoMoney   = await iconB64(FaMoneyBillWave, "#" + C.gold, 256);
  const icoTrophy  = await iconB64(FaTrophy,        "#" + C.gold, 256);
  const icoMap     = await iconB64(FaMapMarkedAlt,  "#" + C.gold, 256);
  const icoGlobe   = await iconB64(FaGlobe,         "#" + C.gray, 256);
  const icoCogs    = await iconB64(FaCogs,          "#" + C.gold, 256);
  const icoStore   = await iconB64(FaStore,         "#" + C.gold, 256);
  const icoBolt    = await iconB64(FaBolt,          "#" + C.gold, 256);
  const icoGrad    = await iconB64(FaGraduationCap, "#" + C.white, 256);

  // Member icons (colored circles per member)
  const memberColors = ["4472C4","7030A0","C00000","375623","0070C0","843C0C"];
  const icoMembers = [];
  for (let i = 0; i < Math.max(members.length, 4); i++) {
    icoMembers.push(await iconB64(FaUser, "#" + (memberColors[i % memberColors.length]), 256));
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 1 — Page de garde
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Top green/gold band
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.18,
      fill: { color: C.gold }, line: { color: C.gold }
    });

    // Deco triangle right
    s.addShape(pres.shapes.RIGHT_TRIANGLE, {
      x: 7.5, y: 0.18, w: 2.5, h: 5.3,
      fill: { color: C.navyL, transparency: 60 },
      line: { color: C.navyL, transparency: 60 }
    });

    // Republic header
    s.addText("République Algérienne Démocratique et Populaire", {
      x: 0.4, y: 0.25, w: 9.2, h: 0.3,
      fontSize: 10, color: C.gold, bold: true,
      fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText("Ministère de l'Enseignement Supérieur et de la Recherche Scientifique", {
      x: 0.4, y: 0.52, w: 9.2, h: 0.25,
      fontSize: 9, color: C.light, fontFace: "Calibri", align: "center", margin: 0
    });

    // University & Faculty
    s.addImage({ data: icoUni, x: 0.5, y: 0.85, w: 0.45, h: 0.45 });
    s.addText(uni, {
      x: 1.05, y: 0.85, w: 8, h: 0.25,
      fontSize: 13, color: C.goldL, bold: true,
      fontFace: "Calibri", margin: 0
    });
    s.addText(`${fac}  |  ${dept}`, {
      x: 1.05, y: 1.1, w: 8, h: 0.22,
      fontSize: 10, color: C.light, fontFace: "Calibri", margin: 0
    });

    // Gold divider
    s.addShape(pres.shapes.LINE, {
      x: 0.4, y: 1.4, w: 9.2, h: 0,
      line: { color: C.gold, width: 2 }
    });

    // Project title card
    addCard(s, pres, 0.5, 1.5, 9, 1.6, "162068");
    s.addText("Intitulé du Projet", {
      x: 0.7, y: 1.55, w: 8.6, h: 0.3,
      fontSize: 10, color: C.gold, italic: true,
      fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText(data.projectTitle || "Titre du projet innovant", {
      x: 0.7, y: 1.85, w: 8.6, h: 1.1,
      fontSize: 26, bold: true, color: C.white,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0
    });

    // Subtitle info
    s.addText(`En vue de l'obtention du diplôme de LICENCE / MASTER`, {
      x: 0.5, y: 3.2, w: 9, h: 0.25,
      fontSize: 10, color: C.light, fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText(`Dans le cadre de la décision 008 : Diplôme Institution Économique`, {
      x: 0.5, y: 3.45, w: 9, h: 0.22,
      fontSize: 9.5, color: C.gold, fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText(`Spécialité : ${spec}`, {
      x: 0.5, y: 3.7, w: 9, h: 0.22,
      fontSize: 10, color: C.light, fontFace: "Calibri", align: "center", margin: 0
    });

    // Incubator badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 3.0, y: 4.0, w: 4.0, h: 0.55,
      fill: { color: C.gold }, line: { color: C.gold }, rectRadius: 0.1
    });
    s.addText(`Incubateur : ${inc}`, {
      x: 3.0, y: 4.0, w: 4.0, h: 0.55,
      fontSize: 11, bold: true, color: C.navy,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0
    });

    s.addText(`Année universitaire : ${year}`, {
      x: 0.5, y: 4.65, w: 9, h: 0.22,
      fontSize: 10, color: C.gray, fontFace: "Calibri", align: "center", margin: 0
    });

    // Bottom gold bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 5.425, w: 10, h: 0.2,
      fill: { color: C.gold }, line: { color: C.gold }
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 2 — Logo & Tagline
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    // Big logo circle
    s.addShape(pres.shapes.OVAL, {
      x: 3.5, y: 0.3, w: 3, h: 3,
      fill: { color: C.navyM }, line: { color: C.gold, width: 3 },
      shadow: makeShadow()
    });
    s.addImage({ data: icoLight, x: 4.25, y: 0.95, w: 1.5, h: 1.5 });
    s.addText(data.projectTitle || "Startup", {
      x: 1, y: 3.4, w: 8, h: 0.65,
      fontSize: 30, bold: true, color: C.gold,
      fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText(`« ${data.tagline || "Innovation • Impact • Solution"} »`, {
      x: 1, y: 4.05, w: 8, h: 0.35,
      fontSize: 14, italic: true, color: C.light,
      fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText(slides.slide2_description || "Une solution innovante qui répond aux besoins réels du marché algérien.", {
      x: 1.5, y: 4.5, w: 7, h: 0.7,
      fontSize: 12, color: C.gray,
      fontFace: "Calibri", align: "center", margin: 0
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 3 — Résumé exécutif (CV du Projet)
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Résumé Exécutif — CV du Projet", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const cv = slides.slide3 || {};
    const boxes = [
      { icon: icoTarget, label: "Problème",          val: cv.problem        || "Description du problème principal",         x: 0.2, y: 0.95 },
      { icon: icoLight,  label: "Solution",           val: cv.solution       || "Description de la solution proposée",        x: 3.55, y: 0.95 },
      { icon: icoBolt,   label: "État d'avancement",  val: cv.status         || "MVP en cours de développement",              x: 6.9, y: 0.95 },
      { icon: icoStore,  label: "Produit",             val: cv.product        || "Description du produit ou service",          x: 0.2, y: 2.85 },
      { icon: icoMoney,  label: "Modèle commercial",  val: cv.businessModel  || "Abonnement / Freemium / Commission",         x: 3.55, y: 2.85 },
      { icon: icoUsers,  label: "Clients cibles",     val: cv.customers      || "Étudiants, Entreprises, Institutions",        x: 6.9, y: 2.85 },
    ];

    for (const b of boxes) {
      addCard(s, pres, b.x, b.y, 3.1, 1.75, C.light);
      s.addImage({ data: b.icon, x: b.x + 0.1, y: b.y + 0.08, w: 0.38, h: 0.38 });
      s.addText(b.label, {
        x: b.x + 0.55, y: b.y + 0.08, w: 2.5, h: 0.38,
        fontSize: 11, bold: true, color: C.navy,
        fontFace: "Calibri", valign: "middle", margin: 0
      });
      s.addShape(pres.shapes.LINE, {
        x: b.x + 0.1, y: b.y + 0.52, w: 2.9, h: 0,
        line: { color: C.gold, width: 1 }
      });
      s.addText(b.val, {
        x: b.x + 0.1, y: b.y + 0.6, w: 2.9, h: 1.08,
        fontSize: 10, color: "444444",
        fontFace: "Calibri", valign: "top", margin: 2
      });
    }

    // Co-founders row
    s.addText("Co-fondateurs : " + (members.map(m => m.name).join("  •  ") || "Équipe à définir"), {
      x: 0.2, y: 4.75, w: 9.4, h: 0.3,
      fontSize: 9.5, color: C.navyM, fontFace: "Calibri",
      align: "center", bold: true, margin: 0
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 4 — Présentation de l'équipe
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Présentation de l'Équipe", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const displayMembers = members.length > 0 ? members.slice(0, 4) : [
      { name: "Étudiant 01", role: "Chef de projet", background: "Spécialité" },
      { name: "Étudiant 02", role: "Développeur", background: "Informatique" },
      { name: "Étudiant 03", role: "Marketing", background: "Commerce" },
    ];

    const cols = Math.min(displayMembers.length, 4);
    const cardW = cols <= 3 ? 2.8 : 2.2;
    const startX = cols <= 3 ? (10 - cols * cardW - (cols - 1) * 0.25) / 2 : 0.2;

    for (let i = 0; i < displayMembers.length && i < 4; i++) {
      const m = displayMembers[i];
      const cx = startX + i * (cardW + 0.25);
      const cy = 1.05;

      addCard(s, pres, cx, cy, cardW, 4.1, C.light);

      // Member icon in colored circle
      const circColor = memberColors[i % memberColors.length];
      s.addShape(pres.shapes.OVAL, {
        x: cx + cardW / 2 - 0.5, y: cy + 0.15, w: 1.0, h: 1.0,
        fill: { color: circColor }, line: { color: circColor },
        shadow: makeShadow()
      });
      s.addImage({ data: icoMembers[i], x: cx + cardW / 2 - 0.35, y: cy + 0.3, w: 0.7, h: 0.7 });

      s.addText(m.name || `Étudiant ${i + 1}`, {
        x: cx + 0.1, y: cy + 1.25, w: cardW - 0.2, h: 0.38,
        fontSize: 12, bold: true, color: C.navy,
        fontFace: "Calibri", align: "center", margin: 0
      });
      s.addShape(pres.shapes.LINE, {
        x: cx + 0.25, y: cy + 1.67, w: cardW - 0.5, h: 0,
        line: { color: C.gold, width: 1.5 }
      });
      s.addText(m.role || "Fonction dans le projet", {
        x: cx + 0.1, y: cy + 1.75, w: cardW - 0.2, h: 0.3,
        fontSize: 10, italic: true, color: C.navyL,
        fontFace: "Calibri", align: "center", margin: 0
      });
      s.addText(m.background || "Spécialité / Background", {
        x: cx + 0.1, y: cy + 2.1, w: cardW - 0.2, h: 0.3,
        fontSize: 9.5, color: "555555",
        fontFace: "Calibri", align: "center", margin: 0
      });
      s.addText(m.faculty || fac, {
        x: cx + 0.1, y: cy + 2.4, w: cardW - 0.2, h: 0.25,
        fontSize: 8.5, color: C.gray,
        fontFace: "Calibri", align: "center", margin: 0
      });
      // Badge icon + université
      s.addImage({ data: icoGrad, x: cx + 0.15, y: cy + 2.75, w: 0.25, h: 0.25 });
      s.addText(uni, {
        x: cx + 0.45, y: cy + 2.75, w: cardW - 0.55, h: 0.25,
        fontSize: 7.5, color: C.gray, fontFace: "Calibri", margin: 0
      });
      s.addText(inc, {
        x: cx + 0.1, y: cy + 3.05, w: cardW - 0.2, h: 0.22,
        fontSize: 7.5, color: C.teal, fontFace: "Calibri", align: "center",
        bold: true, margin: 0
      });
    }
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 5 — Problème
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: "FFF8F8" };

    addHeader(s, "Le Problème", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide5 || {};
    // Big icon left
    addCard(s, pres, 0.2, 0.95, 3.8, 4.3, C.light);
    s.addImage({ data: icoTarget, x: 1.2, y: 1.2, w: 1.8, h: 1.8 });
    s.addText(d.mainProblem || "Le problème principal que nous résolvons", {
      x: 0.3, y: 3.1, w: 3.6, h: 1.3,
      fontSize: 13, bold: true, color: C.navy,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 4
    });

    // Right: sub-problems
    const subs = d.subProblems || ["Sous-problème 1", "Sous-problème 2", "Sous-problème 3"];
    let ry = 1.0;
    subs.forEach((sp, i) => {
      addCard(s, pres, 4.3, ry, 5.5, 0.72, "EEF0FF");
      addCircleBadge(s, pres, i + 1, 4.35, ry + 0.085, 0.52);
      s.addText(sp, {
        x: 5.0, y: ry, w: 4.7, h: 0.72,
        fontSize: 11.5, color: C.navy, fontFace: "Calibri",
        valign: "middle", margin: 4
      });
      ry += 0.82;
    });

    // Stats
    s.addText(`Personnes concernées : ${d.affectedCount || "Millions d'Algériens"}`, {
      x: 4.3, y: ry + 0.15, w: 5.5, h: 0.3,
      fontSize: 11, color: C.navyM, bold: true,
      fontFace: "Calibri", margin: 2
    });
    s.addText(d.stats || "Des statistiques clés illustrant l'ampleur du problème.", {
      x: 4.3, y: ry + 0.5, w: 5.5, h: 0.5,
      fontSize: 10, color: "555555", fontFace: "Calibri", margin: 2
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 6 — Solution
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Notre Solution", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide6 || {};
    s.addText(d.solutionDesc || "Description claire de la solution principale proposée.", {
      x: 0.3, y: 0.95, w: 9.4, h: 0.55,
      fontSize: 13, color: C.navy, fontFace: "Calibri",
      align: "center", margin: 2
    });

    // Benefits
    const benefits = d.benefits || ["Bénéfice 1 — Rapidité", "Bénéfice 2 — Accessibilité", "Bénéfice 3 — Économie"];
    benefits.slice(0, 3).forEach((b, i) => {
      const bx = 0.3 + i * 3.2;
      addCard(s, pres, bx, 1.58, 2.9, 1.0, C.light);
      s.addShape(pres.shapes.OVAL, {
        x: bx + 1.15, y: 1.65, w: 0.6, h: 0.6,
        fill: { color: C.gold }, line: { color: C.gold }
      });
      s.addText(String(i + 1), {
        x: bx + 1.15, y: 1.65, w: 0.6, h: 0.6,
        fontSize: 14, bold: true, color: C.navy,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      s.addText(b, {
        x: bx + 0.1, y: 2.3, w: 2.7, h: 0.25,
        fontSize: 10, color: C.navy, fontFace: "Calibri",
        align: "center", margin: 0
      });
    });

    // How it works
    s.addText("Comment ça marche ?", {
      x: 0.3, y: 2.75, w: 9.4, h: 0.35,
      fontSize: 14, bold: true, color: C.navy,
      fontFace: "Calibri", margin: 0
    });
    s.addShape(pres.shapes.LINE, {
      x: 0.3, y: 3.12, w: 9.4, h: 0,
      line: { color: C.gold, width: 1.5 }
    });

    const steps = d.steps || ["Étape 1", "Étape 2", "Étape 3", "Étape 4"];
    steps.slice(0, 4).forEach((st, i) => {
      const sx = 0.4 + i * 2.35;
      s.addShape(pres.shapes.OVAL, {
        x: sx, y: 3.2, w: 0.6, h: 0.6,
        fill: { color: C.navyM }, line: { color: C.navyM }
      });
      s.addText(String(i + 1), {
        x: sx, y: 3.2, w: 0.6, h: 0.6,
        fontSize: 13, bold: true, color: C.white,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      if (i < 3) {
        s.addShape(pres.shapes.LINE, {
          x: sx + 0.62, y: 3.5, w: 1.7, h: 0,
          line: { color: C.gold, width: 2, dashType: "sysDash" }
        });
      }
      s.addText(st, {
        x: sx - 0.4, y: 3.88, w: 1.4, h: 0.5,
        fontSize: 9.5, color: C.navy, fontFace: "Calibri",
        align: "center", margin: 0
      });
    });
    s.addText(d.mediaNote || "Démonstration disponible via vidéo / prototype numérique.", {
      x: 0.3, y: 4.55, w: 9.4, h: 0.3,
      fontSize: 9.5, color: C.gray, italic: true,
      fontFace: "Calibri", align: "center", margin: 0
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 7 — Notre Produit
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "Notre Produit", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide7 || {};
    // Left card: product digital
    addCard(s, pres, 0.2, 0.95, 4.6, 4.2, C.white);
    s.addImage({ data: icoStore, x: 0.4, y: 1.05, w: 0.6, h: 0.6 });
    s.addText("Produit Numérique", {
      x: 1.1, y: 1.05, w: 3.5, h: 0.6,
      fontSize: 13, bold: true, color: C.navy, fontFace: "Calibri",
      valign: "middle", margin: 0
    });
    s.addText(d.digitalProduct || "Application / Plateforme web · Captures d'écran disponibles · Vidéo de démonstration (max 3 min) · Accès à un compte-test", {
      x: 0.3, y: 1.75, w: 4.4, h: 1.8,
      fontSize: 11, color: "444444", fontFace: "Calibri",
      valign: "top", margin: 4
    });
    s.addText(d.whatsapp || "WhatsApp : +213 ...", {
      x: 0.3, y: 3.7, w: 4.4, h: 0.3,
      fontSize: 10, color: C.teal, bold: true, fontFace: "Calibri", margin: 2
    });

    // Right card: physical
    addCard(s, pres, 5.1, 0.95, 4.6, 4.2, C.white);
    s.addImage({ data: icoCogs, x: 5.3, y: 1.05, w: 0.6, h: 0.6 });
    s.addText("Caractéristiques Clés", {
      x: 6.0, y: 1.05, w: 3.5, h: 0.6,
      fontSize: 13, bold: true, color: C.navy, fontFace: "Calibri",
      valign: "middle", margin: 0
    });
    const features = d.features || ["Fonctionnalité 1", "Fonctionnalité 2", "Fonctionnalité 3", "Fonctionnalité 4"];
    const ftItems = features.slice(0, 5).map((f, i) => ({
      text: `  ${f}`,
      options: { bullet: true, color: "333333", fontSize: 11, fontFace: "Calibri", breakLine: i < features.length - 1 }
    }));
    s.addText(ftItems, { x: 5.2, y: 1.75, w: 4.4, h: 3.1, margin: 4 });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 8 — Plan de production
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Plan de Production (Optionnel)", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide8 || {};
    const steps = d.productionSteps || [
      "Collecte des données / Matières premières",
      "Développement / Fabrication",
      "Conditionnement du produit",
      "Tests et validation",
      "Lancement",
      "Livraison"
    ];

    // Arrow flow
    steps.slice(0, 5).forEach((st, i) => {
      const sx = 0.25 + i * 1.9;
      addCard(s, pres, sx, 1.0, 1.65, 1.5, C.light);
      s.addShape(pres.shapes.OVAL, {
        x: sx + 0.525, y: 1.08, w: 0.6, h: 0.6,
        fill: { color: C.navyM }, line: { color: C.navyM }
      });
      s.addText(String(i + 1), {
        x: sx + 0.525, y: 1.08, w: 0.6, h: 0.6,
        fontSize: 14, bold: true, color: C.white,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      s.addText(st, {
        x: sx + 0.08, y: 1.75, w: 1.5, h: 0.65,
        fontSize: 8.5, color: C.navy, fontFace: "Calibri",
        align: "center", margin: 2
      });
      if (i < 4) {
        s.addShape(pres.shapes.LINE, {
          x: sx + 1.67, y: 1.3, w: 0.2, h: 0,
          line: { color: C.gold, width: 2.5 }
        });
      }
    });

    s.addText(d.productionNote || "Pour les plateformes numériques, préciser les étapes qui permettent de fournir le service.", {
      x: 0.4, y: 2.65, w: 9.2, h: 0.35,
      fontSize: 10, color: "555555", italic: true,
      fontFace: "Calibri", align: "center", margin: 0
    });

    // Main partners
    s.addText("Principaux Partenaires", {
      x: 0.4, y: 3.1, w: 9.2, h: 0.3,
      fontSize: 13, bold: true, color: C.navy, fontFace: "Calibri", margin: 0
    });
    const partners = d.partners || ["Incubateur universitaire", "Fournisseurs", "Banques / Financement", "Collectivités locales"];
    partners.slice(0, 4).forEach((p, i) => {
      addCard(s, pres, 0.3 + i * 2.35, 3.5, 2.1, 0.7, C.light);
      s.addImage({ data: icoHand, x: 0.35 + i * 2.35, y: 3.57, w: 0.35, h: 0.35 });
      s.addText(p, {
        x: 0.75 + i * 2.35, y: 3.5, w: 1.65, h: 0.7,
        fontSize: 9.5, color: C.navy, fontFace: "Calibri",
        valign: "middle", margin: 2
      });
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 9 — Analyse stratégique du marché
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Analyse Stratégique du Marché", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide9 || {};
    const mktBoxes = [
      { num: "01", icon: icoUsers,  title: "Segment du Marché",    val: d.segment || "Marché cible principal et ses caractéristiques démographiques" },
      { num: "02", icon: icoChart,  title: "Taille du Marché",     val: d.marketSize || "National : ...  |  Régional : ...  |  International : ..." },
      { num: "03", icon: icoRocket, title: "Potentiel de Croissance", val: d.growth || "Croissance exponentielle prévue sur 3-5 ans" },
    ];

    mktBoxes.forEach((b, i) => {
      const bx = 0.3 + i * 3.2;
      addCard(s, pres, bx, 0.97, 2.9, 4.2, C.light);
      s.addImage({ data: b.icon, x: bx + 1.1, y: 1.1, w: 0.7, h: 0.7 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: bx + 0.1, y: 1.9, w: 2.7, h: 0.38,
        fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.06
      });
      s.addText(`${b.num}  ${b.title}`, {
        x: bx + 0.1, y: 1.9, w: 2.7, h: 0.38,
        fontSize: 11, bold: true, color: C.gold,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      s.addText(b.val, {
        x: bx + 0.15, y: 2.35, w: 2.6, h: 2.65,
        fontSize: 11, color: "333333", fontFace: "Calibri",
        valign: "top", margin: 4
      });
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 10 — Concurrence
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: "F0F4FF" };

    addHeader(s, "Analyse de la Concurrence", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide10 || {};
    const compBoxes = [
      { num: "01", icon: icoUsers, title: "Nos Concurrents",       val: d.competitors || "Concurrents directs et indirects identifiés sur le marché" },
      { num: "02", icon: icoTrophy, title: "Nos Avantages",        val: d.advantages  || "Avantages distinctifs par rapport aux concurrents existants" },
      { num: "03", icon: icoBolt,   title: "Notre Innovation",     val: d.innovation  || "L'élément différenciateur clé qui rend notre solution unique" },
    ];

    compBoxes.forEach((b, i) => {
      const bx = 0.3 + i * 3.2;
      addCard(s, pres, bx, 0.97, 2.9, 4.2, C.white);
      s.addImage({ data: b.icon, x: bx + 1.1, y: 1.1, w: 0.7, h: 0.7 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: bx + 0.1, y: 1.9, w: 2.7, h: 0.38,
        fill: { color: C.gold }, line: { color: C.gold }, rectRadius: 0.06
      });
      s.addText(`${b.num}  ${b.title}`, {
        x: bx + 0.1, y: 1.9, w: 2.7, h: 0.38,
        fontSize: 11, bold: true, color: C.navy,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      s.addText(b.val, {
        x: bx + 0.15, y: 2.35, w: 2.6, h: 2.65,
        fontSize: 11, color: "333333", fontFace: "Calibri",
        valign: "top", margin: 4
      });
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 11 — Stratégie marketing
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Stratégie Marketing", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide11 || {};
    const mktPillars = [
      { num: "01", icon: icoGlobe,  title: "Stratégie de Visibilité",    val: d.visibility   || "Réseaux sociaux, SEO, presse universitaire, événements" },
      { num: "02", icon: icoStore,  title: "Stratégie de Distribution",  val: d.distribution || "Vente directe, partenariats, plateforme numérique" },
      { num: "03", icon: icoMoney,  title: "Stratégie de Prix",          val: d.pricing      || "Prix compétitif, freemium, abonnement mensuel/annuel" },
    ];

    mktPillars.forEach((b, i) => {
      const bx = 0.3 + i * 3.2;
      addCard(s, pres, bx, 0.97, 2.9, 3.0, C.light);
      s.addImage({ data: b.icon, x: bx + 1.1, y: 1.1, w: 0.7, h: 0.7 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: bx + 0.1, y: 1.9, w: 2.7, h: 0.38,
        fill: { color: C.navyM }, line: { color: C.navyM }, rectRadius: 0.06
      });
      s.addText(`${b.num}  ${b.title}`, {
        x: bx + 0.1, y: 1.9, w: 2.7, h: 0.38,
        fontSize: 10.5, bold: true, color: C.gold,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      s.addText(b.val, {
        x: bx + 0.15, y: 2.35, w: 2.6, h: 1.5,
        fontSize: 11, color: "333333", fontFace: "Calibri",
        valign: "top", margin: 4
      });
    });

    // Mix marketing 4P summary
    s.addText("Mix Marketing (4P)", {
      x: 0.3, y: 4.1, w: 9.4, h: 0.3,
      fontSize: 12, bold: true, color: C.navy, fontFace: "Calibri", margin: 0
    });
    const ps = d.mixMarketing || { product: "...", price: "...", place: "...", promotion: "..." };
    const pLabels = [["Produit", ps.product], ["Prix", ps.price], ["Distribution", ps.place], ["Promotion", ps.promotion]];
    pLabels.forEach(([lbl, val], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.3 + i * 2.35, y: 4.45, w: 2.1, h: 0.7,
        fill: { color: i % 2 === 0 ? C.navy : C.gold },
        line: { color: i % 2 === 0 ? C.navy : C.gold }, rectRadius: 0.08
      });
      s.addText(`${lbl}\n${val || "..."}`, {
        x: 0.3 + i * 2.35, y: 4.45, w: 2.1, h: 0.7,
        fontSize: 9, color: i % 2 === 0 ? C.white : C.navy,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 2
      });
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 12 — Participations / Visibilité (Traction)
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Participations / Concours / Visibilité", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide12 || {};
    const tractBoxes = [
      { num: "01", icon: icoTrophy, title: "Challenges & Programmes",   val: d.challenges  || "Participation aux concours d'innovation nationaux" },
      { num: "02", icon: icoBolt,   title: "Prix obtenus",              val: d.awards      || "Récompenses et reconnaissances obtenues" },
      { num: "03", icon: icoGlobe,  title: "Médias",                    val: d.media       || "Couverture presse, articles, reportages" },
      { num: "04", icon: icoUsers,  title: "Réseaux Sociaux",           val: d.social      || "Communauté en ligne et engagement digital" },
    ];

    tractBoxes.forEach((b, i) => {
      const bx = 0.3 + (i % 2) * 4.8;
      const by = i < 2 ? 1.0 : 3.0;
      addCard(s, pres, bx, by, 4.3, 1.75, C.light);
      s.addImage({ data: b.icon, x: bx + 0.15, y: by + 0.15, w: 0.5, h: 0.5 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: bx + 0.75, y: by + 0.1, w: 3.4, h: 0.35,
        fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.06
      });
      s.addText(`${b.num}  ${b.title}`, {
        x: bx + 0.75, y: by + 0.1, w: 3.4, h: 0.35,
        fontSize: 10.5, bold: true, color: C.gold,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0
      });
      s.addText(b.val, {
        x: bx + 0.15, y: by + 0.55, w: 4.0, h: 1.1,
        fontSize: 11, color: "333333", fontFace: "Calibri",
        valign: "top", margin: 4
      });
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 13 — Roadmap
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: "F5F5FF" };

    addHeader(s, "RoadMap — Stratégie de Développement", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide13 || {};
    s.addText("Feuille de route documentant le lancement et le développement du projet.", {
      x: 0.4, y: 0.92, w: 9.2, h: 0.28,
      fontSize: 10, color: "555555", italic: true,
      fontFace: "Calibri", align: "center", margin: 0
    });

    // Timeline line
    s.addShape(pres.shapes.LINE, {
      x: 0.7, y: 2.4, w: 8.6, h: 0,
      line: { color: C.navy, width: 3 }
    });

    const phases = d.phases || [
      { period: "M1-M2", label: "Étude & Validation", desc: "Validation du concept" },
      { period: "M3-M4", label: "MVP", desc: "Prototype minimum" },
      { period: "M5-M6", label: "Tests", desc: "Beta testeurs" },
      { period: "M7-M9", label: "Lancement", desc: "Mise sur le marché" },
      { period: "M10-M12", label: "Croissance", desc: "Scale & partenariats" },
    ];

    phases.slice(0, 5).forEach((ph, i) => {
      const px = 0.7 + i * 1.82;
      const above = i % 2 === 0;

      s.addShape(pres.shapes.OVAL, {
        x: px - 0.12, y: 2.28, w: 0.26, h: 0.26,
        fill: { color: C.gold }, line: { color: C.navy, width: 1.5 }
      });

      if (above) {
        s.addShape(pres.shapes.LINE, { x: px + 0.01, y: 1.55, w: 0, h: 0.73, line: { color: C.gold, width: 1, dashType: "dash" } });
        addCard(s, pres, px - 0.85, 0.97, 1.8, 0.6, C.white);
        s.addText(`${ph.period}\n${ph.label}`, { x: px - 0.85, y: 0.97, w: 1.8, h: 0.6, fontSize: 9.5, bold: true, color: C.navy, fontFace: "Calibri", align: "center", valign: "middle", margin: 2 });
        s.addText(ph.desc, { x: px - 0.85, y: 1.58, w: 1.8, h: 0.3, fontSize: 8.5, color: "555555", fontFace: "Calibri", align: "center", margin: 0 });
      } else {
        s.addShape(pres.shapes.LINE, { x: px + 0.01, y: 2.56, w: 0, h: 0.73, line: { color: C.gold, width: 1, dashType: "dash" } });
        s.addText(ph.desc, { x: px - 0.85, y: 2.58, w: 1.8, h: 0.3, fontSize: 8.5, color: "555555", fontFace: "Calibri", align: "center", margin: 0 });
        addCard(s, pres, px - 0.85, 2.9, 1.8, 0.6, C.white);
        s.addText(`${ph.period}\n${ph.label}`, { x: px - 0.85, y: 2.9, w: 1.8, h: 0.6, fontSize: 9.5, bold: true, color: C.navy, fontFace: "Calibri", align: "center", valign: "middle", margin: 2 });
      }
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 14 — Plan financier
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Plan Financier — Besoins & Projections", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide14 || {};

    // Funding needs card
    addCard(s, pres, 0.2, 0.97, 4.5, 1.8, C.light);
    s.addImage({ data: icoMoney, x: 0.35, y: 1.05, w: 0.6, h: 0.6 });
    s.addText("Besoins Actuels", {
      x: 1.05, y: 1.05, w: 3.5, h: 0.6,
      fontSize: 13, bold: true, color: C.navy, fontFace: "Calibri", valign: "middle", margin: 0
    });
    const needs = d.needs || ["Financement : montant et type voulu", "Ressources humaines : recrutement", "Formation : expertise métier", "Agrément / certification"];
    const needItems = needs.slice(0, 4).map((n, i) => ({
      text: `  ${n}`,
      options: { bullet: true, color: "333333", fontSize: 10, fontFace: "Calibri", breakLine: i < needs.length - 1 }
    }));
    s.addText(needItems, { x: 0.3, y: 1.72, w: 4.3, h: 0.95, margin: 4 });

    // Revenue projection chart
    const years = ["Année 1", "Année 2", "Année 3"];
    const revenues = d.revenues || [500000, 1500000, 4000000];
    s.addChart(pres.charts.BAR, [{
      name: "Chiffre d'affaires (DZD)",
      labels: years,
      values: revenues
    }], {
      x: 5.0, y: 0.97, w: 4.6, h: 3.5,
      barDir: "col",
      chartColors: [C.navy, C.navyL, C.gold],
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
      catAxisLabelColor: "444444",
      valAxisLabelColor: "444444",
      valGridLine: { color: "E2E8F0", size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true,
      dataLabelColor: "1E293B",
      showLegend: true,
      legendPos: "b",
      title: "Projections CA (DZD)",
      showTitle: true,
    });

    // Investment summary
    const investCards = [
      { label: "Investissement Initial", val: d.initialInvestment || "500 000 DZD" },
      { label: "CA Prévu An 1",          val: d.ca1 || "500 000 DZD" },
      { label: "Seuil de Rentabilité",   val: d.breakEven || "Mois 8-10" },
    ];
    investCards.forEach((ic, i) => {
      addCard(s, pres, 0.2 + i * 1.57, 2.85, 1.35, 0.85, C.navy === "0A1172" ? "E8EAF6" : C.light);
      s.addText(ic.label, {
        x: 0.25 + i * 1.57, y: 2.9, w: 1.25, h: 0.35,
        fontSize: 7.5, color: C.navy, fontFace: "Calibri",
        align: "center", bold: true, margin: 1
      });
      s.addText(ic.val, {
        x: 0.25 + i * 1.57, y: 3.25, w: 1.25, h: 0.35,
        fontSize: 10, color: C.navyM, fontFace: "Calibri",
        align: "center", bold: true, margin: 1
      });
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 15 — Cadre du projet / Attestation
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.white };

    addHeader(s, "Cadre du Projet — Attestation d'Incubation", pres);
    addLogoBadge(s, pres, logo);
    addFooter(s, pres, uni, inc);

    const d = slides.slide15 || {};

    // Incubator large card
    addCard(s, pres, 0.4, 1.0, 4.2, 3.8, "F0F4FF");
    s.addImage({ data: icoUni, x: 1.5, y: 1.2, w: 2.0, h: 2.0 });
    s.addText(inc, {
      x: 0.5, y: 3.3, w: 4.0, h: 0.4,
      fontSize: 14, bold: true, color: C.navy, fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText("Incubateur Universitaire", {
      x: 0.5, y: 3.72, w: 4.0, h: 0.3,
      fontSize: 10, color: C.teal, fontFace: "Calibri", align: "center", bold: true, margin: 0
    });

    // Right info
    addCard(s, pres, 4.9, 1.0, 4.7, 1.7, C.light);
    s.addImage({ data: icoGrad, x: 5.05, y: 1.1, w: 0.5, h: 0.5 });
    s.addText(uni, {
      x: 5.65, y: 1.1, w: 3.8, h: 0.5,
      fontSize: 12, bold: true, color: C.navy, fontFace: "Calibri", valign: "middle", margin: 0
    });
    s.addText(`Faculté : ${fac}\nDépartement : ${dept}\nSpécialité : ${spec}`, {
      x: 5.05, y: 1.68, w: 4.4, h: 0.85,
      fontSize: 10, color: "444444", fontFace: "Calibri", valign: "top", margin: 4
    });

    addCard(s, pres, 4.9, 2.85, 4.7, 1.95, C.light);
    s.addImage({ data: icoHand, x: 5.05, y: 2.95, w: 0.5, h: 0.5 });
    s.addText("Partenaire Industriel / Économique", {
      x: 5.65, y: 2.95, w: 3.8, h: 0.5,
      fontSize: 11, bold: true, color: C.navy, fontFace: "Calibri", valign: "middle", margin: 0
    });
    s.addText(d.industrialPartner || "À compléter — Nom du partenaire industriel ou économique", {
      x: 5.05, y: 3.52, w: 4.4, h: 1.1,
      fontSize: 10, color: "555555", fontFace: "Calibri", valign: "top", margin: 4
    });
  }

  // ════════════════════════════════════════════════════════════════
  // SLIDE 16 — Contact & CTA
  // ════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Top gold bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.18,
      fill: { color: C.gold }, line: { color: C.gold }
    });

    // Circle logo
    s.addShape(pres.shapes.OVAL, {
      x: 3.8, y: 0.35, w: 2.4, h: 2.4,
      fill: { color: C.navyL }, line: { color: C.gold, width: 3 },
      shadow: makeShadow()
    });
    s.addImage({ data: icoLight, x: 4.4, y: 0.75, w: 1.2, h: 1.2 });

    s.addText(data.projectTitle || "Startup", {
      x: 1, y: 2.85, w: 8, h: 0.65,
      fontSize: 32, bold: true, color: C.gold, fontFace: "Calibri", align: "center", margin: 0
    });
    s.addText(`« ${data.tagline || "Innovation • Impact • Solution"} »`, {
      x: 1, y: 3.5, w: 8, h: 0.32,
      fontSize: 13, italic: true, color: C.light, fontFace: "Calibri", align: "center", margin: 0
    });

    const contact = data.contact || {};
    const icoPhone = await iconB64(FaPhone,    "#" + C.gold, 128);
    const icoEmail = await iconB64(FaEnvelope, "#" + C.gold, 128);
    const icoFb    = await iconB64(FaFacebook, "#" + C.gold, 128);
    const icoIg    = await iconB64(FaInstagram,"#" + C.gold, 128);
    const icoYt    = await iconB64(FaYoutube,  "#" + C.gold, 128);
    const icoTw    = await iconB64(FaTwitter,  "#" + C.gold, 128);
    const icoW     = await iconB64(FaGlobe,    "#" + C.gold, 128);

    const contactItems = [
      { ico: icoPhone, val: contact.phone   || "+213 ..." },
      { ico: icoEmail, val: contact.email   || "contact@startup.dz" },
      { ico: icoW,     val: contact.website || "www.startup.dz" },
    ];
    contactItems.forEach((ci, i) => {
      s.addImage({ data: ci.ico, x: 1.5 + i * 2.5, y: 3.92, w: 0.3, h: 0.3 });
      s.addText(ci.val, {
        x: 1.85 + i * 2.5, y: 3.92, w: 2.0, h: 0.3,
        fontSize: 9.5, color: C.light, fontFace: "Calibri", valign: "middle", margin: 0
      });
    });

    const socialIcos = [icoFb, icoIg, icoYt, icoTw];
    const socialLabels = ["@facebook", "@instagram", "@youtube", "@twitter"];
    socialIcos.forEach((ico, i) => {
      s.addImage({ data: ico, x: 1.5 + i * 1.8, y: 4.35, w: 0.3, h: 0.3 });
      s.addText(socialLabels[i], {
        x: 1.85 + i * 1.8, y: 4.35, w: 1.4, h: 0.3,
        fontSize: 8.5, color: C.gray, fontFace: "Calibri", valign: "middle", margin: 0
      });
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 5.425, w: 10, h: 0.2,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(`${uni}  |  ${inc}  |  ${year}`, {
      x: 0, y: 5.425, w: 10, h: 0.2,
      fontSize: 7, color: C.navy, bold: true,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0
    });
  }

  // Write
  await pres.writeFile({ fileName: outputPath });
  // Rezip
  const { execSync } = require("child_process");
  try {
    execSync(`python /mnt/skills/public/pptx/scripts/rezip.py "${outputPath}"`, { stdio: "pipe" });
  } catch (_) {}

  return outputPath;
}

// CLI entry point
if (require.main === module) {
  const dataFile = process.argv[2];
  const outFile  = process.argv[3] || "/tmp/pitch_deck.pptx";
  if (!dataFile) { console.error("Usage: node generate_pptx.js <data.json> <output.pptx>"); process.exit(1); }
  const data = JSON.parse(fs.readFileSync(dataFile, "utf8"));
  generatePPTX(data, outFile)
    .then(() => { console.log("PPTX_OK:" + outFile); })
    .catch(e => { console.error("PPTX_ERR:" + e.message); process.exit(1); });
}

module.exports = { generatePPTX };
