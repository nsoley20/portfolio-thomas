const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "EKANG NSOLEY Sean Thomas";
pres.title = "Portfolio — Thomas EKANG NSOLEY";

// ── PALETTE ──
const C = {
  navy:    "0F172A",
  blue:    "1D4ED8",
  blueMid: "2563EB",
  blueLt:  "3B82F6",
  blue50:  "EFF6FF",
  blue100: "DBEAFE",
  red:     "DC2626",
  redLt:   "EF4444",
  red50:   "FEF2F2",
  white:   "FFFFFF",
  off:     "F8FAFC",
  gray100: "F1F5F9",
  gray200: "E2E8F0",
  gray300: "CBD5E1",
  gray400: "94A3B8",
  gray500: "64748B",
  gray600: "475569",
  gray700: "334155",
  gray800: "1E293B",
};

const makeShadow = () => ({ type:"outer", blur:8, offset:3, angle:135, color:"000000", opacity:0.12 });

// ══════════════════════════════════════════════
// SLIDE 1 — COVER
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Left blue stripe
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.15, h:5.625, fill:{ color:C.blue } });

  // Red bottom stripe
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.1, w:10, h:0.525, fill:{ color:C.red } });
  // White strip above red
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:4.95, w:10, h:0.16, fill:{ color:C.white } });

  // Monogram circle
  s.addShape(pres.shapes.OVAL, { x:0.55, y:0.6, w:1.5, h:1.5, fill:{ color:C.blue }, shadow:makeShadow() });
  s.addText("T", { x:0.55, y:0.6, w:1.5, h:1.5, fontSize:44, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Georgia", italic:true });

  // Name block
  s.addText("EKANG NSOLEY", { x:0.55, y:2.3, w:9, h:0.8, fontSize:42, bold:true, color:C.white, fontFace:"Georgia", margin:0 });
  s.addText("Sean Thomas Patrick Salem", { x:0.55, y:3.0, w:9, h:0.55, fontSize:20, color:C.gray400, fontFace:"Calibri", margin:0 });

  // Divider line
  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:3.62, w:2.2, h:0.04, fill:{ color:C.blueLt } });

  // Role tags row
  const tags = ["Développeur Fullstack", "Réseaux & Sécurité", "Gabon · Libreville"];
  let tx = 0.55;
  tags.forEach((tag, i) => {
    const bg = i === 0 ? C.blue : i === 1 ? C.red : C.gray700;
    const w = i === 2 ? 1.9 : i === 0 ? 2.35 : 2.1;
    s.addShape(pres.shapes.RECTANGLE, { x:tx, y:3.8, w:w, h:0.38, fill:{ color:bg }, rectRadius:0.06 });
    s.addText(tag, { x:tx, y:3.8, w:w, h:0.38, fontSize:11, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
    tx += w + 0.15;
  });

  // URL badge bottom
  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:4.52, w:3.5, h:0.38, fill:{ color:"00000040" }, rectRadius:0.08 });
  s.addText("nsoleyportfolio.netlify.app", { x:0.55, y:4.52, w:3.5, h:0.38, fontSize:11, color:C.blue100, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });

  // Right decorative grid dots
  for (let row = 0; row < 6; row++) {
    for (let col = 0; col < 6; col++) {
      s.addShape(pres.shapes.OVAL, { x:6.8+col*0.42, y:0.4+row*0.42, w:0.08, h:0.08, fill:{ color:"FFFFFF", transparency:80 } });
    }
  }

  // Bottom right label
  s.addText("Portfolio · 2024", { x:7.5, y:5.1, w:2.3, h:0.4, fontSize:10, color:C.white, align:"right", fontFace:"Calibri", opacity:0.6, margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 2 — PROFIL
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.white };

  // Top accent bar
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:C.blue } });

  // Section label
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:0.28, w:0.06, h:0.65, fill:{ color:C.blue } });
  s.addText("PROFIL", { x:0.8, y:0.25, w:2, h:0.35, fontSize:9, bold:true, color:C.blueLt, charSpacing:5, fontFace:"Calibri", margin:0 });
  s.addText("Qui suis-je ?", { x:0.8, y:0.56, w:8, h:0.55, fontSize:28, bold:true, color:C.gray800, fontFace:"Georgia", italic:true, margin:0 });

  // Avatar placeholder card
  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.35, w:2.4, h:3.6, fill:{ color:C.blue }, rectRadius:0.12, shadow:makeShadow() });
  s.addShape(pres.shapes.OVAL, { x:1.05, y:1.7, w:1.4, h:1.4, fill:{ color:"FFFFFF" }, shadow:makeShadow() });
  s.addText("T", { x:1.05, y:1.7, w:1.4, h:1.4, fontSize:40, bold:true, color:C.blue, align:"center", valign:"middle", fontFace:"Georgia", italic:true });
  s.addText("Thomas", { x:0.55, y:3.2, w:2.4, h:0.45, fontSize:15, bold:true, color:C.white, align:"center", fontFace:"Georgia", italic:true });
  s.addText("EKANG NSOLEY", { x:0.55, y:3.6, w:2.4, h:0.35, fontSize:10, color:"BFDBFE", align:"center", fontFace:"Calibri" });
  // online badge
  s.addShape(pres.shapes.OVAL, { x:1.58, y:4.2, w:0.12, h:0.12, fill:{ color:"22C55E" } });
  s.addText("En ligne — nsoleyportfolio.netlify.app", { x:0.5, y:4.1, w:2.5, h:0.4, fontSize:8, color:C.white, align:"center", fontFace:"Calibri", opacity:0.7 });

  // Right content
  const items = [
    { label:"Nom complet", val:"EKANG NSOLEY Sean Thomas Patrick Salem", color:C.gray800 },
    { label:"Nom d'usage", val:"Thomas", color:C.blue },
    { label:"Diplôme", val:"Technicien Supérieur Téléinformatique — Génie Logiciel & Adm. Réseau", color:C.gray700 },
    { label:"Spécialisation", val:"Master 1 — Cybersécurité, Réseaux & Cloud (en cours)", color:C.red },
    { label:"Localisation", val:"Libreville, Gabon", color:C.gray700 },
    { label:"Portfolio", val:"nsoleyportfolio.netlify.app", color:C.blueLt },
  ];
  items.forEach((it, i) => {
    const y = 1.35 + i * 0.57;
    s.addShape(pres.shapes.RECTANGLE, { x:3.25, y, w:6.2, h:0.5, fill:{ color: i%2===0 ? C.off : C.white }, rectRadius:0.06, line:{ color:C.gray200, width:0.5 } });
    s.addText(it.label, { x:3.4, y: y+0.05, w:1.4, h:0.2, fontSize:8, bold:true, color:C.gray400, fontFace:"Calibri", margin:0 });
    s.addText(it.val, { x:3.4, y: y+0.22, w:5.9, h:0.22, fontSize:11, bold: it.label==="Nom d'usage"||it.label==="Spécialisation", color:it.color, fontFace:"Calibri", margin:0 });
  });

  // Bottom red bar
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.35, w:10, h:0.275, fill:{ color:C.red } });
  s.addText("Portfolio · Thomas EKANG NSOLEY", { x:0, y:5.35, w:10, h:0.275, fontSize:9, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 3 — COMPÉTENCES
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.off };

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:C.blue } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:0.28, w:0.06, h:0.65, fill:{ color:C.red } });
  s.addText("COMPÉTENCES", { x:0.8, y:0.25, w:3, h:0.35, fontSize:9, bold:true, color:C.red, charSpacing:5, fontFace:"Calibri", margin:0 });
  s.addText("Stack technique", { x:0.8, y:0.56, w:8, h:0.55, fontSize:28, bold:true, color:C.gray800, fontFace:"Georgia", italic:true, margin:0 });

  const cats = [
    { title:"Frontend", items:["HTML5", "CSS3", "JavaScript", "TypeScript (notions)", "Responsive Design", "UX/UI"], bg:C.blue },
    { title:"Backend", items:["Node.js", "Express.js", "API REST", "Middleware", "Python (notions)"], bg:C.blueMid },
    { title:"Base de données", items:["MongoDB", "Mongoose", "Schémas & CRUD", "Modélisation"], bg:C.blueLt },
    { title:"Cybersécurité", items:["Auth JWT", "Bcrypt / Hachage", "RBAC", "OWASP", "Sécurité réseau"], bg:C.red },
    { title:"Réseaux & Systèmes", items:["TCP/IP, DNS, DHCP", "Firewall · VPN", "Linux (Ubuntu, CentOS)", "Monitoring infra"], bg:"B91C1C" },
    { title:"Architecture", items:["Conception architecture", "Modélisation UML", "Cahiers des charges", "Git / GitHub", "Docker (notions)"], bg:C.gray700 },
  ];

  cats.forEach((cat, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 3.1;
    const y = 1.3 + row * 2.0;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:2.95, h:1.8, fill:{ color:C.white }, rectRadius:0.1, shadow:makeShadow(), line:{ color:C.gray200, width:0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:2.95, h:0.42, fill:{ color:cat.bg }, rectRadius:0.1 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:y+0.28, w:2.95, h:0.15, fill:{ color:cat.bg } });
    s.addText(cat.title, { x: x+0.12, y, w:2.75, h:0.42, fontSize:11, bold:true, color:C.white, valign:"middle", fontFace:"Calibri", margin:0 });
    s.addText(cat.items.map(i => i).join("\n"), {
      x: x+0.12, y: y+0.5, w:2.75, h:1.22,
      fontSize:9.5, color:C.gray600, fontFace:"Calibri",
      bullet:false, margin:0, lineSpacingMultiple:1.3,
    });
    // small dots for each item
    cat.items.forEach((item, j) => {
      if (j < 5) {
        s.addShape(pres.shapes.OVAL, { x: x+0.12, y: y+0.56+j*0.235, w:0.07, h:0.07, fill:{ color:cat.bg } });
      }
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.35, w:10, h:0.275, fill:{ color:C.blue } });
  s.addText("Portfolio · Thomas EKANG NSOLEY", { x:0, y:5.35, w:10, h:0.275, fontSize:9, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 4 — PROJET BAMBOO EMF
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:C.blue } });

  // Project number chip
  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:0.25, w:0.9, h:0.4, fill:{ color:C.blue }, rectRadius:0.06 });
  s.addText("01 / 03", { x:0.55, y:0.25, w:0.9, h:0.4, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:1.6, y:0.25, w:1.6, h:0.4, fill:{ color:C.gray700 }, rectRadius:0.06 });
  s.addText("RH · Fullstack", { x:1.6, y:0.25, w:1.6, h:0.4, fontSize:9, color:C.gray300, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });

  s.addText("BAMBOO EMF", { x:0.55, y:0.82, w:9, h:0.72, fontSize:34, bold:true, color:C.white, fontFace:"Georgia", italic:true, margin:0 });
  s.addText("Gestion de Personnel SFE", { x:0.55, y:1.5, w:9, h:0.45, fontSize:18, color:C.blue100, fontFace:"Calibri", margin:0 });

  // Left description card
  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.1, w:5.0, h:2.8, fill:{ color:"FFFFFF0D" }, rectRadius:0.12, line:{ color:"FFFFFF15", width:0.8 } });
  s.addText("Description", { x:0.75, y:2.25, w:4.6, h:0.3, fontSize:9, bold:true, color:C.blue100, charSpacing:4, fontFace:"Calibri", margin:0 });
  s.addText(
    "Application web complète de gestion RH pour SFE. Centralise toutes les opérations RH : employés, départements, présences et performances.\n\nArchitecture fullstack JavaScript avec backend Node.js/Express sécurisé par JWT et RBAC. Frontend multi-pages avec dashboard analytics (Chart.js) et design responsive.",
    { x:0.75, y:2.62, w:4.6, h:2.1, fontSize:11, color:"CBD5E1", fontFace:"Calibri", margin:0, lineSpacingMultiple:1.5 }
  );

  // Right: features
  const feats = [
    { t:"Dashboard Analytics", d:"5 graphiques Chart.js — temps réel" },
    { t:"Auth JWT + RBAC", d:"Tokens signés HS256, rôles admin/manager" },
    { t:"CRUD Employés", d:"Ajout, édition, suppression, export CSV/PDF" },
    { t:"Suivi de Présence", d:"Marquage quotidien et historique" },
  ];
  feats.forEach((f, i) => {
    const y = 2.1 + i * 0.68;
    s.addShape(pres.shapes.RECTANGLE, { x:5.9, y, w:3.65, h:0.6, fill:{ color:"FFFFFF08" }, rectRadius:0.08, line:{ color:"FFFFFF12", width:0.6 } });
    s.addShape(pres.shapes.OVAL, { x:6.05, y: y+0.2, w:0.2, h:0.2, fill:{ color:C.blue } });
    s.addText(f.t, { x:6.35, y: y+0.05, w:3.1, h:0.25, fontSize:10.5, bold:true, color:C.white, fontFace:"Calibri", margin:0 });
    s.addText(f.d, { x:6.35, y: y+0.3, w:3.1, h:0.22, fontSize:9, color:C.gray400, fontFace:"Calibri", margin:0 });
  });

  // Tech tags
  const tags = ["Node.js","Express","MongoDB","JWT","Chart.js","Bcrypt","Helmet","RBAC"];
  tags.forEach((tag, i) => {
    const x = 0.55 + (i%4)*1.4;
    const y = 5.0 + Math.floor(i/4)*0.0;
    s.addShape(pres.shapes.RECTANGLE, { x, y:4.98, w:1.3, h:0.28, fill:{ color:C.blue, transparency:70 }, rectRadius:0.04 });
    s.addText(tag, { x, y:4.98, w:1.3, h:0.28, fontSize:8.5, bold:true, color:C.blue100, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
  });

  // URL
  s.addShape(pres.shapes.RECTANGLE, { x:5.9, y:4.85, w:3.65, h:0.38, fill:{ color:C.blue }, rectRadius:0.07 });
  s.addText("nsoleyportfolio.netlify.app/bamboo-app", { x:5.9, y:4.85, w:3.65, h:0.38, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.35, w:10, h:0.275, fill:{ color:C.red } });
  s.addText("Portfolio · Thomas EKANG NSOLEY", { x:0, y:5.35, w:10, h:0.275, fontSize:9, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 5 — PROJET LINAF
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:C.red } });

  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:0.25, w:0.9, h:0.4, fill:{ color:C.red }, rectRadius:0.06 });
  s.addText("02 / 03", { x:0.55, y:0.25, w:0.9, h:0.4, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:1.6, y:0.25, w:1.8, h:0.4, fill:{ color:C.gray700 }, rectRadius:0.06 });
  s.addText("Finance · Web App", { x:1.6, y:0.25, w:1.8, h:0.4, fontSize:9, color:C.gray300, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });

  s.addText("LINAF", { x:0.55, y:0.82, w:9, h:0.72, fontSize:42, bold:true, color:C.white, fontFace:"Georgia", italic:true, margin:0 });
  s.addText("Plateforme de Microfinance", { x:0.55, y:1.5, w:9, h:0.45, fontSize:18, color:"FCA5A5", fontFace:"Calibri", margin:0 });

  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.1, w:5.0, h:2.8, fill:{ color:"FFFFFF0D" }, rectRadius:0.12, line:{ color:"FFFFFF15", width:0.8 } });
  s.addText("Description", { x:0.75, y:2.25, w:4.6, h:0.3, fontSize:9, bold:true, color:"FCA5A5", charSpacing:4, fontFace:"Calibri", margin:0 });
  s.addText(
    "Site web institutionnel pour une institution de microfinance. Présentation des services financiers, gestion des demandes de prêts, espace client sécurisé et interface d'administration.\n\nDesign professionnel responsive aux couleurs de l'institution, avec formulaires de contact et pages de services détaillées.",
    { x:0.75, y:2.62, w:4.6, h:2.1, fontSize:11, color:"CBD5E1", fontFace:"Calibri", margin:0, lineSpacingMultiple:1.5 }
  );

  const feats = [
    { t:"Présentation Services", d:"Pages produits et offres de prêts" },
    { t:"Design Responsive", d:"Compatible mobile, tablette, desktop" },
    { t:"Formulaires de Contact", d:"Demandes de prêt intégrées" },
    { t:"Interface Admin", d:"Gestion des demandes clients" },
  ];
  feats.forEach((f, i) => {
    const y = 2.1 + i * 0.68;
    s.addShape(pres.shapes.RECTANGLE, { x:5.9, y, w:3.65, h:0.6, fill:{ color:"FFFFFF08" }, rectRadius:0.08, line:{ color:"FFFFFF12", width:0.6 } });
    s.addShape(pres.shapes.OVAL, { x:6.05, y: y+0.2, w:0.2, h:0.2, fill:{ color:C.red } });
    s.addText(f.t, { x:6.35, y: y+0.05, w:3.1, h:0.25, fontSize:10.5, bold:true, color:C.white, fontFace:"Calibri", margin:0 });
    s.addText(f.d, { x:6.35, y: y+0.3, w:3.1, h:0.22, fontSize:9, color:C.gray400, fontFace:"Calibri", margin:0 });
  });

  const tags = ["HTML5","CSS3","JavaScript","Node.js","Responsive","UX/UI","Git","Design"];
  tags.forEach((tag, i) => {
    const x = 0.55 + (i%4)*1.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y:4.98, w:1.3, h:0.28, fill:{ color:C.red, transparency:70 }, rectRadius:0.04 });
    s.addText(tag, { x, y:4.98, w:1.3, h:0.28, fontSize:8.5, bold:true, color:"FCA5A5", align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:5.9, y:4.85, w:3.65, h:0.38, fill:{ color:C.red }, rectRadius:0.07 });
  s.addText("linafp-billetterie.netlify.app", { x:5.9, y:4.85, w:3.65, h:0.38, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.35, w:10, h:0.275, fill:{ color:C.blue } });
  s.addText("Portfolio · Thomas EKANG NSOLEY", { x:0, y:5.35, w:10, h:0.275, fontSize:9, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 6 — PROJET OTAKU
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:"6D28D9" } });

  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:0.25, w:0.9, h:0.4, fill:{ color:"6D28D9" }, rectRadius:0.06 });
  s.addText("03 / 03", { x:0.55, y:0.25, w:0.9, h:0.4, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:1.6, y:0.25, w:2.0, h:0.4, fill:{ color:C.gray700 }, rectRadius:0.06 });
  s.addText("Anime · Catalogue Web", { x:1.6, y:0.25, w:2.0, h:0.4, fontSize:9, color:C.gray300, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });

  s.addText("OTAKU", { x:0.55, y:0.82, w:9, h:0.72, fontSize:42, bold:true, color:C.white, fontFace:"Georgia", italic:true, margin:0 });
  s.addText("Catalogue Anime & Manga", { x:0.55, y:1.5, w:9, h:0.45, fontSize:18, color:"C4B5FD", fontFace:"Calibri", margin:0 });

  s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.1, w:5.0, h:2.8, fill:{ color:"FFFFFF0D" }, rectRadius:0.12, line:{ color:"FFFFFF15", width:0.8 } });
  s.addText("Description", { x:0.75, y:2.25, w:4.6, h:0.3, fontSize:9, bold:true, color:"C4B5FD", charSpacing:4, fontFace:"Calibri", margin:0 });
  s.addText(
    "Application web de catalogue d'animes et mangas avec moteur de recherche avancé, système de favoris persistants et fiches détaillées.\n\nInterface inspirée de l'esthétique japonaise moderne. Intégration d'API REST pour les données, navigation fluide et expérience utilisateur soignée.",
    { x:0.75, y:2.62, w:4.6, h:2.1, fontSize:11, color:"CBD5E1", fontFace:"Calibri", margin:0, lineSpacingMultiple:1.5 }
  );

  const feats = [
    { t:"Moteur de Recherche", d:"Filtrage par titre, genre, année" },
    { t:"Système de Favoris", d:"Persistance via LocalStorage" },
    { t:"Fiches Détaillées", d:"Synopsis, personnages, scores" },
    { t:"Intégration API", d:"Données temps réel via API REST" },
  ];
  feats.forEach((f, i) => {
    const y = 2.1 + i * 0.68;
    s.addShape(pres.shapes.RECTANGLE, { x:5.9, y, w:3.65, h:0.6, fill:{ color:"FFFFFF08" }, rectRadius:0.08, line:{ color:"FFFFFF12", width:0.6 } });
    s.addShape(pres.shapes.OVAL, { x:6.05, y: y+0.2, w:0.2, h:0.2, fill:{ color:"6D28D9" } });
    s.addText(f.t, { x:6.35, y: y+0.05, w:3.1, h:0.25, fontSize:10.5, bold:true, color:C.white, fontFace:"Calibri", margin:0 });
    s.addText(f.d, { x:6.35, y: y+0.3, w:3.1, h:0.22, fontSize:9, color:C.gray400, fontFace:"Calibri", margin:0 });
  });

  const tags = ["HTML5","CSS3","JavaScript","API REST","LocalStorage","Fetch API","Git","UI Design"];
  tags.forEach((tag, i) => {
    const x = 0.55 + (i%4)*1.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y:4.98, w:1.3, h:0.28, fill:{ color:"6D28D9", transparency:70 }, rectRadius:0.04 });
    s.addText(tag, { x, y:4.98, w:1.3, h:0.28, fontSize:8.5, bold:true, color:"C4B5FD", align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:5.9, y:4.85, w:3.65, h:0.38, fill:{ color:"6D28D9" }, rectRadius:0.07 });
  s.addText("otaku-site.netlify.app", { x:5.9, y:4.85, w:3.65, h:0.38, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.35, w:10, h:0.275, fill:{ color:C.red } });
  s.addText("Portfolio · Thomas EKANG NSOLEY", { x:0, y:5.35, w:10, h:0.275, fontSize:9, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 7 — PARCOURS
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.white };

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:C.blue } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:0.28, w:0.06, h:0.65, fill:{ color:C.blue } });
  s.addText("PARCOURS", { x:0.8, y:0.25, w:3, h:0.35, fontSize:9, bold:true, color:C.blueLt, charSpacing:5, fontFace:"Calibri", margin:0 });
  s.addText("Formation & Expérience", { x:0.8, y:0.56, w:8, h:0.55, fontSize:28, bold:true, color:C.gray800, fontFace:"Georgia", italic:true, margin:0 });

  // Timeline vertical line
  s.addShape(pres.shapes.RECTANGLE, { x:0.95, y:1.35, w:0.04, h:3.8, fill:{ color:C.blue100 } });

  const items = [
    { period:"2024 – 2026", title:"Master 1 — Cybersécurité, Réseaux & Cloud", org:"En spécialisation", desc:"Sécurité applicative, architectures cloud, incident response", dot:C.blue, badge:"En cours", badgeCol:C.blue },
    { period:"2026 (Stage)", title:"Stage pratique — CHUO · Ministère de la Santé", org:"Déploiement logiciel hospitalier · Gabon", desc:"Déploiement et test d'un logiciel pour structures hospitalières", dot:C.red, badge:"Stage", badgeCol:C.red },
    { period:"2023 – Présent", title:"Développeur Fullstack Freelance", org:"Libreville, Gabon", desc:"Apps web complètes, APIs REST sécurisées, systèmes d'information", dot:C.blue, badge:"Actif", badgeCol:C.blue },
    { period:"Diplômé 2024", title:"Technicien Supérieur Téléinformatique", org:"Option Génie Logiciel & Administration Réseau", desc:"Génie logiciel, réseaux, sécurité des systèmes", dot:C.gray400, badge:"Diplôme", badgeCol:C.gray500 },
    { period:"2022 · 2 mois", title:"Stage — Administration Réseau", org:"Opérateur réseau · Libreville", desc:"Linux, pare-feu, monitoring, VLAN", dot:C.gray400, badge:"Stage", badgeCol:C.gray500 },
  ];

  items.forEach((item, i) => {
    const y = 1.35 + i * 0.74;
    s.addShape(pres.shapes.OVAL, { x:0.82, y: y+0.06, w:0.26, h:0.26, fill:{ color:item.dot }, shadow:makeShadow() });
    s.addText(item.period, { x:1.3, y, w:1.8, h:0.28, fontSize:8.5, bold:true, color:item.dot, fontFace:"Consolas", margin:0 });
    s.addText(item.title, { x:1.3, y: y+0.22, w:5.8, h:0.28, fontSize:11, bold:true, color:C.gray800, fontFace:"Calibri", margin:0 });
    s.addText(item.org, { x:1.3, y: y+0.46, w:5.5, h:0.2, fontSize:9, color:C.gray500, fontFace:"Calibri", margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x:7.35, y: y+0.15, w:1.8, h:0.28, fill:{ color:item.badgeCol, transparency:85 }, rectRadius:0.06, line:{ color:item.badgeCol, width:0.5 } });
    s.addText(item.badge, { x:7.35, y: y+0.15, w:1.8, h:0.28, fontSize:8.5, bold:true, color:item.badgeCol, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.35, w:10, h:0.275, fill:{ color:C.red } });
  s.addText("Portfolio · Thomas EKANG NSOLEY", { x:0, y:5.35, w:10, h:0.275, fontSize:9, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 8 — CONTACT / CLOSING
// ══════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color:C.navy };

  // Tricolor bands bottom
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.0, w:3.33, h:0.625, fill:{ color:C.blue } });
  s.addShape(pres.shapes.RECTANGLE, { x:3.33, y:5.0, w:3.34, h:0.625, fill:{ color:C.white } });
  s.addShape(pres.shapes.RECTANGLE, { x:6.67, y:5.0, w:3.33, h:0.625, fill:{ color:C.red } });

  // Blue label
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{ color:C.blue } });

  // Avatar
  s.addShape(pres.shapes.OVAL, { x:0.55, y:0.42, w:1.4, h:1.4, fill:{ color:C.blue }, shadow:makeShadow() });
  s.addText("T", { x:0.55, y:0.42, w:1.4, h:1.4, fontSize:38, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Georgia", italic:true });

  // Title block
  s.addText("Travaillons", { x:2.2, y:0.32, w:7.5, h:0.65, fontSize:34, bold:true, color:C.white, fontFace:"Georgia", italic:true, margin:0 });
  s.addText("ensemble", { x:2.2, y:0.9, w:7.5, h:0.65, fontSize:34, bold:true, color:C.blueLt, fontFace:"Georgia", italic:true, margin:0 });

  // Contact items
  const contacts = [
    { label:"Email", val:"thomasean2002@gmail.com  ·  nsoleypro20@gmail.com", icon:"@", bg:C.blue },
    { label:"LinkedIn", val:"linkedin.com/in/thomas-nsoley-253076247", icon:"in", bg:"0A66C2" },
    { label:"GitHub", val:"github.com/thomasnsoley  ·  github.com/nsoley20", icon:"</>" , bg:C.gray700 },
    { label:"Portfolio", val:"nsoleyportfolio.netlify.app", icon:"www", bg:C.red },
  ];

  contacts.forEach((c, i) => {
    const y = 1.8 + i * 0.72;
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y, w:9.0, h:0.62, fill:{ color:"FFFFFF08" }, rectRadius:0.1, line:{ color:"FFFFFF12", width:0.6 } });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y, w:0.9, h:0.62, fill:{ color:c.bg }, rectRadius:0.1 });
    s.addShape(pres.shapes.RECTANGLE, { x:1.15, y, w:0.32, h:0.62, fill:{ color:c.bg } });
    s.addText(c.icon, { x:0.55, y, w:0.9, h:0.62, fontSize:11, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Consolas", margin:0 });
    s.addText(c.label, { x:1.55, y: y+0.06, w:1.5, h:0.22, fontSize:8, bold:true, color:C.gray400, charSpacing:3, fontFace:"Calibri", margin:0 });
    s.addText(c.val, { x:1.55, y: y+0.3, w:7.8, h:0.26, fontSize:11, color:C.white, fontFace:"Calibri", margin:0 });
  });

  // Footer text on tricolor
  s.addText("DISPONIBLE", { x:0, y:5.0, w:3.33, h:0.625, fontSize:11, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", charSpacing:4, margin:0 });
  s.addText("MERCI", { x:3.33, y:5.0, w:3.34, h:0.625, fontSize:11, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Georgia", italic:true, charSpacing:4, margin:0 });
  s.addText("THOMAS EKANG", { x:6.67, y:5.0, w:3.33, h:0.625, fontSize:9, bold:true, color:C.white, align:"center", valign:"middle", fontFace:"Calibri", charSpacing:3, margin:0 });
}

// ══════════════════════════════════════════════
// WRITE
// ══════════════════════════════════════════════
pres.writeFile({ fileName: "Portfolio_Thomas_EKANG_NSOLEY.pptx" })
  .then(() => console.log("DONE"))
  .catch(e => console.error(e));