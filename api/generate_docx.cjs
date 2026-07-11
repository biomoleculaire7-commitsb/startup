// generate_docx.js — دليل المشروع الرسمي 6 محاور — Word قابل للتعديل
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat,
  BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak, ExternalHyperlink
} = require("docx");
const fs = require("fs");

const NAVY  = "0A1172";
const GOLD  = "C9A840";
const LIGHT = "E8EAF6";
const GRAY  = "757575";
const WHITE = "FFFFFF";
const DARK  = "1A1A2E";

function cell(text, opts = {}) {
  const { bg = WHITE, bold = false, color = "222222", align = AlignmentType.LEFT,
          colSpan, fontSize = 20, italic = false } = opts;
  return new TableCell({
    columnSpan: colSpan,
    verticalAlign: VerticalAlign.CENTER,
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      left:   { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      right:  { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
    },
    children: [
      new Paragraph({
        alignment: align,
        children: [new TextRun({ text, bold, color, size: fontSize, font: "Calibri", italics: italic })],
      })
    ]
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 180 },
    border: { bottom: { style: BorderStyle.THICK, size: 6, color: NAVY, space: 1 } },
    children: [new TextRun({ text, bold: true, color: WHITE, font: "Calibri", size: 36 })],
    shading: { fill: NAVY, type: ShadingType.CLEAR },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, color: NAVY, font: "Calibri", size: 28 })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: GOLD, space: 1 } },
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 140, after: 60 },
    children: [new TextRun({ text, bold: true, color: "444488", font: "Calibri", size: 24 })],
  });
}

function para(text, opts = {}) {
  const { bold = false, color = "222222", align = AlignmentType.LEFT,
          before = 60, after = 60, size = 22, italic = false } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before, after },
    children: [new TextRun({ text, bold, color, size, font: "Calibri", italics: italic })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "main-bullets", level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Calibri", color: "333333" })],
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: GOLD, space: 1 } },
    children: [],
  });
}

function labeledBlock(label, value) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2500, 6526],
    rows: [
      new TableRow({
        children: [
          cell(label, { bg: LIGHT, bold: true, color: NAVY, fontSize: 20, width: 2500 }),
          cell(value || "À compléter", { fontSize: 20, width: 6526 }),
        ]
      })
    ]
  });
}

function sectionHeader(axeNum, title, subtitle) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [
      new TableRow({
        children: [
          cell(`AXE ${axeNum} — ${title.toUpperCase()}`, {
            bg: NAVY, bold: true, color: WHITE, fontSize: 30,
            align: AlignmentType.CENTER, colSpan: 1
          })
        ]
      }),
      new TableRow({
        children: [
          cell(subtitle, { bg: GOLD, color: DARK, fontSize: 20, align: AlignmentType.CENTER })
        ]
      })
    ]
  });
}

// ── MAIN ──────────────────────────────────────────────────────────────────
async function generateDOCX(data, outputPath) {
  const uni  = data.universityName  || "Université ...";
  const inc  = data.incubatorName   || "Incubateur ...";
  const fac  = data.facultyName     || "Faculté ...";
  const dept = data.departmentName  || "Département ...";
  const spec = data.speciality      || "Spécialité ...";
  const year = data.academicYear    || "2025/2026";
  const members = data.teamMembers  || [];
  const g  = data.guide || {};

  const doc = new Document({
    creator: "StartupAI — Diplôme/Startup 1275",
    title: data.projectTitle,
    description: "Guide du Projet — Arrêté Ministériel 1275 — CNCSIIU",
    numbering: {
      config: [
        {
          reference: "main-bullets",
          levels: [
            { level: 0, format: LevelFormat.BULLET, text: "\u25CF", alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
            { level: 1, format: LevelFormat.BULLET, text: "\u2013", alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
          ]
        },
        {
          reference: "numbered",
          levels: [{
            level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }]
        }
      ]
    },
    styles: {
      default: { document: { run: { font: "Calibri", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 36, bold: true, font: "Calibri", color: WHITE },
          paragraph: { spacing: { before: 300, after: 180 }, outlineLevel: 0,
            shading: { fill: NAVY, type: ShadingType.CLEAR } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Calibri", color: NAVY },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Calibri", color: "444488" },
          paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 2 } },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1200, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [
            new Table({
              width: { size: 9026, type: WidthType.DXA },
              columnWidths: [6000, 3026],
              borders: { top: {style:BorderStyle.NONE}, bottom:{style:BorderStyle.SINGLE,size:3,color:GOLD}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} },
              rows: [new TableRow({ children: [
                new TableCell({
                  shading:{fill:WHITE,type:ShadingType.CLEAR},
                  margins:{top:40,bottom:40,left:0,right:40},
                  children:[new Paragraph({children:[new TextRun({text:data.projectTitle||"Guide du Projet",bold:true,size:20,color:NAVY,font:"Calibri"})]})]
                }),
                new TableCell({
                  shading:{fill:WHITE,type:ShadingType.CLEAR},
                  margins:{top:40,bottom:40,left:40,right:0},
                  children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:`${uni} | ${inc}`,size:16,color:GRAY,font:"Calibri"})]})]
                })
              ]})]
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 2, color: GOLD, space: 1 } },
              children: [
                new TextRun({ text: `Guide du Projet — Arrêté Ministériel 1275 — CNCSIIU — ${year}    `, size: 16, color: GRAY, font: "Calibri" }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: NAVY, bold: true, font: "Calibri" }),
                new TextRun({ text: " / ", size: 16, color: GRAY, font: "Calibri" }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: GRAY, font: "Calibri" }),
              ]
            })
          ]
        })
      },
      children: [
        // ── PAGE DE GARDE ────────────────────────────────────────
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [9026],
          rows: [
            new TableRow({ children: [
              cell("République Algérienne Démocratique et Populaire\nMinistère de l'Enseignement Supérieur et de la Recherche Scientifique",
                { bg: NAVY, color: WHITE, bold: true, fontSize: 22, align: AlignmentType.CENTER })
            ]}),
            new TableRow({ children: [
              cell(uni, { bg: "E8EAF6", color: NAVY, bold: true, fontSize: 26, align: AlignmentType.CENTER })
            ]}),
            new TableRow({ children: [
              cell(`Faculté : ${fac}   |   Département : ${dept}`, { fontSize: 20, align: AlignmentType.CENTER, color: "444444" })
            ]}),
          ]
        }),
        new Paragraph({ spacing: { before: 200, after: 200 }, children: [] }),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [9026],
          rows: [
            new TableRow({ children: [
              cell("Guide du Projet", { bg: GOLD, color: DARK, bold: true, fontSize: 36, align: AlignmentType.CENTER })
            ]}),
            new TableRow({ children: [
              cell(data.projectTitle || "Titre du projet innovant", { bg: NAVY, color: WHITE, bold: true, fontSize: 32, align: AlignmentType.CENTER })
            ]}),
            new TableRow({ children: [
              cell(`Spécialité : ${spec}`, { fontSize: 22, align: AlignmentType.CENTER, color: "444444" })
            ]}),
            new TableRow({ children: [
              cell("En vue de l'obtention du diplôme — Décision 008 : Diplôme Institution Économique", { fontSize: 20, align: AlignmentType.CENTER, italic: true, color: "555555" })
            ]}),
            new TableRow({ children: [
              cell(`Incubateur : ${inc}`, { bg: LIGHT, color: NAVY, bold: true, fontSize: 22, align: AlignmentType.CENTER })
            ]}),
          ]
        }),
        new Paragraph({ spacing: { before: 300 }, children: [] }),

        // Équipe de supervision
        new Paragraph({ spacing:{before:100,after:60}, children:[new TextRun({text:"Équipe d'encadrement :", bold:true,size:24,color:NAVY,font:"Calibri"})]}),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[3000,3000,3026],
          rows:[
            new TableRow({children:[
              cell("Encadrant principal", {bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:3000}),
              cell(data.supervisor||"Nom et prénom",{fontSize:20,width:3000}),
              cell(`Spécialité : ${spec}`,{fontSize:18,color:"555555",width:3026}),
            ]}),
            new TableRow({children:[
              cell("Co-encadrant 01",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:3000}),
              cell(data.cosupervisor1||"Nom et prénom",{fontSize:20,width:3000}),
              cell("Spécialité : ...",{fontSize:18,color:"555555",width:3026}),
            ]}),
            new TableRow({children:[
              cell("Co-encadrant 02",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:3000}),
              cell(data.cosupervisor2||"Nom et prénom",{fontSize:20,width:3000}),
              cell("Spécialité : ...",{fontSize:18,color:"555555",width:3026}),
            ]}),
          ]
        }),
        new Paragraph({ spacing:{before:200,after:60}, children:[new TextRun({text:"Équipe de projet :", bold:true,size:24,color:NAVY,font:"Calibri"})]}),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[3000,3000,3026],
          rows:[
            new TableRow({children:[
              cell("Étudiant / Nom et prénom",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3000}),
              cell("Faculté",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3000}),
              cell("Spécialité",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3026}),
            ]}),
            ...(members.length > 0 ? members : [{name:"Étudiant 01",faculty:fac,background:spec},{name:"Étudiant 02",faculty:fac,background:spec},{name:"Étudiant 03",faculty:fac,background:spec}]).map((m,i)=>
              new TableRow({children:[
                cell(m.name||`Étudiant ${i+1}`,{fontSize:20,width:3000}),
                cell(m.faculty||fac,{fontSize:18,color:"444444",width:3000}),
                cell(m.background||spec,{fontSize:18,color:"444444",width:3026}),
              ]})
            ),
          ]
        }),
        para(`Année universitaire : ${year}`, { align: AlignmentType.CENTER, bold: true, color: NAVY, before: 200, after: 100 }),
        new Paragraph({ children: [new PageBreak()] }),

        // ── AXE 1 ────────────────────────────────────────────────
        sectionHeader(1, "Présentation du Projet", "L'idée, les valeurs, l'équipe, les objectifs, le calendrier"),
        new Paragraph({ spacing:{before:200}, children:[] }),

        heading2("1. L'idée de Projet (Solution Proposée)"),
        para("Décrire le projet en abordant :"),
        bullet("Le domaine d'activité (services, numérique, industriel, agricole...)"),
        bullet("Comment l'idée a-t-elle germé et comment s'est-elle développée ?"),
        bullet("Qu'est-ce que vous allez faire ?"),
        bullet("Comment cela se passera-t-il ?"),
        bullet("Qui l'accomplira ?"),
        bullet("Où sera-t-il accompli ?"),
        new Paragraph({ spacing:{before:100,after:60}, children:[] }),
        labeledBlock("Domaine d'activité", g.ax1_domain || "À compléter"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Description de l'idée", g.ax1_idea || "À compléter — Décrivez comment l'idée a émergé et comment elle s'est développée."),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Comment ?", g.ax1_how || "À compléter — Décrivez le processus de réalisation."),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Qui et Où ?", g.ax1_who_where || "À compléter"),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("2. Les Valeurs Proposées"),
        para("Identifier les valeurs proposées ou livrées aux clients selon les éléments suivants :"),
        ...[
          ["Modernité","Répondre à des besoins nouveaux pour lesquels il n'existait pas d'offres similaires auparavant."],
          ["Performance","Le produit/service doit être supérieur ou égal aux attentes du client."],
          ["Flexibilité","Capacité d'adapter les produits et services aux attentes des clients."],
          ["Accomplissement des tâches","Aider le client à accomplir des tâches spécifiques."],
          ["Conception","Rendre les produits conformes aux attentes et au contexte du client."],
          ["Réduction des coûts","Réduire les coûts de production afin de réduire les prix de vente."],
          ["Réduction des risques","Réduire l'exposition aux risques et sécuriser les clients."],
          ["Accessibilité","Mettre les produits à la disposition des clients qui n'y avaient pas accès."],
          ["Facilité d'utilisation","Rendre les produits plus faciles et plus simples à utiliser."],
        ].map(([k,v]) => new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2200,6826],
          rows:[new TableRow({children:[
            cell(`✓  ${k}`,{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:2200}),
            cell(g[`val_${k.replace(/[^a-zA-Z]/g,"").toLowerCase()}`] || v,{fontSize:20,width:6826})
          ]})]
        })).reduce((acc,t)=>[...acc, t, new Paragraph({spacing:{before:30},children:[]})],[]),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Valeurs spécifiques de votre projet", g.ax1_values || "À compléter — Quelles sont les valeurs clés de votre projet ?"),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("3. L'Équipe de Travail"),
        para("Présenter l'équipe qui travaille sur le projet :"),
        bullet("Noms des membres, compétences, qualifications et formations suivies"),
        bullet("Organisation du travail : répartition des tâches et responsabilités"),
        bullet("Modes d'interaction et de communication entre les membres"),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA},
          columnWidths:[2500,2263,2263,2000],
          rows:[
            new TableRow({children:[
              cell("Nom et Prénom",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2500}),
              cell("Rôle dans le projet",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2263}),
              cell("Spécialité / Background",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2263}),
              cell("Contact",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2000}),
            ]}),
            ...(members.length > 0 ? members : [{name:"",role:"",background:"",contact:""}]).map((m,i)=>
              new TableRow({children:[
                cell(m.name||`Étudiant ${i+1}`,{fontSize:20,width:2500}),
                cell(m.role||"Chef de projet / Développeur...",{fontSize:18,color:"444444",width:2263}),
                cell(m.background||spec,{fontSize:18,color:"444444",width:2263}),
                cell(m.contact||"",{fontSize:18,color:"444444",width:2000}),
              ]})
            ),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("4. Objectifs du Projet"),
        para("Définir les objectifs commerciaux du projet et estimer la part de marché cible à court, moyen et long terme."),
        labeledBlock("Objectif à court terme (1 an)", g.ax1_obj_short || "À compléter"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Objectif à moyen terme (3 ans)", g.ax1_obj_mid || "À compléter"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Objectif à long terme (5 ans)", g.ax1_obj_long || "À compléter"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Part de marché cible estimée", g.ax1_market_share || "À compléter (%)"),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("5. Calendrier de Réalisation du Projet"),
        para("Diviser l'objectif final en tâches partielles, déterminer le temps requis et les résultats clés pour chaque tâche."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA},
          columnWidths:[500,2500,900,900,900,900,900,900,628],
          rows:[
            new TableRow({children:[
              cell("N°",{bg:NAVY,bold:true,color:WHITE,fontSize:18,align:AlignmentType.CENTER,width:500}),
              cell("Tâche / Activité",{bg:NAVY,bold:true,color:WHITE,fontSize:18,width:2500}),
              ...["M1","M2","M3","M4","M5","M6","..."].map((m,i)=>
                cell(m,{bg:NAVY,bold:true,color:WHITE,fontSize:18,align:AlignmentType.CENTER,width:i<6?900:628})
              ),
            ]}),
            ...( g.ax1_timeline || [
              ["1","Études préalables et validation du concept"],
              ["2","Développement du prototype / MVP"],
              ["3","Tests utilisateurs et ajustements"],
              ["4","Finalisation et lancement"],
              ["5","Suivi et amélioration continue"],
            ]).map(([n,t])=>new TableRow({children:[
              cell(n,{fontSize:18,align:AlignmentType.CENTER,width:500}),
              cell(t,{fontSize:18,width:2500}),
              ...Array(7).fill(null).map((_,i)=>cell("",{fontSize:18,width:i<6?900:628})),
            ]})),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── AXE 2 ────────────────────────────────────────────────
        sectionHeader(2, "Aspects Innovants", "Nature des innovations et domaines d'innovation"),
        new Paragraph({ spacing:{before:200}, children:[] }),

        heading2("1. Nature des Innovations"),
        para("Préciser la nature des innovations adoptées dans le projet :"),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2500,6526],
          rows:[
            new TableRow({children:[cell("Type d'innovation",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2500}),cell("Description",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:6526})]}),
            ...["Innovations du marché","Innovations radicales","Innovations technologiques","Innovations croissantes"].map(type=>
              new TableRow({children:[cell(type,{bg:LIGHT,bold:true,color:NAVY,fontSize:19,width:2500}),cell("",{fontSize:19,width:6526})]})
            ),
          ]
        }),
        new Paragraph({ spacing:{before:80}, children:[] }),
        labeledBlock("Type d'innovation retenu", g.ax2_innovation_type || "À compléter — Quel type d'innovation votre projet adopte-t-il ?"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Justification", g.ax2_innovation_justification || "À compléter — Pourquoi ce type d'innovation est-il approprié ?"),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("2. Domaines d'Innovation"),
        para("L'innovation peut inclure les domaines suivants :"),
        ...[
          "Nouveaux processus (augmentation de la rentabilité en augmentant l'efficacité des opérations)",
          "Nouvelles fonctionnalités (offre de produits ou services améliorés)",
          "Nouveaux clients (offrir la gamme habituelle de produits ou de services pour de nouveaux segments)",
          "Nouvelles offres (offre de produits innovants)",
          "Nouveaux modèles (modification du modèle d'affaires, adoption d'un nouveau système de création de valeur)",
        ].map(d => bullet(d)),
        new Paragraph({ spacing:{before:80}, children:[] }),
        labeledBlock("Domaines d'innovation de votre projet", g.ax2_domains || "À compléter — Précisez les domaines d'innovation spécifiques à votre projet."),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Aspects innovants spécifiques", g.ax2_specifics || "À compléter — Listez vos innovations concrètes et distinctives."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── AXE 3 ────────────────────────────────────────────────
        sectionHeader(3, "Analyse Stratégique du Marché", "Segment, concurrence et stratégie marketing"),
        new Paragraph({ spacing:{before:200}, children:[] }),

        heading2("1. Le Segment du Marché"),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2800,6226],
          rows:[
            new TableRow({children:[cell("Marché potentiel",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:2800}),cell(g.ax3_potential_market||"À compléter — Qui sont vos clients potentiels ?",{fontSize:20,width:6226})]}),
            new TableRow({children:[cell("Marché cible (segment)",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:2800}),cell(g.ax3_target_market||"À compléter — Quel est votre segment cible précis ?",{fontSize:20,width:6226})]}),
            new TableRow({children:[cell("Justification du choix",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:2800}),cell(g.ax3_segment_why||"À compléter — Pourquoi avez-vous choisi ce segment ?",{fontSize:20,width:6226})]}),
            new TableRow({children:[cell("Clients importants potentiels",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:2800}),cell(g.ax3_key_clients||"À compléter — Y a-t-il des contrats-cadres envisageables ?",{fontSize:20,width:6226})]}),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("2. Mesure de l'Intensité de la Concurrence"),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2500,3263,3263],
          rows:[
            new TableRow({children:[cell("Concurrent",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2500}),cell("Forces",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3263}),cell("Faiblesses",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3263})]}),
            ...(g.ax3_competitors || [
              {name:"Concurrent direct 1",forces:"",weaknesses:""},
              {name:"Concurrent direct 2",forces:"",weaknesses:""},
              {name:"Concurrent indirect",forces:"",weaknesses:""},
            ]).map(c=>new TableRow({children:[cell(c.name,{fontSize:19,width:2500}),cell(c.forces||"À compléter",{fontSize:19,width:3263}),cell(c.weaknesses||"À compléter",{fontSize:19,width:3263})]})),
          ]
        }),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Parts de marché des concurrents", g.ax3_comp_shares || "À compléter — Estimez les parts de marché de vos principaux concurrents."),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("3. La Stratégie Marketing"),
        para("L'ensemble des techniques utilisées afin de sensibiliser les clients potentiels à acheter nos produits."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2500,6526],
          rows:[
            new TableRow({children:[cell("Mix Marketing",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2500}),cell("Détail",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:6526})]}),
            ...["Produit (Product)","Prix (Price)","Distribution (Place)","Promotion"].map((k,i)=>
              new TableRow({children:[cell(k,{bg:LIGHT,bold:true,color:NAVY,fontSize:19,width:2500}),cell(g[`ax3_mix_${i}`]||"À compléter",{fontSize:19,width:6526})]})
            ),
          ]
        }),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Stratégie de visibilité", g.ax3_visibility || "À compléter — Réseaux sociaux, presse, événements..."),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Stratégie de distribution", g.ax3_distribution || "À compléter — Canaux de vente et de distribution."),

        new Paragraph({ children: [new PageBreak()] }),

        // ── AXE 4 ────────────────────────────────────────────────
        sectionHeader(4, "Plan de Production et Organisation", "Processus, approvisionnement, main-d'œuvre, partenaires"),
        new Paragraph({ spacing:{before:200}, children:[] }),

        heading2("1. Le Processus de Production"),
        para("Le processus de production passe par plusieurs étapes. Pour les services/plateformes numériques, préciser les étapes permettant de fournir le service."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[700,3000,3000,2326],
          rows:[
            new TableRow({children:[cell("N°",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:700}),cell("Étape",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3000}),cell("Description",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3000}),cell("Résultat attendu",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2326})]}),
            ...(g.ax4_process || [
              {step:"Collecte / Matières premières",desc:"",result:""},
              {step:"Développement / Fabrication",desc:"",result:""},
              {step:"Conditionnement du produit",desc:"",result:""},
              {step:"Tests et validation",desc:"",result:""},
              {step:"Emballage / Livraison",desc:"",result:""},
            ]).map((p,i)=>new TableRow({children:[
              cell(String(i+1),{fontSize:20,align:AlignmentType.CENTER,width:700}),
              cell(p.step,{fontSize:19,width:3000}),
              cell(p.desc||"À compléter",{fontSize:18,color:"444444",width:3000}),
              cell(p.result||"À compléter",{fontSize:18,color:"444444",width:2326}),
            ]})),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("2. L'Approvisionnement"),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[3000,6026],
          rows:[
            new TableRow({children:[cell("Politique d'achat",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:3000}),cell(g.ax4_purchase_policy||"À compléter (matières premières, équipements, fournitures)",{fontSize:20,width:6026})]}),
            new TableRow({children:[cell("Principaux fournisseurs",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:3000}),cell(g.ax4_suppliers||"À compléter — Identifiez vos fournisseurs clés.",{fontSize:20,width:6026})]}),
            new TableRow({children:[cell("Politique de paiement",{bg:LIGHT,bold:true,color:NAVY,fontSize:20,width:3000}),cell(g.ax4_payment_policy||"À compléter — Délais et modalités de paiement.",{fontSize:20,width:6026})]}),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("3. La Main d'Œuvre"),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2500,1500,2500,2526],
          rows:[
            new TableRow({children:[cell("Poste",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2500}),cell("Nb",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:1500}),cell("Qualifications requises",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2500}),cell("Localisation",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2526})]}),
            ...(g.ax4_workforce||[{post:"Développeur",nb:"2",qual:"Informatique",loc:"Sur site"},{post:"Commercial",nb:"1",qual:"Marketing",loc:"Terrain"},{post:"Directeur",nb:"1",qual:"Gestion",loc:"Siège"}]).map(w=>
              new TableRow({children:[cell(w.post,{fontSize:19,width:2500}),cell(w.nb||"1",{fontSize:19,align:AlignmentType.CENTER,width:1500}),cell(w.qual||"À définir",{fontSize:18,color:"444444",width:2500}),cell(w.loc||"À définir",{fontSize:18,color:"444444",width:2526})]})
            ),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("4. Les Principaux Partenaires"),
        para("Les acteurs qui peuvent aider à la réalisation du projet par accompagnement et intégration."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2800,6226],
          rows:[
            new TableRow({children:[cell("Type de partenaire",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2800}),cell("Nom / Description",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:6226})]}),
            ...["Fournisseurs","Incubateurs universitaires","Banques / Structures de financement","Collectivités locales","Laboratoires"].map(t=>
              new TableRow({children:[cell(t,{bg:LIGHT,bold:true,color:NAVY,fontSize:19,width:2800}),cell("À compléter",{fontSize:19,width:6226})]})
            ),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── AXE 5 ────────────────────────────────────────────────
        sectionHeader(5, "Plan Financier", "Coûts, chiffre d'affaires, résultats escomptés, trésorerie"),
        new Paragraph({ spacing:{before:200}, children:[] }),

        heading2("1. Les Coûts et Charges"),
        para("Tous les coûts du projet et les investissements requis doivent être déterminés avec précision."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[3500,2763,2763],
          rows:[
            new TableRow({children:[cell("Poste de dépense",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3500}),cell("Montant (DZD)",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:2763}),cell("Source de financement",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:2763})]}),
            ...["Immobilisations incorporelles (logiciels, brevets)","Immobilisations corporelles (matériel, équipements)","Frais de constitution","Loyer / Hébergement","Charges de personnel","Charges d'exploitation (marketing, communication)","Fonds de roulement","Autres charges"].map(p=>
              new TableRow({children:[cell(p,{bg:LIGHT,color:NAVY,fontSize:19,width:3500}),cell("À compléter",{fontSize:19,color:"888888",width:2763}),cell("À préciser",{fontSize:19,color:"888888",width:2763})]})
            ),
            new TableRow({children:[cell("TOTAL INVESTISSEMENT",{bg:GOLD,bold:true,color:DARK,fontSize:20,width:3500}),cell(data.slides?.slide14?.initialInvestment||"À compléter",{bg:GOLD,bold:true,color:DARK,fontSize:20,align:AlignmentType.CENTER,width:2763}),cell("Voir besoins ci-dessus",{bg:GOLD,bold:true,color:DARK,fontSize:20,width:2763})]}),
          ]
        }),
        new Paragraph({ spacing:{before:80}, children:[] }),
        ...((data.slides?.slide14?.needs||[]).length>0 ? [
          para("Besoins identifiés :", {bold:true, color:NAVY, before:60, after:40}),
          ...(data.slides?.slide14?.needs||[]).map(n=>bullet("  "+n))
        ] : []),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("2. Le Chiffre d'Affaires"),
        para("Présenter les deux scénarios possibles du chiffre d'affaires attendu (optimiste et pessimiste)."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[2500,1000,1000,1000,1000,1000,1526],
          rows:[
            new TableRow({children:[cell("DÉTAIL CHIFFRE D'AFFAIRES",{bg:NAVY,bold:true,color:WHITE,fontSize:18,width:2500}),cell("N-2",{bg:NAVY,bold:true,color:WHITE,fontSize:18,align:AlignmentType.CENTER,width:1000}),cell("N-1",{bg:NAVY,bold:true,color:WHITE,fontSize:18,align:AlignmentType.CENTER,width:1000}),cell("N",{bg:NAVY,bold:true,color:WHITE,fontSize:18,align:AlignmentType.CENTER,width:1000}),cell("N+1",{bg:GOLD,bold:true,color:DARK,fontSize:18,align:AlignmentType.CENTER,width:1000}),cell("N+2",{bg:GOLD,bold:true,color:DARK,fontSize:18,align:AlignmentType.CENTER,width:1000}),cell("N+3",{bg:GOLD,bold:true,color:DARK,fontSize:18,align:AlignmentType.CENTER,width:1526})]}),
            ...["Produit / Service A","Quantité","Prix HT","Ventes totales","CHIFFRE D'AFFAIRES GLOBAL"].map((r,i)=>{
              const revs = data.slides?.slide14?.revenues||[];
              const vals = i===4
                ? ["—","—","—", revs[0]?String(revs[0]):"-", revs[1]?String(revs[1]):"-", revs[2]?String(revs[2]):"-"]
                : ["—","—","—","—","—","—"];
              return new TableRow({children:[
                cell(r,{bg:i===4?LIGHT:WHITE,bold:i===4,color:i===4?NAVY:"444444",fontSize:18,width:2500}),
                ...vals.map(v=>cell(v,{fontSize:18,align:AlignmentType.CENTER,width:1000,bold:i===4,color:i===4?NAVY:"444444"}))
              ]});
            }),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("3. Les Comptes de Résultats Escomptés"),
        para("Tableau financier résumant le total des ventes et des charges au cours d'une année, se terminant par un solde positif (bénéfice) ou négatif (perte)."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[3000,1000,1000,1000,1000,1000,1026],
          rows:[
            new TableRow({children:[cell("COMPTE DE RÉSULTAT PRÉVISIONNEL",{bg:NAVY,bold:true,color:WHITE,fontSize:18,width:3000}),cell("N-2",{bg:NAVY,bold:true,color:WHITE,fontSize:16,align:AlignmentType.CENTER,width:1000}),cell("N-1",{bg:NAVY,bold:true,color:WHITE,fontSize:16,align:AlignmentType.CENTER,width:1000}),cell("N",{bg:NAVY,bold:true,color:WHITE,fontSize:16,align:AlignmentType.CENTER,width:1000}),cell("N+1",{bg:GOLD,bold:true,color:DARK,fontSize:16,align:AlignmentType.CENTER,width:1000}),cell("N+2",{bg:GOLD,bold:true,color:DARK,fontSize:16,align:AlignmentType.CENTER,width:1000}),cell("N+3",{bg:GOLD,bold:true,color:DARK,fontSize:16,align:AlignmentType.CENTER,width:1026})]}),
            ...["Vente et produits annexes","Production immobilisée","Subvention d'exploitation","PRODUCTION DE L'EXERCICE","Achats consommés","Services extérieurs","CONSOMMATION DE L'EXERCICE","VALEUR AJOUTÉE D'EXPLOITATION","Charges de personnel","Impôts et taxes","EXCÉDENT BRUT D'EXPLOITATION","Résultat opérationnel","Résultat financier","RÉSULTAT ORDINAIRE AVANT IMPÔT","RÉSULTAT NET DE L'EXERCICE"].map((r,i)=>{
              const isBold=r===r.toUpperCase();
              const revs = data.slides?.slide14?.revenues||[];
              // Fill PRODUCTION and NET rows with revenue projections
              let rowVals = ["","","","","",""];
              if(r==="PRODUCTION DE L'EXERCICE"||r==="Vente et produits annexes"){
                rowVals = ["—","—","—", revs[0]?String(revs[0]):"", revs[1]?String(revs[1]):"", revs[2]?String(revs[2]):""];
              }
              return new TableRow({children:[
                cell(r,{bg:isBold?LIGHT:WHITE,bold:isBold,color:isBold?NAVY:"444444",fontSize:i===14?18:17,width:3000}),
                ...rowVals.map(v=>cell(v,{fontSize:17,align:AlignmentType.CENTER,width:1000,color:"333333"}))
              ]});
            }),
          ]
        }),

        new Paragraph({ spacing:{before:120}, children:[] }),
        heading2("4. Le Plan de Trésorerie"),
        para("Document permettant d'identifier toutes les recettes et les dépenses prévues au cours de la première année d'activité. Les revenus et les dépenses sont calculés chaque mois pour toute l'année."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        new Table({
          width:{size:9026,type:WidthType.DXA},
          columnWidths:[2100,...Array(12).fill(576),250].slice(0,14),
          rows:[
            new TableRow({children:[cell("RUBRIQUES",{bg:NAVY,bold:true,color:WHITE,fontSize:16,width:2100}),..."Jan,Fév,Mar,Avr,Mai,Jun,Jui,Aoû,Sep,Oct,Nov,Déc".split(",").map(m=>cell(m,{bg:NAVY,bold:true,color:WHITE,fontSize:13,align:AlignmentType.CENTER,width:576})),cell("Total",{bg:GOLD,bold:true,color:DARK,fontSize:15,align:AlignmentType.CENTER,width:250})]}),
            ...["Recettes d'exploitation","Autres recettes","TOTAL RECETTES","Achats matières","Charges de personnel","Charges diverses","Remboursements","TOTAL DÉPENSES","SOLDE MENSUEL","TRÉSORERIE CUMULATIVE"].map((r,i)=>{
              const isBold=r.startsWith("TOTAL")||r.startsWith("SOLDE")||r.startsWith("TRÉSORERIE");
              const revs = data.slides?.slide14?.revenues||[];
              const monthly = revs[0] ? Math.round(revs[0]/12) : 0;
              const monthlyStr = monthly ? String(monthly) : "";
              // For TOTAL RECETTES row, fill with monthly estimate
              const vals = Array(13).fill("");
              if(r==="TOTAL RECETTES"||r==="Recettes d'exploitation"){
                for(let m=0;m<12;m++) vals[m]=monthlyStr;
                vals[12]=revs[0]?String(revs[0]):"";
              }
              if(r==="TRÉSORERIE CUMULATIVE"){
                vals[12]=revs[0]?String(revs[0]):"";
              }
              return new TableRow({children:[
                cell(r,{bg:isBold?LIGHT:WHITE,bold:isBold,color:isBold?NAVY:"444444",fontSize:15,width:2100}),
                ...vals.map(v=>cell(v,{fontSize:13,align:AlignmentType.CENTER,width:576,color:"333333"}))
              ]});
            }),
          ]
        }),

        new Paragraph({ spacing:{before:80}, children:[] }),
        ...(data.slides?.slide14 ? [
          new Table({
            width:{size:9026,type:WidthType.DXA}, columnWidths:[3008,3009,3009],
            rows:[
              new TableRow({children:[
                cell("💰 Investissement initial",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:3008}),
                cell("📈 CA Prévu An 1",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:3009}),
                cell("⏱ Seuil de rentabilité",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:3009}),
              ]}),
              new TableRow({children:[
                cell(data.slides.slide14.initialInvestment||"À compléter",{bg:GOLD,bold:true,color:DARK,fontSize:22,align:AlignmentType.CENTER,width:3008}),
                cell(data.slides.slide14.ca1||"À compléter",{bg:GOLD,bold:true,color:DARK,fontSize:22,align:AlignmentType.CENTER,width:3009}),
                cell(data.slides.slide14.breakEven||"À compléter",{bg:GOLD,bold:true,color:DARK,fontSize:22,align:AlignmentType.CENTER,width:3009}),
              ]}),
            ]
          }),
        ] : []),

        new Paragraph({ children: [new PageBreak()] }),

        // ── AXE 6 ────────────────────────────────────────────────
        sectionHeader(6, "Prototype Expérimental", "Version initiale du produit ou du service"),
        new Paragraph({ spacing:{before:200}, children:[] }),

        heading2("Description du Prototype"),
        para("Le prototype expérimental est la version initiale d'un produit ou d'un service. Il sert de base pour développer la version définitive qui sera officiellement mise sur le marché."),
        new Paragraph({ spacing:{before:80}, children:[] }),
        labeledBlock("Type de prototype", g.ax6_prototype_type || "Physique / Numérique / Vidéo de démonstration"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Description du prototype", g.ax6_description || "À compléter — Décrivez votre prototype en détail."),
        new Paragraph({ spacing:{before:60}, children:[] }),

        heading2("Étapes de Réalisation du Prototype"),
        new Table({
          width:{size:9026,type:WidthType.DXA}, columnWidths:[700,3000,5326],
          rows:[
            new TableRow({children:[cell("N°",{bg:NAVY,bold:true,color:WHITE,fontSize:20,align:AlignmentType.CENTER,width:700}),cell("Étape",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:3000}),cell("Description",{bg:NAVY,bold:true,color:WHITE,fontSize:20,width:5326})]}),
            ...(g.ax6_steps||[
              {step:"Définition des spécifications",desc:""},
              {step:"Maquette et design",desc:""},
              {step:"Développement / Construction",desc:""},
              {step:"Tests et validation",desc:""},
              {step:"Présentation finale",desc:""},
            ]).map((p,i)=>new TableRow({children:[
              cell(String(i+1),{fontSize:20,align:AlignmentType.CENTER,width:700}),
              cell(p.step,{fontSize:19,width:3000}),
              cell(p.desc||"À compléter",{fontSize:18,color:"444444",width:5326}),
            ]})),
          ]
        }),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Lien vers vidéo / démo en ligne", g.ax6_link || "URL de la vidéo de démonstration ou du prototype numérique"),
        new Paragraph({ spacing:{before:60}, children:[] }),
        labeledBlock("Note de présentation au jury", g.ax6_jury_note || "À compléter — Précisez comment vous allez présenter votre prototype lors de la soutenance."),

        new Paragraph({ spacing:{before:240,after:60}, children:[] }),
        divider(),
        para(`Guide du Projet — Arrêté Ministériel 1275 — CNCSIIU — MHESR — Algérie\n${uni} — ${inc} — Année universitaire ${year}`,
          { align: AlignmentType.CENTER, color: GRAY, before: 60, after: 60, italic: true }),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buf);
  return outputPath;
}

if (require.main === module) {
  const dataFile = process.argv[2];
  const outFile  = process.argv[3] || "/tmp/guide_projet.docx";
  if (!dataFile) { console.error("Usage: node generate_docx.js <data.json> <output.docx>"); process.exit(1); }
  const data = JSON.parse(fs.readFileSync(dataFile, "utf8"));
  generateDOCX(data, outFile)
    .then(() => { console.log("DOCX_OK:" + outFile); })
    .catch(e => { console.error("DOCX_ERR:" + e.message); process.exit(1); });
}

module.exports = { generateDOCX };
