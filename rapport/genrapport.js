const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak,
  Header, Footer, PageNumber, LevelFormat
} = require("docx");
const fs = require("fs");

const NAVY = "1F3864";
const GREY = "595959";
const LIGHT = "D9E1F2";

function h(text, level) {
  return new Paragraph({ heading: level, spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, color: NAVY, bold: true })] });
}
function p(text, opts = {}) {
  return new Paragraph({ spacing: { after: 120 }, ...opts,
    children: [new TextRun({ text, ...opts.run })] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "puces", level: 0 }, spacing: { after: 60 },
    children: [new TextRun(text)] });
}
function label(l, v) {
  return new Paragraph({ spacing: { after: 80 }, children: [
    new TextRun({ text: l + " : ", bold: true }),
    new TextRun({ text: v }) ] });
}
function cell(text, { w, header = false, fill, bold = false } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: fill ? { type: ShadingType.CLEAR, fill } : undefined,
    margins: { top: 50, bottom: 50, left: 70, right: 70 },
    children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: header || bold, color: header ? "FFFFFF" : "000000", size: 17 })] })],
  });
}
function headerRow(cells, widths) {
  return new TableRow({ tableHeader: true,
    children: cells.map((c, i) => cell(c, { w: widths[i], header: true, fill: NAVY })) });
}
function makeTable(widths, headers, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [ headerRow(headers, widths),
      ...rows.map((r, idx) => new TableRow({ children: r.map((c, i) =>
        cell(c, { w: widths[i], fill: idx % 2 ? LIGHT : undefined })) })) ] });
}
const emptyRows = (n, cols) => Array.from({ length: n }, () => Array(cols).fill(""));

// ---------- Inventaire relevé sur les photographies de la Corbeille ----------
// Colonnes : Catégorie | Nom du fichier (tel que lu) | Emplacement d'origine | Type
const INV = [
  ["Ordre de service", "ORDRE DE SERVICE N°64 PRESCRIVANT…", "E:\\PMSP-SANG", "PDF"],
  ["Ordre de service", "OS prolongation de délai MDC publique", "E:\\Dossier MDC publique", "PDF"],
  ["Ordre de service", "OS prolongation MDC publique", "E:\\Dossier MDC publique pour MINESUP", "PDF"],
  ["Ordre de service", "OS (pdf) – 513 Ko", "E:\\Dossier MDC publique pour MINESUP", "PDF"],
  ["Décompte / financier", "fichiers décomptes divers (version 1)", "C:\\Users\\MEBADA\\Desktop\\Dossier télép…", "Fichier"],
  ["Décompte / financier", "fichiers décomptes divers (version 2)", "C:\\Users\\MEBADA\\AppData\\…", "Fichier"],
  ["Décompte / financier", "MERCURIALE 2024", "(à préciser)", "Fichier"],
  ["Décompte / financier", "tableau synoptique des projets de…", "C:\\Users\\MEBADA\\Documents", "Feuille Microsoft"],
  ["Décompte / financier", "TRAVAUX SUPPLEMENTAIRES MAR…", "C:\\Users\\MEBADA\\Documents", "Feuille Microsoft"],
  ["Procès-verbal", "procès-verbal de constatation des travaux", "E:\\PMSP-SANG", "PDF"],
  ["Procès-verbal", "PV ETS BIZAN", "E:\\PMSP-SANG", "PDF"],
  ["Procès-verbal", "PV d'Évaluation d'Avancement des…", "E:\\Dossier MDC publique pour MINESUP", "PDF"],
  ["Procès-verbal", "PV_reunion_de_chantier_du_09_avr…", "E:\\Dossier MDC publique pour MINESUP", "Document Word"],
  ["Plans / CAO", "PLAN BUREAU EBOLOWA FINAL", "C:\\Users\\MEBADA\\Desktop\\UE PLANS", "Plan"],
  ["Plans / CAO", "Plan de charges Route Ebwa si-var", "C:\\Users\\MEBADA\\Desktop", "Plan"],
  ["Plans / CAO", "PLAN SITUATION DEFINITIF A2", "C:\\Users\\MEBADA\\Desktop\\UE PLANS", "Plan"],
  ["Plans / CAO", "ENSET COUPE", "(à préciser)", "Plan / PDF"],
  ["Plans / CAO", "façade principale", "(à préciser)", "Plan"],
  ["Plans / CAO", "Recue1.asd à Recue7.asd (sauvegardes AutoCAD)", "C:\\Users\\MEBADA\\Documents", "Fichier ASD"],
  ["Plans / CAO", "Enregistrements automatiques AutoCAD (nombreux)", "C:\\Users\\MEBADA\\AppData\\…\\Roaming", "Fichier ASD"],
  ["Plans / CAO", "S10-1", "(à préciser)", "Fichier"],
  ["Entreprises / divers", "INTER TRADE", "(à préciser)", "Fichier"],
  ["Entreprises / divers", "MIRAP (version 1)", "(à préciser)", "Fichier"],
  ["Entreprises / divers", "pdf denis abala", "E:\\PMSP-SANG", "PDF"],
  ["Entreprises / divers", "EYEBE PDF", "(à préciser)", "PDF"],
  ["Entreprises / divers", "info_5_2(1) ; Index_6_2(1)", "(à préciser)", "Fichier"],
  ["Médias (WhatsApp)", "VID-20230622-WA0001", "(à préciser)", "Vidéo MP4"],
  ["Médias (WhatsApp)", "VID-20230701-WA0000", "(à préciser)", "Vidéo MP4"],
  ["Médias (WhatsApp)", "VID-202407…-WA…", "(à préciser)", "Vidéo MP4"],
];

const doc = new Document({
  numbering: { config: [{ reference: "puces", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 260 } } } }] }] },
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
    headers: { default: new Header({ children: [ new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [ new TextRun({ text: "Rapport d'intervention – CONFIDENTIEL", color: GREY, size: 16 }) ] }) ] }) },
    footers: { default: new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER,
      children: [ new TextRun({ text: "Page ", size: 16, color: GREY }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GREY }),
        new TextRun({ text: " sur ", size: 16, color: GREY }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: GREY }) ] }) ] }) },
    children: [
      new Paragraph({ spacing: { before: 600, after: 60 }, alignment: AlignmentType.CENTER,
        children: [ new TextRun({ text: "COMPTE RENDU D'INTERVENTION", bold: true, size: 40, color: NAVY }) ] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [ new TextRun({ text: "Examen d'un poste informatique – Recherche de fichiers supprimés", size: 24, color: GREY }) ] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
        children: [ new TextRun({ text: "Document confidentiel – usage interne", italics: true, size: 18, color: GREY }) ] }),
      new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } }, spacing: { after: 200 }, children: [] }),

      label("Objet de l'intervention", "Recherche et sauvegarde des fichiers supprimés présents sur le poste informatique remis pour examen"),
      label("Date de l'intervention", "08 août 2026"),
      label("Réalisée par", "………………………………………… (à compléter)"),
      label("À l'attention de", "M. le Directeur – (à compléter)"),
      label("Poste examiné", "Unité centrale de bureau – compte utilisateur « MEBADA »"),
      label("Volume constaté", "170 éléments présents dans la Corbeille"),

      new Paragraph({ children: [new PageBreak()] }),

      h("1. Contexte", HeadingLevel.HEADING_1),
      p("Le poste informatique concerné a été remis pour examen afin de rechercher les fichiers supprimés par le précédent utilisateur. L'objectif était d'identifier, récupérer et documenter ces éléments, en préservant autant que possible l'emplacement d'origine et la date de suppression."),

      h("2. Méthode employée", HeadingLevel.HEADING_1),
      bullet("Consultation de la Corbeille Windows, affichage « Détails » avec les colonnes « Emplacement d'origine » et « Date de suppression »."),
      bullet("Photographies de l'écran pour conserver, pour chaque élément, son emplacement d'origine et sa date de suppression."),
      bullet("Copie des éléments de la Corbeille vers un support externe (clé USB), sans les restaurer sur le disque du poste."),
      bullet("Tentative d'analyse approfondie (Recuva) : non aboutie en raison de la très grande lenteur du poste."),
      p("Aucune donnée n'a été volontairement modifiée ni restaurée sur le disque du poste examiné.", { run: { italics: true } }),

      h("3. Constats principaux", HeadingLevel.HEADING_1),
      p("La Corbeille contenait 170 éléments supprimés. Les fichiers proviennent principalement :"),
      bullet("du compte utilisateur « C:\\Users\\MEBADA » (Bureau, Documents, AppData) — le précédent utilisateur ;"),
      bullet("de dossiers de projets situés sur un disque externe : « E:\\Dossier MDC publique pour MINESUP » et « E:\\PMSP-SANG »."),
      p("Parmi les éléments supprimés figurent des documents directement liés à l'exécution et au paiement de marchés :"),
      bullet("Ordres de service, dont des OS de prolongation de délai ;"),
      bullet("Décomptes (présents en « version 1 » et « version 2 ») ;"),
      bullet("Procès-verbaux (constatation des travaux, avancement, réunions de chantier) ;"),
      bullet("Documents financiers : mercuriale 2024, tableau synoptique des projets, travaux supplémentaires ;"),
      bullet("Plans (dont plans AutoCAD et leurs sauvegardes automatiques) ;"),
      bullet("Vidéos (fichiers WhatsApp « VID-…-WA… »)."),
      p("Les dates de suppression relevées sur les photographies s'échelonnent de 2023 à mai 2026 ; plusieurs suppressions se concentrent sur quelques journées (points à vérifier).", { run: { italics: true } }),

      new Paragraph({ children: [new PageBreak()] }),

      h("4. Inventaire des fichiers supprimés relevés", HeadingLevel.HEADING_1),
      p("Éléments lus sur les photographies de la Corbeille. Certains noms sont partiellement lisibles (indiqués par « … ») et sont à confirmer sur les fichiers copiés sur la clé. Les dates de suppression figurent sur les photographies.", { run: { italics: true, size: 18 } }),
      makeTable([1500, 3000, 3000, 1140],
        ["Catégorie", "Nom du fichier (lu)", "Emplacement d'origine", "Type"],
        INV.map((r) => [r[0], r[1], r[2], r[3]])),

      h("5. Points méritant vérification", HeadingLevel.HEADING_1),
      p("À examiner par une personne compétente, sans préjuger de quoi que ce soit :", { run: { italics: true } }),
      bullet("La présence de décomptes en deux versions (« version 1 » / « version 2 ») : comparer leur contenu."),
      bullet("Les ordres de service de prolongation de délai : cohérence avec les marchés concernés."),
      bullet("Le rapprochement entre PV de réception/avancement, décomptes et paiements effectifs."),
      bullet("La raison de la suppression de ces pièces et le moment où elle est intervenue (voir dates sur les photos)."),

      h("6. Limites de l'intervention", HeadingLevel.HEADING_1),
      bullet("Seule la Corbeille a pu être traitée : elle ne contient que les fichiers supprimés de façon ordinaire."),
      bullet("Les fichiers d'une Corbeille vidée ou supprimés définitivement (Maj+Suppr) n'y figurent pas."),
      bullet("Le poste étant très lent, l'analyse approfondie n'a pas abouti ; une récupération complète nécessiterait d'extraire le disque et de l'analyser sur une machine fiable (adaptateur USB-SATA)."),
      bullet("Le démarrage du poste sur son système a pu écraser une partie des données supprimées non encore récupérées."),
      bullet("Les noms de fichiers ont été relevés à partir de photographies ; ils doivent être confirmés sur les fichiers réellement copiés."),

      h("7. Recommandations", HeadingLevel.HEADING_1),
      bullet("Conserver le disque du poste en l'état (ne plus utiliser le poste) pour préserver les données encore récupérables."),
      bullet("Faire réaliser une copie intégrale (image) du disque puis une récupération approfondie, si possible par une personne qualifiée."),
      bullet("Si des irrégularités se confirment, les faire remonter par les voies officielles (audit interne, autorités compétentes) afin de préserver la valeur probante des éléments."),

      new Paragraph({ spacing: { before: 400 }, children: [] }),
      makeTable([4520, 4520], ["Fait le", "Signature"], [["", ""]]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/home/user/SITE-INTERNET/rapport/Compte-rendu_intervention.docx", buf);
  console.log("written");
});
