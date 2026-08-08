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

// ---- Table helpers ----
function cell(text, { w, header = false, fill } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: fill ? { type: ShadingType.CLEAR, fill } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: header, color: header ? "FFFFFF" : "000000", size: 18 })] })],
  });
}
function headerRow(cells, widths) {
  return new TableRow({ tableHeader: true,
    children: cells.map((c, i) => cell(c, { w: widths[i], header: true, fill: NAVY })) });
}
function dataRow(cells, widths, fill) {
  return new TableRow({ children: cells.map((c, i) => cell(c, { w: widths[i], fill })) });
}
function makeTable(widths, headers, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [ headerRow(headers, widths),
      ...rows.map((r, idx) => dataRow(r, widths, idx % 2 ? LIGHT : undefined)) ],
  });
}

const emptyRows = (n, cols) => Array.from({ length: n }, () => Array(cols).fill(""));

const doc = new Document({
  numbering: { config: [{
    reference: "puces", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
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
      // ---- Title block ----
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
      label("Poste examiné", "Unité centrale de bureau (à préciser : marque / modèle / n° série si connus)"),

      new Paragraph({ children: [new PageBreak()] }),

      // ---- 1. Contexte ----
      h("1. Contexte", HeadingLevel.HEADING_1),
      p("Le poste informatique concerné a été remis pour examen afin de rechercher les fichiers qui auraient été supprimés par le précédent utilisateur. L'objectif était d'identifier, récupérer et documenter ces éléments, en préservant autant que possible les informations d'origine (emplacement et date de suppression)."),

      // ---- 2. Methode ----
      h("2. Méthode employée", HeadingLevel.HEADING_1),
      p("Les opérations suivantes ont été réalisées :"),
      bullet("Consultation de la Corbeille Windows du poste, affichage en mode « Détails » faisant apparaître les colonnes « Emplacement d'origine » et « Date de suppression »."),
      bullet("Prise de photographies de l'écran afin de conserver, pour chaque élément, son emplacement d'origine et sa date de suppression (ces informations n'étant pas conservées dans le fichier lui-même lors de la copie)."),
      bullet("Copie des éléments de la Corbeille vers un support externe (clé USB), sans les « restaurer » sur le disque du poste."),
      bullet("Tentative d'analyse approfondie à l'aide de l'outil de récupération Recuva (version portable) : non aboutie en raison de la très grande lenteur du poste (voir « Limites »)."),
      p("Aucune donnée n'a été volontairement modifiée, ni restaurée sur le disque du poste examiné.", { run: { italics: true } }),

      // ---- 3. Journal ----
      h("3. Journal des opérations", HeadingLevel.HEADING_1),
      p("À compléter au fur et à mesure (heure réelle de chaque action) :"),
      makeTable([1400, 3600, 4040],
        ["Heure", "Action réalisée", "Observations"],
        [
          ["", "Mise sous tension du poste", ""],
          ["", "Ouverture de la Corbeille (mode Détails)", ""],
          ["", "Photographies de la liste des fichiers supprimés", ""],
          ["", "Copie des fichiers vers la clé USB", ""],
          ["", "Arrêt du poste", ""],
          ...emptyRows(2, 3),
        ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ---- 4. Inventaire ----
      h("4. Inventaire des fichiers récupérés", HeadingLevel.HEADING_1),
      p("Reporter ici les éléments visibles sur les photographies de la Corbeille et présents sur la clé. Les dates de suppression proviennent des photographies ; les dates de modification proviennent des fichiers copiés."),
      makeTable([420, 2600, 2600, 1500, 1520],
        ["N°", "Nom du fichier / dossier", "Emplacement d'origine", "Date de suppression", "Observation"],
        emptyRows(12, 5)),

      // ---- 5. Limites ----
      h("5. Limites de l'intervention", HeadingLevel.HEADING_1),
      bullet("Seule la Corbeille a pu être traitée. Elle ne contient que les fichiers supprimés de façon « ordinaire »."),
      bullet("Les fichiers d'une Corbeille déjà vidée, ou supprimés définitivement (Maj+Suppr), ne figurent PAS dans la Corbeille et n'ont donc pas pu être récupérés par cette méthode."),
      bullet("Le poste étant très lent, l'analyse approfondie (Recuva) n'a pas abouti. Une récupération complète nécessiterait d'extraire le disque et de l'analyser sur une machine fiable, via un adaptateur USB-SATA."),
      bullet("Le fait d'avoir démarré le poste sur son système a pu, inévitablement, écraser une partie des données supprimées non encore récupérées."),

      // ---- 6. Recommandations ----
      h("6. Recommandations", HeadingLevel.HEADING_1),
      bullet("Conserver le disque du poste en l'état (ne plus utiliser le poste) afin de préserver les données encore récupérables."),
      bullet("Pour une analyse complète et exploitable, faire réaliser une copie intégrale (image) du disque puis une récupération approfondie, si possible par une personne qualifiée."),
      bullet("Si des éléments laissant présumer des irrégularités sont confirmés, les faire remonter par les voies officielles (audit interne, autorités compétentes) plutôt que de constituer un dossier de manière informelle, afin d'en préserver la valeur.")

      ,new Paragraph({ spacing: { before: 500 }, children: [] }),
      makeTable([4520, 4520],
        ["Fait le", "Signature"],
        [["", ""]]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/home/user/SITE-INTERNET/rapport/Compte-rendu_intervention.docx", buf);
  console.log("written");
});
