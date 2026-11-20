const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, BorderStyle, TabStopType, ExternalHyperlink } = require('docx');

const FONT = "Calibri";
const ACCENT = "1F3864";

const section = (t) => new Paragraph({
  spacing: { before: 240, after: 100 },
  keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 2 } },
  children: [new TextRun({ text: t.toUpperCase(), bold: true, size: 24, color: ACCENT, font: FONT })]
});

const entry = (left, right) => new Paragraph({
  spacing: { before: 100, after: 0 },
  keepNext: true,
  tabStops: [{ type: TabStopType.RIGHT, position: 9706 }],
  children: [
    new TextRun({ text: left, bold: true, size: 22, font: FONT }),
    new TextRun({ text: "\t" + right, size: 21, italics: true, font: FONT })
  ]
});

const sub = (t) => new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: t, size: 21, font: FONT })] });

const bullet = (runs) => new Paragraph({
  numbering: { reference: "b", level: 0 },
  spacing: { after: 40 },
  children: runs
});

const link = (text, url) => new ExternalHyperlink({
  link: url,
  children: [new TextRun({ text, size: 21, font: FONT, color: "0563C1", underline: {} })]
});

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "UDAYANRAJE BALIRAM WAYKAR", bold: true, size: 40, color: ACCENT, font: FONT })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
        new TextRun({ text: "Phone: 9370667060  |  Email: udayanrajewaykar22@gmail.com", size: 21, font: FONT })
      ]}),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [
        link("LinkedIn", "https://www.linkedin.com/in/udayanraje-waykar-b26208428"),
        new TextRun({ text: "  |  ", size: 21, font: FONT }),
        link("GitHub: udayanrajewaykar22-dotcom", "https://github.com/udayanrajewaykar22-dotcom")
      ]}),

      section("Career Objective"),
      new Paragraph({ spacing: { after: 40 }, alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "Motivated B.Tech Computer Science (Cloud Computing) student with a foundation in cloud concepts and coding, seeking internship and learning opportunities to apply technical skills, grow in the cloud domain, and contribute to real-world projects.", size: 21, font: FONT })] }),

      section("Education"),
      entry("B.Tech, Computer Science Engineering (Cloud Computing)", "Pursuing (1st Year)"),
      sub("Parul University, Vadodara, Gujarat"),
      entry("Higher Secondary (11th–12th), Pune Board", ""),
      sub("Priyadarshani Public School & Junior College, Ahilyanagar"),
      entry("Secondary (1st–10th), Pune Board", ""),
      sub("Sacred Heart Convent High School, Ahilyanagar"),

      section("Certifications"),
      bullet([new TextRun({ text: "AWS Academy Cloud Foundations", bold: true, size: 21, font: FONT }), new TextRun({ text: " – AWS Academy", size: 21, font: FONT })]),

      section("Skills"),
      bullet([new TextRun({ text: "Technical: ", bold: true, size: 21, font: FONT }), new TextRun({ text: "Coding, Cloud Computing fundamentals (AWS)", size: 21, font: FONT })]),
      bullet([new TextRun({ text: "Soft skills: ", bold: true, size: 21, font: FONT }), new TextRun({ text: "Public speaking, communication", size: 21, font: FONT })]),
    ]
  }]
});

Packer.toBuffer(doc).then(b => fs.writeFileSync('/mnt/user-data/outputs/Udayanraje_Waykar_CV.docx', b));
