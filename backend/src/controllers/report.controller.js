const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  HeadingLevel,
} = require("docx");

const createFrontendReport = async (req, res, next) => {
  try {
    const {
      attainmentType = "theory",
      facultyName = "Faculty Member",
      academicYear = "2025-26",
      department = "Computer Science Engineering",
      courseCode = "",
      courseTitle = "",
      targetCA1 = 60,
      targetCA2 = 60,
      cos = {},
      mapping = {},
    } = req.body;

    const coKeys = Object.keys(cos).filter((k) => cos[k] && cos[k].trim() !== "");
    const activeCos = coKeys.length > 0 ? coKeys : ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"];

    const poList = [
      "PO1", "PO2", "PO3", "PO4", "PO5", "PO6",
      "PO7", "PO8", "PO9", "PO10", "PO11", "PO12",
      "PSO1", "PSO2", "PSO3",
    ];

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            // Header
            new Paragraph({
              text: "INSTITUTE OF TECHNOLOGY & MANAGEMENT",
              heading: HeadingLevel.HEADING_1,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: `DEPARTMENT OF ${String(department).toUpperCase()}`,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: `NBA ACCREDITATION - ${String(attainmentType).toUpperCase()} ATTAINMENT REPORT`,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "" }),

            // Metadata
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("Course Code:")] }),
                    new TableCell({ children: [new Paragraph(courseCode ? String(courseCode) : "N/A")] }),
                    new TableCell({ children: [new Paragraph("Course Title:")] }),
                    new TableCell({ children: [new Paragraph(courseTitle ? String(courseTitle) : "Course Outcome Evaluation")] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("Faculty In-Charge:")] }),
                    new TableCell({ children: [new Paragraph(String(facultyName))] }),
                    new TableCell({ children: [new Paragraph("Academic Year:")] }),
                    new TableCell({ children: [new Paragraph(String(academicYear))] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("Target Cutoff CA-1:")] }),
                    new TableCell({ children: [new Paragraph(`${targetCA1}%`)] }),
                    new TableCell({ children: [new Paragraph("Target Cutoff CA-2:")] }),
                    new TableCell({ children: [new Paragraph(`${targetCA2}%`)] }),
                  ],
                }),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "1. Course Outcomes (COs) Defined",
              heading: HeadingLevel.HEADING_2,
            }),

            // CO statements
            ...activeCos.map(
              (k) =>
                new Paragraph({
                  children: [
                    new TextRun({ text: `${k}: `, bold: true }),
                    new TextRun({ text: cos[k] || "Course outcome statement" }),
                  ],
                })
            ),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "2. CO-PO / PSO Correlation Matrix",
              heading: HeadingLevel.HEADING_2,
            }),

            // Correlation Matrix Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("CO")] }),
                    ...poList.map((po) => new TableCell({ children: [new Paragraph(po)] })),
                  ],
                }),
                ...activeCos.map(
                  (co) =>
                    new TableRow({
                      children: [
                        new TableCell({ children: [new Paragraph(co)] }),
                        ...poList.map((po) => {
                          const val = (mapping[co] && mapping[co][po]) || "-";
                          return new TableCell({ children: [new Paragraph(String(val))] });
                        }),
                      ],
                    })
                ),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "3. Direct & Indirect Attainment Evaluation Summary",
              heading: HeadingLevel.HEADING_2,
            }),

            // Sample Direct Attainment Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("CO")] }),
                    new TableCell({ children: [new Paragraph("Direct Attainment Level (0-3)")] }),
                    new TableCell({ children: [new Paragraph("CES Indirect Score (0-3)")] }),
                    new TableCell({ children: [new Paragraph("Overall Attainment (90% + 10%)")] }),
                  ],
                }),
                ...activeCos.map(
                  (co, idx) =>
                    new TableRow({
                      children: [
                        new TableCell({ children: [new Paragraph(co)] }),
                        new TableCell({ children: [new Paragraph("2.0")] }),
                        new TableCell({ children: [new Paragraph("2.5")] }),
                        new TableCell({ children: [new Paragraph("2.05")] }),
                      ],
                    })
                ),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({ text: "" }),

            // Signature block
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({ text: "Prepared By:" }),
                        new Paragraph({ text: String(facultyName) }),
                        new Paragraph({ text: "Course Faculty" }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({ text: "Verified By:" }),
                        new Paragraph({ text: "Module Coordinator" }),
                        new Paragraph({ text: "Academic Committee" }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({ text: "Approved By:" }),
                        new Paragraph({ text: "Head of Department" }),
                        new Paragraph({ text: String(department) }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=Attainment_Report_${attainmentType}.docx`
    );

    res.send(buffer);
  } catch (error) {
    next(error);
  }
};

module.exports = { createFrontendReport };
