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
  BorderStyle,
} = require("docx");

const {
  parseMarksheetBuffer,
  calculateAttainmentFromData,
  loadSampleMarksheet,
} = require("../services/liveAttainment.service");

/**
 * Controller: Calculate live attainment from uploaded file or JSON payload
 */
const calculateLive = async (req, res, next) => {
  try {
    let {
      cos = "{}",
      mapping = "{}",
      targetCA1 = 60,
      targetCA2 = 60,
      directWeight = 0.9,
      indirectWeight = 0.1,
      students = null,
      surveyCounts = null,
      useSample = false,
    } = req.body;

    // Handle stringified JSON from multipart form-data
    if (typeof cos === "string") {
      try { cos = JSON.parse(cos); } catch (e) { cos = {}; }
    }
    if (typeof mapping === "string") {
      try { mapping = JSON.parse(mapping); } catch (e) { mapping = {}; }
    }
    if (typeof students === "string") {
      try { students = JSON.parse(students); } catch (e) { students = null; }
    }
    if (typeof surveyCounts === "string") {
      try { surveyCounts = JSON.parse(surveyCounts); } catch (e) { surveyCounts = null; }
    }

    let parsedMarksheet = null;

    if (req.file) {
      parsedMarksheet = await parseMarksheetBuffer(req.file.buffer);
    } else if (students && Array.isArray(students) && students.length > 0) {
      parsedMarksheet = {
        totalStudents: students.length,
        students,
        surveyCounts: surveyCounts || { X: 0, Y: 0, Z: 0 },
        detectedColumns: [],
      };
    } else if (useSample || req.query.sample === "true") {
      parsedMarksheet = await loadSampleMarksheet();
    } else {
      // Default to sample if nothing provided
      parsedMarksheet = await loadSampleMarksheet();
    }

    const calculationResult = calculateAttainmentFromData({
      students: parsedMarksheet.students,
      surveyCounts: parsedMarksheet.surveyCounts,
      cos,
      mapping,
      targetCA1: Number(targetCA1),
      targetCA2: Number(targetCA2),
      directWeight: Number(directWeight),
      indirectWeight: Number(indirectWeight),
    });

    res.json({
      success: true,
      message: "Attainment calculated successfully.",
      data: {
        totalStudents: parsedMarksheet.totalStudents,
        detectedColumns: parsedMarksheet.detectedColumns || [],
        ...calculationResult,
      },
    });
  } catch (error) {
    next(error);
  }
};

/**
 * Controller: Get sample marksheet metadata and preview
 */
const getSampleMarks = async (req, res, next) => {
  try {
    const sample = await loadSampleMarksheet();
    res.json({
      success: true,
      totalStudents: sample.totalStudents,
      surveyCounts: sample.surveyCounts,
      detectedColumns: sample.detectedColumns,
      preview: sample.students.slice(0, 10),
    });
  } catch (error) {
    next(error);
  }
};

/**
 * Controller: Generate full NBA DOCX Attainment Report
 */
const createFrontendReport = async (req, res, next) => {
  try {
    let {
      attainmentType = "theory",
      facultyName = "Faculty Member",
      academicYear = "2025-26",
      department = "Computer Science Engineering",
      courseCode = "",
      courseTitle = "",
      targetCA1 = 60,
      targetCA2 = 60,
      directWeight = 0.9,
      indirectWeight = 0.1,
      cos = {},
      mapping = {},
      calculatedAttainment = null,
      students = null,
    } = req.body;

    // Handle form-data JSON fields
    if (typeof cos === "string") {
      try { cos = JSON.parse(cos); } catch (e) { cos = {}; }
    }
    if (typeof mapping === "string") {
      try { mapping = JSON.parse(mapping); } catch (e) { mapping = {}; }
    }
    if (typeof calculatedAttainment === "string") {
      try { calculatedAttainment = JSON.parse(calculatedAttainment); } catch (e) { calculatedAttainment = null; }
    }
    if (typeof students === "string") {
      try { students = JSON.parse(students); } catch (e) { students = null; }
    }

    // If calculation wasn't passed directly, compute it now!
    if (!calculatedAttainment) {
      let parsedMarksheet = null;
      if (req.file) {
        parsedMarksheet = await parseMarksheetBuffer(req.file.buffer);
      } else if (students && Array.isArray(students)) {
        parsedMarksheet = {
          totalStudents: students.length,
          students,
          surveyCounts: { X: 0, Y: 0, Z: 0 },
        };
      } else {
        parsedMarksheet = await loadSampleMarksheet();
      }

      calculatedAttainment = calculateAttainmentFromData({
        students: parsedMarksheet.students,
        surveyCounts: parsedMarksheet.surveyCounts,
        cos,
        mapping,
        targetCA1: Number(targetCA1),
        targetCA2: Number(targetCA2),
        directWeight: Number(directWeight),
        indirectWeight: Number(indirectWeight),
      });
    }

    const coKeys = Object.keys(cos).filter((k) => cos[k] && String(cos[k]).trim() !== "");
    const activeCos = coKeys.length > 0 ? coKeys : ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"];

    const poList = [
      "PO1", "PO2", "PO3", "PO4", "PO5", "PO6",
      "PO7", "PO8", "PO9", "PO10", "PO11", "PO12",
      "PSO1", "PSO2", "PSO3",
    ];

    const coResults = calculatedAttainment.coResults || [];
    const poResults = calculatedAttainment.poResults || {};
    const psoResults = calculatedAttainment.psoResults || {};
    const survey = calculatedAttainment.survey || { countX: 0, countY: 0, countZ: 0, cesScore: 2.0 };
    const summary = calculatedAttainment.summary || {};

    const tableBorder = {
      top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
    };

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            // Institutional Header
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
              text: `NBA ACCREDITATION - ${String(attainmentType).toUpperCase()} ATTAINMENT EVALUATION REPORT`,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "" }),

            // Metadata Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "Course Code:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(courseCode ? String(courseCode) : "N/A")] }),
                    new TableCell({ children: [new Paragraph({ text: "Course Title:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(courseTitle ? String(courseTitle) : "Course Outcome Evaluation")] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "Faculty In-Charge:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(String(facultyName))] }),
                    new TableCell({ children: [new Paragraph({ text: "Academic Year:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(String(academicYear))] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "Target Cutoff CA-1:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(`${targetCA1}%`)] }),
                    new TableCell({ children: [new Paragraph({ text: "Target Cutoff CA-2:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(`${targetCA2}%`)] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "Total Students Analyzed:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(`${calculatedAttainment.totalStudents || 0} enrolled`)] }),
                    new TableCell({ children: [new Paragraph({ text: "Direct / Indirect Weight:", bold: true })] }),
                    new TableCell({ children: [new Paragraph(`${Math.round((summary.weights?.direct || directWeight) * 100)}% Direct : ${Math.round((summary.weights?.indirect || indirectWeight) * 100)}% Indirect`)] }),
                  ],
                }),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "1. Course Outcomes (COs) Defined",
              heading: HeadingLevel.HEADING_2,
            }),

            // CO statements list
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
              text: "2. CO-PO & PSO Correlation Matrix",
              heading: HeadingLevel.HEADING_2,
            }),

            // Correlation Matrix Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "CO", bold: true })] }),
                    ...poList.map((po) => new TableCell({ children: [new Paragraph({ text: po, bold: true })] })),
                  ],
                }),
                ...activeCos.map(
                  (co) =>
                    new TableRow({
                      children: [
                        new TableCell({ children: [new Paragraph({ text: co, bold: true })] }),
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
              text: "3. Course Outcome (CO) Direct Attainment Summary",
              heading: HeadingLevel.HEADING_2,
            }),

            // REAL Direct Attainment Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "CO", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Group", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Target %", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Max Marks", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Cutoff", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Attempted", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: ">= Cutoff", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Attainment %", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "NBA Level (0-3)", bold: true })] }),
                  ],
                }),
                ...coResults.map(
                  (co) =>
                    new TableRow({
                      children: [
                        new TableCell({ children: [new Paragraph({ text: co.coCode, bold: true })] }),
                        new TableCell({ children: [new Paragraph(co.assessmentGroup || "-")] }),
                        new TableCell({ children: [new Paragraph(`${co.targetPercentage}%`)] }),
                        new TableCell({ children: [new Paragraph(String(co.maxMarks))] }),
                        new TableCell({ children: [new Paragraph(String(co.cutoffMarks))] }),
                        new TableCell({ children: [new Paragraph(String(co.studentsAttempted))] }),
                        new TableCell({ children: [new Paragraph(String(co.studentsAboveCutoff))] }),
                        new TableCell({ children: [new Paragraph(`${co.attainmentPercentage}%`)] }),
                        new TableCell({ children: [new Paragraph({ text: `Level ${co.directLevel}`, bold: true })] }),
                      ],
                    })
                ),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "4. Course Exit Survey (CES) Indirect Attainment Evaluation",
              heading: HeadingLevel.HEADING_2,
            }),

            // CES Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "Rating Level", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Satisfaction Criteria", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Weightage Points", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Student Response Count", bold: true })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("X (High)")] }),
                    new TableCell({ children: [new Paragraph("> 75% Satisfied")] }),
                    new TableCell({ children: [new Paragraph("3")] }),
                    new TableCell({ children: [new Paragraph(String(survey.countX || 0))] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("Y (Medium)")] }),
                    new TableCell({ children: [new Paragraph("50% - 75% Satisfied")] }),
                    new TableCell({ children: [new Paragraph("2")] }),
                    new TableCell({ children: [new Paragraph(String(survey.countY || 0))] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("Z (Low)")] }),
                    new TableCell({ children: [new Paragraph("< 50% Satisfied")] }),
                    new TableCell({ children: [new Paragraph("1")] }),
                    new TableCell({ children: [new Paragraph(String(survey.countZ || 0))] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "CES Indirect Score:", bold: true })] }),
                    new TableCell({
                      columnSpan: 3,
                      children: [
                        new Paragraph({
                          text: `${survey.cesScore || 2.0} / 3.0 (Formula: [3*X + 2*Y + 1*Z] / Total)`,
                          bold: true,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "5. Combined Final Course Outcome (CO) Attainment",
              heading: HeadingLevel.HEADING_2,
            }),

            // Combined Overall CO Attainment Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "CO", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Direct Level (0-3)", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "CES Indirect (0-3)", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Direct (90%)", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Indirect (10%)", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Final Combined Attainment", bold: true })] }),
                  ],
                }),
                ...coResults.map(
                  (co) =>
                    new TableRow({
                      children: [
                        new TableCell({ children: [new Paragraph({ text: co.coCode, bold: true })] }),
                        new TableCell({ children: [new Paragraph(String(co.directLevel))] }),
                        new TableCell({ children: [new Paragraph(String(co.cesScore))] }),
                        new TableCell({ children: [new Paragraph((0.9 * co.directLevel).toFixed(2))] }),
                        new TableCell({ children: [new Paragraph((0.1 * co.cesScore).toFixed(2))] }),
                        new TableCell({ children: [new Paragraph({ text: String(co.finalCoScore), bold: true })] }),
                      ],
                    })
                ),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({
              text: "6. Program Outcome (PO) & PSO Attainment Summary",
              heading: HeadingLevel.HEADING_2,
            }),

            // PO Attainment Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ text: "Outcome Code", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Mapped COs Count", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Direct PO Attainment Score (0-3)", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Weighted Attainment %", bold: true })] }),
                  ],
                }),
                ...poList.map((code) => {
                  const data = (code.startsWith("PSO") ? psoResults[code] : poResults[code]) || {
                    mappedCoCount: 0,
                    directPoScore: 0,
                    weightedPercentage: 0,
                  };
                  return new TableRow({
                    children: [
                      new TableCell({ children: [new Paragraph({ text: code, bold: true })] }),
                      new TableCell({ children: [new Paragraph(String(data.mappedCoCount))] }),
                      new TableCell({ children: [new Paragraph(String(data.directPoScore || (data.directPsoScore || 0)))] }),
                      new TableCell({ children: [new Paragraph(`${data.weightedPercentage}%`)] }),
                    ],
                  });
                }),
              ],
            }),

            new Paragraph({ text: "" }),
            new Paragraph({ text: "" }),

            // Academic Signatures block
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({ text: "Prepared By:", bold: true }),
                        new Paragraph({ text: String(facultyName) }),
                        new Paragraph({ text: "Course Faculty In-Charge" }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({ text: "Verified By:", bold: true }),
                        new Paragraph({ text: "Module Coordinator" }),
                        new Paragraph({ text: "Academic Audit Committee" }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({ text: "Approved By:", bold: true }),
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
      `attachment; filename=Attainment_Report_${courseCode || attainmentType}.docx`
    );

    res.send(buffer);
  } catch (error) {
    next(error);
  }
};

module.exports = {
  calculateLive,
  getSampleMarks,
  createFrontendReport,
};
