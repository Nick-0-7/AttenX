const {
  parseMarksheetBuffer,
  calculateAttainmentFromData,
  loadSampleMarksheet,
} = require("../services/liveAttainment.service");
const { generateTemplateReport } = require("../services/templateReport.service");

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
 * Controller: Generate full NBA DOCX Attainment Report based on institutional template
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

    // If calculation wasn't passed directly, compute it now from uploaded file, students, or sample
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

    const buffer = generateTemplateReport({
      attainmentType,
      facultyName,
      academicYear,
      department,
      courseCode,
      courseTitle,
      className: req.body.className || department,
      cos,
      mapping,
      calculatedAttainment,
    });

    const fileName = `Attainx_${attainmentType}_Report.docx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${fileName}"`
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
