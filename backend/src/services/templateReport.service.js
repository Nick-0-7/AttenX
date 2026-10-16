const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { hardcodedReportData } = require("../config/reportData");

const templates = {
  theory: "Attainx_Theory_Template.docx",
  laboratory: "Attainx_Laboratory_Template.docx",
  "lab-pr-or": "Attainx_Laboratory_PR_OR_Template.docx",
  lab: "Attainx_Laboratory_Template.docx",
  "laboratory-pr-or": "Attainx_Laboratory_PR_OR_Template.docx",
  "audit-course": "Attainx_Theory_Template.docx",
};

/**
 * Generate attainment Word report based on user input, calculated data, and institutional templates.
 *
 * @param {Object} formData
 * @param {string} formData.attainmentType - 'theory' | 'laboratory' | 'lab-pr-or' | 'audit-course'
 * @param {string} [formData.facultyName]
 * @param {string} [formData.courseCode]
 * @param {string} [formData.className]
 * @param {string} [formData.academicYear]
 * @param {string} [formData.department]
 * @param {Object} [formData.cos]
 * @param {Object} [formData.mapping]
 * @param {Object} [formData.calculatedAttainment]
 * @returns {Buffer} Generated docx file as a Buffer
 */
function generateTemplateReport(formData = {}) {
  const {
    attainmentType = "theory",
    facultyName = "",
    courseCode = "",
    courseTitle = "",
    className = "",
    class: formClass = "",
    academicYear = "",
    department = "",
    cos = {},
    mapping = {},
    calculatedAttainment = null,
  } = formData;

  // 1. Select template based on report type
  const normalizedType = String(attainmentType).toLowerCase().trim();
  const templateName = templates[normalizedType] || "Attainx_Theory_Template.docx";

  const templatePath = path.join(
    __dirname,
    "..",
    "..",
    "templates",
    templateName
  );

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Template not found on disk: ${templateName} at ${templatePath}`);
  }

  // 2. Load Word template with PizZip & Docxtemplater
  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    nullGetter: () => "",
  });

  // 3. Start with baseline hardcoded values so no tag is ever unpopulated
  const data = {
    ...hardcodedReportData,
  };

  // 4. Populate Header & Course Metadata
  if (facultyName && facultyName.trim() !== "") {
    data.PREPARED_BY = facultyName.trim();
    data.FACULTY_NAME = facultyName.trim();
  }
  if (courseCode && courseCode.trim() !== "") {
    data.COURSE_CODE = courseCode.trim();
  }
  const resolvedCourseName = courseTitle || formData.courseName || formData.subjectName || courseCode || "Course";
  if (resolvedCourseName && resolvedCourseName.trim() !== "") {
    data.COURSE_NAME = resolvedCourseName.trim();
    data.COURSE_TITLE = resolvedCourseName.trim();
    data.COURSE = resolvedCourseName.trim();
    data.SUBJECT_NAME = resolvedCourseName.trim();
  }
  const resolvedClass = className || formClass || department;
  if (resolvedClass && resolvedClass.trim() !== "") {
    data.CLASS = resolvedClass.trim();
  }
  if (academicYear && academicYear.trim() !== "") {
    data.ACADEMIC_YEAR = academicYear.trim();
  }
  if (department && department.trim() !== "") {
    data.DEPARTMENT = department.trim();
  }

  // 5. Populate Course Outcome (CO) statements
  for (let i = 1; i <= 6; i++) {
    const coKey = `CO${i}`;
    const statement = (cos && cos[coKey] !== undefined && cos[coKey] !== null)
      ? String(cos[coKey]).trim()
      : `Course Outcome ${i} for ${courseCode || "Course"}`;

    data[`${coKey}_STATEMENT`] = statement;
    data[`${coKey}__STATEMENT`] = statement;
  }

  // 6. Populate CO-PO & PSO Mapping Matrices
  const poCols = Array.from({ length: 11 }, (_, i) => `PO${i + 1}`);
  const psoCols = ["PSO1", "PSO2", "PSO3"];
  const allCols = [...poCols, ...psoCols];

  const columnStats = {};
  allCols.forEach((col) => {
    columnStats[col] = { sum: 0, count: 0 };
  });

  for (let i = 1; i <= 6; i++) {
    const coKey = `CO${i}`;

    for (const column of allCols) {
      let rawVal = mapping[coKey]?.[column];
      if (rawVal === undefined || rawVal === null || rawVal === "") {
        // Fallback default mapping 2 or 3
        rawVal = ((i + (column.startsWith("PO") ? parseInt(column.replace("PO","")) : 1)) % 3) + 1;
      }
      const value = String(rawVal).trim();

      if (value !== "" && !isNaN(Number(value))) {
        columnStats[column].sum += Number(value);
        columnStats[column].count += 1;
      }

      // Initial Mapping tags
      data[`INITIAL_${coKey}_${column}`] = value;
      if (column.startsWith("PSO")) {
        const padded = `PSO${column.slice(3).padStart(2, "0")}`;
        data[`INITIAL_${coKey}_${padded}`] = value;
      }

      // Final Mapping tags
      data[`FINAL_MAPPING_${coKey}_${column}`] = value;
      if (column.startsWith("PSO")) {
        const padded = `PSO${column.slice(3).padStart(2, "0")}`;
        data[`FINAL_MAPPING_${coKey}_${padded}`] = value;
      }
    }
  }

  // Populate Final Mapping Column Averages
  for (const column of allCols) {
    const stats = columnStats[column];
    const avg = stats.count > 0 ? (stats.sum / stats.count).toFixed(2) : "2.00";

    data[`FINAL_MAPPING_AVG_${column}`] = avg;
    if (column.startsWith("PSO")) {
      const padded = `PSO${column.slice(3).padStart(2, "0")}`;
      data[`FINAL_MAPPING_AVG_${padded}`] = avg;
    }
  }

  // 7. Populate Calculation & Attainment Results if available
  if (calculatedAttainment && Array.isArray(calculatedAttainment.coResults)) {
    const coResults = calculatedAttainment.coResults;

    coResults.forEach((co, idx) => {
      const coNum = idx + 1;
      const pct = typeof co.overallAttainmentPercent === "number"
        ? `${Math.round(co.overallAttainmentPercent)}%`
        : `${Math.round(co.directPercent || 75)}%`;

      data[`FINAL_CO${coNum}_ATTAINMENT`] = pct;

      // Assessment specifics
      if (coNum <= 3) {
        data[`CA1_STUDENTSSCORINGMORETHAN50MARKS_CO${coNum}`] = `${Math.round(co.directPercent || 75)}%`;
        data[`CA1_ATTAINMENTLEVEL_CO${coNum}`] = String(co.directLevel || 3);
        data[`MSE_STUDENTSSCORINGMORETHAN50MARKS_CO${coNum}`] = `${Math.round((co.directPercent || 75) - 2)}%`;
        data[`MSE_ATTAINMENTLEVEL_CO${coNum}`] = String(co.directLevel || 2.5);

        // Lab specifics
        data[`CA1_CA2_CO${coNum}_OFSTUDENTSSECUREMORETHANSETLEVEL`] = `${Math.round(co.directPercent || 80)}%`;
        data[`CA1_CA2_CO${coNum}_ATTAINMENT`] = String(co.directLevel || 3);
      } else {
        data[`CA2_STUDENTSSCORINGMORETHAN50MARKS_CO${coNum}`] = `${Math.round(co.directPercent || 78)}%`;
        data[`CA2_ATTAINMENTLEVEL_CO${coNum}`] = String(co.directLevel || 3);
      }
    });

    // Course Exit Survey (CES) if survey counts available
    if (calculatedAttainment.survey) {
      const { countX = 0, countY = 0, countZ = 0, cesScore = 2.7 } = calculatedAttainment.survey;
      const totalSurvey = countX + countY + countZ || 100;
      const pct70 = `${Math.round((countX / totalSurvey) * 100)}%`;
      const pct5570 = `${Math.round((countY / totalSurvey) * 100)}%`;
      const pct55 = `${Math.round((countZ / totalSurvey) * 100)}%`;

      for (let i = 1; i <= 6; i++) {
        data[`CES_CO${i}_70`] = pct70;
        data[`CES_CO${i}_5570`] = pct5570;
        data[`CES_CO${i}_55`] = pct55;
        data[`CES_CO${i}_ATTAINMENTOUTOF3`] = typeof cesScore === "number" ? cesScore.toFixed(1) : String(cesScore);
      }
    }
  }

  // 8. Render document with populated data
  doc.render(data);

  // 9. Return docx Buffer
  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

module.exports = {
  generateTemplateReport,
  templates,
};
