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
};

/**
 * Generate attainment Word report based on user input and hardcoded values.
 *
 * @param {Object} formData
 * @param {string} formData.attainmentType - 'theory' | 'laboratory' | 'lab-pr-or'
 * @param {string} [formData.facultyName] - Name of the faculty
 * @param {string} [formData.courseCode] - Course Code (e.g. CSE301)
 * @param {string} [formData.className] - Class (e.g. Third Year CSE)
 * @param {string} [formData.academicYear] - Academic Year (e.g. 2026-27)
 * @param {Object} [formData.cos] - { CO1: '...', CO2: '...', ... }
 * @param {Object} [formData.mapping] - { CO1: { PO1: 3, ... }, ... }
 * @returns {Buffer} Generated docx file as a Buffer
 */
function generateReport(formData = {}) {
  const {
    attainmentType = "theory",
    facultyName = "",
    courseCode = "",
    className = "",
    class: formClass = "",
    academicYear = "",
    cos = {},
    mapping = {},
  } = formData;

  // 1. Select the template based on report type
  const templateName = templates[attainmentType];

  if (!templateName) {
    throw new Error(`Invalid attainment type selected: "${attainmentType}". Allowed types: theory, laboratory, lab-pr-or`);
  }

  const templatePath = path.join(
    __dirname,
    "..",
    "templates",
    templateName
  );

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Template not found on disk: ${templateName}`);
  }

  // 2. Load the Word template with PizZip & Docxtemplater using {{ and }} delimiters
  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    nullGetter: () => "",
  });

  // 3. Start with hardcoded values defined in app.js / reportData.js
  const data = {
    ...hardcodedReportData,
  };

  // 4. Replace placeholders with input values taken from frontend
  // Faculty name
  if (facultyName && facultyName.trim() !== "") {
    data.PREPARED_BY = facultyName.trim();
    data.FACULTY_NAME = facultyName.trim();
  }

  // Course code, Class, Academic year
  if (courseCode && courseCode.trim() !== "") {
    data.COURSE_CODE = courseCode.trim();
  }
  const resolvedClass = className || formClass;
  if (resolvedClass && resolvedClass.trim() !== "") {
    data.CLASS = resolvedClass.trim();
  }
  if (academicYear && academicYear.trim() !== "") {
    data.ACADEMIC_YEAR = academicYear.trim();
  }

  // 5. Course Outcome statements from frontend
  // In the templates, placeholders are named CO1__STATEMENT, CO2__STATEMENT, etc.
  for (let i = 1; i <= 6; i++) {
    const coKey = `CO${i}`;
    const statement = (cos && cos[coKey] !== undefined) ? String(cos[coKey]).trim() : "";

    data[`${coKey}_STATEMENT`] = statement;
    data[`${coKey}__STATEMENT`] = statement;
  }

  // 6. CO-PO / PSO Mapping from frontend
  const columns = [
    ...Array.from({ length: 11 }, (_, i) => `PO${i + 1}`),
    "PSO1",
    "PSO2",
    "PSO3",
  ];

  // Object to accumulate column sums and counts for calculating averages
  const columnStats = {};
  columns.forEach((col) => {
    columnStats[col] = { sum: 0, count: 0 };
  });

  for (let i = 1; i <= 6; i++) {
    const coKey = `CO${i}`;

    for (const column of columns) {
      const rawVal = mapping[coKey]?.[column];
      const value = (rawVal !== undefined && rawVal !== null) ? String(rawVal).trim() : "";

      // Track numeric values for calculating column averages
      if (value !== "" && !isNaN(Number(value))) {
        columnStats[column].sum += Number(value);
        columnStats[column].count += 1;
      }

      // Initial mapping placeholders: INITIAL_CO1_PO1 ...
      data[`INITIAL_${coKey}_${column}`] = value;

      // Zero-padded PSO names: INITIAL_CO1_PSO01, PSO02, PSO03
      if (column.startsWith("PSO")) {
        const paddedColumn = `PSO${column.slice(3).padStart(2, "0")}`;
        data[`INITIAL_${coKey}_${paddedColumn}`] = value;
      }

      // Final mapping placeholders: FINAL_MAPPING_CO1_PO1 ...
      data[`FINAL_MAPPING_${coKey}_${column}`] = value;

      if (column.startsWith("PSO")) {
        const paddedColumn = `PSO${column.slice(3).padStart(2, "0")}`;
        data[`FINAL_MAPPING_${coKey}_${paddedColumn}`] = value;
      }
    }
  }

  // 7. Calculate and replace Final Mapping Average row
  for (const column of columns) {
    const stats = columnStats[column];
    const avg = stats.count > 0 ? (stats.sum / stats.count).toFixed(2) : "";

    data[`FINAL_MAPPING_AVG_${column}`] = avg;

    if (column.startsWith("PSO")) {
      const paddedColumn = `PSO${column.slice(3).padStart(2, "0")}`;
      data[`FINAL_MAPPING_AVG_${paddedColumn}`] = avg;
    }
  }

  // 8. Render document with populated data
  doc.render(data);

  // 9. Generate and return the updated Word file
  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

module.exports = { generateReport };