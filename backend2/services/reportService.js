
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const templates = {
  theory: "Attainx_Theory_Template.docx",
  laboratory: "Attainx_Laboratory_Template.docx",
  "lab-pr-or": "Attainx_Laboratory_PR_OR_Template.docx",
  "audit-course": "Attainx_Theory_Template.docx",
  audit: "Attainx_Theory_Template.docx",
};

// Backend values: keep hardcoded for now
const hardcodedData = {
  COURSE_CODE: "CSE301",
  CLASS: "Third Year CSE",
  ACADEMIC_YEAR: "2026-27",

  CO1_ATTAINMENT: "75%",
  CO2_ATTAINMENT: "80%",
  CO3_ATTAINMENT: "70%",
  CO4_ATTAINMENT: "85%",
  CO5_ATTAINMENT: "78%",
  CO6_ATTAINMENT: "82%",

  FINAL_CO1_ATTAINMENT: "75%",
  FINAL_CO2_ATTAINMENT: "80%",
  FINAL_CO3_ATTAINMENT: "70%",
  FINAL_CO4_ATTAINMENT: "85%",
  FINAL_CO5_ATTAINMENT: "78%",
  FINAL_CO6_ATTAINMENT: "82%",

  // Signature fields
  VERIFIED_BY: "Department Verification",
  HOD_NAME: "Head of Department",
};

function generateReport(formData) {
  const {
    attainmentType,
    facultyName = "",
    cos = {},
    mapping = {},
  } = formData;

  // 1. Select the template
  const templateName = templates[attainmentType];

  if (!templateName) {
    throw new Error("Invalid attainment type selected");
  }

  const templatePath = path.join(
    __dirname,
    "..",
    "templates",
    templateName
  );

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Template not found: ${templateName}`);
  }

  // 2. Load the existing Word template
  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => "0",
  });

  // 3. Backend data + faculty name from frontend
  const data = {
    ...hardcodedData,
    PREPARED_BY: facultyName,
    FACULTY_NAME: facultyName,
  };

  // 4. CO statements from frontend
  // Templates use double underscores, e.g. CO2__STATEMENT.
  for (let i = 1; i <= 6; i++) {
    const coKey = `CO${i}`;
    const statement = cos[coKey] ?? "";

    data[`${coKey}_STATEMENT`] = statement;
    data[`${coKey}__STATEMENT`] = statement;
  }

  // 5. CO-PO/PSO mapping from frontend
  const columns = [
    ...Array.from({ length: 11 }, (_, i) => `PO${i + 1}`),
    "PSO1",
    "PSO2",
    "PSO3",
  ];

  for (let i = 1; i <= 6; i++) {
    const coKey = `CO${i}`;

    for (const column of columns) {
      const value = mapping[coKey]?.[column] ?? "";

      // Support the initial mapping placeholders in the templates.
      // Examples: INITIAL_CO1_PO1, INITIAL_CO1_PSO01
      data[`INITIAL_${coKey}_${column}`] = value;

      // Templates use zero-padded PSO names: PSO01, PSO02, PSO03.
      if (column.startsWith("PSO")) {
        const paddedColumn = `PSO${column.slice(3).padStart(2, "0")}`;
        data[`INITIAL_${coKey}_${paddedColumn}`] = value;
      }

      // Support final mapping placeholders.
      // Examples: FINAL_MAPPING_CO1_PO1, FINAL_MAPPING_CO1_PSO01
      data[`FINAL_MAPPING_${coKey}_${column}`] = value;

      if (column.startsWith("PSO")) {
        const paddedColumn = `PSO${column.slice(3).padStart(2, "0")}`;
        data[`FINAL_MAPPING_${coKey}_${paddedColumn}`] = value;
      }
    }
  }

  // 6. Generate the report without modifying the template
  doc.render(data);

  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

module.exports = { generateReport };