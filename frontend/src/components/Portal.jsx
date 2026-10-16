import React, { useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import { curriculumSubjects } from "../data/curriculumData";
import "../templates/Portal.css";
import attainxLogo from "../assets/attainx-logo.png";

// =========================================
// SVG ICONS (Crisp, themeable, zero-emoji)
// =========================================
const IconCheck = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <polyline points="20 6 9 17 4 12" />
  </svg>
);

const IconClipboard = ({ size = 18, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
    <rect x="8" y="2" width="8" height="4" rx="1" ry="1" />
  </svg>
);

const IconUploadCloud = ({ size = 22, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <path d="M16 16l-4-4-4 4" />
    <path d="M12 12v9" />
    <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3" />
  </svg>
);

const IconArrowRight = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <line x1="5" y1="12" x2="19" y2="12" />
    <polyline points="12 5 19 12 12 19" />
  </svg>
);

const IconArrowLeft = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <line x1="19" y1="12" x2="5" y2="12" />
    <polyline points="12 19 5 12 12 5" />
  </svg>
);

const IconDownload = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);

const IconHistory = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <circle cx="12" cy="12" r="10" />
    <polyline points="12 6 12 12 16 14" />
  </svg>
);

const IconSettings = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z" />
  </svg>
);

const IconAlert = ({ size = 18, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
    <line x1="12" y1="9" x2="12" y2="13" />
    <line x1="12" y1="17" x2="12.01" y2="17" />
  </svg>
);

const IconSpinner = ({ size = 18, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={`icon-spin ${className}`}
    aria-hidden="true"
  >
    <path d="M21 12a9 9 0 1 1-6.219-8.56" />
  </svg>
);

const IconExternalLink = ({ size = 14, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
    <polyline points="15 3 21 3 21 9" />
    <line x1="10" y1="14" x2="21" y2="3" />
  </svg>
);

const IconClose = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <line x1="18" y1="6" x2="6" y2="18" />
    <line x1="6" y1="6" x2="18" y2="18" />
  </svg>
);

const IconSparkles = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z" />
    <path d="M5 3v4" />
    <path d="M19 17v4" />
    <path d="M3 5h4" />
    <path d="M17 19h4" />
  </svg>
);

const IconChevronDown = ({ size = 14, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    <polyline points="6 9 12 15 18 9" />
  </svg>
);

function Portal() {
  const navigate = useNavigate();

  // 4-Step Wizard State: 1 | 2 | 3 | 4
  const [currentStep, setCurrentStep] = useState(1);
  const [stepError, setStepError] = useState("");

  // Temporary hardcoded logged-in user
  const [user, setUser] = useState({
    name: "Admin User",
    department: "Computer Science Engineering",
    email: "admin@example.com",
    password: "admin123",
    image: null,
  });

  // Light mode by default as requested
  const [darkMode, setDarkMode] = useState(false);
  const [showProfile, setShowProfile] = useState(false);

  // =========================================
  // STEP 1: TEMPLATE GENERATOR STATE
  // =========================================
  const [templateCoCount, setTemplateCoCount] = useState(6);
  const [templateCourseName, setTemplateCourseName] = useState("");
  const [templateAcademicYear, setTemplateAcademicYear] = useState("2026-27");
  const [templateDownloaded, setTemplateDownloaded] = useState(false);
  const [hasExistingTemplate, setHasExistingTemplate] = useState(false);

  // Empty rows generator for blank template
  const createEmptyTemplateRows = (count = 5) => {
    return Array.from({ length: count }, () => {
      const row = { name: "", prn: "", ca1: "", mse: "", ca2: "", ese: "" };
      for (let i = 1; i <= 6; i++) {
        row[`co${i}`] = "";
      }
      return row;
    });
  };

  const [templateData, setTemplateData] = useState(createEmptyTemplateRows(5));

  // =========================================
  // STEP 2: ACADEMIC DETAILS STATE
  // =========================================
  const [facultyName, setFacultyName] = useState(user.name);
  const [attainmentType, setAttainmentType] = useState("theory");
  const [academicYear, setAcademicYear] = useState("TY");

  const departmentOptions = [
    "Computer Science Engineering",
    "Information Technology",
    "Electronics & Telecommunication Engineering",
    "Mechanical Engineering",
    "Civil Engineering",
    "Electrical Engineering",
    "Artificial Intelligence & Data Science",
  ];

  const [department, setDepartment] = useState(
    user.department || "Computer Science Engineering"
  );

  // Uploaded mark sheet file and validation state
  const [uploadedFile, setUploadedFile] = useState(null);
  const [fileStats, setFileStats] = useState(null);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadError, setUploadError] = useState("");

  // Subject Selection and Auto-fetch state
  const [selectedSubjectCode, setSelectedSubjectCode] = useState("");
  const [selectedSubjectTitle, setSelectedSubjectTitle] = useState("");
  const [isAutoLoaded, setIsAutoLoaded] = useState(false);

  // Available subjects filtered by Academic Year
  const availableSubjects = curriculumSubjects.filter((s) => {
    if (!academicYear) return true;
    return s.academicYear === academicYear;
  });

  // =========================================
  // STEP 3: COURSE OUTCOMES & MAPPING STATE
  // =========================================
  const rows = ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"];
  const columns = [
    "PO1",
    "PO2",
    "PO3",
    "PO4",
    "PO5",
    "PO6",
    "PO7",
    "PO8",
    "PO9",
    "PO10",
    "PO11",
    "PSO1",
    "PSO2",
    "PSO3",
  ];

  const createEmptyMapping = () => {
    return rows.reduce((obj, row) => {
      obj[row] = {};
      columns.forEach((column) => {
        obj[row][column] = "";
      });
      return obj;
    }, {});
  };

  const [cos, setCos] = useState({
    CO1: "",
    CO2: "",
    CO3: "",
    CO4: "",
    CO5: "",
    CO6: "",
  });

  const [mapping, setMapping] = useState(createEmptyMapping());

  // =========================================
  // STEP 4: ATTAINMENT & REPORT STATE
  // =========================================
  const [targetCA1, setTargetCA1] = useState(60);
  const [targetCA2, setTargetCA2] = useState(60);
  const [isGenerating, setIsGenerating] = useState(false);
  const [reportHtml, setReportHtml] = useState("");
  const [reportBlob, setReportBlob] = useState(null);
  const [reportError, setReportError] = useState("");

  // References for keyboard navigation
  const facultyRef = useRef(null);
  const coRefs = useRef({});
  const inputRefs = useRef({});
  const templateGridRefs = useRef({});

  // =========================================
  // AUTO-FETCH SUBJECT DATA (COs & Mapping)
  // =========================================
  const handleSubjectSelect = (code) => {
    setSelectedSubjectCode(code);
    setStepError("");
    if (!code) {
      setSelectedSubjectTitle("");
      setIsAutoLoaded(false);
      return;
    }

    const found = curriculumSubjects.find((s) => s.code === code);
    if (found) {
      setSelectedSubjectTitle(found.title);
      setIsAutoLoaded(true);

      // 1. Auto-fill Course Outcomes
      setCos({
        CO1: found.cos?.CO1 || "",
        CO2: found.cos?.CO2 || "",
        CO3: found.cos?.CO3 || "",
        CO4: found.cos?.CO4 || "",
        CO5: found.cos?.CO5 || "",
        CO6: found.cos?.CO6 || "",
      });

      // 2. Auto-fill CO-PO mapping matrix
      const newMapping = createEmptyMapping();
      rows.forEach((r) => {
        if (found.mapping && found.mapping[r]) {
          columns.forEach((col) => {
            const val =
              found.mapping[r][col] ??
              found.mapping[r][col.replace("PO", "PO ")] ??
              found.mapping[r][col.replace("PSO", "PSO ")] ??
              "";
            if (val !== undefined && val !== null && val !== "") {
              newMapping[r][col] = String(val);
            }
          });
        }
      });
      setMapping(newMapping);
    }
  };

  const handleAcademicYearChange = (newYear) => {
    setAcademicYear(newYear);
    if (selectedSubjectCode) {
      const currentSub = curriculumSubjects.find((s) => s.code === selectedSubjectCode);
      if (currentSub && currentSub.academicYear !== newYear) {
        setSelectedSubjectCode("");
        setSelectedSubjectTitle("");
        setIsAutoLoaded(false);
      }
    }
  };

  // =========================================
  // STEP VALIDATION HELPERS
  // =========================================
  // Step 1: User has downloaded the template OR marked they already have one
  const isStep1Complete = templateDownloaded || hasExistingTemplate;

  // Step 2: Faculty, Attainment Type, Academic Year, Department, Subject, and Marksheet uploaded
  const getStep2MissingFields = () => {
    const missing = [];
    if (!facultyName.trim()) missing.push("Faculty Name");
    if (!attainmentType) missing.push("Attainment Type");
    if (!academicYear) missing.push("Academic Year");
    if (!department) missing.push("Department");
    if (!selectedSubjectCode) missing.push("Curriculum Subject");
    if (!uploadedFile) missing.push("Student Marksheet File");
    return missing;
  };

  const isStep2Complete = getStep2MissingFields().length === 0;

  // Step 3: At least active COs are filled & mapping has entries
  const getStep3MissingFields = () => {
    const missing = [];
    const filledCos = Object.values(cos).filter((c) => c && c.trim().length > 0);
    if (filledCos.length === 0) {
      missing.push("Course Outcome Statements (at least 1 CO statement)");
    }
    const hasAnyMapping = Object.values(mapping).some((row) =>
      Object.values(row).some((val) => val && String(val).trim() !== "")
    );
    if (!hasAnyMapping) {
      missing.push("CO / PO Mapping Matrix correlation values");
    }
    return missing;
  };

  const isStep3Complete = isStep2Complete && getStep3MissingFields().length === 0;

  // Step 4: Ready if steps 1-3 complete
  const isStep4Complete = isStep1Complete && isStep2Complete && isStep3Complete;

  // =========================================
  // FEATURE: SKIP TEMPLATE & GO AUTO TO STEP 2
  // =========================================
  const handleSkipTemplateAndProceed = () => {
    setHasExistingTemplate(true);
    setStepError("");
    setCurrentStep(2);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  // Navigation handlers with rigorous validation
  const goToStep = (stepNumber) => {
    setStepError("");

    if (stepNumber === 1) {
      setCurrentStep(1);
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }

    if (stepNumber === 2) {
      if (!isStep1Complete) {
        setStepError(
          "Please download a template or choose 'I already have a template' to complete Step 1."
        );
        return;
      }
      setCurrentStep(2);
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }

    if (stepNumber === 3) {
      if (!isStep1Complete) {
        setStepError("Please complete Step 1 first.");
        return;
      }
      const missing = getStep2MissingFields();
      if (missing.length > 0) {
        setStepError(
          `Please fill all required details in Step 2: ${missing.join(", ")}`
        );
        return;
      }
      setCurrentStep(3);
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }

    if (stepNumber === 4) {
      if (!isStep1Complete) {
        setStepError("Please complete Step 1 first.");
        return;
      }
      const missing2 = getStep2MissingFields();
      if (missing2.length > 0) {
        setStepError(
          `Please complete Step 2 first. Missing: ${missing2.join(", ")}`
        );
        return;
      }
      const missing3 = getStep3MissingFields();
      if (missing3.length > 0) {
        setStepError(
          `Please complete Step 3 first. Missing: ${missing3.join(", ")}`
        );
        return;
      }
      setCurrentStep(4);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const handleNextStep = () => {
    if (currentStep === 1) {
      if (!isStep1Complete) {
        setStepError(
          "Please download a template or click 'I already have a template' to proceed."
        );
        return;
      }
      setStepError("");
      setCurrentStep(2);
      window.scrollTo({ top: 0, behavior: "smooth" });
    } else if (currentStep === 2) {
      const missing = getStep2MissingFields();
      if (missing.length > 0) {
        setStepError(
          `Cannot proceed. Please complete all required fields: ${missing.join(
            ", "
          )}`
        );
        return;
      }
      setStepError("");
      setCurrentStep(3);
      window.scrollTo({ top: 0, behavior: "smooth" });
    } else if (currentStep === 3) {
      const missing = getStep3MissingFields();
      if (missing.length > 0) {
        setStepError(
          `Cannot proceed. Please fill: ${missing.join(", ")}`
        );
        return;
      }
      setStepError("");
      setCurrentStep(4);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const handlePrevStep = () => {
    setStepError("");
    if (currentStep > 1) {
      setCurrentStep(currentStep - 1);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  // =========================================
  // COURSE OUTCOME CHANGE
  // =========================================
  const handleCOChange = (co, value) => {
    setCos({
      ...cos,
      [co]: value,
    });
    setStepError("");
  };

  // =========================================
  // MAPPING CHANGE
  // =========================================
  const handleMappingChange = (row, column, value) => {
    setMapping({
      ...mapping,
      [row]: {
        ...mapping[row],
        [column]: value,
      },
    });
    setStepError("");
  };

  // =========================================
  // KEYBOARD NAVIGATION: FACULTY INPUT
  // =========================================
  const handleFacultyKeyDown = (e) => {
    if (e.key === "ArrowDown" || e.key === "Enter") {
      e.preventDefault();
      document.getElementById("academic-year-select")?.focus();
    }
  };

  // =========================================
  // KEYBOARD NAVIGATION: CO INPUTS
  // =========================================
  const handleCOKeyDown = (e, index) => {
    if (e.key === "ArrowDown" || e.key === "Enter") {
      e.preventDefault();
      if (index < rows.length - 1) {
        coRefs.current[rows[index + 1]]?.focus();
      } else {
        inputRefs.current["0-0"]?.focus();
      }
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (index > 0) {
        coRefs.current[rows[index - 1]]?.focus();
      }
    }
  };

  // =========================================
  // KEYBOARD NAVIGATION: MAPPING TABLE
  // =========================================
  const handleKeyDown = (e, rowIndex, columnIndex) => {
    let newRow = rowIndex;
    let newColumn = columnIndex;

    if (e.key === "ArrowRight") {
      if (columnIndex < columns.length - 1) {
        newColumn++;
      } else if (rowIndex < rows.length - 1) {
        newRow++;
        newColumn = 0;
      } else {
        return;
      }
    } else if (e.key === "ArrowLeft") {
      if (columnIndex > 0) {
        newColumn--;
      } else if (rowIndex > 0) {
        newRow--;
        newColumn = columns.length - 1;
      } else {
        return;
      }
    } else if (e.key === "ArrowDown" || e.key === "Enter") {
      if (rowIndex < rows.length - 1) {
        newRow++;
      } else {
        return;
      }
    } else if (e.key === "ArrowUp") {
      if (rowIndex > 0) {
        newRow--;
      } else {
        e.preventDefault();
        coRefs.current[rows[rows.length - 1]]?.focus();
        return;
      }
    } else {
      return;
    }

    e.preventDefault();

    if (
      newRow >= 0 &&
      newRow < rows.length &&
      newColumn >= 0 &&
      newColumn < columns.length
    ) {
      inputRefs.current[`${newRow}-${newColumn}`]?.focus();
    }
  };

  // =========================================
  // KEYBOARD NAVIGATION: TEMPLATE GENERATOR PREVIEW
  // =========================================
  const handleTemplateKeyDown = (e, rowIndex, colKey, totalCols, colIndex) => {
    let targetRow = rowIndex;
    let targetColIndex = colIndex;

    if (e.key === "ArrowRight") {
      if (colIndex < totalCols - 1) {
        targetColIndex++;
      } else if (rowIndex < templateData.length - 1) {
        targetRow++;
        targetColIndex = 0;
      } else {
        return;
      }
    } else if (e.key === "ArrowLeft") {
      if (colIndex > 0) {
        targetColIndex--;
      } else if (rowIndex > 0) {
        targetRow--;
        targetColIndex = totalCols - 1;
      } else {
        return;
      }
    } else if (e.key === "ArrowDown" || e.key === "Enter") {
      if (rowIndex < templateData.length - 1) {
        targetRow++;
      } else {
        return;
      }
    } else if (e.key === "ArrowUp") {
      if (rowIndex > 0) {
        targetRow--;
      } else {
        return;
      }
    } else {
      return;
    }

    e.preventDefault();
    templateGridRefs.current[`${targetRow}-${targetColIndex}`]?.focus();
  };

  // =========================================
  // CLEAR ALL MAPPING
  // =========================================
  const clearAllMapping = () => {
    setMapping(createEmptyMapping());
  };

  // =========================================
  // GENERATE ATTAINMENT REPORT
  // =========================================
  const calculateAttainment = async () => {
    setIsGenerating(true);
    setReportError("");
    setReportHtml("");
    setReportBlob(null);

    try {
      let response;
      if (uploadedFile) {
        const formData = new FormData();
        formData.append("file", uploadedFile);
        formData.append("attainmentType", attainmentType);
        formData.append("facultyName", facultyName);
        formData.append("academicYear", academicYear);
        formData.append("department", department);
        formData.append("courseCode", selectedSubjectCode || "");
        formData.append("courseTitle", selectedSubjectTitle || "");
        formData.append("targetCA1", targetCA1);
        formData.append("targetCA2", targetCA2);
        formData.append("directWeight", 0.9);
        formData.append("indirectWeight", 0.1);
        formData.append("cos", JSON.stringify(cos));
        formData.append("mapping", JSON.stringify(mapping));

        response = await fetch(
          "http://localhost:5000/api/reports/create-report",
          {
            method: "POST",
            body: formData,
          }
        );
      } else {
        response = await fetch(
          "http://localhost:5000/api/reports/create-report",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              attainmentType,
              facultyName,
              academicYear,
              department,
              courseCode: selectedSubjectCode,
              courseTitle: selectedSubjectTitle,
              targetCA1,
              targetCA2,
              cos,
              mapping,
            }),
          }
        );
      }

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.message || "Failed to generate report.");
      }

      const blob = await response.blob();
      const arrayBuffer = await blob.arrayBuffer();

      const result = await mammoth.convertToHtml({
        arrayBuffer,
      });

      setReportBlob(blob);
      setReportHtml(result.value);

      setTimeout(() => {
        document
          .getElementById("generated-report")
          ?.scrollIntoView({
            behavior: "smooth",
            block: "start",
          });
      }, 100);
    } catch (error) {
      console.error("Report generation error:", error);
      setReportError(
        error.message ||
        "Could not connect to the backend. Make sure the server is running."
      );
    } finally {
      setIsGenerating(false);
    }
  };

  // =========================================
  // FILE UPLOAD & VALIDATION
  // =========================================
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const extension = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(extension)) {
      alert("Only CSV and Excel (.xlsx, .csv) files are allowed.");
      e.target.value = "";
      return;
    }

    setUploadedFile(file);
    setUploadError("");
    setStepError("");
    setFileStats(null);
    setIsUploading(true);

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("targetCA1", targetCA1);
      formData.append("targetCA2", targetCA2);
      formData.append("cos", JSON.stringify(cos));
      formData.append("mapping", JSON.stringify(mapping));

      const res = await fetch("http://localhost:5000/api/reports/calculate-live", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setFileStats({
          totalStudents: data.data?.totalStudents || 0,
          detectedColumns: data.data?.detectedColumns || [],
        });
      } else {
        setUploadError(data.message || "Failed to parse uploaded marksheet.");
      }
    } catch (err) {
      console.warn("Live file preview error:", err);
    } finally {
      setIsUploading(false);
    }
  };

  // =========================================
  // PROFILE IMAGE
  // =========================================
  const handleProfileImage = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUser({
      ...user,
      image: URL.createObjectURL(file),
    });
  };

  // =========================================
  // LOGOUT
  // =========================================
  const logout = () => {
    navigate("/");
  };

  // =========================================
  // EXCEL / CSV TEMPLATE GENERATION
  // =========================================
  const handleCoCountChange = (val) => {
    const num = Math.min(6, Math.max(1, parseInt(val, 10) || 1));
    setTemplateCoCount(num);
  };

  const getTemplateColumns = () => {
    const baseCols = ["Name", "PRN", "CA1", "MSE", "CA2", "ESE"];
    const coCols = Array.from({ length: templateCoCount }, (_, i) => `CO${i + 1}`);
    return [...baseCols, ...coCols];
  };

  const handleTemplateCellChange = (rowIndex, colKey, val) => {
    const updated = [...templateData];
    updated[rowIndex] = {
      ...updated[rowIndex],
      [colKey]: val,
    };
    setTemplateData(updated);
  };

  const handleAddTemplateRow = () => {
    const newRow = {
      name: "",
      prn: "",
      ca1: "",
      mse: "",
      ca2: "",
      ese: "",
    };
    for (let i = 1; i <= 6; i++) {
      newRow[`co${i}`] = "";
    }
    setTemplateData([...templateData, newRow]);
  };

  const downloadExcelTemplate = () => {
    const cols = getTemplateColumns();
    const rowsData = templateData.map((item) => {
      const row = [
        item.name || "",
        item.prn || "",
        item.ca1 ?? "",
        item.mse ?? "",
        item.ca2 ?? "",
        item.ese ?? "",
      ];
      for (let i = 1; i <= templateCoCount; i++) {
        row.push(item[`co${i}`] ?? "");
      }
      return row;
    });

    const emptyRowsToAdd = Math.max(0, 15 - rowsData.length);
    for (let i = 0; i < emptyRowsToAdd; i++) {
      const emptyRow = new Array(cols.length).fill("");
      rowsData.push(emptyRow);
    }

    const worksheetData = [cols, ...rowsData];
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    worksheet["!cols"] = [
      { wch: 22 },
      { wch: 16 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      ...Array.from({ length: templateCoCount }, () => ({ wch: 10 })),
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Student Attainment");

    const fileName = templateCourseName
      ? `Attainx_${templateCourseName.replace(/\s+/g, "_")}_${templateCoCount}COs_Template.xlsx`
      : `Attainx_Student_Attainment_${templateCoCount}COs_Template.xlsx`;

    XLSX.writeFile(workbook, fileName);
    setTemplateDownloaded(true);
    setStepError("");
  };

  const downloadCsvTemplate = () => {
    const cols = getTemplateColumns();
    const rowsData = templateData.map((item) => {
      const row = [
        `"${item.name || ""}"`,
        `"${item.prn || ""}"`,
        item.ca1 ?? "",
        item.mse ?? "",
        item.ca2 ?? "",
        item.ese ?? "",
      ];
      for (let i = 1; i <= templateCoCount; i++) {
        row.push(item[`co${i}`] ?? "");
      }
      return row.join(",");
    });

    const csvContent = [cols.join(","), ...rowsData].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = templateCourseName
      ? `Attainx_${templateCourseName.replace(/\s+/g, "_")}_${templateCoCount}COs_Template.csv`
      : `Attainx_Student_Attainment_${templateCoCount}COs_Template.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setTemplateDownloaded(true);
    setStepError("");
  };

  const initial = user.name.charAt(0).toUpperCase();

  const stepItems = [
    {
      id: 1,
      num: "01",
      title: "Generate Template",
      sub: isStep1Complete ? "Completed" : "Download or Skip",
      isComplete: isStep1Complete,
    },
    {
      id: 2,
      num: "02",
      title: "Academic Details",
      sub: isStep2Complete ? "Completed" : "Faculty, Subject & Upload",
      isComplete: isStep2Complete,
    },
    {
      id: 3,
      num: "03",
      title: "COs & Mapping",
      sub: isStep3Complete ? "Completed" : "Syllabus Matrix",
      isComplete: isStep3Complete,
    },
    {
      id: 4,
      num: "04",
      title: "Calculate Attainment",
      sub: reportHtml
        ? "Report Generated"
        : isStep4Complete
          ? "Ready to Analyze"
          : "Pending",
      isComplete: !!reportHtml,
    },
  ];

  return (
    <div className={`portal ${darkMode ? "dark" : "light"}`}>
      {/* =====================================
          SIDEBAR
      ===================================== */}
      <aside className="sidebar">
        <div className="logo">
          <img src={attainxLogo} alt="Attainx" />
          <span>Attainx</span>
        </div>

        {/* USER */}
        <div className="sidebar-user" onClick={() => setShowProfile(true)}>
          <div className="avatar">
            {user.image ? <img src={user.image} alt="profile" /> : initial}
          </div>

          <div className="user-details">
            <strong>{user.name}</strong>
            <span>Faculty</span>
          </div>

          <span className="dots">•••</span>
        </div>

        {/* 4-STEP WORKFLOW MENU */}
        <div className="sidebar-menu">
          <div style={{ padding: "0 14px 6px", fontSize: "11px", fontWeight: "700", color: "var(--muted)", letterSpacing: "1px" }}>
            WORKFLOW STEPS
          </div>

          {stepItems.map((step) => {
            const isClickable =
              step.id === 1 ||
              (step.id === 2 && isStep1Complete) ||
              (step.id === 3 && isStep1Complete && isStep2Complete) ||
              (step.id === 4 && isStep1Complete && isStep2Complete && isStep3Complete);

            return (
              <button
                key={step.id}
                className={`menu-item ${currentStep === step.id ? "active" : ""}`}
                onClick={() => goToStep(step.id)}
                style={{ opacity: isClickable ? 1 : 0.6 }}
              >
                <span>{step.num}</span>
                {step.title}
                {step.isComplete && (
                  <span className="menu-step-badge">
                    <IconCheck size={12} />
                  </span>
                )}
              </button>
            );
          })}

          <div style={{ height: "1px", background: "var(--border)", margin: "14px 0" }} />

          <button
            className="menu-item"
            onClick={() => alert("Attainment calculation history will be archived here.")}
          >
            <span>
              <IconHistory size={16} />
            </span>
            Attainment History
          </button>
        </div>

        {/* SETTINGS */}
        <div className="sidebar-bottom">
          <button className="menu-item" onClick={() => setShowProfile(true)}>
            <span>
              <IconSettings size={16} />
            </span>
            Settings
          </button>
        </div>
      </aside>

      {/* =====================================
          PROFILE PANEL
      ===================================== */}
      {showProfile && (
        <div className="profile-panel">
          <div className="profile-top">
            <h2>Profile</h2>

            <button
              className="close-profile"
              onClick={() => setShowProfile(false)}
              aria-label="Close Profile"
            >
              <IconClose size={16} />
            </button>
          </div>

          <div className="profile-picture">
            <div className="large-avatar">
              {user.image ? <img src={user.image} alt="profile" /> : initial}
            </div>

            <label className="change-picture">
              Change picture
              <input
                type="file"
                accept="image/*"
                onChange={handleProfileImage}
              />
            </label>
          </div>

          <label>NAME</label>
          <input
            className="profile-input"
            value={user.name}
            onChange={(e) =>
              setUser({
                ...user,
                name: e.target.value,
              })
            }
          />

          <label>DEPARTMENT</label>
          <input
            className="profile-input"
            value={user.department}
            onChange={(e) =>
              setUser({
                ...user,
                department: e.target.value,
              })
            }
          />

          <label>EMAIL</label>
          <input
            className="profile-input"
            value={user.email}
            onChange={(e) =>
              setUser({
                ...user,
                email: e.target.value,
              })
            }
          />

          <label>PASSWORD</label>
          <input
            className="profile-input"
            type="password"
            value={user.password}
            onChange={(e) =>
              setUser({
                ...user,
                password: e.target.value,
              })
            }
          />

          <div className="theme-row">
            <div>
              <strong>Appearance</strong>
              <span>{darkMode ? "Dark mode" : "Light mode"}</span>
            </div>

            <button
              className={`toggle ${darkMode ? "on" : ""}`}
              onClick={() => setDarkMode(!darkMode)}
              aria-label="Toggle theme"
            >
              <span />
            </button>
          </div>

          <button className="logout-button" onClick={logout}>
            Logout
            <IconExternalLink size={14} />
          </button>
        </div>
      )}

      {/* =====================================
          MAIN CONTENT AREA
      ===================================== */}
      <main className="portal-main">
        {/* HEADER */}
        <header className="portal-header">
          <div>
            <span className="eyebrow">ACADEMIC ATTAINMENT WORKBENCH</span>
            <h1>
              {currentStep === 1 && "Generate Template"}
              {currentStep === 2 && "Academic Details"}
              {currentStep === 3 && "COs & Mapping Matrix"}
              {currentStep === 4 && "Calculate Attainment"}
            </h1>
            <p>
              {currentStep === 1 && "Step 1 of 4: Download assessment template or skip if you already have one ready"}
              {currentStep === 2 && "Step 2 of 4: Enter faculty information, curriculum selection, and upload student marksheet"}
              {currentStep === 3 && "Step 3 of 4: Review auto-fetched Course Outcomes & correlation matrix (editable)"}
              {currentStep === 4 && "Step 4 of 4: Set target thresholds and compute comprehensive outcome attainment"}
            </p>
          </div>
        </header>

        {/* =====================================
            WIZARD STEPPER TRACK
        ===================================== */}
        <div className="wizard-stepper-container">
          <div className="wizard-stepper-track">
            {stepItems.map((step, idx) => {
              const isActive = currentStep === step.id;
              const isComplete = step.isComplete;
              const isAccessible =
                step.id === 1 ||
                (step.id === 2 && isStep1Complete) ||
                (step.id === 3 && isStep1Complete && isStep2Complete) ||
                (step.id === 4 && isStep1Complete && isStep2Complete && isStep3Complete);

              return (
                <React.Fragment key={step.id}>
                  <button
                    type="button"
                    className={`wizard-step-node ${isActive ? "active" : ""} ${isComplete ? "completed" : ""
                      }`}
                    onClick={() => goToStep(step.id)}
                    disabled={!isAccessible && !isActive}
                    aria-current={isActive ? "step" : undefined}
                  >
                    <div className="wizard-step-badge">
                      {isComplete && !isActive ? (
                        <IconCheck size={14} />
                      ) : (
                        step.num
                      )}
                    </div>

                    <div className="wizard-step-info">
                      <span className="wizard-step-index">Step {step.id}</span>
                      <span className="wizard-step-title">{step.title}</span>
                      <span className="wizard-step-status">
                        {isActive
                          ? "In Progress"
                          : isComplete
                            ? "Completed"
                            : "Pending"}
                      </span>
                    </div>
                  </button>

                  {idx < stepItems.length - 1 && (
                    <div
                      className={`wizard-step-connector ${isComplete ? "completed" : ""
                        }`}
                    />
                  )}
                </React.Fragment>
              );
            })}
          </div>
        </div>

        {/* VALIDATION ERROR BANNER */}
        {stepError && (
          <div className="step-validation-banner" role="alert">
            <span className="step-validation-icon">
              <IconAlert size={20} />
            </span>
            <div className="step-validation-content">
              <div className="step-validation-title">Action Required to Proceed</div>
              <p style={{ margin: 0, fontSize: "13px", color: "var(--text)" }}>{stepError}</p>
            </div>
          </div>
        )}

        {/* =====================================
            STEP 1: GENERATE TEMPLATE
        ===================================== */}
        {currentStep === 1 && (
          <div className="step-content-container">
            {/* FEATURE: SKIP CALLOUT CARD */}
            <div className={`skip-template-card ${hasExistingTemplate ? "active-skip" : ""}`}>
              <div className="skip-template-left">
                <div className="skip-template-icon">
                  {hasExistingTemplate ? (
                    <IconCheck size={22} />
                  ) : (
                    <IconClipboard size={22} />
                  )}
                </div>
                <div className="skip-template-text">
                  <h3>Already have your student marksheet template?</h3>
                  <p>
                    If you have an existing Excel/CSV marksheet with Name, PRN, and test marks (CA1, MSE, CA2, ESE, COs), clicking below will immediately mark this step completed and redirect you straight to Academic Details.
                  </p>
                </div>
              </div>

              <button
                type="button"
                className={`skip-template-btn ${hasExistingTemplate ? "is-skipped" : ""}`}
                onClick={handleSkipTemplateAndProceed}
              >
                {hasExistingTemplate ? (
                  <>
                    <IconCheck size={15} />
                    Template Marked as Ready (Proceed to Step 2)
                    <IconArrowRight size={14} />
                  </>
                ) : (
                  <>
                    <IconClipboard size={15} />
                    I already have a template — Proceed to Next Step
                    <IconArrowRight size={14} />
                  </>
                )}
              </button>
            </div>

            {/* TEMPLATE CONFIGURATION */}
            <section className="portal-section">
              <div className="section-title">
                <span>01</span>
                <div>
                  <h2>Configure Marksheet Template</h2>
                  <p>Customize the number of Course Outcomes (COs) for your course (Maximum 6 COs).</p>
                </div>
              </div>

              <div className="template-config-card">
                <div className="config-group">
                  <div className="config-label-row">
                    <label>NUMBER OF COURSE OUTCOMES (COs)</label>
                    <span className="co-badge">Max 6 COs Allowed</span>
                  </div>

                  <div className="co-counter-row">
                    <div className="co-selector-buttons">
                      {[1, 2, 3, 4, 5, 6].map((num) => (
                        <button
                          key={num}
                          type="button"
                          className={`co-pill-btn ${templateCoCount === num ? "active" : ""}`}
                          onClick={() => setTemplateCoCount(num)}
                        >
                          {num} {num === 1 ? "CO" : "COs"}
                        </button>
                      ))}
                    </div>

                    <div className="co-number-input-wrap">
                      <input
                        type="number"
                        min="1"
                        max="6"
                        value={templateCoCount}
                        onChange={(e) => handleCoCountChange(e.target.value)}
                        className="co-number-input"
                      />
                      <span className="co-input-hint">COs</span>
                    </div>
                  </div>
                </div>

                <div className="config-grid-two">
                  <div className="config-field">
                    <label>COURSE NAME / CODE (OPTIONAL)</label>
                    <input
                      type="text"
                      placeholder="e.g. Distributed Systems (CS401)"
                      value={templateCourseName}
                      onChange={(e) => setTemplateCourseName(e.target.value)}
                      className="template-meta-input"
                    />
                  </div>

                  <div className="config-field">
                    <label>ACADEMIC YEAR (OPTIONAL)</label>
                    <input
                      type="text"
                      placeholder="e.g. 2026-2027"
                      value={templateAcademicYear}
                      onChange={(e) => setTemplateAcademicYear(e.target.value)}
                      className="template-meta-input"
                    />
                  </div>
                </div>

                <div className="generated-columns-summary">
                  <span className="summary-title">Columns generated in Excel template:</span>
                  <div className="columns-badges">
                    <span className="col-badge fixed">Name</span>
                    <span className="col-badge fixed">PRN</span>
                    <span className="col-badge fixed">CA1</span>
                    <span className="col-badge fixed">MSE</span>
                    <span className="col-badge fixed">CA2</span>
                    <span className="col-badge fixed">ESE</span>
                    {Array.from({ length: templateCoCount }, (_, i) => (
                      <span key={i} className="col-badge dynamic">
                        CO{i + 1}
                      </span>
                    ))}
                  </div>
                </div>
              </div>
            </section>

            {/* LIVE PREVIEW TABLE */}
            <section className="portal-section">
              <div className="section-title">
                <span>02</span>
                <div>
                  <h2>Live Template Preview</h2>
                  <p>Preview and edit columns before downloading. Use arrow keys (↑ ↓ ← →) to navigate cells.</p>
                </div>
              </div>

              <div className="table-wrapper template-preview-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th style={{ width: "60px" }}>#</th>
                      <th style={{ width: "200px" }}>Name</th>
                      <th style={{ width: "140px" }}>PRN Number</th>
                      <th style={{ width: "90px" }}>CA1</th>
                      <th style={{ width: "90px" }}>MSE</th>
                      <th style={{ width: "90px" }}>CA2</th>
                      <th style={{ width: "90px" }}>ESE</th>
                      {Array.from({ length: templateCoCount }, (_, i) => (
                        <th key={i} style={{ width: "90px" }} className="co-header-th">
                          CO{i + 1}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {templateData.map((row, rIdx) => {
                      const cols = getTemplateColumns();
                      return (
                        <tr key={rIdx}>
                          <td>{rIdx + 1}</td>
                          <td>
                            <input
                              ref={(el) => {
                                templateGridRefs.current[`${rIdx}-0`] = el;
                              }}
                              value={row.name}
                              onChange={(e) => handleTemplateCellChange(rIdx, "name", e.target.value)}
                              onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, "name", cols.length, 0)}
                            />
                          </td>
                          <td>
                            <input
                              ref={(el) => {
                                templateGridRefs.current[`${rIdx}-1`] = el;
                              }}
                              value={row.prn}
                              onChange={(e) => handleTemplateCellChange(rIdx, "prn", e.target.value)}
                              onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, "prn", cols.length, 1)}
                            />
                          </td>
                          <td>
                            <input
                              ref={(el) => {
                                templateGridRefs.current[`${rIdx}-2`] = el;
                              }}
                              value={row.ca1}
                              onChange={(e) => handleTemplateCellChange(rIdx, "ca1", e.target.value)}
                              onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, "ca1", cols.length, 2)}
                              inputMode="numeric"
                            />
                          </td>
                          <td>
                            <input
                              ref={(el) => {
                                templateGridRefs.current[`${rIdx}-3`] = el;
                              }}
                              value={row.mse}
                              onChange={(e) => handleTemplateCellChange(rIdx, "mse", e.target.value)}
                              onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, "mse", cols.length, 3)}
                              inputMode="numeric"
                            />
                          </td>
                          <td>
                            <input
                              ref={(el) => {
                                templateGridRefs.current[`${rIdx}-4`] = el;
                              }}
                              value={row.ca2}
                              onChange={(e) => handleTemplateCellChange(rIdx, "ca2", e.target.value)}
                              onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, "ca2", cols.length, 4)}
                              inputMode="numeric"
                            />
                          </td>
                          <td>
                            <input
                              ref={(el) => {
                                templateGridRefs.current[`${rIdx}-5`] = el;
                              }}
                              value={row.ese}
                              onChange={(e) => handleTemplateCellChange(rIdx, "ese", e.target.value)}
                              onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, "ese", cols.length, 5)}
                              inputMode="numeric"
                            />
                          </td>
                          {Array.from({ length: templateCoCount }, (_, i) => {
                            const coKey = `co${i + 1}`;
                            const cIdx = 6 + i;
                            return (
                              <td key={coKey}>
                                <input
                                  ref={(el) => {
                                    templateGridRefs.current[`${rIdx}-${cIdx}`] = el;
                                  }}
                                  value={row[coKey] ?? ""}
                                  onChange={(e) => handleTemplateCellChange(rIdx, coKey, e.target.value)}
                                  onKeyDown={(e) => handleTemplateKeyDown(e, rIdx, coKey, cols.length, cIdx)}
                                  inputMode="numeric"
                                />
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div className="template-preview-actions">
                <div style={{ display: "flex", gap: "10px" }}>
                  <button
                    type="button"
                    className="add-row-btn"
                    onClick={handleAddTemplateRow}
                  >
                    + Add Empty Row
                  </button>

                  <button
                    type="button"
                    className="clear-button"
                    onClick={() => setTemplateData(createEmptyTemplateRows(5))}
                  >
                    Clear Rows
                  </button>
                </div>

                <small className="keyboard-help">
                  Navigate seamlessly across all fields using ↑ ↓ ← → arrow keys
                </small>
              </div>
            </section>

            {/* DOWNLOAD ACTIONS */}
            <section className="calculate-section template-download-section">
              <div>
                <span className="calculate-label">DOWNLOAD TEMPLATE</span>
                <h2>
                  Export
                  <br />
                  <span>Excel Sheet.</span>
                </h2>
                <p className="template-download-sub">
                  Downloading marks Step 1 as completed so you can upload your filled sheet in Step 2.
                </p>
              </div>

              <div className="template-action-buttons">
                <button
                  type="button"
                  className="calculate-button download-excel-main-btn"
                  onClick={downloadExcelTemplate}
                >
                  Download Excel (.xlsx)
                  <IconDownload size={15} />
                </button>

                <button
                  type="button"
                  className="clear-button download-csv-btn"
                  onClick={downloadCsvTemplate}
                >
                  <IconDownload size={15} />
                  Download CSV (.csv)
                </button>
              </div>
            </section>

            {templateDownloaded && (
              <div className="download-toast">
                <IconCheck size={16} />
                Template downloaded! Step 1 completed. You can now proceed to Academic Details.
              </div>
            )}
          </div>
        )}

        {/* =====================================
            STEP 2: DETAILS TILL ACADEMIC DETAILS & MARKSHEET
        ===================================== */}
        {currentStep === 2 && (
          <div className="step-content-container">
            {/* FACULTY INFORMATION */}
            <section className="portal-section">
              <div className="section-title">
                <span>01</span>
                <div>
                  <h2>
                    Faculty Information <span className="required-field-star">*</span>
                  </h2>
                  <p>Confirm the course instructor name responsible for this evaluation.</p>
                </div>
              </div>

              <div className="faculty-input">
                <label>FACULTY NAME</label>
                <input
                  ref={facultyRef}
                  value={facultyName}
                  onChange={(e) => {
                    setFacultyName(e.target.value);
                    setStepError("");
                  }}
                  onKeyDown={handleFacultyKeyDown}
                  placeholder="Enter faculty name"
                  className={!facultyName.trim() && stepError ? "input-has-error" : ""}
                />
                <small>Automatically loaded from your profile · Press Enter/Down to continue</small>
              </div>
            </section>

            {/* ATTAINMENT TYPE */}
            <section className="portal-section">
              <div className="section-title">
                <span>02</span>
                <div>
                  <h2>
                    Attainment Type <span className="required-field-star">*</span>
                  </h2>
                  <p>Select the assessment mode according to course structure.</p>
                </div>
              </div>

              <div className="attainment-type-grid">
                <button
                  type="button"
                  className={`attainment-type-card ${attainmentType === "theory" ? "selected" : ""
                    }`}
                  aria-pressed={attainmentType === "theory"}
                  onClick={() => {
                    setAttainmentType("theory");
                    setStepError("");
                  }}
                >
                  <span className="attainment-type-number">01</span>
                  <strong>Theory Attainment</strong>
                  <span>CA-1, MSE, CA-2, ESE and CES</span>
                  <span className="attainment-type-check">
                    {attainmentType === "theory" ? (
                      <>
                        <IconCheck size={13} /> Selected
                      </>
                    ) : (
                      "Select"
                    )}
                  </span>
                </button>

                <button
                  type="button"
                  className={`attainment-type-card ${attainmentType === "laboratory" ? "selected" : ""
                    }`}
                  aria-pressed={attainmentType === "laboratory"}
                  onClick={() => {
                    setAttainmentType("laboratory");
                    setStepError("");
                  }}
                >
                  <span className="attainment-type-number">02</span>
                  <strong>Laboratory Attainment</strong>
                  <span>Laboratory course assessment and CES</span>
                  <span className="attainment-type-check">
                    {attainmentType === "laboratory" ? (
                      <>
                        <IconCheck size={13} /> Selected
                      </>
                    ) : (
                      "Select"
                    )}
                  </span>
                </button>

                <button
                  type="button"
                  className={`attainment-type-card ${attainmentType === "lab-pr-or" ? "selected" : ""
                    }`}
                  aria-pressed={attainmentType === "lab-pr-or"}
                  onClick={() => {
                    setAttainmentType("lab-pr-or");
                    setStepError("");
                  }}
                >
                  <span className="attainment-type-number">03</span>
                  <strong>Laboratory with PR/OR</strong>
                  <span>Practical/oral examination and CES</span>
                  <span className="attainment-type-check">
                    {attainmentType === "lab-pr-or" ? (
                      <>
                        <IconCheck size={13} /> Selected
                      </>
                    ) : (
                      "Select"
                    )}
                  </span>
                </button>

                <button
                  type="button"
                  className={`attainment-type-card ${attainmentType === "audit-course" ? "selected" : ""
                    }`}
                  aria-pressed={attainmentType === "audit-course"}
                  onClick={() => {
                    setAttainmentType("audit-course");
                    setStepError("");
                  }}
                >
                  <span className="attainment-type-number">04</span>
                  <strong>Audit Course</strong>
                  <span>Audit course assessment and outcome evaluation</span>
                  <span className="attainment-type-check">
                    {attainmentType === "audit-course" ? (
                      <>
                        <IconCheck size={13} /> Selected
                      </>
                    ) : (
                      "Select"
                    )}
                  </span>
                </button>
              </div>
            </section>

            {/* ACADEMIC DETAILS & SUBJECT AUTO-FETCH */}
            <section className="portal-section">
              <div className="section-title">
                <span>03</span>
                <div>
                  <h2>
                    Academic Details <span className="required-field-star">*</span>
                  </h2>
                  <p>Choose academic year, department, and curriculum subject for auto-populating COs & mapping.</p>
                </div>
              </div>

              <div className="academic-selection-grid">
                <div className="academic-select-card">
                  <div className="academic-select-label-wrap">
                    <label htmlFor="academic-year-select">ACADEMIC YEAR</label>
                    <span className="selection-badge">Class</span>
                  </div>

                  <div className="custom-select-container">
                    <select
                      id="academic-year-select"
                      className="portal-select"
                      value={academicYear}
                      onChange={(e) => handleAcademicYearChange(e.target.value)}
                    >
                      <option value="FY">FY</option>
                      <option value="SY">SY</option>
                      <option value="TY">TY</option>
                      <option value="BTech">BTech</option>
                    </select>
                    <span className="select-dropdown-indicator">
                      <IconChevronDown size={14} />
                    </span>
                  </div>
                  <small className="academic-select-hint">
                    Select from FY, SY, TY, or BTech
                  </small>
                </div>

                <div className="academic-select-card">
                  <div className="academic-select-label-wrap">
                    <label htmlFor="academic-dept-select">SELECT DEPARTMENT</label>
                    <span className="selection-badge">Branch</span>
                  </div>

                  <div className="custom-select-container">
                    <select
                      id="academic-dept-select"
                      className="portal-select"
                      value={department}
                      onChange={(e) => setDepartment(e.target.value)}
                    >
                      {departmentOptions.map((dept) => (
                        <option key={dept} value={dept}>
                          {dept}
                        </option>
                      ))}
                    </select>
                    <span className="select-dropdown-indicator">
                      <IconChevronDown size={14} />
                    </span>
                  </div>
                  <small className="academic-select-hint">
                    Select department for this course
                  </small>
                </div>

                <div className="academic-select-card">
                  <div className="academic-select-label-wrap">
                    <label htmlFor="academic-subject-select">SELECT SUBJECT</label>
                    <span className="selection-badge">Curriculum</span>
                  </div>

                  <div className="custom-select-container">
                    <select
                      id="academic-subject-select"
                      className={`portal-select ${!selectedSubjectCode && stepError ? "input-has-error" : ""
                        }`}
                      value={selectedSubjectCode}
                      onChange={(e) => handleSubjectSelect(e.target.value)}
                    >
                      <option value="">-- Choose Subject (Auto-fill) --</option>
                      {availableSubjects.map((sub) => (
                        <option key={sub.code} value={sub.code}>
                          {sub.code} - {sub.title}
                        </option>
                      ))}
                    </select>
                    <span className="select-dropdown-indicator">
                      <IconChevronDown size={14} />
                    </span>
                  </div>
                  <small className="academic-select-hint">
                    Auto-populates COs & correlation matrix for Step 3
                  </small>
                </div>
              </div>

              {isAutoLoaded && (
                <div className="curriculum-badge-banner" style={{ marginTop: "20px" }}>
                  <span className="curriculum-badge-icon">
                    <IconCheck size={16} />
                  </span>
                  <div>
                    <strong>
                      Loaded Curriculum for {selectedSubjectCode}: {selectedSubjectTitle}
                    </strong>
                    <p>
                      Course Outcomes and CO-PO Correlation Matrix have been auto-fetched from NEP syllabus and are ready in Step 3.
                    </p>
                  </div>
                  <button
                    type="button"
                    className="reset-curriculum-btn"
                    onClick={() => {
                      setSelectedSubjectCode("");
                      setSelectedSubjectTitle("");
                      setIsAutoLoaded(false);
                    }}
                  >
                    Reset / Custom
                  </button>
                </div>
              )}
            </section>

            {/* MARKSHEET FILE UPLOAD */}
            <section className="portal-section">
              <div className="section-title">
                <span>04</span>
                <div>
                  <h2>
                    Upload Student Marksheet <span className="required-field-star">*</span>
                  </h2>
                  <p>Upload your filled Excel (.xlsx) or CSV (.csv) marksheet to evaluate attainment.</p>
                </div>
              </div>

              <label
                className={`upload-box ${uploadedFile ? "upload-box--done" : ""} ${!uploadedFile && stepError ? "input-has-error" : ""
                  }`}
              >
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={handleFileUpload}
                />

                <div className="upload-arrow">
                  {isUploading ? (
                    <IconSpinner size={24} />
                  ) : uploadedFile ? (
                    <IconCheck size={24} />
                  ) : (
                    <IconUploadCloud size={24} />
                  )}
                </div>

                {isUploading ? (
                  <>
                    <strong>Analyzing Marksheet...</strong>
                    <span>Validating student rows and columns</span>
                  </>
                ) : uploadedFile ? (
                  <>
                    <strong style={{ color: "#22c55e" }}>
                      {uploadedFile.name}
                    </strong>
                    <span>
                      {(uploadedFile.size / 1024).toFixed(1)} KB
                      {fileStats?.totalStudents
                        ? ` • ${fileStats.totalStudents} students detected`
                        : ""}
                      {" • Click to replace file"}
                    </span>
                  </>
                ) : (
                  <>
                    <strong>Upload CSV or Excel marksheet</strong>
                    <span>CSV · XLS · XLSX (Required to proceed)</span>
                  </>
                )}
              </label>

              {uploadError && (
                <p className="report-error" style={{ marginTop: "12px" }}>
                  {uploadError}
                </p>
              )}
            </section>
          </div>
        )}

        {/* =====================================
            STEP 3: SHOW COs AND MAPPING TABLE
        ===================================== */}
        {currentStep === 3 && (
          <div className="step-content-container">
            {isAutoLoaded && (
              <div className="curriculum-badge-banner" style={{ marginBottom: "35px" }}>
                <span className="curriculum-badge-icon">
                  <IconCheck size={16} />
                </span>
                <div>
                  <strong>
                    Curriculum Auto-Fetched: {selectedSubjectCode} — {selectedSubjectTitle}
                  </strong>
                  <p>
                    All Course Outcomes and correlation mappings below were automatically loaded from the curriculum. You can edit any statement or cell values directly.
                  </p>
                </div>
                <button
                  type="button"
                  className="reset-curriculum-btn"
                  onClick={() => {
                    setSelectedSubjectCode("");
                    setSelectedSubjectTitle("");
                    setIsAutoLoaded(false);
                  }}
                >
                  Clear Selection
                </button>
              </div>
            )}

            {/* COURSE OUTCOMES */}
            <section className="portal-section">
              <div className="section-title">
                <span>01</span>
                <div>
                  <h2>
                    Course Outcomes (COs) <span className="required-field-star">*</span>
                  </h2>
                  <p>
                    {isAutoLoaded
                      ? "Auto-fetched from syllabus. Click and edit any statement as needed."
                      : "Enter Course Outcome statements (use ↑ ↓ arrow keys to navigate)."}
                  </p>
                </div>
              </div>

              <div className="co-container">
                {rows.map((co, index) => (
                  <div className="co-row" key={co}>
                    <span>{co}</span>

                    <input
                      ref={(el) => {
                        coRefs.current[co] = el;
                      }}
                      value={cos[co]}
                      placeholder={`Enter ${co} statement (e.g. Understand core concepts...)`}
                      onChange={(e) => handleCOChange(co, e.target.value)}
                      onKeyDown={(e) => handleCOKeyDown(e, index)}
                    />
                  </div>
                ))}
              </div>
            </section>

            {/* CO / PO MAPPING MATRIX */}
            <section className="portal-section">
              <div className="section-title">
                <span>02</span>
                <div>
                  <h2>
                    CO / PO Correlation Matrix <span className="required-field-star">*</span>
                  </h2>
                  <p>
                    {isAutoLoaded
                      ? "Auto-fetched correlation values (0: None, 1: Low, 2: Medium, 3: High). Click any cell to customize."
                      : "Enter correlation matrix values from 0 to 3."}
                  </p>
                </div>
              </div>

              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>CO / PO</th>
                      {columns.map((column) => (
                        <th key={column}>{column}</th>
                      ))}
                    </tr>
                  </thead>

                  <tbody>
                    {rows.map((row, rowIndex) => (
                      <tr key={row}>
                        <th>{row}</th>

                        {columns.map((column, columnIndex) => (
                          <td key={column}>
                            <input
                              ref={(element) => {
                                inputRefs.current[`${rowIndex}-${columnIndex}`] =
                                  element;
                              }}
                              value={mapping[row][column]}
                              onChange={(e) =>
                                handleMappingChange(row, column, e.target.value)
                              }
                              onKeyDown={(e) =>
                                handleKeyDown(e, rowIndex, columnIndex)
                              }
                              inputMode="numeric"
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="mapping-actions">
                <button className="clear-button" onClick={clearAllMapping}>
                  Clear All
                </button>

                <small className="keyboard-help">
                  Use ↑ ↓ ← → arrow keys to move seamlessly between matrix cells
                </small>
              </div>
            </section>
          </div>
        )}

        {/* =====================================
            STEP 4: CALCULATE ATTAINMENT
        ===================================== */}
        {currentStep === 4 && (
          <div className="step-content-container">
            {/* REVIEW SUMMARY CARD */}
            <div className="step-review-summary-card">
              <div className="step-review-header">
                <h3>Verification Summary</h3>
                <span className="step-review-badge">
                  <IconCheck size={12} /> Steps 1, 2, and 3 Verified
                </span>
              </div>

              <div className="step-review-grid">
                <div className="step-review-item">
                  <label>FACULTY</label>
                  <strong>{facultyName || "Not specified"}</strong>
                </div>

                <div className="step-review-item">
                  <label>DEPARTMENT</label>
                  <strong>{department}</strong>
                </div>

                <div className="step-review-item">
                  <label>CLASS / YEAR</label>
                  <strong>{academicYear}</strong>
                </div>

                <div className="step-review-item">
                  <label>SUBJECT</label>
                  <strong>
                    {selectedSubjectCode
                      ? `${selectedSubjectCode} - ${selectedSubjectTitle}`
                      : "Custom Subject"}
                  </strong>
                </div>

                <div className="step-review-item">
                  <label>ATTAINMENT TYPE</label>
                  <strong style={{ textTransform: "capitalize" }}>
                    {attainmentType.replace("-", " ")}
                  </strong>
                </div>

                <div className="step-review-item">
                  <label>MARKSHEET FILE</label>
                  <strong style={{ color: uploadedFile ? "#22c55e" : "inherit" }}>
                    {uploadedFile
                      ? `${uploadedFile.name} (${(
                        uploadedFile.size / 1024
                      ).toFixed(1)} KB)`
                      : "No file"}
                  </strong>
                </div>
              </div>
            </div>

            {/* TARGETS & CALCULATE BUTTON */}
            <section className="calculate-section">
              <div>
                <span className="calculate-label">FINAL STEP</span>
                <h2>
                  Calculate
                  <br />
                  <span>Attainment.</span>
                </h2>
                <p style={{ margin: "8px 0 0", color: "var(--soft)", fontSize: "14px" }}>
                  Adjust target percentage thresholds if needed and click to generate the report.
                </p>
              </div>

              <div className="calculate-actions-container">
                <div className="set-target-card">
                  <div className="target-header">
                    <span className="target-title-label">SET TARGET</span>
                    <span className="target-default-badge">Default 60%</span>
                  </div>

                  <div className="target-inputs-row">
                    <div className="target-field">
                      <label htmlFor="target-ca1-input">CA1</label>
                      <div className="target-input-box">
                        <input
                          id="target-ca1-input"
                          type="number"
                          min="0"
                          max="100"
                          value={targetCA1}
                          onChange={(e) => setTargetCA1(e.target.value)}
                          className="target-input"
                        />
                        <span className="target-suffix">%</span>
                      </div>
                    </div>

                    <div className="target-field">
                      <label htmlFor="target-ca2-input">CA2</label>
                      <div className="target-input-box">
                        <input
                          id="target-ca2-input"
                          type="number"
                          min="0"
                          max="100"
                          value={targetCA2}
                          onChange={(e) => setTargetCA2(e.target.value)}
                          className="target-input"
                        />
                        <span className="target-suffix">%</span>
                      </div>
                    </div>
                  </div>
                </div>

                <button
                  className="calculate-button"
                  onClick={calculateAttainment}
                  disabled={isGenerating}
                >
                  {isGenerating ? (
                    <>
                      <IconSpinner size={16} /> Computing Attainment...
                    </>
                  ) : (
                    <>
                      Calculate Attainment
                      <IconSparkles size={16} />
                    </>
                  )}
                </button>
              </div>
            </section>

            {/* GENERATED REPORT */}
            <section
              className="portal-section generated-report-section"
              id="generated-report"
              style={{ marginTop: "50px" }}
            >
              <div className="section-title">
                <span>02</span>
                <div>
                  <h2>Generated Attainment Report</h2>
                  <p>Preview and export your comprehensive outcome attainment report.</p>
                </div>
              </div>

              {isGenerating && (
                <div style={{ padding: "40px", textAlign: "center", background: "var(--card-bg)", borderRadius: "12px", border: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "center", gap: "10px" }}>
                  <IconSpinner size={20} />
                  <p style={{ margin: 0, fontSize: "15px", color: "var(--text)" }}>
                    Computing direct and indirect attainment from uploaded marksheet...
                  </p>
                </div>
              )}

              {reportError && (
                <p className="report-error" role="alert">
                  {reportError}
                </p>
              )}

              {!isGenerating && !reportError && !reportHtml && (
                <p className="report-empty">
                  No report generated yet. Click <strong>Calculate Attainment</strong> above to run analysis.
                </p>
              )}

              {reportHtml && (
                <>
                  <div className="report-preview">
                    <div
                      dangerouslySetInnerHTML={{
                        __html: reportHtml,
                      }}
                    />
                  </div>

                  <button
                    type="button"
                    className="calculate-button"
                    onClick={() => {
                      if (!reportBlob) return;
                      const url = URL.createObjectURL(reportBlob);
                      const link = document.createElement("a");
                      link.href = url;
                      link.download = `Attainx_${attainmentType}_Report.docx`;
                      document.body.appendChild(link);
                      link.click();
                      link.remove();
                      URL.revokeObjectURL(url);
                    }}
                  >
                    Download Word Report (.docx)
                    <IconDownload size={16} />
                  </button>
                </>
              )}
            </section>
          </div>
        )}

        {/* =====================================
            PERSISTENT STEP FOOTER NAVIGATION
        ===================================== */}
        <div className="step-nav-footer">
          <div className="step-nav-left">
            {currentStep > 1 ? (
              <button
                type="button"
                className="step-nav-btn step-nav-btn-prev"
                onClick={handlePrevStep}
              >
                <IconArrowLeft size={16} />
                Back to Step {currentStep - 1}
              </button>
            ) : (
              <span style={{ fontSize: "12px", color: "var(--muted)" }}>
                Start with Step 1
              </span>
            )}
          </div>

          <div className="step-nav-center">
            <span className="step-progress-indicator">Step {currentStep} of 4</span>
            <span className="step-progress-name">
              {currentStep === 1 && "Generate Marksheet Template"}
              {currentStep === 2 && "Course & Academic Details"}
              {currentStep === 3 && "Course Outcomes & Correlation Matrix"}
              {currentStep === 4 && "Calculate Attainment & Export Report"}
            </span>
          </div>

          <div className="step-nav-right">
            {currentStep < 4 ? (
              <button
                type="button"
                className="step-nav-btn step-nav-btn-next"
                onClick={handleNextStep}
              >
                {currentStep === 1 && "Continue to Step 2: Academic Details"}
                {currentStep === 2 && "Continue to Step 3: COs & Mapping"}
                {currentStep === 3 && "Continue to Step 4: Calculate Attainment"}
                <IconArrowRight size={16} />
              </button>
            ) : (
              <button
                type="button"
                className="step-nav-btn step-nav-btn-next"
                onClick={calculateAttainment}
                disabled={isGenerating}
              >
                {isGenerating ? (
                  <>
                    <IconSpinner size={16} /> Calculating...
                  </>
                ) : (
                  <>
                    Calculate Attainment
                    <IconSparkles size={16} />
                  </>
                )}
              </button>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}

export default Portal;