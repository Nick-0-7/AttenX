import React, { useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import "../templates/Portal.css";
import attainxLogo from "../assets/attainx-logo.png";

function Portal() {
  const navigate = useNavigate();

  // Active tab: 'attainment' | 'template'
  const [activeTab, setActiveTab] = useState("attainment");

  const [isGenerating, setIsGenerating] = useState(false);
  const [reportHtml, setReportHtml] = useState("");
  const [reportBlob, setReportBlob] = useState(null);
  const [reportError, setReportError] = useState("");

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

  // Selected attainment type (theory | laboratory | lab-pr-or | audit-course)
  const [attainmentType, setAttainmentType] = useState("theory");

  // Faculty name
  const [facultyName, setFacultyName] = useState(user.name);

  // Academic Year options and selection: FY | SY | TY | BTech
  const [academicYear, setAcademicYear] = useState("TY");

  // Department options and selection
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

  // Attainment Targets for CA1 and CA2 (default 60%)
  const [targetCA1, setTargetCA1] = useState(60);
  const [targetCA2, setTargetCA2] = useState(60);

  // Course Outcomes
  const [cos, setCos] = useState({
    CO1: "",
    CO2: "",
    CO3: "",
    CO4: "",
    CO5: "",
    CO6: "",
  });

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

  // CO-PO Mapping
  const createEmptyMapping = () => {
    return rows.reduce((obj, row) => {
      obj[row] = {};

      columns.forEach((column) => {
        obj[row][column] = "";
      });

      return obj;
    }, {});
  };

  const [mapping, setMapping] = useState(createEmptyMapping());

  // =========================================
  // TEMPLATE GENERATOR STATE
  // =========================================
  const [templateCoCount, setTemplateCoCount] = useState(6);
  const [templateCourseName, setTemplateCourseName] = useState("");
  const [templateAcademicYear, setTemplateAcademicYear] = useState("2026-27");
  const [templateDownloaded, setTemplateDownloaded] = useState(false);

  // Empty rows generator for blank template (no dummy student names)
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



  // References for keyboard navigation
  const facultyRef = useRef(null);
  const coRefs = useRef({});
  const inputRefs = useRef({});
  const templateGridRefs = useRef({});

  // =========================================
  // COURSE OUTCOME CHANGE
  // =========================================
  const handleCOChange = (co, value) => {
    setCos({
      ...cos,
      [co]: value,
    });
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
  };

  // =========================================
  // KEYBOARD NAVIGATION: FACULTY INPUT
  // =========================================
  const handleFacultyKeyDown = (e) => {
    if (e.key === "ArrowDown" || e.key === "Enter") {
      e.preventDefault();
      coRefs.current["CO1"]?.focus();
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
        // Move to first cell of mapping table
        inputRefs.current["0-0"]?.focus();
      }
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (index > 0) {
        coRefs.current[rows[index - 1]]?.focus();
      } else {
        facultyRef.current?.focus();
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
      const response = await fetch(
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
            targetCA1,
            targetCA2,
            cos,
            mapping,
          }),
        }
      );

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(
          errorData.message || "Failed to generate report."
        );
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
  // FILE UPLOAD
  // =========================================
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const extension = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(extension)) {
      alert("Only CSV and Excel files are allowed.");
      e.target.value = "";
      return;
    }
    console.log("Uploaded file:", file.name);
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

    // Add empty formatted rows for teacher to fill in
    const emptyRowsToAdd = Math.max(0, 15 - rowsData.length);
    for (let i = 0; i < emptyRowsToAdd; i++) {
      const emptyRow = new Array(cols.length).fill("");
      rowsData.push(emptyRow);
    }

    const worksheetData = [cols, ...rowsData];
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    // Set nice column widths
    worksheet["!cols"] = [
      { wch: 22 }, // Name
      { wch: 16 }, // PRN
      { wch: 10 }, // CA1
      { wch: 10 }, // MSE
      { wch: 10 }, // CA2
      { wch: 10 }, // ESE
      ...Array.from({ length: templateCoCount }, () => ({ wch: 10 })),
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Student Attainment");

    const fileName = templateCourseName
      ? `Attainx_${templateCourseName.replace(/\s+/g, "_")}_${templateCoCount}COs_Template.xlsx`
      : `Attainx_Student_Attainment_${templateCoCount}COs_Template.xlsx`;

    XLSX.writeFile(workbook, fileName);
    setTemplateDownloaded(true);
    setTimeout(() => setTemplateDownloaded(false), 4000);
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
    setTimeout(() => setTemplateDownloaded(false), 4000);
  };

  const initial = user.name.charAt(0).toUpperCase();

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

        {/* NAVIGATION */}
        <div className="sidebar-menu">
          <button
            className={`menu-item ${activeTab === "attainment" ? "active" : ""}`}
            onClick={() => setActiveTab("attainment")}
          >
            <span>01</span>
            Attainment
          </button>

          <button
            className={`menu-item ${activeTab === "template" ? "active" : ""}`}
            onClick={() => setActiveTab("template")}
          >
            <span>02</span>
            Generate Template
          </button>

          <button
            className="menu-item"
            onClick={() => alert("Attainment history will be available here.")}
          >
            <span>03</span>
            Attainment History
          </button>
        </div>

        {/* SETTINGS */}
        <div className="sidebar-bottom">
          <button className="menu-item" onClick={() => setShowProfile(true)}>
            <span>⚙</span>
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
            >
              ×
            </button>
          </div>

          {/* PROFILE IMAGE */}
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

          {/* NAME */}
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

          {/* DEPARTMENT */}
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

          {/* EMAIL */}
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

          {/* PASSWORD */}
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

          {/* THEME */}
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

          {/* LOGOUT */}
          <button className="logout-button" onClick={logout}>
            Logout
            <span>↗</span>
          </button>
        </div>
      )}

      {/* =====================================
          MAIN CONTENT AREA
      ===================================== */}
      <main className="portal-main">
        {activeTab === "attainment" ? (
          <>
            {/* HEADER */}
            <header className="portal-header">
              <div>
                <span className="eyebrow">ACADEMIC ANALYTICS</span>
                <h1>Student Attainment</h1>
                <p>Course Outcome Mapping & Analysis</p>
              </div>
            </header>

            {/* =====================================
                FACULTY INFORMATION
            ===================================== */}
            <section className="portal-section">
              <div className="section-title">
                <span>01</span>
                <h2>Faculty Information</h2>
              </div>

              <div className="faculty-input">
                <label>FACULTY NAME</label>

                <input
                  ref={facultyRef}
                  value={facultyName}
                  onChange={(e) => setFacultyName(e.target.value)}
                  onKeyDown={handleFacultyKeyDown}
                  placeholder="Enter faculty name"
                />

                <small>Automatically loaded from your profile · Press ↓ to navigate</small>
              </div>
            </section>

            {/* =====================================
                ATTAINMENT TYPE
            ===================================== */}
            <section className="portal-section">
              <div className="section-title">
                <span>02</span>
                <div>
                  <h2>Attainment Type</h2>
                  <p>Select the report type you want to calculate.</p>
                </div>
              </div>

              <div className="attainment-type-grid">
                <button
                  type="button"
                  className={`attainment-type-card ${
                    attainmentType === "theory" ? "selected" : ""
                  }`}
                  aria-pressed={attainmentType === "theory"}
                  onClick={() => setAttainmentType("theory")}
                >
                  <span className="attainment-type-number">01</span>
                  <strong>Theory Attainment</strong>
                  <span>CA-1, MSE, CA-2, ESE and CES</span>
                  <span className="attainment-type-check">
                    {attainmentType === "theory" ? "Selected ✓" : "Select"}
                  </span>
                </button>

                <button
                  type="button"
                  className={`attainment-type-card ${
                    attainmentType === "laboratory" ? "selected" : ""
                  }`}
                  aria-pressed={attainmentType === "laboratory"}
                  onClick={() => setAttainmentType("laboratory")}
                >
                  <span className="attainment-type-number">02</span>
                  <strong>Laboratory Attainment</strong>
                  <span>Laboratory course assessment and CES</span>
                  <span className="attainment-type-check">
                    {attainmentType === "laboratory" ? "Selected ✓" : "Select"}
                  </span>
                </button>

                <button
                  type="button"
                  className={`attainment-type-card ${
                    attainmentType === "lab-pr-or" ? "selected" : ""
                  }`}
                  aria-pressed={attainmentType === "lab-pr-or"}
                  onClick={() => setAttainmentType("lab-pr-or")}
                >
                  <span className="attainment-type-number">03</span>
                  <strong>Laboratory with PR/OR</strong>
                  <span>Practical/oral examination and CES</span>
                  <span className="attainment-type-check">
                    {attainmentType === "lab-pr-or" ? "Selected ✓" : "Select"}
                  </span>
                </button>

                <button
                  type="button"
                  className={`attainment-type-card ${
                    attainmentType === "audit-course" ? "selected" : ""
                  }`}
                  aria-pressed={attainmentType === "audit-course"}
                  onClick={() => setAttainmentType("audit-course")}
                >
                  <span className="attainment-type-number">04</span>
                  <strong>Audit Course</strong>
                  <span>Audit course assessment and outcome evaluation</span>
                  <span className="attainment-type-check">
                    {attainmentType === "audit-course" ? "Selected ✓" : "Select"}
                  </span>
                </button>
              </div>
            </section>

            {/* =====================================
                ACADEMIC YEAR & DEPARTMENT
            ===================================== */}
            <section className="portal-section">
              <div className="section-title">
                <span>03</span>
                <div>
                  <h2>Academic Details</h2>
                  <p>Select academic year and department for attainment evaluation.</p>
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
                      onChange={(e) => setAcademicYear(e.target.value)}
                    >
                      <option value="FY">FY</option>
                      <option value="SY">SY</option>
                      <option value="TY">TY</option>
                      <option value="BTech">BTech</option>
                    </select>
                    <span className="select-dropdown-indicator">▾</span>
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
                    <span className="select-dropdown-indicator">▾</span>
                  </div>
                  <small className="academic-select-hint">
                    Select department for this course
                  </small>
                </div>
              </div>
            </section>

            {/* =====================================
                COURSE OUTCOMES
            ===================================== */}
            <section className="portal-section">
              <div className="section-title">
                <span>04</span>
                <div>
                  <h2>Course Outcomes</h2>
                  <p>Enter course outcome statements (use ↑ ↓ arrow keys to navigate).</p>
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
                      placeholder={`Enter ${co} statement`}
                      onChange={(e) => handleCOChange(co, e.target.value)}
                      onKeyDown={(e) => handleCOKeyDown(e, index)}
                    />
                  </div>
                ))}
              </div>
            </section>

            {/* =====================================
                CO / PO MAPPING
            ===================================== */}
            <section className="portal-section">
              <div className="section-title">
                <span>05</span>
                <div>
                  <h2>CO / PO Mapping</h2>
                  <p>Enter the mapping values (0 to 3).</p>
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

              {/* MAPPING ACTIONS */}
              <div className="mapping-actions">
                <button className="clear-button" onClick={clearAllMapping}>
                  Clear All
                </button>

                <small className="keyboard-help">
                  Use ↑ ↓ ← → arrow keys to move between cells
                </small>
              </div>
            </section>

            {/* =====================================
                IMPORT DATA
            ===================================== */}
            <section className="portal-section">
              <div className="section-title">
                <span>06</span>
                <h2>Import Data</h2>
              </div>

              <label className="upload-box">
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={handleFileUpload}
                />

                <div className="upload-arrow">↑</div>
                <strong>Upload CSV or Excel file</strong>
                <span>CSV · XLS · XLSX</span>
              </label>
            </section>

            {/* =====================================
                CALCULATE ATTAINMENT
            ===================================== */}
            <section className="calculate-section">
              <div>
                <span className="calculate-label">READY TO ANALYZE?</span>

                <h2>
                  Calculate
                  <br />
                  <span>Attainment.</span>
                </h2>
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
                  {isGenerating ? "Generating Report..." : "Calculate Attainment"}
                  <span>↗</span>
                </button>
              </div>
            </section>

            {/* =====================================
                GENERATED REPORT
            ===================================== */}
            <section
              className="portal-section generated-report-section"
              id="generated-report"
            >
              <div className="section-title">
                <span>07</span>

                <div>
                  <h2>Generated Report</h2>
                  <p>Preview and download your generated report.</p>
                </div>
              </div>

              {isGenerating && (
                <p>Generating your report. Please wait...</p>
              )}

              {reportError && (
                <p className="report-error" role="alert">
                  {reportError}
                </p>
              )}

              {!isGenerating && !reportError && !reportHtml && (
                <p className="report-empty">No report generated yet. Click Calculate Attainment above.</p>
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
                    Download Word Report <span>↓</span>
                  </button>
                </>
              )}
            </section>
          </>
        ) : (
          /* =====================================
             GENERATE TEMPLATE TAB
          ===================================== */
          <div className="template-tab-container">
            <header className="portal-header">
              <div>
                <span className="eyebrow">ACADEMIC TEMPLATES</span>
                <h1>Generate Template</h1>
                <p>Configure and download assessment Excel sheets customized for your Course Outcomes.</p>
              </div>
            </header>

            {/* TEMPLATE CONFIGURATION SECTION */}
            <section className="portal-section">
              <div className="section-title">
                <span>01</span>
                <div>
                  <h2>Template Configuration</h2>
                  <p>Specify the number of Course Outcomes (COs) for your course (Maximum 6 COs).</p>
                </div>
              </div>

              <div className="template-config-card">
                {/* NUMBER OF COs */}
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

                {/* OPTIONAL METADATA */}
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

                {/* COLUMNS GENERATED SUMMARY */}
                <div className="generated-columns-summary">
                  <span className="summary-title">Columns to be created in Excel sheet:</span>
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

            {/* INTERACTIVE TABLE PREVIEW */}
            <section className="portal-section">
              <div className="section-title">
                <span>02</span>
                <div>
                  <h2>Live Template Preview</h2>
                  <p>Preview sample rows before downloading. Use arrow keys (↑ ↓ ← →) to edit & navigate cells.</p>
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
                          {/* Name */}
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
                          {/* PRN */}
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
                          {/* CA1 */}
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
                          {/* MSE */}
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
                          {/* CA2 */}
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
                          {/* ESE */}
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
                          {/* Dynamic COs */}
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

            {/* DOWNLOAD ACTIONS SECTION */}
            <section className="calculate-section template-download-section">
              <div>
                <span className="calculate-label">READY TO EXPORT?</span>
                <h2>
                  Download
                  <br />
                  <span>Excel Sheet.</span>
                </h2>
                <p className="template-download-sub">
                  Generated template includes columns for {templateCoCount} Course Outcomes ready for direct data entry.
                </p>
              </div>

              <div className="template-action-buttons">
                <button
                  type="button"
                  className="calculate-button download-excel-main-btn"
                  onClick={downloadExcelTemplate}
                >
                  Download Excel (.xlsx)
                  <span>↓</span>
                </button>

                <button
                  type="button"
                  className="clear-button download-csv-btn"
                  onClick={downloadCsvTemplate}
                >
                  Download CSV (.csv)
                </button>
              </div>
            </section>

            {templateDownloaded && (
              <div className="download-toast">
                ✓ Template with {templateCoCount} COs downloaded successfully!
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

export default Portal;