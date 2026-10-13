# CO-PO Automated Attainment Workbench: Complete Architecture & Backend Specification Report

---

## 1. Executive Summary & System Overview

The **CO-PO Automated Attainment System** is an Outcome-Based Education (OBE) evaluation platform compliant with National Board of Accreditation (NBA) guidelines. It automates the calculation of:
1. **Course Outcome (CO) Direct Attainment** from continuous assessments (CA-1, CA-2, Mid-Sem, End-Sem, Lab exams).
2. **Course Exit Survey (CES) Indirect Attainment** from student satisfaction metrics ($X, Y, Z$ levels).
3. **Combined Final CO Attainment** (90% Direct + 10% Indirect).
4. **Program Outcome (PO1–PO12) & Program Specific Outcome (PSO1–PSO3) Attainment** using correlation mapping matrices (Levels 0–3) and weighted formulas.
5. **Bulk Data Ingestion & Reporting**: Flexible Excel parser for various college marksheets, plus automated export to multi-sheet Excel workbooks and NBA-formatted `.docx` institutional reports.

If you are porting this backend and logic to a **new project with a different frontend**, this document provides the complete blueprint: database schema, mathematical algorithms, API contracts, file parsing logic, and migration strategies.

---

## 2. Technology Stack & Dependencies

| Layer | Technology | Key Packages | Notes |
|---|---|---|---|
| **Runtime & Server** | Node.js (CommonJS) | `express` (^5.2.1), `cors` (^2.8.6), `dotenv` (^17.4.2) | REST API structure |
| **Database & ORM** | MySQL | `@prisma/client` (^6.19.0), `prisma` (^6.19.0) | Relational model with migrations |
| **Authentication** | JWT & Bcrypt | `jsonwebtoken` (^9.0.3), `bcrypt` (^6.0.0) | RBAC (`ADMIN`, `TEACHER`, `COORDINATOR`) |
| **File Processing** | Multer & ExcelJS | `multer` (^2.3.0), `exceljs` (^4.4.0) | In-memory stream parsing and template generation |
| **Report Generation** | Docx | `docx` (^9.7.1) | Generates official Word documents with formatting |

---

## 3. Database Schema (Prisma / MySQL)

### 3.1 Entity Relationship Diagram

```mermaid
erDiagram
    Department ||--o{ Subject : "offers"
    Teacher ||--o{ Subject : "instructed_by"
    User ||--o{ Teacher : "profile_link"
    Subject ||--o{ CourseOutcome : "defines"
    Subject ||--o{ CoPoMapping : "maps"
    Subject ||--o{ StudentMark : "contains"
    Subject ||--o{ CesFeedback : "aggregates"
    CourseOutcome ||--o{ CoPoMapping : "mapped_to"
    CourseOutcome ||--o{ StudentMark : "evaluated_in"
    CourseOutcome ||--o{ CesFeedback : "surveyed_in"
    ProgramOutcome ||--o{ CoPoMapping : "correlates"
    ProgramSpecificOutcome ||--o{ CoPoMapping : "correlates"
```

### 3.2 Data Models

```prisma
enum Role {
  ADMIN
  TEACHER
  COORDINATOR
}

model User {
  id           Int      @id @default(autoincrement())
  name         String
  email        String   @unique
  passwordHash String
  role         Role     @default(TEACHER)
  createdAt    DateTime @default(now())
  updatedAt    DateTime @updatedAt
}

model Teacher {
  id        Int      @id @default(autoincrement())
  name      String
  email     String   @unique
  createdAt DateTime @default(now())
  subjects  Subject[]
}

model Department {
  id        Int      @id @default(autoincrement())
  name      String   @unique
  code      String   @unique
  createdAt DateTime @default(now())
  subjects  Subject[]
}

model Subject {
  id           Int        @id @default(autoincrement())
  code         String     @unique
  name         String
  semester     Int        // 1 to 8
  academicYear String     // e.g. "2025-26"
  teacherId    Int
  departmentId Int

  teacher      Teacher    @relation(fields: [teacherId], references: [id])
  department   Department @relation(fields: [departmentId], references: [id])

  cos          CourseOutcome[]
  coPoMappings CoPoMapping[]
  studentMarks StudentMark[]
  cesFeedbacks CesFeedback[]
}

model CourseOutcome {
  id        Int    @id @default(autoincrement())
  code      String // e.g. "CO1", "CO2"
  statement String
  subjectId Int
  subject   Subject @relation(fields: [subjectId], references: [id])

  coPoMappings CoPoMapping[]
  studentMarks StudentMark[]
  cesFeedbacks CesFeedback[]

  @@unique([subjectId, code])
}

model ProgramOutcome {
  id        Int    @id @default(autoincrement())
  code      String @unique // PO1 to PO12
  statement String
  coPoMappings CoPoMapping[]
}

model ProgramSpecificOutcome {
  id        Int    @id @default(autoincrement())
  code      String @unique // PSO1 to PSO3
  statement String
  coPoMappings CoPoMapping[]
}

model CoPoMapping {
  id        Int      @id @default(autoincrement())
  subjectId Int
  coId      Int
  poId      Int?
  psoId     Int?
  level     Float    // 0 = No mapping, 1 = Low, 2 = Medium, 3 = High

  subject   Subject                 @relation(fields: [subjectId], references: [id])
  co        CourseOutcome           @relation(fields: [coId], references: [id])
  po        ProgramOutcome?         @relation(fields: [poId], references: [id])
  pso       ProgramSpecificOutcome? @relation(fields: [psoId], references: [id])

  @@index([coId])
  @@index([poId])
  @@index([psoId])
}

model StudentMark {
  id            Int     @id @default(autoincrement())
  studentRollNo String
  marks         Float
  subjectId     Int
  coId          Int

  subject       Subject       @relation(fields: [subjectId], references: [id])
  co            CourseOutcome @relation(fields: [coId], references: [id])

  @@index([studentRollNo])
  @@index([subjectId])
  @@index([coId])
}

model CesFeedback {
  id        Int @id @default(autoincrement())
  subjectId Int
  coId      Int

  countX    Int @default(0) // > 70% satisfaction
  countY    Int @default(0) // 55% - 70% satisfaction
  countZ    Int @default(0) // < 55% satisfaction

  subject   Subject       @relation(fields: [subjectId], references: [id])
  co        CourseOutcome @relation(fields: [coId], references: [id])

  @@unique([subjectId, coId])
  @@index([subjectId])
  @@index([coId])
}
```

---

## 4. Mathematical Calculations & Business Logic

### 4.1 Max Marks Scaling (Dynamic Scale Detection)
When marks are ingested, the system determines the maximum possible mark for each CO dynamically by checking the highest observed student score:
$$\text{Max Observed} = \max(\text{marks for that CO})$$
$$\text{Scale Target} = \begin{cases} 
10 & \text{if } \max \le 10 \\
15 & \text{if } \max \le 15 \\
20 & \text{if } \max \le 20 \\
25 & \text{if } \max \le 25 \\
30 & \text{if } \max \le 30 \\
40 & \text{if } \max \le 40 \\
50 & \text{if } \max \le 50 \\
75 & \text{if } \max \le 75 \\
100 & \text{otherwise}
\end{cases}$$

### 4.2 Direct CO Attainment & Target Cutoff
- **Assessment Groups**:
  - The first half of COs ($\text{index} < \lceil N_{\text{CO}} / 2 \rceil$) belongs to **CA-1**.
  - The second half belongs to **CA-2**.
- Teachers can specify distinct target cutoff percentages:
  - $\text{Threshold}_{\text{CA1}}$ (default 60%)
  - $\text{Threshold}_{\text{CA2}}$ (default 60% or 70%)
- **Cutoff Marks per CO**:
  $$\text{Cutoff} = \left(\frac{\text{Target \%}}{100}\right) \times \text{coMaxMarks}$$
- **Direct CO Attainment Percentage**:
  $$\text{CO Attainment \%} = \left(\frac{\text{Students with Marks} \ge \text{Cutoff}}{\text{Total Students Enrolled}}\right) \times 100$$
- **NBA Attainment Level Mapping (0 to 3 scale)**:
  $$\text{Level}_{\text{CA}} = \begin{cases}
  3 & \text{if } \text{CO Attainment \%} \ge 80\% \\
  2 & \text{if } 70\% \le \text{CO Attainment \%} < 80\% \\
  1 & \text{if } 60\% \le \text{CO Attainment \%} < 70\% \\
  0 & \text{if } \text{CO Attainment \%} < 60\%
  \end{cases}$$

### 4.3 Indirect Attainment: Course Exit Survey (CES)
Feedback is classified into three survey bins:
- $X$: High satisfaction ($>70\%$ score or Rating 3)
- $Y$: Moderate satisfaction ($55\% - 70\%$ score or Rating 2)
- $Z$: Low satisfaction ($<55\%$ score or Rating 1)

$$\text{CES Attainment} = \frac{3 \cdot X + 2 \cdot Y + 1 \cdot Z}{X + Y + Z}$$

### 4.4 Final (Overall) CO Attainment
Combines Direct Continuous Assessment ($90\%$) and Indirect Course Exit Survey ($10\%$):
$$\text{Final CO Attainment} = (0.90 \times \text{Level}_{\text{CA}}) + (0.10 \times \text{CES Attainment})$$

### 4.5 PO / PSO Attainment (Direct & Weighted)
There are two complementary methods supported:
1. **Matrix Cell Level Calculation**:
   For each $\text{CO}_i$ and $\text{PO}_j$ with mapping correlation level $L_{i,j} \in \{1, 2, 3\}$:
   $$\text{Cell Attainment}_{i, j} = \left\lceil \frac{L_{i,j} \times \text{Final CO}_i}{3.0} \times 100 \right\rceil \div 100$$
   The direct PO score is the arithmetic average across all mapped COs:
   $$\text{PO}_j = \frac{\sum_{i \in \text{Mapped}} \text{Cell Attainment}_{i, j}}{N_{\text{Mapped COs}}}$$

2. **Weighted Average Method (Direct Percentage)**:
   $$\text{PO Attainment \%} = \frac{\sum_{i} (\text{CO Attainment \%}_i \times L_{i,j})}{\sum_{i} L_{i,j}} \quad (\text{for } L_{i,j} > 0)$$

---

## 5. API Endpoints Specification

### 5.1 Authentication (`/api/auth`)
| Method | Endpoint | Access | Body / Params | Description |
|---|---|---|---|---|
| `POST` | `/api/auth/register` | Public (or Admin for roles) | `{ name, email, password, role }` | Registers user (Admin creates Admin/Coord) |
| `POST` | `/api/auth/login` | Public | `{ email, password }` | Authenticates and returns `{ token, user }` |
| `GET` | `/api/auth/me` | Authenticated | Bearer Token Header | Returns active user profile |
| `PUT` | `/api/auth/update-password` | Authenticated | `{ currentPassword, newPassword }` | Changes password |

### 5.2 Academic Master Data
| Method | Endpoint | Description |
|---|---|---|
| `GET / POST / PUT / DELETE` | `/api/departments` | Full CRUD for departments (`name`, `code`) |
| `GET / POST / PUT / DELETE` | `/api/teachers` | Full CRUD for teachers (`name`, `email`) |
| `GET / POST / PUT / DELETE` | `/api/subjects` | Full CRUD for subjects (`code`, `name`, `semester`, `academicYear`, `teacherId`, `departmentId`) |
| `GET / POST / PUT / DELETE` | `/api/course-outcomes` | CRUD for COs (`code`, `statement`, `subjectId`) |
| `GET / POST / PUT / DELETE` | `/api/program-outcomes` | CRUD for PO1–PO12 (`code`, `statement`) |
| `GET / POST / PUT / DELETE` | `/api/program-specific-outcomes` | CRUD for PSO1–PSO3 (`code`, `statement`) |

### 5.3 Mappings & Student Marks
| Method | Endpoint | Body / Query | Description |
|---|---|---|---|
| `GET` | `/api/co-po-mappings` | `?subjectId=X` | Get all correlation mappings for a subject |
| `POST` | `/api/co-po-mappings` | `{ subjectId, mappings: [{ coId, poId, psoId, level }] }` | Bulk save entire correlation matrix |
| `GET` | `/api/student-marks` | `?subjectId=X` | Retrieve all student mark records for subject |
| `POST` | `/api/student-marks` | `{ studentRollNo, marks, subjectId, coId }` | Save single student mark |

### 5.4 Course Exit Survey (CES) (`/api/ces`)
| Method | Endpoint | Body | Description |
|---|---|---|---|
| `GET` | `/api/ces/:subjectId` | — | Get feedback counts ($X, Y, Z$) per CO |
| `POST` | `/api/ces/save` | `{ subjectId, coId, countX, countY, countZ }` | Upsert single CO feedback |
| `POST` | `/api/ces/bulk` | `{ subjectId, feedbacks: [{ coId, countX, countY, countZ }] }` | Bulk upsert feedback for all COs |
| `POST` | `/api/ces/generate-from-marks/:subjectId` | — | Auto-generate $X, Y, Z$ based on student marks distribution ($\ge 70\% \to X, 55\text{--}69\% \to Y, <55\% \to Z$) |

### 5.5 Attainment Engine (`/api/attainment`)
| Method | Endpoint | Query Parameters | Response |
|---|---|---|---|
| `GET` | `/api/attainment/co/:subjectId` | `threshold`, `thresholdCa1`, `thresholdCa2` | CO attainment stats (total, above threshold, %) |
| `GET` | `/api/attainment/po/:subjectId` | `threshold`, `thresholdCa1`, `thresholdCa2` | PO attainment weighted averages |
| `GET` | `/api/attainment/pso/:subjectId` | `threshold`, `thresholdCa1`, `thresholdCa2` | PSO attainment weighted averages |
| `GET` | `/api/attainment/matrix/:subjectId` | `threshold`, `thresholdCa1`, `thresholdCa2` | Full integrated matrix (COs, POs, PSOs, Levels, Attainment) |

### 5.6 File Upload & Document Generation (`/api/excel`)
| Method | Endpoint | Query / Body | Output |
|---|---|---|---|
| `GET` | `/api/excel/template/student-marks` | `?subjectId=X` | Downloads `.xlsx` student mark entry template |
| `GET` | `/api/excel/template/ces` | `?subjectId=X` | Downloads `.xlsx` CES survey input template |
| `POST` | `/api/excel/upload-marks` | `multipart/form-data`: `file`, `subjectId` | Parses and bulk-upserts student marks |
| `POST` | `/api/excel/upload-ces` | `multipart/form-data`: `file`, `subjectId` | Parses aggregated or raw Google Form survey data |
| `GET` | `/api/excel/export-attainment/:subjectId` | `?targetPercentage=60&targetPercentageCa1=60&targetPercentageCa2=70` | Generates full `.xlsx` workbook (Summary + Matrix) |
| `GET` | `/api/excel/export-word-report/:subjectId` | `?targetPercentage=60&targetPercentageCa1=60&targetPercentageCa2=70` | Generates official institutional `.docx` report |

---

## 6. Excel & Document Processing Mechanics

### 6.1 Student Marks Excel Parser
The parser handles college spreadsheets without strict column indices:
1. **Header Row Detection**: Scans first 10 rows for keywords (`PRN`, `ROLL`, `ENROLLMENT`, `USN`, `STUDENT`, `CA1`, `CA2`, `CO1`, `MARKS`).
2. **Identifier Extraction**: Strips prefixes like `PRN: 2310104005` $\to$ `2310104005`. Filters out summary/aggregate rows (`TOTAL`, `AVERAGE`, `MEAN`, `MAX`, `MIN`).
3. **Column to CO Mapping**: Maps columns matching `CO1..CO6`, `CA1`, `CA2`, `MSE`, `ESE`, `IA1` to corresponding subject CO IDs in sorted order.
4. **Bulk Upsert Performance**: Uses Prisma transaction with `deleteMany` and `createMany` on unique `(studentRollNo, coId, subjectId)` combinations for fast bulk operations.

### 6.2 Course Exit Survey (CES) Parser
Supports two distinct input formats:
- **Aggregated Sheet**: Columns for `CO`, `X (>70%)`, `Y (55-70%)`, `Z (<55%)`.
- **Raw Student Responses (Google Forms export)**: Rows containing student ratings (scale 1–3, 1–5, or marks). The parser bins each response into $X, Y, Z$ automatically.

### 6.3 Word Document Export (`docx`)
Produces an accreditation-ready document featuring:
- Official Institutional Header (Institute Name, Trust, Department).
- Course details (Course Code, Name, Semester, Academic Year, Faculty In-Charge).
- Course Outcome Statements list.
- Formatted tables with colored header cells for Direct Attainment, Indirect Attainment, and CO-PO Correlation Matrix.
- Signature Block at the bottom: **Prepared By** (Teacher) — **Verified By** (Module Coordinator) — **Head of Department**.

---

## 7. Migration & Porting Guide for Your New Project

When connecting a **different frontend** to this backend:

```mermaid
flowchart TD
    A[New Frontend] -->|Auth Token| B[API Gateway / Express Server]
    B --> C[Auth Middleware]
    B --> D[Domain Controllers]
    D --> E[Attainment Calculation Engine]
    D --> F[Excel / Word Generation Engine]
    D --> G[Prisma ORM Layer]
    G --> H[(MySQL Database)]
```

### 7.1 Backend Setup Checklist
1. **Environment Variables (`.env`)**:
   ```env
   PORT=5000
   DATABASE_URL="mysql://username:password@localhost:3306/copo_db"
   JWT_SECRET="your_strong_jwt_secret"
   ```
2. **Database Initialization**:
   ```bash
   npx prisma generate
   npx prisma migrate dev --name init
   npm run seed
   ```
3. **Seed Data**: Run `prisma/seed.js` to populate:
   - Default users (`admin@college.edu`, `teacher@college.edu`).
   - PO1 to PO12 (NBA Engineering graduate attributes).
   - PSO1 to PSO3.
   - Departments (`CSE`, `IT`).

### 7.2 Frontend Implementation Steps (Step-by-Step Flow)
For your new frontend, the user journey is structured into **6 clear steps**:

| Step | Screen / View | Associated APIs |
|---|---|---|
| **Step 1: Profile & Auth** | Login & User State | `POST /api/auth/login`, `GET /api/auth/me` |
| **Step 2: Department Selection** | Department Dropdown / Cards | `GET /api/departments` |
| **Step 3: Subject Selection** | Subject Picker (filter by Semester / Year) | `GET /api/subjects`, `GET /api/subjects/:id` |
| **Step 4: CO-PO Matrix Editor** | Matrix Grid ($N_{\text{CO}} \times 15$ cells for PO1–12, PSO1–3) | `GET /api/program-outcomes`, `GET /api/co-po-mappings?subjectId=X`, `POST /api/co-po-mappings` |
| **Step 5: Student Marks Ingestion** | Drag & Drop Excel Upload + Manual Entry Table | `GET /api/excel/template/student-marks?subjectId=X`, `POST /api/excel/upload-marks`, `GET /api/student-marks?subjectId=X` |
| **Step 6: Attainment Dashboard & Reports** | Target Slider (CA-1, CA-2), CES Table, Combined Results, Export Buttons | `GET /api/attainment/matrix/:subjectId`, `GET /api/ces/:subjectId`, `POST /api/ces/generate-from-marks/:subjectId`, `GET /api/excel/export-attainment/:subjectId`, `GET /api/excel/export-word-report/:subjectId` |

---

## 8. Summary of Reusable Assets

All of the core logic files can be reused directly in your new system without modification:
1. `backend/src/services/attainment.service.js` — Core mathematical formulas for Direct, PO, and PSO attainment.
2. `backend/src/controllers/excel.controller.js` — All Excel/CSV ingestion, template creation, and Word document report generation.
3. `backend/src/controllers/ces.controller.js` — Course Exit Survey ingestion, calculation, and mark-based simulation.
4. `backend/prisma/schema.prisma` — Normalized relational database schema.
