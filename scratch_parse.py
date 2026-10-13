import openpyxl
import json

excel_path = r"C:\Users\nikhi\Documents\5th sem project\CO-PO 24-25-NEP.xlsx"
wb = openpyxl.load_workbook(excel_path, data_only=True)

subjects = {}

def get_or_create(code, title=None, year=None):
    code = str(code).strip()
    if not code or code in ["Course Code", "None", "statements"]:
        return None
    if code not in subjects:
        subjects[code] = {
            "code": code,
            "title": str(title).strip() if title else "",
            "academicYear": year or "TY",
            "cos": {},
            "mapping": {}
        }
    elif title and not subjects[code]["title"]:
        subjects[code]["title"] = str(title).strip()
    return subjects[code]

# 1. Parse CO sheets:
# 'SY TY & B.Tech CO Statement (25', 'SY TY & B.Tech CO Statements', ' B.Tech CO Statement '
co_sheets = [
    "SY TY & B.Tech CO Statement (25",
    "SY TY & B.Tech CO Statements",
    " B.Tech CO Statement "
]

for sheet_name in co_sheets:
    if sheet_name not in wb.sheetnames:
        continue
    sheet = wb[sheet_name]
    curr_sub = None
    for r in range(10, sheet.max_row + 1):
        c2 = sheet.cell(r, 2).value
        c3 = sheet.cell(r, 3).value
        c4 = sheet.cell(r, 4).value
        c5 = sheet.cell(r, 5).value

        if c2 and str(c2).strip() and str(c2).strip() not in ["Course Code", "None"]:
            curr_sub = get_or_create(c2, c3)

        if curr_sub and c4 and str(c4).strip().startswith("CO"):
            co_num = str(c4).strip()
            co_stmt = str(c5).strip() if c5 else ""
            if co_num not in curr_sub["cos"] or not curr_sub["cos"][co_num]:
                curr_sub["cos"][co_num] = co_stmt

# 2. Parse Mapping sheets:
# 'SY TY & B.Tech CO CO PO Mapping', 'co po mapping B.Tech'
mapping_sheets = [
    "SY TY & B.Tech CO CO PO Mapping",
    "co po mapping B.Tech"
]

for sheet_name in mapping_sheets:
    if sheet_name not in wb.sheetnames:
        continue
    sheet = wb[sheet_name]

    # Find headers in row 9
    headers = {}
    for col in range(5, sheet.max_column + 1):
        val = sheet.cell(9, col).value
        if val:
            clean_hdr = str(val).strip().replace(" ", "").upper()
            headers[col] = clean_hdr

    curr_sub = None
    for r in range(10, sheet.max_row + 1):
        c2 = sheet.cell(r, 2).value
        c3 = sheet.cell(r, 3).value
        c4 = sheet.cell(r, 4).value

        if c2 and str(c2).strip() and str(c2).strip() not in ["Course Code", "None"]:
            curr_sub = get_or_create(c2, c3)

        if curr_sub and c4 and str(c4).strip().startswith("CO"):
            co_num = str(c4).strip()
            if co_num not in curr_sub["mapping"]:
                curr_sub["mapping"][co_num] = {}
            for col, po_key in headers.items():
                val = sheet.cell(r, col).value
                if val is not None:
                    try:
                        v = int(float(val))
                        if v > 0:
                            curr_sub["mapping"][co_num][po_key] = str(v)
                    except (ValueError, TypeError):
                        pass

# 3. Classify Academic Year accurately:
# SY: 23CS23xx (Sem 3), 23CS24xx (Sem 4), 23CSMDA1/2, 23OECS21, 23HSSM01/02
# TY: 23CS35xx (Sem 5), 23CS36xx (Sem 6), CS5xx, CS6xx, 23CSMDA3/4, 23HSSM05/07
# BTech: CS7xx (Sem 7), CS8xx (Sem 8), PRJ06/07, HMS09, IFT03, 23CS47xx, 23CS48xx
for code, sub in subjects.items():
    c = code.upper()
    if (
        "23CS23" in c or "23CS24" in c or "23CS2" in c
        or c in ["23CSMDA1", "23CSMDA2", "23OECS21", "23HSSM01", "23HSSM02"]
        or "SY" in sub["title"].upper()
    ):
        sub["academicYear"] = "SY"
    elif (
        "23CS35" in c or "23CS36" in c or "23CS3" in c
        or c.startswith("CS5") or c.startswith("CS6")
        or c in ["23CSMDA3", "23CSMDA4", "23HSSM05", "23HSSM07"]
        or "TY" in sub["title"].upper()
    ):
        sub["academicYear"] = "TY"
    elif (
        c.startswith("CS7") or c.startswith("CS8")
        or "PRJ" in c or "HMS" in c or "IFT" in c
        or "23CS4" in c or "B.TECH" in sub["title"].upper()
    ):
        sub["academicYear"] = "BTech"
    else:
        sub["academicYear"] = "TY"

# Sort subjects by code
sorted_subjects = sorted(list(subjects.values()), key=lambda x: (x["academicYear"], x["code"]))

print(f"Extracted {len(sorted_subjects)} valid subjects.")
by_year = {}
for s in sorted_subjects:
    by_year[s["academicYear"]] = by_year.get(s["academicYear"], 0) + 1
print("Counts by Academic Year:", by_year)

with open("extracted_curriculum.json", "w", encoding="utf-8") as f:
    json.dump(sorted_subjects, f, indent=2, ensure_ascii=False)

# Also write directly to frontend/src/data/curriculumData.json and .js
import os
os.makedirs("frontend/src/data", exist_ok=True)
with open("frontend/src/data/curriculumData.json", "w", encoding="utf-8") as f:
    json.dump(sorted_subjects, f, indent=2, ensure_ascii=False)

with open("frontend/src/data/curriculumData.js", "w", encoding="utf-8") as f:
    f.write("// Auto-generated from CO-PO 24-25-NEP.xlsx\n")
    f.write("export const curriculumSubjects = ")
    json.dump(sorted_subjects, f, indent=2, ensure_ascii=False)
    f.write(";\n\nexport default curriculumSubjects;\n")

print("Generated frontend/src/data/curriculumData.js and json!")
