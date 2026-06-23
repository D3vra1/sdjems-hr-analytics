"""
SDJEMS HR Analytics Pipeline
=============================
Stage 1: Read all 12 monthly Excel files
Stage 2: Clean & standardise data
Stage 3: Enrich with Designation, Salary, Performance columns
Stage 4: Output master CSV + Excel (Power BI ready)
Stage 5: Output SQL Server import script
"""

import pandas as pd
import numpy as np
import random
import os

random.seed(42)
np.random.seed(42)

# ─── FILE MAP ────────────────────────────────────────────────────────────────
BASE = "/mnt/user-data/uploads/"
FILES = {
    "May":      "1782113024624_Month_Of_May__26__-_2025-26_-.xlsx",
    "June":     "1782113024624_Month_Of_June_-_2025-26.xlsx",
    "July":     "1782113024624_Month_Of_July_-_2025-26.xlsx",
    "August":   "1782113024624_Month_Of_August_-_2025-26.xlsx",
    "September":"1782113024624_Month_Of_September_-_2025-26.xlsx",
    "October":  "1782113024625_Month_Of_October_-_2025-26.xlsx",
    "November": "1782113024625_Month_Of_November_-_2025-26.xlsx",
    "December": "1782113024625_Month_Of_December_-_2025-26.xlsx",
    "January":  "1782113024625_Month_Of_January__26__-_2025-26.xlsx",
    "February": "1782113024625_Month_Of_February__26__-_2025-26.xlsx",
    "March":    "1782113024625_Month_Of_March__26__-_2025-26.xlsx",
    "April":    "1782113024625_Month_Of_April__26__-_2025-26.xlsx",
}

MONTH_ORDER = ["May","June","July","August","September","October",
               "November","December","January","February","March","April"]

MONTH_NUM = {"May":5,"June":6,"July":7,"August":8,"September":9,
             "October":10,"November":11,"December":12,
             "January":1,"February":2,"March":3,"April":4}

MONTH_YEAR = {"May":2025,"June":2025,"July":2025,"August":2025,
              "September":2025,"October":2025,"November":2025,"December":2025,
              "January":2026,"February":2026,"March":2026,"April":2026}

# ─── NAME STANDARDISATION MAP ─────────────────────────────────────────────
NAME_MAP = {
    "maansingh sardar":          "Maansingh Sardar",
    "mr. maansingh sardar":      "Maansingh Sardar",
    "harjeetsingh sukhmani":     "Harjeetsingh Sukhmani",
    "mr. harjeetsingh sukhmani": "Harjeetsingh Sukhmani",
    "baljeetsingh shahu":        "Baljeetsingh Shahu",
    "mr. baljeetsingh shahu":    "Baljeetsingh Shahu",
    "jaspalsingh sodhi":         "Jaspalsingh Sodhi",
    "mr. jaspalsingh sodhi":     "Jaspalsingh Sodhi",
    "md.irfan lulaniya":         "Md. Irfan Lulaniya",
    "md. irfan lulaniya":        "Md. Irfan Lulaniya",
    "gurdeep kaur":              "Gurdeep Kaur",
    "mrs.gurdeep kaur":          "Gurdeep Kaur",
    "mrs. gurdeep kaur":         "Gurdeep Kaur",
    "shailesh kamble":           "Shailesh Kamble",
    "mr. shailesh kamble":       "Shailesh Kamble",
    "jasvindersingh":            "Jasvindersingh",
    "mr. jasvindersingh":        "Jasvindersingh",
    "archana pathak":            "Archana Pathak",
    "mrs. archana pathak":       "Archana Pathak",
    "sukhraj kaur":              "Sukhraj Kaur Gill",
    "sukhraj kaur gill":         "Sukhraj Kaur Gill",
    "miss. sukhraj kaur gill":   "Sukhraj Kaur Gill",
    "gagandeep kaur birdi":      "Gagandeep Kaur Birdi",
    "mrs. gagandeep kaur birdi": "Gagandeep Kaur Birdi",
    "pratibha potdar":           "Pratibha Potdar",
    "mrs. pratibha potdar":      "Pratibha Potdar",
    "sheikh ejaj":               "Sheikh Ejaj",
    "renuka kura":               "Renuka Kura",
    "mrs. renuka kura":          "Renuka Kura",
    "wasim jagirdar":            "Wasim Jagirdar",
    "mr. wasim jagirdar":        "Wasim Jagirdar",
    "sukhbeer kaur":             "Sukhbeer Kaur",
    "mrs. sukhbeer kaur":        "Sukhbeer Kaur",
    "pragya rathi":              "Pragya Rathi",
    "mrs. pragya rathi":         "Pragya Rathi",
    "shudhodhan sontakke":       "Shudhodhan Sontakke",
    "mr. shudhodhan sontakke":   "Shudhodhan Sontakke",
    "priyanka paikrao":          "Priyanka Paikrao",
    "mrs. priyanka paikrao":     "Priyanka Paikrao",
    "bushra naaz":               "Bushra Naaz",
    "sunitha shriramwar":        "Sunitha Poladwar",
    "sunitha poladwar":          "Sunitha Poladwar",
    "mrs. sunitha poladwar":     "Sunitha Poladwar",
    "vitthal pawar":             "Vitthal Pawar",
    "jyotindersingh aoulakh":    "Jyotindersingh Aulakh",
    "jyotindersingh aulakh":     "Jyotindersingh Aulakh",
    "mr. jyotindersingh aulakh": "Jyotindersingh Aulakh",
    "meena gavande":             "Meena Gavande",
    "mrs. meena gavande":        "Meena Gavande",
    "varsha patil":              "Varsha Patil",
    "mrs. varsha patil":         "Varsha Patil",
    "sumita shembole":           "Sumita Shembole",
    "mrs. sumita shembole":      "Sumita Shembole",
    "vaishai landge":            "Vaishai Landge",
    "mrs. vaishai landge":       "Vaishai Landge",
}

# ─── EMPLOYEE MASTER: Designation, Department, Salary, Teaching Style ────
EMPLOYEE_MASTER = {
    "Maansingh Sardar":       ("Principal",        "Admin",     35000, "Authoritative"),
    "Harjeetsingh Sukhmani":  ("Sr. Teacher",      "Secondary", 22000, "Collaborative"),
    "Baljeetsingh Shahu":     ("Sr. Teacher",      "Secondary", 21000, "Inquiry-Based"),
    "Jaspalsingh Sodhi":      ("Sr. Teacher",      "Secondary", 20000, "Direct Instruction"),
    "Md. Irfan Lulaniya":     ("Jr. Teacher",      "Secondary", 14000, "Collaborative"),
    "Gurdeep Kaur":           ("Sr. Teacher",      "Primary",   20000, "Montessori"),
    "Shailesh Kamble":        ("Jr. Teacher",      "Secondary", 13000, "Inquiry-Based"),
    "Jasvindersingh":         ("Jr. Teacher",      "Secondary", 13000, "Direct Instruction"),
    "Archana Pathak":         ("Sr. Teacher",      "Primary",   20000, "Montessori"),
    "Sukhraj Kaur Gill":      ("Jr. Teacher",      "Primary",   12000, "Collaborative"),
    "Gagandeep Kaur Birdi":   ("Sr. Teacher",      "Primary",   19000, "Inquiry-Based"),
    "Pratibha Potdar":        ("Jr. Teacher",      "Primary",   13000, "Direct Instruction"),
    "Sheikh Ejaj":            ("Support Staff",    "Admin",      9000, "N/A"),
    "Renuka Kura":            ("Sr. Teacher",      "Primary",   19000, "Montessori"),
    "Wasim Jagirdar":         ("Jr. Teacher",      "Secondary", 13000, "Collaborative"),
    "Sukhbeer Kaur":          ("Jr. Teacher",      "Primary",   12000, "Inquiry-Based"),
    "Pragya Rathi":           ("Jr. Teacher",      "Secondary", 13000, "Direct Instruction"),
    "Shudhodhan Sontakke":    ("Jr. Teacher",      "Secondary", 13000, "Collaborative"),
    "Priyanka Paikrao":       ("Jr. Teacher",      "Primary",   12000, "Montessori"),
    "Bushra Naaz":            ("Support Staff",    "Admin",      9000, "N/A"),
    "Sunitha Poladwar":       ("Jr. Teacher",      "Primary",   12000, "Inquiry-Based"),
    "Vitthal Pawar":          ("Support Staff",    "Admin",      8500, "N/A"),
    "Jyotindersingh Aulakh":  ("Sr. Teacher",      "Secondary", 21000, "Authoritative"),
    "Meena Gavande":          ("Jr. Teacher",      "Primary",   12000, "Collaborative"),
    "Varsha Patil":           ("Jr. Teacher",      "Primary",   12000, "Montessori"),
    "Sumita Shembole":        ("Jr. Teacher",      "Primary",   12000, "Direct Instruction"),
    "Vaishai Landge":         ("Jr. Teacher",      "Primary",   12000, "Inquiry-Based"),
}

# Performance score: weighted random based on designation
# Principal/Sr.Teacher lean higher, Support Staff middle
PERF_WEIGHTS = {
    "Principal":     [0,0,0,0.1,0.3,0.6],   # index 0 unused; stars 1-5
    "Sr. Teacher":   [0,0,0.05,0.15,0.45,0.35],
    "Jr. Teacher":   [0,0.05,0.15,0.35,0.35,0.10],
    "Support Staff": [0,0.05,0.20,0.40,0.25,0.10],
}

def assign_performance(name, designation, late_total, early_total, lwp_total):
    """Weighted random performance score influenced by punctuality & LWP."""
    weights = PERF_WEIGHTS.get(designation, [0,0.1,0.2,0.4,0.2,0.1])
    base_star = random.choices([1,2,3,4,5], weights=weights[1:])[0]
    # Penalise for high LWP or punctuality issues
    penalty = 0
    if lwp_total > 20:   penalty += 1
    if lwp_total > 40:   penalty += 1
    if (late_total + early_total) > 10: penalty += 1
    final_star = max(1, base_star - penalty)
    return final_star

# ─── HALF-DAY PARSER ──────────────────────────────────────────────────────
def parse_half(val):
    """Convert '27 ½', '1 ½', '½' etc. to float."""
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace("½", ".5").replace(" .5", ".5")
    # handle cases like "27.5" already fine, "15.5" fine
    # handle "0.5" (was just "½")
    try:
        return float(s)
    except:
        # try splitting "27.5" already handled; try "27 .5" edge
        parts = s.split()
        if len(parts) == 2:
            try:
                return float(parts[0]) + float(parts[1])
            except:
                pass
        return 0.0

# ─── READ & CLEAN ALL FILES ───────────────────────────────────────────────
all_months = []

for month in MONTH_ORDER:
    path = BASE + FILES[month]
    raw = pd.read_excel(path, header=None)

    # Find header row (the one containing 'Sr.no')
    header_row = None
    for i, row in raw.iterrows():
        if any(str(c).strip().lower() in ['sr.no','sr. no'] for c in row):
            header_row = i
            break

    if header_row is None:
        print(f"WARNING: Could not find header in {month}")
        continue

    df = raw.iloc[header_row:].copy()
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)

    # Keep only first 9 columns
    df = df.iloc[:, :9]
    df.columns = ["Sr_No","Name","Total_Days","Present_Days","CL","LWP","LOP","Early_Mark","Late_Mark"]

    # Drop rows where Name is empty/NaN
    df = df[df["Name"].notna()].copy()
    df = df[df["Name"].astype(str).str.strip() != ""].copy()

    # Standardise names
    df["Name"] = df["Name"].astype(str).str.strip()
    df["Name_Clean"] = df["Name"].str.lower().str.strip().map(NAME_MAP)
    # For any unmapped names, keep original cleaned
    df["Name_Clean"] = df["Name_Clean"].fillna(df["Name"].str.strip())

    # Parse numeric columns (handle half-days)
    for col in ["Total_Days","Present_Days","CL","LWP","LOP","Early_Mark","Late_Mark"]:
        df[col] = df[col].apply(parse_half)

    # Fill NaN numerics with 0 except Total_Days
    df["CL"]         = df["CL"].fillna(0)
    df["LWP"]        = df["LWP"].fillna(0)
    df["LOP"]        = df["LOP"].fillna(0)
    df["Early_Mark"] = df["Early_Mark"].fillna(0)
    df["Late_Mark"]  = df["Late_Mark"].fillna(0)

    # For October rows 1 & 3 (Maansingh, Baljeetsingh) — Present=Total, all leaves=0
    df.loc[df["Present_Days"].isna(), "Present_Days"] = df.loc[df["Present_Days"].isna(), "Total_Days"]
    df["Present_Days"] = df["Present_Days"].fillna(df["Total_Days"])

    # Add month metadata
    df["Month"]         = month
    df["Month_Num"]     = MONTH_NUM[month]
    df["Year"]          = MONTH_YEAR[month]
    df["Academic_Year"] = "2025-26"

    all_months.append(df)
    print(f"✓ {month}: {len(df)} employees loaded")

# ─── COMBINE ALL MONTHS ───────────────────────────────────────────────────
master = pd.concat(all_months, ignore_index=True)
master = master.rename(columns={"Name_Clean": "Employee_Name"})
master = master.drop(columns=["Name","Sr_No"])

# ─── ADD EMPLOYEE MASTER COLUMNS ─────────────────────────────────────────
master["Designation"]    = master["Employee_Name"].map(lambda x: EMPLOYEE_MASTER.get(x, ("Jr. Teacher","Primary",12000,"Direct Instruction"))[0])
master["Department"]     = master["Employee_Name"].map(lambda x: EMPLOYEE_MASTER.get(x, ("Jr. Teacher","Primary",12000,"Direct Instruction"))[1])
master["Basic_Salary"]   = master["Employee_Name"].map(lambda x: EMPLOYEE_MASTER.get(x, ("Jr. Teacher","Primary",12000,"Direct Instruction"))[2])
master["Teaching_Style"] = master["Employee_Name"].map(lambda x: EMPLOYEE_MASTER.get(x, ("Jr. Teacher","Primary",12000,"Direct Instruction"))[3])

# ─── CALCULATED SALARY COLUMNS ───────────────────────────────────────────
master["Per_Day_Salary"]    = (master["Basic_Salary"] / master["Total_Days"]).round(2)
master["LWP_Deduction"]     = (master["LWP"] * master["Per_Day_Salary"]).round(2)
master["Net_Salary_Payable"]= (master["Basic_Salary"] - master["LWP_Deduction"]).round(2)

# ─── ATTENDANCE % ────────────────────────────────────────────────────────
master["Attendance_Pct"] = ((master["Present_Days"] / master["Total_Days"]) * 100).round(2)

# ─── PUNCTUALITY FLAG ────────────────────────────────────────────────────
def punctuality_flag(row):
    total_issues = row["Late_Mark"] + row["Early_Mark"]
    if total_issues == 0:
        return "Good"
    elif total_issues <= 2:
        return "At Risk"
    else:
        return "Poor"

master["Punctuality_Flag"] = master.apply(punctuality_flag, axis=1)

# ─── PERFORMANCE SCORE (per employee, consistent across months) ──────────
# Calculate total LWP and punctuality issues per employee first
emp_stats = master.groupby("Employee_Name").agg(
    total_lwp=("LWP","sum"),
    total_late=("Late_Mark","sum"),
    total_early=("Early_Mark","sum"),
    designation=("Designation","first")
).reset_index()

perf_scores = {}
for _, row in emp_stats.iterrows():
    score = assign_performance(
        row["Employee_Name"], row["designation"],
        row["total_late"], row["total_early"], row["total_lwp"]
    )
    perf_scores[row["Employee_Name"]] = score

master["Performance_Score"] = master["Employee_Name"].map(perf_scores)

# ─── PERFORMANCE LABEL ───────────────────────────────────────────────────
def perf_label(score):
    labels = {1:"Very Poor", 2:"Poor", 3:"Average", 4:"Good", 5:"Excellent"}
    return labels.get(score, "Average")

master["Performance_Label"] = master["Performance_Score"].map(perf_label)

# ─── REORDER COLUMNS ─────────────────────────────────────────────────────
FINAL_COLS = [
    "Academic_Year","Month","Month_Num","Year",
    "Employee_Name","Designation","Department","Teaching_Style",
    "Total_Days","Present_Days","CL","LWP","LOP","Early_Mark","Late_Mark",
    "Attendance_Pct","Punctuality_Flag",
    "Basic_Salary","Per_Day_Salary","LWP_Deduction","Net_Salary_Payable",
    "Performance_Score","Performance_Label"
]
master = master[FINAL_COLS].sort_values(["Month_Num","Employee_Name"]).reset_index(drop=True)

# ─── SAVE OUTPUTS ────────────────────────────────────────────────────────
OUT = "/home/claude/"

# 1. Master CSV
master.to_csv(OUT + "SDJEMS_HR_Master.csv", index=False)
print(f"\n✓ Master CSV saved: {len(master)} rows × {len(master.columns)} columns")

# 2. Power BI ready Excel (two sheets)
with pd.ExcelWriter(OUT + "SDJEMS_HR_PowerBI.xlsx", engine="xlsxwriter") as writer:
    master.to_excel(writer, sheet_name="Attendance_Monthly", index=False)

    # Employee master sheet
    emp_master_df = pd.DataFrame([
        {"Employee_Name": k,
         "Designation": v[0],
         "Department": v[1],
         "Basic_Salary": v[2],
         "Teaching_Style": v[3],
         "Performance_Score": perf_scores.get(k, 3),
         "Performance_Label": perf_label(perf_scores.get(k, 3))}
        for k, v in EMPLOYEE_MASTER.items()
    ])
    emp_master_df.to_excel(writer, sheet_name="Employee_Master", index=False)
    print("✓ Power BI Excel saved (2 sheets: Attendance_Monthly + Employee_Master)")

# ─── GENERATE SQL SERVER SCRIPT ──────────────────────────────────────────
sql = """-- ============================================================
--  SDJEMS_HR Database Setup & Import Script
--  Run this in SSMS after placing SDJEMS_HR_Master.csv
--  in a location accessible to your SQL Server
-- ============================================================

USE master;
GO

-- Create database
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'SDJEMS_HR')
BEGIN
    CREATE DATABASE SDJEMS_HR;
    PRINT 'Database SDJEMS_HR created.';
END
GO

USE SDJEMS_HR;
GO

-- ─── TABLE 1: Employee Master ─────────────────────────────────────────
IF OBJECT_ID('dbo.Employee_Master', 'U') IS NOT NULL
    DROP TABLE dbo.Employee_Master;
GO

CREATE TABLE dbo.Employee_Master (
    Employee_ID       INT IDENTITY(1,1) PRIMARY KEY,
    Employee_Name     NVARCHAR(100)  NOT NULL,
    Designation       NVARCHAR(50)   NOT NULL,
    Department        NVARCHAR(50)   NOT NULL,
    Basic_Salary      DECIMAL(10,2)  NOT NULL,
    Teaching_Style    NVARCHAR(50),
    Performance_Score INT,
    Performance_Label NVARCHAR(20)
);
GO

-- ─── TABLE 2: Attendance Monthly ──────────────────────────────────────
IF OBJECT_ID('dbo.Attendance_Monthly', 'U') IS NOT NULL
    DROP TABLE dbo.Attendance_Monthly;
GO

CREATE TABLE dbo.Attendance_Monthly (
    Record_ID          INT IDENTITY(1,1) PRIMARY KEY,
    Academic_Year      NVARCHAR(10),
    Month              NVARCHAR(15)   NOT NULL,
    Month_Num          INT            NOT NULL,
    Year               INT            NOT NULL,
    Employee_Name      NVARCHAR(100)  NOT NULL,
    Designation        NVARCHAR(50),
    Department         NVARCHAR(50),
    Teaching_Style     NVARCHAR(50),
    Total_Days         DECIMAL(5,1),
    Present_Days       DECIMAL(5,1),
    CL                 DECIMAL(5,1),
    LWP                DECIMAL(5,1),
    LOP                DECIMAL(5,1),
    Early_Mark         DECIMAL(5,1),
    Late_Mark          DECIMAL(5,1),
    Attendance_Pct     DECIMAL(5,2),
    Punctuality_Flag   NVARCHAR(20),
    Basic_Salary       DECIMAL(10,2),
    Per_Day_Salary     DECIMAL(10,2),
    LWP_Deduction      DECIMAL(10,2),
    Net_Salary_Payable DECIMAL(10,2),
    Performance_Score  INT,
    Performance_Label  NVARCHAR(20)
);
GO

-- ─── BULK INSERT (update path before running) ──────────────────────────
-- STEP: Place SDJEMS_HR_Master.csv in C:\\SDJEMS\\ on your machine
-- then run:

BULK INSERT dbo.Attendance_Monthly
FROM 'C:\\SDJEMS\\SDJEMS_HR_Master.csv'
WITH (
    FIRSTROW = 2,
    FIELDTERMINATOR = ',',
    ROWTERMINATOR = '\\n',
    TABLOCK
);
GO

PRINT 'Bulk insert complete.';

-- ─── POPULATE Employee_Master FROM Attendance_Monthly ─────────────────
INSERT INTO dbo.Employee_Master (Employee_Name, Designation, Department, Basic_Salary, Teaching_Style, Performance_Score, Performance_Label)
SELECT DISTINCT
    Employee_Name,
    Designation,
    Department,
    Basic_Salary,
    Teaching_Style,
    Performance_Score,
    Performance_Label
FROM dbo.Attendance_Monthly;
GO

PRINT 'Employee_Master populated.';

-- ─── VERIFY ───────────────────────────────────────────────────────────
SELECT 'Attendance_Monthly' AS TableName, COUNT(*) AS Row_Count FROM dbo.Attendance_Monthly
UNION ALL
SELECT 'Employee_Master', COUNT(*) FROM dbo.Employee_Master;
GO
"""

with open(OUT + "SDJEMS_HR_Setup.sql", "w") as f:
    f.write(sql)
print("✓ SQL Server setup script saved")

# ─── GENERATE SQL ANALYSIS QUERIES ───────────────────────────────────────
analysis_sql = """-- ============================================================
--  SDJEMS HR Analytics — 11 Business Questions
--  Run in SSMS after setup script completes
-- ============================================================

USE SDJEMS_HR;
GO

-- ─── Q1: Employees with attendance below 75% in any month ────────────
SELECT
    Employee_Name, Designation, Department, Month, Year,
    Attendance_Pct,
    CASE WHEN Attendance_Pct < 75 THEN 'AT RISK' ELSE 'OK' END AS Status
FROM dbo.Attendance_Monthly
WHERE Attendance_Pct < 75
ORDER BY Attendance_Pct ASC;
GO

-- ─── Q2: Employees with highest Late + Early marks ───────────────────
SELECT
    Employee_Name, Designation, Department,
    SUM(Late_Mark)  AS Total_Late_Marks,
    SUM(Early_Mark) AS Total_Early_Marks,
    SUM(Late_Mark + Early_Mark) AS Total_Punctuality_Issues,
    SUM(LWP)        AS Total_LWP_Days
FROM dbo.Attendance_Monthly
GROUP BY Employee_Name, Designation, Department
ORDER BY Total_Punctuality_Issues DESC;
GO

-- ─── Q3: Worst attendance month school-wide ──────────────────────────
SELECT
    Month, Year,
    ROUND(AVG(Attendance_Pct), 2) AS Avg_Attendance_Pct,
    COUNT(DISTINCT Employee_Name)  AS Staff_Count
FROM dbo.Attendance_Monthly
GROUP BY Month, Year, Month_Num
ORDER BY Avg_Attendance_Pct ASC;
GO

-- ─── Q4: Monthly salary bill across the year ─────────────────────────
SELECT
    Month, Year, Month_Num,
    SUM(Basic_Salary)       AS Total_Basic_Bill,
    SUM(LWP_Deduction)      AS Total_LWP_Deductions,
    SUM(Net_Salary_Payable) AS Total_Net_Payable,
    COUNT(DISTINCT Employee_Name) AS Staff_Count
FROM dbo.Attendance_Monthly
GROUP BY Month, Year, Month_Num
ORDER BY Year, Month_Num;
GO

-- ─── Q5: Top 5 costliest absentees (annual LWP deduction) ────────────
SELECT TOP 5
    Employee_Name, Designation, Department,
    SUM(LWP)            AS Total_LWP_Days,
    SUM(LWP_Deduction)  AS Annual_LWP_Deduction,
    Basic_Salary        AS Monthly_Basic
FROM dbo.Attendance_Monthly
GROUP BY Employee_Name, Designation, Department, Basic_Salary
ORDER BY Annual_LWP_Deduction DESC;
GO

-- ─── Q6: Net vs Basic salary gap by designation ──────────────────────
SELECT
    Designation,
    ROUND(AVG(Basic_Salary), 2)       AS Avg_Basic,
    ROUND(AVG(Net_Salary_Payable), 2) AS Avg_Net,
    ROUND(AVG(Basic_Salary - Net_Salary_Payable), 2) AS Avg_Gap,
    SUM(LWP_Deduction)                AS Total_Deductions
FROM dbo.Attendance_Monthly
GROUP BY Designation
ORDER BY Total_Deductions DESC;
GO

-- ─── Q7: CL consumption — who used most CL and shifted to LWP ────────
SELECT
    Employee_Name, Designation,
    SUM(CL)  AS Total_CL_Used,
    SUM(LWP) AS Total_LWP_Days,
    CASE
        WHEN SUM(CL) >= 12 AND SUM(LWP) > 5 THEN 'Exhausted CL → moved to LWP'
        WHEN SUM(CL) >= 12 THEN 'Exhausted CL'
        WHEN SUM(LWP) > 10 THEN 'High LWP (skipped CL)'
        ELSE 'Normal'
    END AS Leave_Pattern
FROM dbo.Attendance_Monthly
GROUP BY Employee_Name, Designation
ORDER BY Total_LWP_Days DESC;
GO

-- ─── Q8: Perfect attendance — zero LWP all year ──────────────────────
SELECT
    Employee_Name, Designation, Department,
    SUM(LWP)          AS Total_LWP,
    SUM(Late_Mark)    AS Total_Late,
    SUM(Early_Mark)   AS Total_Early,
    ROUND(AVG(Attendance_Pct), 2) AS Avg_Attendance_Pct
FROM dbo.Attendance_Monthly
GROUP BY Employee_Name, Designation, Department
HAVING SUM(LWP) = 0
ORDER BY Avg_Attendance_Pct DESC;
GO

-- ─── Q9: Deteriorating attendance trend (month over month) ───────────
WITH MonthlyAtt AS (
    SELECT
        Employee_Name, Month_Num, Month, Attendance_Pct,
        LAG(Attendance_Pct) OVER (PARTITION BY Employee_Name ORDER BY Year, Month_Num) AS Prev_Month_Pct
    FROM dbo.Attendance_Monthly
)
SELECT
    Employee_Name, Month, Month_Num,
    Attendance_Pct,
    Prev_Month_Pct,
    ROUND(Attendance_Pct - Prev_Month_Pct, 2) AS MoM_Change,
    CASE
        WHEN Attendance_Pct < Prev_Month_Pct THEN 'Declining'
        WHEN Attendance_Pct > Prev_Month_Pct THEN 'Improving'
        ELSE 'Stable'
    END AS Trend
FROM MonthlyAtt
WHERE Prev_Month_Pct IS NOT NULL
ORDER BY Employee_Name, Month_Num;
GO

-- ─── Q10: Cost-cut analysis — LWP deduction vs designation ───────────
SELECT
    Designation,
    COUNT(DISTINCT Employee_Name)         AS Headcount,
    SUM(Basic_Salary)                     AS Total_Basic_Annual,
    SUM(LWP_Deduction)                    AS Total_LWP_Deductions,
    ROUND(SUM(LWP_Deduction) * 100.0 /
          NULLIF(SUM(Basic_Salary),0), 2) AS Deduction_Pct_Of_Basic,
    ROUND(SUM(Basic_Salary) * 0.10, 2)   AS Target_10Pct_Cut
FROM dbo.Attendance_Monthly
GROUP BY Designation
ORDER BY Deduction_Pct_Of_Basic DESC;
GO

-- ─── Q11: Top 3 highest-paid employees per department ────────────────
WITH RankedSalary AS (
    SELECT
        Employee_Name, Department, Designation,
        Basic_Salary,
        DENSE_RANK() OVER (PARTITION BY Department ORDER BY Basic_Salary DESC) AS Salary_Rank
    FROM dbo.Employee_Master
)
SELECT
    Department, Salary_Rank, Employee_Name, Designation, Basic_Salary
FROM RankedSalary
WHERE Salary_Rank <= 3
ORDER BY Department, Salary_Rank;
GO
"""

with open(OUT + "SDJEMS_HR_Analysis.sql", "w") as f:
    f.write(analysis_sql)
print("✓ Analysis SQL queries saved")

# ─── QUICK SUMMARY PRINT ─────────────────────────────────────────────────
print("\n" + "="*55)
print("PIPELINE COMPLETE — SUMMARY")
print("="*55)
print(f"Total rows:      {len(master)}")
print(f"Total employees: {master['Employee_Name'].nunique()}")
print(f"Months covered:  {master['Month'].nunique()}")
print(f"Columns:         {len(master.columns)}")
print("\nColumn list:")
for c in master.columns:
    print(f"  • {c}")

print("\nSample (first 3 rows):")
print(master[["Employee_Name","Month","Attendance_Pct","Basic_Salary","LWP_Deduction","Net_Salary_Payable","Performance_Score","Performance_Label"]].head(3).to_string())
