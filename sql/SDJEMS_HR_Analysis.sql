-- ============================================================
--  SDJEMS HR Analytics — 11 Business Questions
--  Run in SSMS after setup script completes
-- ============================================================

USE SDJEMS_HR_DB;
GO

-- ─── Q1: Employees with attendance below 75% in any month ────────────
SELECT
    Employee_Name, Designation, Department, Month, Year,
    Attendance_Pct,
    CASE WHEN Attendance_Pct < 75 THEN 'AT RISK' ELSE 'OK' END AS Status
FROM dbo.SDJEMS_HR_Master
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
FROM dbo.SDJEMS_HR_Master
GROUP BY Employee_Name, Designation, Department
ORDER BY Total_Punctuality_Issues DESC;
GO

-- ─── Q3: Worst attendance month school-wide ──────────────────────────
SELECT
    Month, Year,
    ROUND(AVG(Attendance_Pct), 2) AS Avg_Attendance_Pct,
    COUNT(DISTINCT Employee_Name)  AS Staff_Count
FROM dbo.SDJEMS_HR_Master
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
FROM dbo.SDJEMS_HR_Master
GROUP BY Month, Year, Month_Num
ORDER BY Year, Month_Num;
GO

-- ─── Q5: Top 5 costliest absentees (annual LWP deduction) ────────────
SELECT TOP 5
    Employee_Name, Designation, Department,
    SUM(LWP)            AS Total_LWP_Days,
    SUM(LWP_Deduction)  AS Annual_LWP_Deduction,
    Basic_Salary        AS Monthly_Basic
FROM dbo.SDJEMS_HR_Master
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
FROM dbo.SDJEMS_HR_Master
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
FROM dbo.SDJEMS_HR_Master
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
FROM dbo.SDJEMS_HR_Master
GROUP BY Employee_Name, Designation, Department
HAVING SUM(LWP) = 0
ORDER BY Avg_Attendance_Pct DESC;
GO

-- ─── Q9: Deteriorating attendance trend (month over month) ───────────
WITH MonthlyAtt AS (
    SELECT
        Employee_Name, Month_Num, Month, Attendance_Pct,
        LAG(Attendance_Pct) OVER (PARTITION BY Employee_Name ORDER BY Year, Month_Num) AS Prev_Month_Pct
    FROM dbo.SDJEMS_HR_Master
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
FROM dbo.SDJEMS_HR_Master
GROUP BY Designation
ORDER BY Deduction_Pct_Of_Basic DESC;
GO

-- ─── Q11: Top 3 highest-paid employees per department ────────────────
WITH UniqueEmployeeSalary AS (
    SELECT 
        Employee_Name, 
        Department, 
        Designation,
        MAX(Basic_Salary) AS Basic_Salary -- Gets the base salary rate without monthly duplication
    FROM dbo.SDJEMS_HR_Master
    GROUP BY Employee_Name, Department, Designation
),
RankedSalary AS (
    SELECT 
        Employee_Name, 
        Department, 
        Designation,
        Basic_Salary,
        DENSE_RANK() OVER (PARTITION BY Department ORDER BY Basic_Salary DESC) AS Salary_Rank
    FROM UniqueEmployeeSalary
)
SELECT 
    Department, 
    Salary_Rank, 
    Employee_Name, 
    Designation, 
    Basic_Salary
FROM RankedSalary
WHERE Salary_Rank <= 3
ORDER BY Department, Salary_Rank;
GO

-- Using Top 3 with ties to include all employees with the same salary as the 3rd highest in each department

SELECT TOP 3 WITH TIES
    Department,
    ROW_NUMBER() OVER (PARTITION BY Department ORDER BY MAX(Basic_Salary) DESC) AS Salary_Rank,
    Employee_Name,
    Designation,
    MAX(Basic_Salary) AS Basic_Salary
FROM dbo.SDJEMS_HR_Master
GROUP BY Department, Employee_Name, Designation
ORDER BY ROW_NUMBER() OVER (PARTITION BY Department ORDER BY MAX(Basic_Salary) DESC);
