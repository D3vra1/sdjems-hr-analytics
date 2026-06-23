
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
