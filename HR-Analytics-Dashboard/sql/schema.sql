-- 🧠 HR Analytics Dashboard - SQL Schema (PostgreSQL / Supabase compatible)

-- 1. Dimension: Department
CREATE TABLE IF NOT EXISTS dim_department (
    "DepartmentID" SERIAL PRIMARY KEY,
    "Department" TEXT UNIQUE NOT NULL,
    "AvgSalary" REAL,
    "HeadCount" INTEGER
);

-- 2. Dimension: Job Role
CREATE TABLE IF NOT EXISTS dim_jobrole (
    "JobRole" TEXT PRIMARY KEY,
    "AvgSkillScore" REAL,
    "AvgSatisfaction" REAL
);

-- 3. AI Recommendations
CREATE TABLE IF NOT EXISTS ai_recommendations (
    "RecommendationID" INTEGER PRIMARY KEY,
    "Title" TEXT NOT NULL,
    "Insight" TEXT,
    "ActionableAdvice" TEXT,
    "Priority" TEXT,
    "ImpactArea" TEXT
);

-- 4. Fact Table: Employees
CREATE TABLE IF NOT EXISTS fact_employees (
    "EmployeeNumber" INTEGER PRIMARY KEY,
    "Age" INTEGER,
    "Attrition" TEXT,
    "AttritionBinary" INTEGER,
    "BusinessTravel" TEXT,
    "DailyRate" INTEGER,
    "Department" TEXT,
    "DistanceFromHome" INTEGER,
    "Education" INTEGER,
    "EducationField" TEXT,
    "EmployeeCount" INTEGER,
    "EnvironmentSatisfaction" INTEGER,
    "Gender" TEXT,
    "HourlyRate" INTEGER,
    "JobInvolvement" INTEGER,
    "JobLevel" INTEGER,
    "JobRole" TEXT,
    "JobSatisfaction" INTEGER,
    "MaritalStatus" TEXT,
    "MonthlyIncome" INTEGER,
    "MonthlyRate" INTEGER,
    "NumCompaniesWorked" INTEGER,
    "Over18" TEXT,
    "OverTime" TEXT,
    "PercentSalaryHike" INTEGER,
    "PerformanceRating" INTEGER,
    "RelationshipSatisfaction" INTEGER,
    "StandardHours" INTEGER,
    "StockOptionLevel" INTEGER,
    "TotalWorkingYears" INTEGER,
    "TrainingTimesLastYear" INTEGER,
    "WorkLifeBalance" INTEGER,
    "YearsAtCompany" INTEGER,
    "YearsInCurrentRole" INTEGER,
    "YearsSinceLastPromotion" INTEGER,
    "YearsWithCurrManager" INTEGER,
    "HireDate" TEXT,
    "SkillScore" REAL,
    "AttritionRisk" TEXT,
    "HiringChannel" TEXT,
    "SalaryBand" TEXT,
    "TenureGroup" TEXT,
    "YearHired" INTEGER,
    "PromotionFlag" INTEGER,
    
    -- Foreign Key Constraints
    FOREIGN KEY ("Department") REFERENCES dim_department("Department"),
    FOREIGN KEY ("JobRole") REFERENCES dim_jobrole("JobRole")
);

-- Create Indexes for Performance
CREATE INDEX IF NOT EXISTS idx_attrition ON fact_employees("AttritionBinary");
CREATE INDEX IF NOT EXISTS idx_dept ON fact_employees("Department");
CREATE INDEX IF NOT EXISTS idx_jobrole ON fact_employees("JobRole");
