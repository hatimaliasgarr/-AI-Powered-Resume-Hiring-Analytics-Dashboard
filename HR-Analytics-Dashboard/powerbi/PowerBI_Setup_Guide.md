# Power BI Implementation Guide — HR Analytics Dashboard

This guide provides the exact specifications to build the **HR Analytics Dashboard** using the processed Excel files.

## ⚙️ Data Model Setup
1. **Import Data**: Load the 4 Excel files from `data/processed/`:
   - `fact_employees.xlsx`
   - `dim_department.xlsx`
   - `dim_jobrole.xlsx`
   - `ai_recommendations.xlsx`
2. **Relationships**:
   - `fact_employees[Department]` ↔ `dim_department[Department]` (Many-to-One)
   - `fact_employees[JobRole]` ↔ `dim_jobrole[JobRole]` (Many-to-One)
   - `ai_recommendations` (Keep as a standalone table)

## 🧮 DAX Measures
Create a table named `_Measures` and add the following:

```dax
Total Employees = COUNTROWS(fact_employees)

Attrition Rate = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_employees), fact_employees[Attrition] = "Yes"),
    COUNTROWS(fact_employees)
) * 100

Avg Monthly Salary = AVERAGE(fact_employees[MonthlyIncome])

Avg Job Satisfaction = AVERAGE(fact_employees[JobSatisfaction])

Avg Skill Score = AVERAGE(fact_employees[SkillScore])

High Risk Employees = 
CALCULATE(COUNTROWS(fact_employees), fact_employees[AttritionRisk] = "High")

Retention Rate = 100 - [Attrition Rate]

Avg Tenure Years = AVERAGE(fact_employees[YearsAtCompany])

Promotion Rate = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_employees), fact_employees[PromotionFlag] = 1),
    COUNTROWS(fact_employees)
) * 100

Salary Gap % = 
VAR MaleSalary = CALCULATE(AVERAGE(fact_employees[MonthlyIncome]), fact_employees[Gender] = "Male")
VAR FemaleSalary = CALCULATE(AVERAGE(fact_employees[MonthlyIncome]), fact_employees[Gender] = "Female")
RETURN DIVIDE(MaleSalary - FemaleSalary, MaleSalary) * 100
```

## 🎨 Theme & Styling
- **Primary Color**: `#0F3460` (Navy)
- **Accent Color**: `#E94560` (Red)
- **Success Color**: `#16C79A` (Teal)
- **Warning Color**: `#F5A623` (Orange)
- **Background**: `#F0F4F8`
- **Font**: Segoe UI

## 📄 Page Layouts
Follow the visual specifications provided in the main project prompt for all 5 pages:
1. **Executive Overview**: Focus on high-level KPIs and trends.
2. **Attrition Deep Dive**: Use Waterfall and Heatmap visuals.
3. **Skill Gap & Performance**: Focus on Radar or Bullet charts.
4. **Salary & Hiring**: Use Box Plots and Funnels.
5. **AI Recommendation Panel**: Use a Table or Multi-row card visual to display the `ai_recommendations` table.

## 🤖 AI Recommendation Card Styling (Page 5)
To create the cards on Page 5, use a **Table visual** with:
- **Insight**: Conditional formatting for font color (dark gray).
- **ActionableAdvice**: Font style set to *Italic*, color `#16C79A`.
- **Priority**: Use background color rules (Red for High, Orange for Medium, Blue for Low).
