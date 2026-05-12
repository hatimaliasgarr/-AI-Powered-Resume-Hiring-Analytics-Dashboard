# HR Analytics Professional Power BI Dashboard

Generated artifact: `HR_Analytics_Professional_Dashboard.pbix`

## Pages
1. Executive Overview
2. Attrition Deep Dive
3. Performance Intelligence
4. Compensation and Hiring
5. AI Recommendations

## Model
- `fact_employees` connects to `dim_department` on `Department`.
- `fact_employees` connects to `dim_jobrole` on `JobRole`.
- `fact_employees` connects to `dim_calendar` on `HireDate` -> `Date`.
- `ai_recommendations` remains standalone for the recommendation panel.

## Interaction Design
- Every analytical page includes Department, Job Role, and Attrition Risk slicers.
- The AI page includes Priority and Impact Area slicers.
- Visuals are native Power BI visuals, so cross-highlighting and page-level filtering work after opening in Power BI Desktop.

## Rebuild
Run this script from the project root:

```powershell
py -3.13 scripts/build_powerbi_dashboard.py
```

If `pbix-mcp` is missing:

```powershell
py -3.13 -m pip install pandas numpy openpyxl pbix-mcp
```
