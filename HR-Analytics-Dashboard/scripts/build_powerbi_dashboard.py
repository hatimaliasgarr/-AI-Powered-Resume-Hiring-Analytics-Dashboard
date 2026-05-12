"""Build a professional Power BI dashboard for the HR analytics project.

This script reads the processed Excel tables, creates a reusable calendar
dimension, writes Power BI support assets, and generates a multi-page PBIX
dashboard with slicers, KPI cards, charts, tables, relationships, and DAX
measures.

Dependency note:
    pip install pandas numpy openpyxl pbix-mcp

Python 3.13 is recommended for pbix-mcp on Windows because its xpress9
dependency currently ships a Windows wheel for Python 3.13.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parents[1]
PROCESSED_DIR = PROJECT_DIR / "data" / "processed"
POWERBI_DIR = PROJECT_DIR / "powerbi"

PBIX_NAME = "HR_Analytics_Professional_Dashboard.pbix"
THEME_NAME = "HR_Analytics_Professional_Theme.json"
DAX_NAME = "HR_Analytics_DAX_Measures.dax"
GUIDE_NAME = "HR_Analytics_Professional_Build_Guide.md"
PBIP_DIR_NAME = "HR_Analytics_Professional_Dashboard_PBIP"


PALETTE = {
    "navy": "#102A43",
    "blue": "#2F80ED",
    "teal": "#00A6A6",
    "coral": "#F95738",
    "amber": "#F2C94C",
    "green": "#2F855A",
    "slate": "#546A7B",
    "ink": "#101828",
    "muted": "#667085",
    "line": "#E3E8EF",
    "paper": "#FFFFFF",
    "canvas": "#F6F8FB",
}


THEME = {
    "name": "HR Analytics Professional",
    "dataColors": [
        PALETTE["blue"],
        PALETTE["teal"],
        PALETTE["coral"],
        PALETTE["amber"],
        PALETTE["green"],
        "#7C3AED",
        "#06AED5",
        "#B83280",
        "#6B7280",
    ],
    "background": PALETTE["canvas"],
    "foreground": PALETTE["ink"],
    "tableAccent": PALETTE["teal"],
    "good": PALETTE["green"],
    "neutral": PALETTE["amber"],
    "bad": PALETTE["coral"],
    "textClasses": {
        "title": {
            "fontFace": "Segoe UI Semibold",
            "fontSize": 16,
            "color": PALETTE["navy"],
        },
        "label": {
            "fontFace": "Segoe UI",
            "fontSize": 10,
            "color": PALETTE["muted"],
        },
        "callout": {
            "fontFace": "Segoe UI Semibold",
            "fontSize": 26,
            "color": PALETTE["ink"],
        },
    },
    "visualStyles": {
        "*": {
            "*": {
                "visualHeader": [{"show": True}],
                "border": [{"show": True, "color": {"solid": {"color": PALETTE["line"]}}}],
                "background": [
                    {
                        "show": True,
                        "color": {"solid": {"color": PALETTE["paper"]}},
                        "transparency": 0,
                    }
                ],
            }
        }
    },
}


MEASURES: list[dict[str, str]] = [
    {
        "table": "fact_employees",
        "name": "Total Employees",
        "expression": "COUNTROWS('fact_employees')",
        "description": "Current employee records in filter context.",
    },
    {
        "table": "fact_employees",
        "name": "Active Employees",
        "expression": "CALCULATE([Total Employees], 'fact_employees'[Attrition] = \"No\")",
        "description": "Employees retained in the company.",
    },
    {
        "table": "fact_employees",
        "name": "Attrition Employees",
        "expression": "CALCULATE([Total Employees], 'fact_employees'[Attrition] = \"Yes\")",
        "description": "Employees marked as attrited.",
    },
    {
        "table": "fact_employees",
        "name": "Attrition Rate",
        "expression": "DIVIDE([Attrition Employees], [Total Employees]) * 100",
        "description": "Percentage of employees who attrited.",
    },
    {
        "table": "fact_employees",
        "name": "Retention Rate",
        "expression": "100 - [Attrition Rate]",
        "description": "Percentage of employees retained.",
    },
    {
        "table": "fact_employees",
        "name": "High Risk Employees",
        "expression": "CALCULATE([Total Employees], 'fact_employees'[AttritionRisk] = \"High\")",
        "description": "Employees classified as high attrition risk.",
    },
    {
        "table": "fact_employees",
        "name": "High Risk Rate",
        "expression": "DIVIDE([High Risk Employees], [Total Employees]) * 100",
        "description": "Share of employees classified as high risk.",
    },
    {
        "table": "fact_employees",
        "name": "Avg Monthly Salary",
        "expression": "AVERAGE('fact_employees'[MonthlyIncome])",
        "description": "Average monthly income.",
    },
    {
        "table": "fact_employees",
        "name": "Avg Job Satisfaction",
        "expression": "AVERAGE('fact_employees'[JobSatisfaction])",
        "description": "Average job satisfaction score.",
    },
    {
        "table": "fact_employees",
        "name": "Avg Skill Score",
        "expression": "AVERAGE('fact_employees'[SkillScore])",
        "description": "Average engineered skill score.",
    },
    {
        "table": "fact_employees",
        "name": "Avg Tenure Years",
        "expression": "AVERAGE('fact_employees'[YearsAtCompany])",
        "description": "Average employee tenure in years.",
    },
    {
        "table": "fact_employees",
        "name": "Promotion Rate",
        "expression": (
            "DIVIDE(CALCULATE([Total Employees], "
            "'fact_employees'[PromotionFlag] = 1), [Total Employees]) * 100"
        ),
        "description": "Share of employees promoted in the latest period flag.",
    },
    {
        "table": "fact_employees",
        "name": "Overtime Attrition Rate",
        "expression": (
            "DIVIDE("
            "CALCULATE([Attrition Employees], 'fact_employees'[OverTime] = \"Yes\"), "
            "CALCULATE([Total Employees], 'fact_employees'[OverTime] = \"Yes\")"
            ") * 100"
        ),
        "description": "Attrition rate among employees working overtime.",
    },
    {
        "table": "fact_employees",
        "name": "Salary Gap",
        "expression": (
            "VAR MaleSalary = CALCULATE(AVERAGE('fact_employees'[MonthlyIncome]), "
            "'fact_employees'[Gender] = \"Male\")\n"
            "VAR FemaleSalary = CALCULATE(AVERAGE('fact_employees'[MonthlyIncome]), "
            "'fact_employees'[Gender] = \"Female\")\n"
            "RETURN DIVIDE(MaleSalary - FemaleSalary, MaleSalary) * 100"
        ),
        "description": "Male minus female average salary as a percent of male salary.",
    },
    {
        "table": "fact_employees",
        "name": "Avg Training",
        "expression": "AVERAGE('fact_employees'[TrainingTimesLastYear])",
        "description": "Average trainings attended last year.",
    },
    {
        "table": "fact_employees",
        "name": "Avg Performance",
        "expression": "AVERAGE('fact_employees'[PerformanceRating])",
        "description": "Average performance rating.",
    },
    {
        "table": "ai_recommendations",
        "name": "Total Recommendations",
        "expression": "COUNTROWS('ai_recommendations')",
        "description": "Total AI recommendation records.",
    },
    {
        "table": "ai_recommendations",
        "name": "High Priority Recommendations",
        "expression": (
            "CALCULATE([Total Recommendations], "
            "'ai_recommendations'[Priority] = \"High\")"
        ),
        "description": "AI recommendations marked high priority.",
    },
]


def read_inputs() -> dict[str, pd.DataFrame]:
    fact = pd.read_excel(PROCESSED_DIR / "fact_employees.xlsx")
    fact["HireDate"] = pd.to_datetime(fact["HireDate"])

    calendar = pd.DataFrame(
        {"Date": pd.date_range(fact["HireDate"].min(), fact["HireDate"].max(), freq="D")}
    )
    calendar["Year"] = calendar["Date"].dt.year
    calendar["Quarter"] = "Q" + calendar["Date"].dt.quarter.astype(str)
    calendar["MonthNumber"] = calendar["Date"].dt.month
    calendar["MonthName"] = calendar["Date"].dt.strftime("%b")
    calendar["YearMonth"] = calendar["Date"].dt.strftime("%Y-%m")
    calendar_path = PROCESSED_DIR / "dim_calendar.xlsx"
    calendar.to_excel(calendar_path, index=False)

    return {
        "fact_employees": fact,
        "dim_department": pd.read_excel(PROCESSED_DIR / "dim_department.xlsx"),
        "dim_jobrole": pd.read_excel(PROCESSED_DIR / "dim_jobrole.xlsx"),
        "ai_recommendations": pd.read_excel(PROCESSED_DIR / "ai_recommendations.xlsx"),
        "dim_calendar": calendar,
    }


def infer_columns(df: pd.DataFrame, overrides: dict[str, str] | None = None) -> list[dict[str, str]]:
    overrides = overrides or {}
    columns = []
    for col in df.columns:
        if col in overrides:
            data_type = overrides[col]
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            data_type = "DateTime"
        elif pd.api.types.is_integer_dtype(df[col]):
            data_type = "Int64"
        elif pd.api.types.is_float_dtype(df[col]):
            data_type = "Double"
        elif pd.api.types.is_bool_dtype(df[col]):
            data_type = "Boolean"
        else:
            data_type = "String"
        columns.append({"name": str(col), "data_type": data_type})
    return columns


def rows_from_frame(df: pd.DataFrame) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for raw in df.to_dict("records"):
        row: dict[str, Any] = {}
        for key, value in raw.items():
            if value is None or pd.isna(value):
                row[key] = None
            elif isinstance(value, pd.Timestamp):
                row[key] = value.to_pydatetime()
            elif isinstance(value, np.integer):
                row[key] = int(value)
            elif isinstance(value, np.floating):
                row[key] = float(value)
            elif isinstance(value, np.bool_):
                row[key] = bool(value)
            else:
                row[key] = value
        rows.append(row)
    return rows


def field(table: str, column: str) -> dict[str, str]:
    return {"table": table, "column": column}


def measure(name: str) -> dict[str, str]:
    return {"measure": name}


def card(name: str, x: int, y: int, measure_name: str) -> dict[str, Any]:
    return {
        "name": name,
        "type": "card",
        "x": x,
        "y": y,
        "width": 182,
        "height": 92,
        "config": {"measure": measure_name},
    }


def slicer(name: str, x: int, y: int, column: dict[str, str], width: int = 150) -> dict[str, Any]:
    return {
        "name": name,
        "type": "slicer",
        "x": x,
        "y": y,
        "width": width,
        "height": 58,
        "config": {"column": column},
    }


def chart(
    name: str,
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    category: dict[str, str],
    measure_name: str,
) -> dict[str, Any]:
    return {
        "name": name,
        "type": visual_type,
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "config": {"category": category, "measure": measure_name},
    }


def table_visual(
    name: str,
    x: int,
    y: int,
    width: int,
    height: int,
    columns: list[dict[str, str]],
) -> dict[str, Any]:
    return {
        "name": name,
        "type": "tableEx",
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "config": {"columns": columns},
    }


def build_pages() -> list[dict[str, Any]]:
    common_slicers = [
        slicer("slc_department", 742, 18, field("fact_employees", "Department")),
        slicer("slc_job_role", 904, 18, field("fact_employees", "JobRole"), 170),
        slicer("slc_risk", 1086, 18, field("fact_employees", "AttritionRisk"), 148),
    ]

    return [
        {
            "name": "Executive Overview",
            "visuals": [
                card("kpi_total", 24, 92, "Total Employees"),
                card("kpi_attrition", 222, 92, "Attrition Rate"),
                card("kpi_retention", 420, 92, "Retention Rate"),
                card("kpi_salary", 618, 92, "Avg Monthly Salary"),
                card("kpi_high_risk", 816, 92, "High Risk Employees"),
                card("kpi_tenure", 1014, 92, "Avg Tenure Years"),
                chart(
                    "trend_hires",
                    "lineChart",
                    24,
                    210,
                    520,
                    220,
                    field("dim_calendar", "Year"),
                    "Total Employees",
                ),
                chart(
                    "donut_dept",
                    "donutChart",
                    564,
                    210,
                    310,
                    220,
                    field("fact_employees", "Department"),
                    "Total Employees",
                ),
                chart(
                    "attrition_dept",
                    "clusteredColumnChart",
                    894,
                    210,
                    338,
                    220,
                    field("fact_employees", "Department"),
                    "Attrition Rate",
                ),
                chart(
                    "headcount_role",
                    "clusteredBarChart",
                    24,
                    456,
                    824,
                    222,
                    field("fact_employees", "JobRole"),
                    "Total Employees",
                ),
                table_visual(
                    "dept_scorecard",
                    872,
                    456,
                    360,
                    222,
                    [
                        field("fact_employees", "Department"),
                        measure("Total Employees"),
                        measure("Attrition Rate"),
                        measure("High Risk Employees"),
                        measure("Avg Monthly Salary"),
                    ],
                ),
                *common_slicers,
            ],
        },
        {
            "name": "Attrition Deep Dive",
            "visuals": [
                card("kpi_attrition_2", 24, 92, "Attrition Rate"),
                card("kpi_attrition_count", 222, 92, "Attrition Employees"),
                card("kpi_high_risk_rate", 420, 92, "High Risk Rate"),
                card("kpi_overtime_attrition", 618, 92, "Overtime Attrition Rate"),
                card("kpi_satisfaction", 816, 92, "Avg Job Satisfaction"),
                chart(
                    "attrition_role",
                    "clusteredBarChart",
                    24,
                    210,
                    560,
                    250,
                    field("fact_employees", "JobRole"),
                    "Attrition Rate",
                ),
                chart(
                    "risk_department",
                    "clusteredColumnChart",
                    608,
                    210,
                    300,
                    250,
                    field("fact_employees", "Department"),
                    "High Risk Employees",
                ),
                chart(
                    "attrition_overtime",
                    "clusteredColumnChart",
                    932,
                    210,
                    300,
                    250,
                    field("fact_employees", "OverTime"),
                    "Attrition Rate",
                ),
                chart(
                    "attrition_tenure",
                    "clusteredColumnChart",
                    24,
                    488,
                    380,
                    190,
                    field("fact_employees", "TenureGroup"),
                    "Attrition Rate",
                ),
                chart(
                    "attrition_year",
                    "lineChart",
                    428,
                    488,
                    380,
                    190,
                    field("dim_calendar", "Year"),
                    "Attrition Rate",
                ),
                table_visual(
                    "risk_role_table",
                    832,
                    488,
                    400,
                    190,
                    [
                        field("fact_employees", "JobRole"),
                        measure("Total Employees"),
                        measure("Attrition Employees"),
                        measure("Attrition Rate"),
                        measure("High Risk Employees"),
                    ],
                ),
                *common_slicers,
            ],
        },
        {
            "name": "Performance Intelligence",
            "visuals": [
                card("kpi_skill", 24, 92, "Avg Skill Score"),
                card("kpi_performance", 222, 92, "Avg Performance"),
                card("kpi_training", 420, 92, "Avg Training"),
                card("kpi_satisfaction_3", 618, 92, "Avg Job Satisfaction"),
                card("kpi_promotion", 816, 92, "Promotion Rate"),
                chart(
                    "skill_role",
                    "clusteredBarChart",
                    24,
                    210,
                    540,
                    250,
                    field("fact_employees", "JobRole"),
                    "Avg Skill Score",
                ),
                chart(
                    "performance_dept",
                    "clusteredColumnChart",
                    588,
                    210,
                    312,
                    250,
                    field("fact_employees", "Department"),
                    "Avg Performance",
                ),
                chart(
                    "training_role",
                    "clusteredBarChart",
                    924,
                    210,
                    308,
                    250,
                    field("fact_employees", "JobRole"),
                    "Avg Training",
                ),
                chart(
                    "satisfaction_role",
                    "clusteredColumnChart",
                    24,
                    488,
                    460,
                    190,
                    field("fact_employees", "JobRole"),
                    "Avg Job Satisfaction",
                ),
                table_visual(
                    "role_perf_table",
                    508,
                    488,
                    724,
                    190,
                    [
                        field("fact_employees", "JobRole"),
                        measure("Total Employees"),
                        measure("Avg Skill Score"),
                        measure("Avg Performance"),
                        measure("Avg Training"),
                        measure("Avg Job Satisfaction"),
                    ],
                ),
                *common_slicers,
            ],
        },
        {
            "name": "Compensation and Hiring",
            "visuals": [
                card("kpi_salary_4", 24, 92, "Avg Monthly Salary"),
                card("kpi_salary_gap", 222, 92, "Salary Gap"),
                card("kpi_promotion_4", 420, 92, "Promotion Rate"),
                card("kpi_skill_4", 618, 92, "Avg Skill Score"),
                card("kpi_tenure_4", 816, 92, "Avg Tenure Years"),
                chart(
                    "salary_role",
                    "clusteredBarChart",
                    24,
                    210,
                    530,
                    250,
                    field("fact_employees", "JobRole"),
                    "Avg Monthly Salary",
                ),
                chart(
                    "salary_band",
                    "donutChart",
                    578,
                    210,
                    300,
                    250,
                    field("fact_employees", "SalaryBand"),
                    "Total Employees",
                ),
                chart(
                    "hiring_channel_skill",
                    "clusteredColumnChart",
                    902,
                    210,
                    330,
                    250,
                    field("fact_employees", "HiringChannel"),
                    "Avg Skill Score",
                ),
                chart(
                    "promotion_dept",
                    "clusteredColumnChart",
                    24,
                    488,
                    420,
                    190,
                    field("fact_employees", "Department"),
                    "Promotion Rate",
                ),
                table_visual(
                    "comp_role_table",
                    468,
                    488,
                    764,
                    190,
                    [
                        field("fact_employees", "JobRole"),
                        measure("Avg Monthly Salary"),
                        measure("Salary Gap"),
                        measure("Promotion Rate"),
                        measure("Avg Tenure Years"),
                        measure("Total Employees"),
                    ],
                ),
                *common_slicers,
            ],
        },
        {
            "name": "AI Recommendations",
            "visuals": [
                card("kpi_recommendations", 24, 92, "Total Recommendations"),
                card("kpi_high_priority", 222, 92, "High Priority Recommendations"),
                chart(
                    "recommendations_impact",
                    "clusteredColumnChart",
                    420,
                    92,
                    360,
                    170,
                    field("ai_recommendations", "ImpactArea"),
                    "Total Recommendations",
                ),
                chart(
                    "recommendations_priority",
                    "donutChart",
                    804,
                    92,
                    428,
                    170,
                    field("ai_recommendations", "Priority"),
                    "Total Recommendations",
                ),
                table_visual(
                    "ai_recommendation_table",
                    24,
                    292,
                    1208,
                    386,
                    [
                        field("ai_recommendations", "RecommendationID"),
                        field("ai_recommendations", "Priority"),
                        field("ai_recommendations", "ImpactArea"),
                        field("ai_recommendations", "Title"),
                        field("ai_recommendations", "Insight"),
                        field("ai_recommendations", "ActionableAdvice"),
                    ],
                ),
                slicer("slc_priority", 742, 18, field("ai_recommendations", "Priority")),
                slicer("slc_impact", 904, 18, field("ai_recommendations", "ImpactArea"), 170),
            ],
        },
    ]


def write_support_assets() -> None:
    POWERBI_DIR.mkdir(parents=True, exist_ok=True)
    (POWERBI_DIR / THEME_NAME).write_text(json.dumps(THEME, indent=2), encoding="utf-8")

    dax_lines = []
    for item in MEASURES:
        dax_lines.append(f"-- {item['description']}")
        dax_lines.append(f"{item['name']} =")
        dax_lines.append(item["expression"])
        dax_lines.append("")
    (POWERBI_DIR / DAX_NAME).write_text("\n".join(dax_lines), encoding="utf-8")

    guide = f"""# HR Analytics Professional Power BI Dashboard

Generated artifact: `{PBIX_NAME}`

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
"""
    (POWERBI_DIR / GUIDE_NAME).write_text(guide, encoding="utf-8")


def add_tables(builder: Any, frames: dict[str, pd.DataFrame]) -> None:
    source_paths = {
        "fact_employees": PROCESSED_DIR / "fact_employees.xlsx",
        "dim_department": PROCESSED_DIR / "dim_department.xlsx",
        "dim_jobrole": PROCESSED_DIR / "dim_jobrole.xlsx",
        "ai_recommendations": PROCESSED_DIR / "ai_recommendations.xlsx",
        "dim_calendar": PROCESSED_DIR / "dim_calendar.xlsx",
    }

    overrides = {
        "fact_employees": {"HireDate": "DateTime", "SkillScore": "Double"},
        "dim_department": {"AvgSalary": "Double"},
        "dim_jobrole": {"AvgSkillScore": "Double", "AvgSatisfaction": "Double"},
        "dim_calendar": {"Date": "DateTime"},
    }

    for table_name, df in frames.items():
        builder.add_table(
            table_name,
            infer_columns(df, overrides.get(table_name)),
            rows=rows_from_frame(df),
            source_db={
                "type": "excel",
                "path": str(source_paths[table_name]),
                "sheet": "Sheet1",
            },
        )


def add_model(builder: Any) -> None:
    for item in MEASURES:
        builder.add_measure(
            item["table"],
            item["name"],
            item["expression"],
            item["description"],
        )

    builder.add_relationship("fact_employees", "Department", "dim_department", "Department")
    builder.add_relationship("fact_employees", "JobRole", "dim_jobrole", "JobRole")
    builder.add_relationship("fact_employees", "HireDate", "dim_calendar", "Date")

    builder.add_user_hierarchy(
        "fact_employees",
        "Workforce Drilldown",
        [
            {"name": "Department", "column": "Department"},
            {"name": "Job Role", "column": "JobRole"},
            {"name": "Salary Band", "column": "SalaryBand"},
        ],
    )
    builder.add_user_hierarchy(
        "dim_calendar",
        "Calendar",
        [
            {"name": "Year", "column": "Year"},
            {"name": "Quarter", "column": "Quarter"},
            {"name": "Month", "column": "MonthName"},
        ],
    )


def base_visual_format(title: str) -> dict[str, Any]:
    return {
        "title": {
            "show": True,
            "text": title,
            "fontSize": 12,
            "color": PALETTE["navy"],
            "fontFamily": "Segoe UI Semibold",
        },
        "background": {"show": True, "color": PALETTE["paper"], "transparency": 0},
        "border": {"show": True, "color": PALETTE["line"], "radius": 7, "width": 1},
        "dropShadow": {
            "show": True,
            "color": "#D7DEE8",
            "angle": 45,
            "blur": 8,
            "distance": 2,
            "transparency": 55,
        },
        "visualHeader": {"show": True, "showFocusModeButton": True, "showOptionsMenu": True},
        "visualTooltip": {
            "show": True,
            "fontSize": 10,
            "titleFontColor": PALETTE["paper"],
            "valueFontColor": PALETTE["paper"],
            "background": PALETTE["navy"],
        },
    }


def chart_format(title: str, value_title: str = "") -> dict[str, Any]:
    fmt = base_visual_format(title)
    fmt.update(
        {
            "dataLabels": {"show": True, "fontSize": 9, "color": PALETTE["ink"]},
            "categoryAxis": {
                "show": True,
                "fontSize": 9,
                "color": PALETTE["muted"],
                "gridlineShow": False,
            },
            "valueAxis": {
                "show": True,
                "fontSize": 9,
                "color": PALETTE["muted"],
                "displayUnits": "auto",
                "gridlineShow": True,
            },
            "legend": {"show": True, "position": "right", "fontSize": 9, "color": PALETTE["muted"]},
            "dataColors": THEME["dataColors"],
        }
    )
    if value_title:
        fmt["valueAxis"]["title"] = value_title
    return fmt


def card_format(title: str, accent: str = PALETTE["blue"]) -> dict[str, Any]:
    fmt = base_visual_format(title)
    fmt.update(
        {
            "title": {
                "show": True,
                "text": title,
                "fontSize": 10,
                "color": PALETTE["muted"],
                "fontFamily": "Segoe UI Semibold",
            },
            "dataColors": [accent],
            "categoryLabels": {
                "show": True,
                "fontSize": 9,
                "color": PALETTE["muted"],
                "fontFamily": "Segoe UI",
            },
            "padding": {"top": 8, "bottom": 8, "left": 10, "right": 10},
        }
    )
    return fmt


def table_format(title: str) -> dict[str, Any]:
    fmt = base_visual_format(title)
    fmt.update(
        {
            "grid": {
                "gridVertical": False,
                "gridHorizontal": True,
                "gridHorizontalColor": PALETTE["line"],
                "rowPadding": 7,
                "textSize": 9,
            },
            "columnHeaders": {
                "bold": True,
                "fontSize": 9,
                "fontFamily": "Segoe UI Semibold",
                "fontColor": PALETTE["paper"],
                "backColor": PALETTE["navy"],
                "wordWrap": True,
            },
            "values": {
                "fontSize": 9,
                "fontFamily": "Segoe UI",
                "fontColorPrimary": PALETTE["ink"],
                "fontColorSecondary": PALETTE["ink"],
                "backColorPrimary": PALETTE["paper"],
                "backColorSecondary": "#F2F6FA",
                "wordWrap": True,
            },
            "total": {
                "show": True,
                "bold": True,
                "fontSize": 9,
                "fontColor": PALETTE["navy"],
                "backColor": "#EAF4F4",
            },
        }
    )
    return fmt


def slicer_format(title: str) -> dict[str, Any]:
    fmt = base_visual_format(title)
    fmt.update(
        {
            "title": {
                "show": True,
                "text": title,
                "fontSize": 9,
                "color": PALETTE["navy"],
                "fontFamily": "Segoe UI Semibold",
            },
            "dropShadow": {"show": False},
            "padding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
        }
    )
    return fmt


PAGE_TITLES = [
    (
        "Executive HR Overview",
        "Board-level workforce health, retention pressure, and headcount mix.",
    ),
    (
        "Attrition Deep Dive",
        "Risk segments, turnover drivers, and where retention work should focus first.",
    ),
    (
        "Performance Intelligence",
        "Skill, training, performance, and satisfaction signals by role and department.",
    ),
    (
        "Compensation and Hiring",
        "Pay equity, salary bands, promotion momentum, and sourcing quality.",
    ),
    (
        "AI-Powered HR Recommendations",
        "Prioritized, data-backed actions for retention, compensation, L&D, and hiring.",
    ),
]


VISUAL_TITLES = {
    0: {
        0: "Total Employees",
        1: "Attrition Rate %",
        2: "Retention Rate %",
        3: "Avg Monthly Salary",
        4: "High Risk Employees",
        5: "Avg Tenure",
        6: "Hiring Trend by Year",
        7: "Workforce by Department",
        8: "Attrition Rate by Department",
        9: "Headcount by Job Role",
        10: "Department Scorecard",
        11: "Department",
        12: "Job Role",
        13: "Risk",
    },
    1: {
        0: "Attrition Rate %",
        1: "Attrition Count",
        2: "High Risk Rate %",
        3: "Overtime Attrition %",
        4: "Avg Job Satisfaction",
        5: "Attrition Rate by Job Role",
        6: "High-Risk Employees by Department",
        7: "Attrition Rate by Overtime",
        8: "Attrition Rate by Tenure",
        9: "Attrition Trend by Hire Year",
        10: "Role-Level Risk Table",
        11: "Department",
        12: "Job Role",
        13: "Risk",
    },
    2: {
        0: "Avg Skill Score",
        1: "Avg Performance",
        2: "Avg Training",
        3: "Avg Job Satisfaction",
        4: "Promotion Rate %",
        5: "Skill Score by Job Role",
        6: "Performance by Department",
        7: "Training by Job Role",
        8: "Satisfaction by Job Role",
        9: "Role Performance Table",
        10: "Department",
        11: "Job Role",
        12: "Risk",
    },
    3: {
        0: "Avg Monthly Salary",
        1: "Salary Gap %",
        2: "Promotion Rate %",
        3: "Avg Skill Score",
        4: "Avg Tenure",
        5: "Average Salary by Job Role",
        6: "Headcount by Salary Band",
        7: "Hiring Channel Skill Quality",
        8: "Promotion Rate by Department",
        9: "Compensation by Role",
        10: "Department",
        11: "Job Role",
        12: "Risk",
    },
    4: {
        0: "Total Recommendations",
        1: "High Priority",
        2: "Recommendations by Impact Area",
        3: "Recommendations by Priority",
        4: "AI Recommendation Action Register",
        5: "Priority",
        6: "Impact Area",
    },
}


def textbox_config(title: str, subtitle: str) -> str:
    config = {
        "singleVisual": {
            "objects": {
                "general": [
                    {
                        "properties": {
                            "paragraphs": [
                                {
                                    "textRuns": [
                                        {
                                            "value": title,
                                            "textStyle": {
                                                "fontFamily": "Segoe UI Semibold",
                                                "fontSize": "24pt",
                                                "color": PALETTE["navy"],
                                            },
                                        }
                                    ]
                                },
                                {
                                    "textRuns": [
                                        {
                                            "value": subtitle,
                                            "textStyle": {
                                                "fontFamily": "Segoe UI",
                                                "fontSize": "10pt",
                                                "color": PALETTE["muted"],
                                            },
                                        }
                                    ]
                                },
                            ]
                        }
                    }
                ]
            }
        }
    }
    return json.dumps(config)


def polish_report(pbix_path: Path) -> None:
    try:
        import pbix_mcp.server as pbi
    except Exception as exc:
        print(f"PBIX generated, but styling/export polish was skipped: {exc}")
        return

    alias = "hr_analytics_pro"
    print(pbi.pbix_open(str(pbix_path), alias))
    print(pbi.pbix_set_theme(alias, json.dumps(THEME), THEME_NAME))

    card_accents = [PALETTE["blue"], PALETTE["coral"], PALETTE["teal"], PALETTE["green"], PALETTE["amber"]]
    for page_index, titles in VISUAL_TITLES.items():
        for visual_index, title in titles.items():
            if page_index == 4 and visual_index in (5, 6):
                fmt = slicer_format(title)
            elif visual_index >= (11 if page_index in (0, 1) else 10 if page_index in (2, 3) else 5):
                fmt = slicer_format(title)
            elif page_index == 4 and visual_index == 4:
                fmt = table_format(title)
            elif page_index == 4 and visual_index <= 1:
                fmt = card_format(title, card_accents[visual_index % len(card_accents)])
            elif page_index != 4 and visual_index <= (5 if page_index in (0, 1) else 4):
                fmt = card_format(title, card_accents[visual_index % len(card_accents)])
            elif title.lower().endswith("table") or "scorecard" in title.lower() or "register" in title.lower():
                fmt = table_format(title)
            else:
                fmt = chart_format(title)
            result = pbi.pbix_format_visual(alias, page_index, visual_index, json.dumps(fmt))
            if "ERROR" in result.upper():
                print(result)

    for page_index, (title, subtitle) in enumerate(PAGE_TITLES):
        print(
            pbi.pbix_add_visual(
                alias,
                page_index,
                "textbox",
                x=24,
                y=14,
                width=690,
                height=62,
                config_json=textbox_config(title, subtitle),
            )
        )

    print(pbi.pbix_save(alias, str(pbix_path), overwrite=True, backup=False))

    pbip_output = POWERBI_DIR / PBIP_DIR_NAME
    try:
        print(pbi.pbix_export_pbip(alias, str(pbip_output)))
    except Exception as exc:
        print(f"PBIP export skipped: {exc}")

    try:
        print(pbi.pbix_doctor(alias))
    except Exception as exc:
        print(f"PBIX diagnostics skipped: {exc}")

    try:
        print(pbi.pbix_close(alias, force=True))
    except Exception:
        pass


def build_dashboard() -> Path:
    try:
        from pbix_mcp.builder import PBIXBuilder
    except Exception as exc:
        raise RuntimeError(
            "Missing Power BI builder dependency. Install with: "
            "py -3.13 -m pip install pandas numpy openpyxl pbix-mcp"
        ) from exc

    frames = read_inputs()
    write_support_assets()

    builder = PBIXBuilder(name="HR Analytics Professional Dashboard")
    add_tables(builder, frames)
    add_model(builder)

    for page in build_pages():
        builder.add_page(page["name"], page["visuals"])

    pbix_path = POWERBI_DIR / PBIX_NAME
    print(f"Building PBIX at {pbix_path}")
    builder.save(str(pbix_path), validate=True)
    polish_report(pbix_path)
    return pbix_path


if __name__ == "__main__":
    start = datetime.now()
    output = build_dashboard()
    elapsed = (datetime.now() - start).total_seconds()
    print(f"Done: {output} ({elapsed:.1f}s)")
