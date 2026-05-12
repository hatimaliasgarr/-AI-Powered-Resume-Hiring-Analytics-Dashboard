# 🧠 AI-Powered HR & Hiring Analytics Dashboard

---

## 🎯 Project Overview
This project delivers a comprehensive, end-to-end **AI-Powered HR & Hiring Analytics** solution. Leveraging the IBM HR Analytics dataset, it provides deep insights into employee attrition, performance, salary equity, and recruitment efficiency. The dashboard features an **AI Recommendation Engine** that translates complex data patterns into actionable HR strategies, enabling data-driven decision-making for workforce management.

## 📊 Dashboard Pages
1.  **Executive HR Overview**: High-level KPIs including Attrition Rate, Retention, and Headcount trends by department.
2.  **Attrition Deep Dive**: Detailed analysis of attrition drivers, high-risk groups, and tenure-based exit patterns.
3.  **Skill Gap & Performance Intelligence**: Evaluation of skill scores vs. targets, training effectiveness, and performance distribution.
4.  **Salary Distribution & Hiring Analytics**: Insights into pay equity, salary bands, and recruitment funnel efficiency.
5.  **🤖 AI Recommendation Panel**: A dedicated panel displaying 8 data-backed recommendations with priority badges and actionable advice.

## 📦 Dataset Source
The project uses the **IBM HR Analytics Employee Attrition & Performance** dataset from Kaggle.
🔗 [Kaggle Dataset Link](https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset)

## 🛠️ Tech Stack
- **Python**: Data processing, feature engineering, and AI insight generation.
- **Pandas & NumPy**: Data manipulation and statistical analysis.
- **Power BI Desktop**: Data modeling, DAX measures, and interactive visualization.
- **SQLite / SQL**: Relational database storage for local reporting.
- **Supabase**: Cloud PostgreSQL database for scalable workforce analytics.

## 🚀 How to Run
1.  **Install Requirements**:
    ```bash
    pip install pandas numpy
    ```
2.  **Run Data Scripts**:
    ```bash
    python scripts/prepare_hr_data.py
    python scripts/generate_ai_insights.py
    python scripts/load_to_sql.py
    python scripts/load_to_supabase_pg.py
    ```
3.  **Open Power BI**:
    - Open `powerbi/HR_Analytics_Professional_Dashboard.pbix`.
    - Optional: open `powerbi/HR_Analytics_Professional_Dashboard_PBIP/HR_Analytics_Professional_Dashboard.pbip` for the source-controlled Power BI project version.
    - The report is already connected to the processed Excel tables in `data/processed/`; click **Refresh** after regenerating data.
    - To rebuild the dashboard from code, run `py -3.13 scripts/build_powerbi_dashboard.py`.

## 💡 Key Insights
- **Attrition Drivers**: Overtime and low job satisfaction are the primary predictors of turnover in the Research & Development department.
- **Pay Equity**: Identified specific job roles with salary gaps, allowing for targeted compensation adjustments.
- **Training ROI**: Roles like "Sales Representative" show the highest performance boost following training sessions.
- **Hiring Efficiency**: Referral and LinkedIn channels consistently provide candidates with higher initial skill scores.
- **Promotion Bottlenecks**: Sales department shows a significant delay in career progression, impacting long-term retention.

## 🤖 AI Recommendation Panel
The recommendation panel utilizes a custom Python engine to identify patterns across multiple dimensions (Performance, Satisfaction, Salary, Tenure). It generates:
- **Insight**: A data-backed finding explaining "what is happening".
- **Actionable Advice**: A specific HR strategy for "what to do next".
- **Priority**: Color-coded badges (High/Medium/Low) based on the severity of the finding.

## 🎓 Skills Demonstrated
- **HR Analytics & Workforce Planning**
- **Advanced DAX & Data Modeling**
- **Feature Engineering & Synthetic Data Generation**
- **Automated AI Insight Engines**
- **Professional Dashboard Design & UI/UX**
- **Relational Database Design & SQL ETL**
- **Cloud Database Integration (Supabase / PostgreSQL)**

## 🏛️ Database Architecture
The project utilizes a dual-database approach:
1. **Local SQL (SQLite)**: Best for local development and Power BI Desktop reporting.
2. **Cloud SQL (Supabase)**: Provides a production-ready PostgreSQL environment for collaborative analytics.

### Supabase Tables:
- `fact_employees`: Main employee fact table.
- `dim_department`: Department dimension.
- `dim_jobrole`: Job role dimension.
- `ai_recommendations`: AI-generated insights.

### How to Sync to Supabase:
1. Add your `SUPABASE_PG_URL` (direct PostgreSQL connection string) to the `.env` file.
2. Run `python scripts/load_to_supabase_pg.py`. This script automatically creates the tables and populates the data from the `.xlsx` files.

---
*Created as a production-ready HR Analytics Portfolio Project.*
