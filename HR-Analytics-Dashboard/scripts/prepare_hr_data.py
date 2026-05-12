import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os

# Set paths
RAW_DATA_PATH = os.path.join('data', 'raw', 'WA_Fn-UseC_-HR-Employee-Attrition.csv')
PROCESSED_DIR = os.path.join('data', 'processed')

def prepare_data():
    # 2A. Load & Inspect
    # Check current directory and set path accordingly
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    raw_path = os.path.join(base_dir, 'data', 'raw', 'WA_Fn-UseC_-HR-Employee-Attrition.csv')
    processed_dir = os.path.join(base_dir, 'data', 'processed')

    if not os.path.exists(raw_path):
        print(f"Error: Could not find {raw_path}")
        return

    df = pd.read_csv(raw_path)
    print(f"Initial Shape: {df.shape}")

    # 2B. Add Missing Columns
    # HireDate -> Generate randomly between 2010-01-01 and 2022-12-31, ensuring YearsAtCompany is consistent
    ref_date = datetime(2023, 1, 1)
    random.seed(42) # For reproducibility
    df['HireDate'] = df['YearsAtCompany'].apply(lambda x: ref_date - timedelta(days=int(x*365 + random.randint(0, 364))))
    df['HireDate'] = df['HireDate'].dt.strftime('%Y-%m-%d')

    # SkillScore -> (PerformanceRating * 20) + (TrainingTimesLastYear * 5), capped at 100
    df['SkillScore'] = (df['PerformanceRating'] * 20) + (df['TrainingTimesLastYear'] * 5)
    df['SkillScore'] = df['SkillScore'].clip(upper=100)

    # AttritionRisk
    def get_attrition_risk(row):
        if row['JobSatisfaction'] <= 2 and row['OverTime'] == 'Yes':
            return 'High'
        elif row['JobSatisfaction'] == 3 or row['YearsSinceLastPromotion'] >= 4:
            return 'Medium'
        else:
            return 'Low'
    
    df['AttritionRisk'] = df.apply(get_attrition_risk, axis=1)

    # HiringChannel
    channels = ['LinkedIn', 'Referral', 'Job Portal', 'Campus', 'Agency']
    df['HiringChannel'] = [random.choice(channels) for _ in range(len(df))]

    # SalaryBand
    def get_salary_band(income):
        if income < 3000: return 'Entry'
        elif income < 7000: return 'Mid'
        elif income < 12000: return 'Senior'
        else: return 'Executive'
    
    df['SalaryBand'] = df['MonthlyIncome'].apply(get_salary_band)

    # TenureGroup
    def get_tenure_group(years):
        if years <= 2: return '0-2 yrs'
        elif years <= 5: return '3-5 yrs'
        elif years <= 10: return '6-10 yrs'
        else: return '10+ yrs'
    
    df['TenureGroup'] = df['YearsAtCompany'].apply(get_tenure_group)

    # YearHired
    df['YearHired'] = pd.to_datetime(df['HireDate']).dt.year

    # PromotionFlag
    df['PromotionFlag'] = (df['YearsSinceLastPromotion'] == 0).astype(int)

    # 2C. Clean Data
    df = df.drop_duplicates()
    
    # Fill nulls in numeric columns with median
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].median())

    # Standardize text columns (strip whitespace, title case)
    text_cols = df.select_dtypes(include=['object']).columns
    for col in text_cols:
        df[col] = df[col].astype(str).str.strip().str.title()

    # Convert Attrition to binary: AttritionBinary -> Yes = 1, No = 0
    df['AttritionBinary'] = df['Attrition'].apply(lambda x: 1 if x == 'Yes' else 0)

    # Ensure MonthlyIncome is numeric integer
    df['MonthlyIncome'] = df['MonthlyIncome'].astype(int)

    # 2D. Export Clean Files
    os.makedirs(processed_dir, exist_ok=True)
    
    # fact_employees.xlsx
    df.to_excel(os.path.join(processed_dir, 'fact_employees.xlsx'), index=False)

    # dim_department.xlsx
    dim_dept = df.groupby('Department').agg(
        AvgSalary=('MonthlyIncome', 'mean'),
        HeadCount=('EmployeeNumber', 'count')
    ).reset_index()
    dim_dept['DepartmentID'] = range(1, len(dim_dept) + 1)
    dim_dept.to_excel(os.path.join(processed_dir, 'dim_department.xlsx'), index=False)

    # dim_jobrole.xlsx
    dim_role = df.groupby('JobRole').agg(
        AvgSkillScore=('SkillScore', 'mean'),
        AvgSatisfaction=('JobSatisfaction', 'mean')
    ).reset_index()
    dim_role.to_excel(os.path.join(processed_dir, 'dim_jobrole.xlsx'), index=False)

    print(f"Data preparation complete. Files exported to: {processed_dir}")

if __name__ == "__main__":
    prepare_data()
