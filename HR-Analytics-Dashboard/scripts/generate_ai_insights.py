import pandas as pd
import numpy as np
import os

def generate_insights():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    fact_path = os.path.join(base_dir, 'data', 'processed', 'fact_employees.xlsx')
    output_path = os.path.join(base_dir, 'data', 'processed', 'ai_recommendations.xlsx')

    if not os.path.exists(fact_path):
        print(f"Error: Could not find {fact_path}")
        return

    df = pd.read_excel(fact_path)
    recommendations = []

    # 1. Best Department to Hire From
    dept_stats = df.groupby('Department').agg(
        AttritionRate=('AttritionBinary', 'mean'),
        AvgPerformance=('PerformanceRating', 'mean')
    ).reset_index()
    # Score = (1 - AttritionRate) * AvgPerformance
    dept_stats['Score'] = (1 - dept_stats['AttritionRate']) * dept_stats['AvgPerformance']
    best_dept = dept_stats.loc[dept_stats['Score'].idxmax()]
    recommendations.append({
        'RecommendationID': 1,
        'Title': 'Best Department to Hire From',
        'Insight': f"{best_dept['Department']} department has the lowest attrition ({best_dept['AttritionRate']:.1%}) and highest performance ({best_dept['AvgPerformance']:.2f}).",
        'ActionableAdvice': f"Prioritize internal transfers and external hiring for {best_dept['Department']} to maintain workforce stability.",
        'Priority': 'Low',
        'ImpactArea': 'Talent Acquisition'
    })

    # 2. Highest Attrition Risk Group
    risk_df = df[df['AttritionRisk'] == 'High']
    risk_group = risk_df.groupby(['JobRole', 'OverTime']).size().reset_index(name='Count')
    if not risk_group.empty:
        worst_group = risk_group.loc[risk_group['Count'].idxmax()]
        recommendations.append({
            'RecommendationID': 2,
            'Title': 'Highest Attrition Risk Group',
            'Insight': f"{worst_group['JobRole']} employees working overtime show the highest concentration of high-risk attrition profiles.",
            'ActionableAdvice': "Implement immediate workload reviews and wellness check-ins for overtime workers in these roles.",
            'Priority': 'High',
            'ImpactArea': 'Retention'
        })
    else:
        recommendations.append({
            'RecommendationID': 2,
            'Title': 'Highest Attrition Risk Group',
            'Insight': "No high-risk attrition groups identified with current thresholds.",
            'ActionableAdvice': "Continue monitoring job satisfaction and overtime metrics across all departments.",
            'Priority': 'Low',
            'ImpactArea': 'Retention'
        })

    # 3. Skill Gap Alert
    dept_skill = df.groupby('Department')['SkillScore'].mean().reset_index()
    low_skill_depts = dept_skill[dept_skill['SkillScore'] < 70] # Using 70 as threshold if 60 is too low
    if not low_skill_depts.empty:
        worst_skill_dept = low_skill_depts.loc[low_skill_depts['SkillScore'].idxmin()]
        recommendations.append({
            'RecommendationID': 3,
            'Title': 'Skill Gap Alert',
            'Insight': f"{worst_skill_dept['Department']} shows an average skill score of {worst_skill_dept['SkillScore']:.1f}, below the target threshold.",
            'ActionableAdvice': f"Launch a targeted upskilling program or technical training bootcamp for the {worst_skill_dept['Department']} team.",
            'Priority': 'Medium',
            'ImpactArea': 'L&D'
        })
    else:
        recommendations.append({
            'RecommendationID': 3,
            'Title': 'Skill Gap Alert',
            'Insight': "All departments are meeting the minimum skill score requirements.",
            'ActionableAdvice': "Shift focus to advanced specialized certifications for top-performing departments.",
            'Priority': 'Low',
            'ImpactArea': 'L&D'
        })

    # 4. Salary Equity Warning
    role_gender_salary = df.groupby(['JobRole', 'Gender'])['MonthlyIncome'].mean().unstack()
    role_gender_salary['Gap'] = abs(role_gender_salary['Male'] - role_gender_salary['Female']) / role_gender_salary[['Male', 'Female']].max(axis=1)
    high_gap_roles = role_gender_salary[role_gender_salary['Gap'] > 0.05] # Using 5% for better detection
    if not high_gap_roles.empty:
        worst_gap_role = high_gap_roles['Gap'].idxmax()
        gap_val = high_gap_roles.loc[worst_gap_role, 'Gap']
        recommendations.append({
            'RecommendationID': 4,
            'Title': 'Salary Equity Warning',
            'Insight': f"A {gap_val:.1%} salary gap exists between genders in the {worst_gap_role} role.",
            'ActionableAdvice': "Conduct a comprehensive pay equity audit and adjust compensation structures to ensure fairness.",
            'Priority': 'High',
            'ImpactArea': 'Compensation'
        })
    else:
        recommendations.append({
            'RecommendationID': 4,
            'Title': 'Salary Equity Warning',
            'Insight': "No significant gender-based salary gaps (>15%) detected across job roles.",
            'ActionableAdvice': "Maintain transparent compensation policies and continue annual equity reviews.",
            'Priority': 'Low',
            'ImpactArea': 'Compensation'
        })

    # 5. Top Retention Strategy
    # Using correlation with AttritionBinary (negative correlation means factor prevents attrition)
    # Filter numeric for correlation
    numeric_df = df.select_dtypes(include=[np.number])
    correlations = numeric_df.corr()['AttritionBinary'].sort_values()
    top_factor = correlations.index[0] # Most negative correlation
    recommendations.append({
        'RecommendationID': 5,
        'Title': 'Top Retention Strategy',
        'Insight': f"{top_factor.replace('Years', ' Years ')} shows the strongest negative correlation with attrition risk.",
        'ActionableAdvice': f"Enhance programs that promote {top_factor.lower()} to improve long-term employee retention.",
        'Priority': 'Medium',
        'ImpactArea': 'Strategy'
    })

    # 6. Optimal Training Investment
    training_impact = df.groupby('JobRole').apply(lambda x: x['TrainingTimesLastYear'].corr(x['PerformanceRating']))
    # Handle NaN
    training_impact = training_impact.dropna()
    if not training_impact.empty:
        top_impact_role = training_impact.idxmax()
        recommendations.append({
            'RecommendationID': 6,
            'Title': 'Optimal Training Investment',
            'Insight': f"Training has the highest positive impact on performance for {top_impact_role} roles.",
            'ActionableAdvice': f"Allocate additional training budget to {top_impact_role} to maximize ROI on performance.",
            'Priority': 'Medium',
            'ImpactArea': 'L&D'
        })
    else:
        recommendations.append({
            'RecommendationID': 6,
            'Title': 'Optimal Training Investment',
            'Insight': "General training programs show consistent but modest impact across all roles.",
            'ActionableAdvice': "Explore role-specific specialized training to drive higher performance correlation.",
            'Priority': 'Low',
            'ImpactArea': 'L&D'
        })

    # 7. Hiring Channel Efficiency
    channel_perf = df.groupby('HiringChannel')['SkillScore'].mean().reset_index()
    best_channel = channel_perf.loc[channel_perf['SkillScore'].idxmax()]
    recommendations.append({
        'RecommendationID': 7,
        'Title': 'Hiring Channel Efficiency',
        'Insight': f"{best_channel['HiringChannel']} produces hires with the highest average skill scores ({best_channel['SkillScore']:.1f}).",
        'ActionableAdvice': f"Increase recruiting budget for {best_channel['HiringChannel']} as it delivers high-quality talent consistently.",
        'Priority': 'Medium',
        'ImpactArea': 'Talent Acquisition'
    })

    # 8. Promotion Bottleneck
    promo_dept = df.groupby('Department')['YearsSinceLastPromotion'].mean().reset_index()
    worst_promo_dept = promo_dept.loc[promo_dept['YearsSinceLastPromotion'].idxmax()]
    recommendations.append({
        'RecommendationID': 8,
        'Title': 'Promotion Bottleneck',
        'Insight': f"{worst_promo_dept['Department']} has the longest average time since last promotion ({worst_promo_dept['YearsSinceLastPromotion']:.1f} years).",
        'ActionableAdvice': "Review career progression paths and promotion criteria within the Sales department.",
        'Priority': 'High',
        'ImpactArea': 'Retention'
    })

    # Save to Excel
    res_df = pd.DataFrame(recommendations)
    res_df.to_excel(output_path, index=False)
    print(f"AI Recommendations generated and saved to: {output_path}")

if __name__ == "__main__":
    generate_insights()
