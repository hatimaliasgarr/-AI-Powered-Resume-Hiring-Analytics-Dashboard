import pandas as pd
import sqlite3
import os

def load_data_to_sql():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    processed_dir = os.path.join(base_dir, 'data', 'processed')
    sql_dir = os.path.join(base_dir, 'sql')
    db_path = os.path.join(sql_dir, 'hr_analytics.db')
    schema_path = os.path.join(sql_dir, 'schema.sql')

    # Connect to (or create) the database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 1. Execute Schema
    print("Creating tables...")
    with open(schema_path, 'r') as f:
        schema_sql = f.read()
    cursor.executescript(schema_sql)

    # 2. Load Excel files into DataFrames and then to SQL
    files_to_load = {
        'dim_department.xlsx': 'dim_department',
        'dim_jobrole.xlsx': 'dim_jobrole',
        'ai_recommendations.xlsx': 'ai_recommendations',
        'fact_employees.xlsx': 'fact_employees'
    }

    for excel_file, table_name in files_to_load.items():
        excel_path = os.path.join(processed_dir, excel_file)
        if os.path.exists(excel_path):
            print(f"Loading {excel_file} into table {table_name}...")
            df = pd.read_excel(excel_path)
            # Use 'replace' for dimensions and AI recs to avoid PK conflicts on rerun
            # Use 'append' for fact if needed, but here we replace to ensure clean state
            df.to_sql(table_name, conn, if_exists='replace', index=False)
        else:
            print(f"Warning: {csv_path} not found.")

    conn.commit()
    conn.close()
    print(f"Database successfully created and populated at: {db_path}")

if __name__ == "__main__":
    load_data_to_sql()
