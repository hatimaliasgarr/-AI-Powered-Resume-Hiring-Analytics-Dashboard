import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import os
from dotenv import load_dotenv

def load_to_supabase_pg():
    # Load environment variables
    load_dotenv()
    
    pg_url = os.getenv('SUPABASE_PG_URL')

    if not pg_url or '[YOUR-PASSWORD]' in pg_url:
        print("Error: SUPABASE_PG_URL not set correctly in .env file.")
        return

    try:
        # Connect to PostgreSQL
        print("Connecting to Supabase PostgreSQL...")
        conn = psycopg2.connect(pg_url)
        cursor = conn.cursor()
        
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        processed_dir = os.path.join(base_dir, 'data', 'processed')
        schema_path = os.path.join(base_dir, 'sql', 'schema.sql')

        # 1. Execute Schema
        print("Executing schema.sql...")
        with open(schema_path, 'r') as f:
            schema_sql = f.read()
        cursor.execute(schema_sql)
        conn.commit()

        # 2. Load Excel files
        files_to_tables = {
            'dim_department.xlsx': 'dim_department',
            'dim_jobrole.xlsx': 'dim_jobrole',
            'ai_recommendations.xlsx': 'ai_recommendations',
            'fact_employees.xlsx': 'fact_employees'
        }

        for excel_file, table_name in files_to_tables.items():
            excel_path = os.path.join(processed_dir, excel_file)
            if os.path.exists(excel_path):
                print(f"Loading {excel_file} into '{table_name}'...")
                df = pd.read_excel(excel_path)
                
                # Prepare data for insertion
                columns = [f'"{col}"' for col in df.columns]
                table_cols = ', '.join(columns)
                
                # Convert NaN to None for SQL NULL
                values = [tuple(x) for x in df.where(pd.notnull(df), None).values]
                
                # Truncate table before loading (ensures clean state)
                cursor.execute(f'TRUNCATE TABLE "{table_name}" CASCADE;')
                
                # Bulk insert
                insert_query = f'INSERT INTO "{table_name}" ({table_cols}) VALUES %s'
                execute_values(cursor, insert_query, values)
                conn.commit()
                
                print(f"Successfully loaded {len(df)} rows into '{table_name}'.")
            else:
                print(f"Warning: {csv_path} not found.")

        cursor.close()
        conn.close()
        print("\nAll data successfully synced to Supabase PostgreSQL!")
        
    except Exception as e:
        print(f"Failed to connect or load data: {e}")

if __name__ == "__main__":
    load_to_supabase_pg()
