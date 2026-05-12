import pandas as pd
from supabase import create_client, Client
import os
from dotenv import load_dotenv

def load_to_supabase():
    # Load environment variables
    load_dotenv()
    
    url = os.getenv('SUPABASE_URL')
    key = os.getenv('SUPABASE_KEY')

    if not url or not key or 'your-project' in url:
        print("Error: SUPABASE_URL or SUPABASE_KEY not set correctly in .env file.")
        return

    try:
        supabase: Client = create_client(url, key)
        
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        processed_dir = os.path.join(base_dir, 'data', 'processed')

        files_to_tables = {
            'dim_department.xlsx': 'dim_department',
            'dim_jobrole.xlsx': 'dim_jobrole',
            'ai_recommendations.xlsx': 'ai_recommendations',
            'fact_employees.xlsx': 'fact_employees'
        }

        for excel_file, table_name in files_to_tables.items():
            excel_path = os.path.join(processed_dir, excel_file)
            if os.path.exists(excel_path):
                print(f"Loading {excel_file} into Supabase table '{table_name}'...")
                df = pd.read_excel(excel_path)
                
                # Convert DataFrame to list of dicts
                # Fill NaN with None for SQL compatibility
                data = df.where(pd.notnull(df), None).to_dict(orient='records')
                
                # Supabase upsert requires a primary key in the data if we want to replace
                # For simplicity, we'll delete and insert or just upsert if PKs match
                # Note: Supabase tables must be created first!
                
                try:
                    # Clear table first (optional, but ensures clean state like previous scripts)
                    # Note: This requires appropriate permissions
                    # supabase.table(table_name).delete().neq("id", -1).execute() 
                    
                    # Upsert data (batch size 500 to avoid timeouts)
                    batch_size = 500
                    for i in range(0, len(data), batch_size):
                        batch = data[i:i + batch_size]
                        supabase.table(table_name).upsert(batch).execute()
                    
                    print(f"Successfully loaded {len(data)} rows into '{table_name}'.")
                except Exception as e:
                    print(f"Error loading into '{table_name}': {e}")
                    print("Ensure the table exists in Supabase with correct column names.")
            else:
                print(f"Warning: {csv_path} not found.")

        print("\nAll data successfully synced to Supabase!")
        
    except Exception as e:
        print(f"Failed to connect to Supabase: {e}")

if __name__ == "__main__":
    load_to_supabase()
