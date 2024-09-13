import sqlite3
import pandas as pd
from openpyxl import load_workbook

def export_to_excel():
    conn = sqlite3.connect('project_data.db')
    
    tasks_df = pd.read_sql_query('SELECT * FROM tasks', conn)
    resources_df = pd.read_sql_query('SELECT * FROM resources', conn)
    custom_fields_df = pd.read_sql_query('SELECT * FROM custom_fields', conn)

    excel_file = 'project_data_output.xlsx'
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        tasks_df.to_excel(writer, sheet_name='Tasks', index=False)
        resources_df.to_excel(writer, sheet_name='Resources', index=False)
        custom_fields_df.to_excel(writer, sheet_name='CustomFields', index=False)

    conn.close()
