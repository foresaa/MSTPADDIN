from f_dialog import ask_for_file
from msp_extract import extract_fields_from_msp
from db_ops import populate_sqlite_db
from excel_export import export_to_excel

def main():
    file_path = ask_for_file()  # Ask for the Microsoft Project file
    if file_path:
        tasks_data, resources_data, custom_fields_data = extract_fields_from_msp(file_path)
        populate_sqlite_db(tasks_data, custom_fields_data, resources_data)
        export_to_excel()
        print("ETL Complete")
    else:
        print("No file selected.")

if __name__ == "__main__":
    main()
