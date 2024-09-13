import sqlite3

    # Create tables and insert data as in your original code
def populate_sqlite_db(tasks_data, custom_fields_data, resources_data):
    conn = sqlite3.connect('project_data.db')
    cursor = conn.cursor()

    # Create tasks table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            start DATE,
            finish DATE,
            duration REAL,  -- Duration as decimal days
            status TEXT,
            summary TEXT,
            critical TEXT,
            stop DATE,
            resume DATE
        )
    ''')

    # Insert task data
    for task in tasks_data:
        cursor.execute('''
            INSERT INTO tasks (name, start, finish, duration, status, summary, critical, stop, resume)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            task['Name'], task['Start'], task['Finish'], task['Duration'],
            task['Status'], task['Summary'], task['Critical'], task['Stop'], task['Resume']
        ))

    # Create custom fields table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS custom_fields (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER,
            field_type TEXT,
            field_name TEXT,
            field_value TEXT,
            FOREIGN KEY(task_id) REFERENCES tasks(id)
        )
    ''')

    # Insert custom fields data
    for field in custom_fields_data:
        cursor.execute('''
            INSERT INTO custom_fields (task_id, field_type, field_name, field_value)
            VALUES (?, ?, ?, ?)
        ''', (None, field['FieldType'], field['FieldName'], field['FieldValue']))

    # Create resources table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS resources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            max_units INTEGER,
            cost REAL,
            resource_group TEXT
        )
    ''')

    # Insert resource data
    for resource in resources_data:
        cursor.execute('''
            INSERT INTO resources (name, max_units, cost, resource_group)
            VALUES (?, ?, ?, ?)
        ''', (
            resource['Name'], resource['MaxUnits'], resource['Cost'], resource['Group']
        ))

    conn.commit()
    conn.close()
