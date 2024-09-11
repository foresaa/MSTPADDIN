from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import sqlite3

app = FastAPI()

# SQLite database connection
def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn

# Example Pydantic model
class Item(BaseModel):
    name: str
    description: str

# Initialize database and create table
@app.on_event("startup")
def startup():
    conn = get_db_connection()
    conn.execute('''CREATE TABLE IF NOT EXISTS items (id INTEGER PRIMARY KEY, name TEXT, description TEXT)''')
    conn.commit()
    conn.close()


@app.get("/")
def read_root():
    return {"message": "Welcome to the backend service for Project Pilot!"}

# Endpoint to get all items
@app.get("/items/")
def read_items():
    conn = get_db_connection()
    items = conn.execute("SELECT * FROM items").fetchall()
    conn.close()
    return [dict(item) for item in items]

# Endpoint to create a new item
@app.post("/items/")
def create_item(item: Item):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO items (name, description) VALUES (?, ?)", (item.name, item.description))
    conn.commit()
    new_item_id = cursor.lastrowid
    conn.close()
    return {"id": new_item_id, "name": item.name, "description": item.description}

# Endpoint to get an item by ID
@app.get("/items/{item_id}")
def read_item(item_id: int):
    conn = get_db_connection()
    item = conn.execute("SELECT * FROM items WHERE id = ?", (item_id,)).fetchone()
    conn.close()
    if item is None:
        raise HTTPException(status_code=404, detail="Item not found")
    return dict(item)

# Endpoint to delete an item by ID
@app.delete("/items/{item_id}")
def delete_item(item_id: int):
    conn = get_db_connection()
    conn.execute("DELETE FROM items WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()
    return {"detail": "Item deleted"}

# ADDIN FUNCTIONALITY ENDPOINTS

#____________________________________Load Current Project Data _______________________________
#
# This basically executes the process of using win32 API to read the active project in MSP where the Add-In resides to 
# load all project data temporarily into the SQLite db residing on Render as 'repository to support all further functions of th Add-In


@app.post("/load-current-project-data/")
def load_project_data():
    # Run the external Python script (mainEX.py)
    try:
        result = subprocess.run(["python", "mainEX.py"], capture_output=True, text=True)
        if result.returncode == 0:
            return {"message": "Project data loaded successfully!", "output": result.stdout}
        else:
            return {"message": "Error loading project data", "error": result.stderr}
    except Exception as e:
        return {"message": "An error occurred", "error": str(e)}
#_____________________________________________________________________________________________
