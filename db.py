import pyodbc
import pandas as pd
import warnings

def table_exists(table_name, cursor):
    try:
        cursor.execute(f"SELECT TOP 1 * FROM {table_name}")
        return True
    except:
        return False

def update_product(category, productID, name, current_date):
    table_name = category
    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                        r'DBQ=\pathYourDb\db.accdb;')
    cursor = conn.cursor()

    if not table_exists(table_name, cursor):
        cursor.execute(f"""
        CREATE TABLE {table_name} (
            No COUNTER PRIMARY KEY,
            ID INTEGER,
            ProductName TEXT(255),
        )
        """)
        conn.commit()

    cursor.execute(f"""
        INSERT INTO {table_name} (ID, ProductName, RecordedAt)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        productID, name, current_date)

    conn.commit()
    conn.close()

def excel_table(table_name):
    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                        r'DBQ=\pathYourDb\db.accdb;')
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    conn.close()

    excel_wr = r".\Links_output.xlsx"

    try:
        existing_df = pd.read_excel(excel_wr, sheet_name=None)

        existing_df[table_name] = df

        with pd.ExcelWriter(excel_wr) as writer:
            for sheet_name, data in existing_df.items():
                data.to_excel(writer, index=False, sheet_name=sheet_name)

    except FileNotFoundError:
        with pd.ExcelWriter(excel_wr) as writer:
            df.to_excel(writer, index=False, sheet_name=table_name)