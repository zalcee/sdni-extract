import pyodbc
import pandas as pd
from datetime import datetime
from tqdm import tqdm
import argparse
import math
from time import sleep
import os

def main(xls):
    try:
        print("Connecting to server...")
        server = 'tcp:192.168.1.171' 
        database = 'SUSDATA' 
        username = 'sa' 
        password = 'root@$pruc3~#'
        # Check OS Name
        if os.name == 'nt':
            cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        elif os.name == 'posix':
            if os.uname().sysname == 'Linux':
                cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
            
        cursor = cnxn.cursor()

        if 'BARCODE_EXTRACT' in xls.sheet_names:
            du = pd.read_excel(xls, "BARCODE_EXTRACT", converters={'barcode': str})

            # List to store the results
            results_list = []

            sql_query = """
            SELECT 
                i.barcode AS BARCODE,
                i.[desc] AS DESCRIPTION,
                s.subcated AS SUBCATEGORY,
                dept.deptd AS DEPARTMENT,
                cate.categoryd AS CATEGORY,
                si.sized AS SIZE,
                c.colord AS COLOR,
                i.price AS PRICE,
                area.aread AS AREA,
                brand.brandd AS BRAND,
                i.size AS SIZE,
                sup.name AS SUPPLIER,
                grp.groupd AS COLLECTION,
                unit.unitd AS UNIT_OF_MEASUREMENT
            FROM items i
            INNER JOIN group_ grp ON grp.[group] = i.[group]
            INNER JOIN color c ON c.color = i.color
            INNER JOIN subcate s ON s.subcate = i.subcate
            INNER JOIN size si ON si.size = i.size
            INNER JOIN dept ON dept.dept = i.dept
            INNER JOIN category cate ON cate.category = i.category
            INNER JOIN area ON area.area = i.area
            INNER JOIN brand ON brand.brand = i.brand
            INNER JOIN supplier sup ON sup.suppcode = i.suppcode
            INNER JOIN unit ON unit.unit = i.unit
            WHERE i.barcode = ?
            """

            # Iterate through the barcodes in the Excel sheet and check in the database
            for index, row in tqdm(du.iterrows(), total=du.shape[0]):
                barcode = row['barcode']
                cursor.execute(sql_query, (barcode,))
                result = cursor.fetchone()
                
                if result:
                    # Convert the result to a dictionary
                    col_names = [desc[0] for desc in cursor.description]
                    result_dict = dict(zip(col_names, result))
                    results_list.append(result_dict)
                    print(f"Barcode {barcode} found in database.")
                else:
                    print(f"Barcode {barcode} not found in database.")

            # Create a DataFrame from the results
            results_df = pd.DataFrame(results_list)

            # Save the results to a new Excel file
            with pd.ExcelWriter("query_results.xlsx") as writer:
                results_df.to_excel(writer, index=False, sheet_name="Query_Results")

        cursor.close()
    except pyodbc.OperationalError:
        print("Couldn't connect to Server or Access Denied!")

if __name__ == "__main__":
    # Initialize Parser
    parser = argparse.ArgumentParser()
    # Arguments
    parser.add_argument('-f', '--file', help="Select a .xls or .xlsx file!")
    # Read arguments from Console / Terminal
    args = parser.parse_args()
    # Global Variable
    global xls
    # Arguments Condition for Excel File
    if args.file:
        # Pandas Excel File
        try:
            with pd.ExcelFile(args.file) as xls:
                main(xls)
        except FileNotFoundError:
            print(f"The file '{args.file}' couldn't be found within this directory!")
    else:
        filename = input("Filename: ")
        # Pandas Excel File
        try:
            with pd.ExcelFile(filename) as xls:
                main(xls)
        except FileNotFoundError:
            print(f"The file '{filename}' couldn't be found within this directory!")
        except ValueError:
            print(f"The file '{filename}' can't be determined as an Excel file.")
