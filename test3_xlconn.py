# Windows Authentication

import pyodbc
import xlrd

try:
    conn = pyodbc.connect(
        "Driver={SQL Server Native Client 11.0};"
        "Server=SERVER_NAME;"
        "Database=Database_NAME;"
        "Trusted_Connection=yes;"
        )
except pyodbc.OperationalError as err:
    print("Could not establish connection to SQL Server: " + str(err))



def f1(conn):
    workbook = xlrd.open_workbook('wb.xlsx')
    first_sheet = workbook.sheet_by_index(0)
    cell = first_sheet.cell(0,0)
    print(cell.value)
    value=str(cell.value)
    print(type(value))

    print("-----Fetching table data-----","\n")
    cursor = conn.cursor()

    try:
        cursor.execute(value)
        rows = cursor.fetchall()
        print("Total rows fetched : ",len(rows),"\n")
        print(rows[0],"\n")
        print(rows[-1],"\n")
        print('-----Records fetched successfully-----',"\n")
    
        sql='''
                SET IDENTITY_INSERT Demo_Harshad ON
            
                INSERT into Demo_Harshad(empId,empName,empDate,SchemaName) values(?,?,?,?);
            
                SET IDENTITY_INSERT Demo_Harshad OFF
            '''
        cursor.executemany(sql, (rows))
        conn.commit()
        cursor.close()
        print('-----Records inserted successfully-----',"\n")

    except ValueError as e:
        print(str(e))
        print("---Unable to show data---")

    finally:
        if conn:
            conn.close()
            print("-----The connection is closed-----")

    

f1(conn)


