import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd
#importing the required libraries

class employees:
    def employees_details(self):
        #connecting the python file with the database
        try:
            connection = psycopg2.connect(
                host="localhost",
                database="PythonANDSQL",
                user="postgres",
                password="1234")
            cursor_object = connection.cursor()
            #creating a object

            query_command = """SELECT e1.empno, e1.ename, (case when mgr is not null then (select ename from emp as e2 where e1.mgr=e2.empno limit 1) else null end) as manager
            from emp as e1"""
            #sql command to show the desired data

            cursor_object.execute(query_command)

            columns = [desc[0] for desc in cursor_object.description]
            data = cursor_object.fetchall()
            df = pd.DataFrame(list(data), columns=columns)

            writer = pd.ExcelWriter('number1.py.xlsx')     #converting the code to xlsx file
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("The Program has not run successfully", e)
            #Run if the program has any exceptions
        finally:
            #this will run in all test cases
            if connection is not None:
                cursor_object.close()
                connection.close()
            #Closing the connections created after the program has run


if __name__=='__main__':
    connection = None
    cursor_object = None
    emp = employees()          #Create a object of employees class
    emp.employees_details()