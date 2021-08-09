import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd
#importing the required libraries

class employees:
    def employee_details(self):
        #connecting the python file with the database
        try:
            connection = psycopg2.connect(

                database="PythonANDSQL",
                user="postgres",
                password="1234")
            cursor_object = connection.cursor()
            query_command = """
                    select dept.deptno, dept_name, sum(total_compensation) from Compensation, dept
                    where Compensation.dept_name=dept.dname
                    group by dept_name, dept.deptno
                    """
            # sql command to show the desired data

            cursor_object.execute(query_command)

            columns = [desc[0] for desc in cursor_object.description]
            data = cursor_object.fetchall()
            df = pd.DataFrame(list(data), columns=columns)

            writer = pd.ExcelWriter('number4.py.xlsx')
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("The Program has not run successfully", e)
            #Run if the program has any exceptions
        finally:
            #this will run in all the cases
            if connection is not None:
                cursor_object.close()
                connection.close()
            # Closing the connections created after the program has run


if __name__ == '__main__':
    connection = None
    cursor_object = None
    emp = employees()                   #Create a object of employees class
    emp.employee_details()