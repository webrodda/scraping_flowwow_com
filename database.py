import pymysql
from pymysql import cursors
from config import username, password, db_name, port, host, table, fields
import openpyxl


def database(query):
    connection = pymysql.connect(user=username, password=password, db=db_name, port=port,
                                 host=host, cursorclass=cursors.DictCursor)
    try:
        cursor = connection.cursor()
        cursor.execute(query=query)
        result = cursor.fetchall()
        connection.commit()
        return result
    except Exception as ex:
        print(f"[ERROR] Execute not successful because {ex}")
        return False
    finally:
        connection.close()


def input_data_to_table():
    workbook = openpyxl.load_workbook("result.xlsx")
    worksheet = workbook.active
    result = list()
    for index_row in range(1, worksheet.max_row):
        sub_result = list()
        for column in worksheet.iter_cols(1, worksheet.max_column):
            sub_result.append(column[index_row].value)
        result.append(sub_result)
    for record in result:
        query = f"""INSERT INTO {table} ({fields}) 
        VALUES ({record[0]}, "{record[1]}", {record[2]}, "{record[3]}", "{record[4]}", "{record[5]}");"""
        database(query=query)
    print("[INFO] Данные успешно записаны в базу данных")


if __name__ == "__main__":
    input_data_to_table()
