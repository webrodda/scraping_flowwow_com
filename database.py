import pymysql
from pymysql import cursors
from config import username, password, db_name, port, host, table, fields
import csv


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
    with open("result.csv", "r", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            query = f"""INSERT INTO {table} ({fields}) 
            VALUES ("{row['Название']}", "{row['Полная цена']}", "{row['Цена со скидкой']}", "{row['Описание']}", 
            "{row['Состав']}", "{row['Размер']}", "{row['Ссылки на фотографии']}");"""
            database(query=query)
    print("[INFO] Данные успешно записаны в базу данных")


if __name__ == "__main__":
    input_data_to_table()
