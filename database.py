import pymysql
from pymysql import cursors
from config import username, password, db_name, port, host
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


def get_clean_value(workbook, formula):
    item_list = formula[1:].split("!")
    page = item_list[0]
    cell = item_list[1]
    result = workbook[page][cell].value
    return result


def input_data_to_table():
    workbook = openpyxl.load_workbook("result.xlsx")
    oc_product = workbook["oc_product"]
    oc_product_description = workbook["oc_product_description"]
    oc_product_image = workbook["oc_product_image"]
    oc_product_to_category = workbook["oc_product_to_category"]
    oc_seo_url = workbook["oc_seo_url"]

    # select oc_product
    for index in range(2, oc_product.max_row + 1):
        product_id = oc_product[f"A{index}"].value
        quantity = oc_product[f"B{index}"].value
        stock_status_id = oc_product[f"C{index}"].value
        image = oc_product[f"D{index}"].value
        shipping = oc_product[f"E{index}"].value
        price = oc_product[f"F{index}"].value
        query_to_database = f"""INSERT INTO oc_product (product_id, quantity, stock_status_id, image, shipping, price) 
                                VALUES ({product_id}, {quantity}, {stock_status_id}, '{image}', {shipping}, {price});"""
        database(query=query_to_database)

    # select oc_product_description
    for index in range(2, oc_product_description.max_row + 1):
        product_id = get_clean_value(workbook=workbook, formula=oc_product_description[f"A{index}"].value)
        language_id = oc_product_description[f"B{index}"].value
        name = str(oc_product_description[f"C{index}"].value).replace('"', "'")
        description = str(oc_product_description[f"D{index}"].value).replace('"', "'")
        meta_title = oc_product_description[f"E{index}"].value
        meta_description = oc_product_description[f"F{index}"].value
        query_to_database = f"""INSERT INTO oc_product_description (product_id, language_id, name, description, 
                                                                    meta_title, meta_description) 
                                VALUES ({product_id}, {language_id}, "{name}", "{description}", 
                                        '{meta_title}', '{meta_description}');"""
        database(query=query_to_database)

    # select oc_product_image
    for index in range(2, oc_product_image.max_row + 1):
        product_image_id = oc_product_image[f"A{index}"].value
        product_id = get_clean_value(workbook=workbook, formula=oc_product_image[f"B{index}"].value)
        image = oc_product_image[f"C{index}"].value
        sort_order = oc_product_image[f"D{index}"].value
        query_to_database = f"""INSERT INTO oc_product_image (product_image_id, product_id, image, sort_order) 
                                VALUES ({product_image_id}, {product_id}, '{image}', {sort_order});"""
        database(query=query_to_database)

    # select oc_product_to_category
    for index in range(2, oc_product_to_category.max_row + 1):
        product_id = get_clean_value(workbook=workbook, formula=oc_product_to_category[f"A{index}"].value)
        category_id = oc_product_to_category[f"B{index}"].value
        query_to_database = f"""INSERT INTO oc_product_to_category (product_id, category_id) 
                                VALUES ({product_id}, {category_id});"""

        database(query=query_to_database)

    # select oc_seo_url
    for index in range(2, oc_seo_url.max_row + 1):
        seo_url_id = oc_seo_url[f"A{index}"].value
        store_id = oc_seo_url[f"B{index}"].value
        language_id = oc_seo_url[f"C{index}"].value
        query = oc_seo_url[f"D{index}"].value
        keyword = oc_seo_url[f"E{index}"].value
        query_to_database = f"""INSERT INTO oc_seo_url (seo_url_id, store_id, language_id, query, keyword) 
                                VALUES ({seo_url_id}, {store_id}, {language_id}, '{query}', '{keyword}');"""
        database(query=query_to_database)

    print("[INFO] Данные успешно записаны в базу данных")


def create_database():
    create_oc_product = """CREATE TABLE oc_product(product_id INT UNIQUE NOT NULL, quantity INT NOT NULL,
                                                   stock_status_id INT NOT NULL, image VARCHAR(1000),
                                                   shipping INT NOT NULL, price INT NOT NULL);"""
    database(query=create_oc_product)

    create_oc_product_description = """CREATE TABLE oc_product_description(product_id INT UNIQUE NOT NULL,
                                                                           language_id INT NOT NULL,
                                                                           name VARCHAR(500),
                                                                           description TEXT,
                                                                           meta_title VARCHAR(500),
                                                                           meta_description VARCHAR(500));"""
    database(query=create_oc_product_description)

    create_oc_product_image = """CREATE TABLE oc_product_image(product_image_id INT UNIQUE NOT NULL,
                                                               product_id INT NOT NULL,
                                                               image VARCHAR(500),
                                                               sort_order INT NOT NULL);"""
    database(query=create_oc_product_image)

    create_oc_product_to_category = """CREATE TABLE oc_product_to_category(product_id INT NOT NULL UNIQUE,
                                                                           category_id INT NOT NULL);"""
    database(query=create_oc_product_to_category)

    create_oc_seo_url = """CREATE TABLE oc_seo_url(seo_url_id INT UNIQUE NOT NULL, store_id INT NOT NULL,
                                                   language_id INT NOT NULL, query VARCHAR(100), 
                                                   keyword VARCHAR(500));"""
    database(query=create_oc_seo_url)


if __name__ == "__main__":
    create_database()
    input_data_to_table()
