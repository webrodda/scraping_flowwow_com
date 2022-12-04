import os.path
import shutil
from openpyxl import Workbook
import pathlib
import requests
from config import new_domain
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from bs4 import BeautifulSoup
from database import input_data_to_table
import json
import time
import re


ALPHA = {"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "j", "з": "z", "и": "i",
         "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
         "у": "u", "ф": "f", "х": "x", "ц": "c", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "", "ы": "i", "ь": "",
         "э": "e", "ю": "yu", "я": "ya"}

workbook = Workbook()
workbook_page = workbook.active

ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
options = webdriver.ChromeOptions()
options.add_argument(f"user-agent={ua}")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--headless")
timeout = 300


def get_product_ids():
    browser = webdriver.Chrome(options=options)
    browser.set_page_load_timeout(time_to_wait=timeout)
    wait = WebDriverWait(browser, 200)
    try:
        url = "https://flowwow.com/shop/flower-bag/"
        browser.get(url=url)
        time.sleep(2)
        wait.until(ec.visibility_of_element_located((By.CLASS_NAME, "store_readmore")))
        read_more_buttons = browser.find_elements(By.LINK_TEXT, "Показать ещё")
        for read_more_button in read_more_buttons:
            wait.until(ec.element_to_be_clickable((By.CLASS_NAME, "store_readmore")))
            browser.execute_script("arguments[0].scrollIntoView();", read_more_button)
            browser.execute_script("arguments[0].click();", read_more_button)
            time.sleep(1)
        last_block = browser.find_elements(By.TAG_NAME, "div")[-1]
        browser.execute_script("arguments[0].scrollIntoView();", last_block)
        time.sleep(2)
        response = browser.page_source
        bs_object = BeautifulSoup(response, "lxml")
        products = bs_object.find_all(name="a", class_=re.compile(r"shop-product js-product-popu\w*"))
        product_ids = set([product["data-id"] for product in products])
        print(f"[INFO] В магазине найдено {len(product_ids)} уникальных товаров")
        return product_ids
    finally:
        browser.close()
        browser.quit()


def write_data(data, index):
    global workbook_page
    workbook_page[f"A{index}"].value = data[0]
    workbook_page[f"B{index}"].value = data[1]
    workbook_page[f"C{index}"].value = data[2]
    workbook_page[f"D{index}"].value = data[3]
    workbook_page[f"E{index}"].value = data[4]
    workbook_page[f"F{index}"].value = data[5]


def create_seo_url(title):
    result = list()
    sub_result = list()
    for symbol in title.lower():
        if symbol.isalnum():
            sub_result.append(symbol)
        else:
            sub_result.append("-")
    for symbol in sub_result:
        if symbol == "-" or symbol.isdigit():
            result.append(symbol)
        else:
            if symbol in ALPHA.keys():
                result.append(ALPHA[symbol])
            else:
                result.append(symbol)
    result = "".join(result).replace("--", "-")
    return result


def download_photos(photos, title):
    index = 0
    result = list()
    for photo in photos:
        index += 1
        response = requests.get(photo)
        path = pathlib.Path("photos", f"{title} {index}.jpg")
        with open(path, "wb") as file:
            file.write(response.content)
        correct_url = f'{new_domain}{photo.split("data")[1]}'
        result.append(correct_url)
    result = "\n".join(result)

    return result


def get_data(product_ids):
    browser = webdriver.Chrome(options=options)
    browser.set_page_load_timeout(time_to_wait=timeout)
    index = 0
    index_field = 1
    try:
        for product_id in product_ids:
            print(f"[INFO] Собираем данные о товаре номер {product_id}")
            index += 1
            index_field += 1
            url = f"https://flowwow.com/moscow/data/getProductInfo/?id={product_id}&from=direct&lang=ru&currency=RUB"
            browser.get(url=url)
            response = browser.page_source
            bs_object = BeautifulSoup(response, "lxml")
            json_object = json.loads(bs_object.text)["data"]
            photo_objects = json_object["photos"]
            photos = list()
            for photo_object in photo_objects:
                if "img" in photo_object.keys():
                    photos.append(photo_object["img"].replace('"', ""))
                else:
                    video_object = photo_object["html"].replace("\n", "").replace("\\", "")
                    bs_object = BeautifulSoup(video_object, "lxml")
                    link = bs_object.source["src"].replace('"', "")
                    photos.append(link)

            if "base" in json_object.keys():
                full_price = json_object["base"]
            else:
                full_price = json_object["cost"]

            full_info = json_object["fullInfo"]
            full_info = full_info.replace("\n", "").replace("\\", "")
            bs_object = BeautifulSoup(full_info, "lxml")
            title = bs_object.find(name="div", class_='pp-title').text.strip().replace('"', "")
            photos = download_photos(photos=photos, title=title)

            composition = bs_object.find(name="div", class_="product-desc-line").find(name="div", class_='desc')
            if composition is not None:
                composition = composition.text.strip().split(".")
                composition = ", ".join(element.strip() for element in composition)
                composition = composition.replace("Показать ещё", "")
                composition = f"Состав: {composition}"
                if len(composition) < 1:
                    composition = ""
            else:
                composition = ""

            size = bs_object.find_all(name="div", class_="product-desc-line")
            if len(size) > 1:
                size = size[1].find(name="div", class_='desc')
                if size is not None:
                    size = size.text.strip().replace(" ", "")
                    size_index = size.find("Ш")
                    size = size[:size_index] + "; " + size[size_index:]
                    size = f"Размер: {size}"
                    if len(size) < 1:
                        size = ""
                else:
                    size = ""
            else:
                size = ""

            description = bs_object.find(name="div", class_='product-describe')
            if description is not None:
                description = description.text.strip().replace('"', "")
                if len(description) < 1:
                    description = ""
            else:
                description = ""

            description = "\n".join([description, composition, size])
            seo_url = create_seo_url(title)
            result = [index, title, full_price, description, seo_url, photos]
            write_data(data=result, index=index_field)
            print(f"[INFO] Собрано и записано {index}/{len(product_ids)} товаров")
    finally:
        browser.close()
        browser.quit()


def new_excel():
    global workbook_page
    workbook_page.title = "Page 1"
    workbook_page["A1"].value = "ID"
    workbook_page["B1"].value = "Название"
    workbook_page["C1"].value = "Цена"
    workbook_page["D1"].value = "Описание"
    workbook_page["E1"].value = "SEO URL"
    workbook_page["F1"].value = "Фотографии"


def parsing():
    print("[INFO] Программа запущена")
    print("[INFO] Идет сбор информации из магазина Flower Bag. Это займет не более 20 минут")
    start_time = time.time()
    product_ids = get_product_ids()
    get_data(product_ids=product_ids)
    print("[INFO] Сбор информации закончен")
    stop_time = time.time()
    total_time = stop_time - start_time
    print("[INFO] Программа завершена")
    print("[INFO] Результат работы парсера хранится в файле result.csv в той же папке, где хранится программа")
    print(f"[INFO] На работу программы ушло {total_time} секунд")


def create_photo_dir():
    if os.path.isdir("photos"):
        shutil.rmtree("photos")
        os.mkdir("photos")
    else:
        os.mkdir("photos")
    print("[INFO] Папка для фотографий photos успешно создана")


def main():
    global workbook
    print("[INFO] Привет! Я автоматизированный помощник по сбору и записи данных")
    print("[INFO] У меня есть подробная инструкция в файле ИНСТРУКЦИЯ.txt и я очень советую сначала ее прочитать")
    print('[INFO] Чтобы включить режим парсинга, введи 1 в консоль. Для выбора режима записи, введи 2')
    while True:
        mode = input("[INPUT] Выбери режим работы: >>> ")
        if mode == "1":
            new_excel()
            create_photo_dir()
            parsing()
            workbook.save("result.xlsx")
            break
        elif mode == "2":
            input_data_to_table()
            break
        else:
            print("[ERROR] Прости, похоже ты ввел не ту команду. Попробуй еще раз")


if __name__ == "__main__":
    main()
