from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from database import input_data_to_table
import json
import time
import csv
import re


ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
options = webdriver.ChromeOptions()
options.add_argument(f"user-agent={ua}")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--headless")
timeout = 60


def get_product_ids():
    browser = webdriver.Chrome(options=options)
    browser.set_page_load_timeout(time_to_wait=timeout)
    try:
        url = "https://flowwow.com/shop/flower-bag/"
        browser.get(url=url)
        time.sleep(2)
        read_more_buttons = browser.find_elements(By.LINK_TEXT, "Показать ещё")
        for read_more_button in read_more_buttons:
            browser.execute_script("arguments[0].click();", read_more_button)
            time.sleep(1)
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
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


def write_data(data):
    with open("result.csv", "a", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(data)


def get_data(product_ids):
    browser = webdriver.Chrome(options=options)
    browser.set_page_load_timeout(time_to_wait=timeout)
    index = 0
    try:
        for product_id in product_ids:
            print(f"[INFO] Собираем данные о товаре номер {product_id}")
            index += 1
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
            photos = "; ".join(photos)

            purchase_price = json_object["cost"]
            if "base" in json_object.keys():
                full_price = json_object["base"]
            else:
                full_price = purchase_price

            full_info = json_object["fullInfo"]
            full_info = full_info.replace("\n", "").replace("\\", "")
            bs_object = BeautifulSoup(full_info, "lxml")
            title = bs_object.find(name="div", class_='pp-title').text.strip().replace('"', "")

            composition = bs_object.find(name="div", class_="product-desc-line").find(name="div", class_='desc')
            if composition is not None:
                composition = composition.text.strip().split(".")
                composition = "; ".join(element.strip() for element in composition)
                composition = composition.replace("Показать ещё", "")
                if len(composition) < 1:
                    composition = "Нет данных на сайте"
            else:
                composition = "Нет данных на сайте"

            size = bs_object.find_all(name="div", class_="product-desc-line")
            if len(size) > 1:
                size = size[1].find(name="div", class_='desc')
                if size is not None:
                    size = size.text.strip().replace(" ", "")
                    size_index = size.find("Ш")
                    size = size[:size_index] + "; " + size[size_index:]
                    if len(size) < 1:
                        size = "Нет данных на сайте"
                else:
                    size = "Нет данных на сайте"
            else:
                size = "Нет данных на сайте"

            description = bs_object.find(name="div", class_='product-describe')
            if description is not None:
                description = description.text.strip().replace('"', "").replace(",", ";")
                if len(description) < 1:
                    description = "Нет данных на сайте"
            else:
                description = "Нет данных на сайте"

            result = [title, full_price, purchase_price, description, composition, size, photos]
            write_data(data=result)
            print(f"[INFO] Собрано и записано {index}/{len(product_ids)} товаров")
    finally:
        browser.close()
        browser.quit()


def create_csv_file():
    with open("result.csv", "w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Название", "Полная цена", "Цена со скидкой", "Описание",
                         "Состав", "Размер", "Ссылки на фотографии"])


def parsing():
    print("[INFO] Программа запущена")
    print("[INFO] Идет сбор информации из магазина Flower Bag. Это займет не более 2 минут")
    start_time = time.time()
    create_csv_file()
    product_ids = get_product_ids()
    get_data(product_ids=product_ids)
    print("[INFO] Сбор информации закончен")
    stop_time = time.time()
    total_time = stop_time - start_time
    print("[INFO] Программа завершена")
    print("[INFO] Результат работы парсера хранится в файле result.csv в той же папке, где хранится программа")
    print(f"[INFO] На работу программы ушло {total_time} секунд")


def main():
    print("[INFO] Привет! Я автоматизированный помощник по сбору и записи данных")
    print("[INFO] У меня есть подробная инструкция в файле ИНСТРУКЦИЯ.txt и я очень советую сначала ее прочитать")
    print('[INFO] Чтобы включить режим парсинга, введи 1 в консоль. Для выбора режима записи, введи 2')
    while True:
        mode = input("[INPUT] Выбери режим работы: >>> ")
        if mode == "1":
            parsing()
            break
        elif mode == "2":
            input_data_to_table()
            break
        else:
            print("Прости, похоже ты ввел не ту команду. Попробуй еще раз")


if __name__ == "__main__":
    main()
