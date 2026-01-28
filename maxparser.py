import os
import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

BASE_URL = "https://maximum.md/ru/kompyuternaya-tehnika/monitory/monitory/"
IMAGES_DIR = "images"
os.makedirs(IMAGES_DIR, exist_ok=True)

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def download_image(url, filename):
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            with open(filename, "wb") as f:
                f.write(resp.content)
            return True
    except Exception as e:
        print(f"Ошибка при скачивании {url}: {e}")
    return False

def parse_page(driver):
    wait = WebDriverWait(driver, 15)  # увеличенный таймаут
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".js-content.product__item")))
    cards = driver.find_elements(By.CSS_SELECTOR, ".js-content.product__item")
    print(f"Найдено карточек: {len(cards)}")

    products = []
    for idx, card in enumerate(cards, 1):
        try:
            # Пропускаем "Нет в наличии"
            try:
                not_in_stock = card.find_element(By.CSS_SELECTOR, ".not_in_shops")
                if not_in_stock.is_displayed():
                    continue
            except:
                pass

            # Фото
            img_elem = card.find_element(By.CSS_SELECTOR, ".product__item__image img")
            img_url = img_elem.get_attribute("data-src") or img_elem.get_attribute("src")
            img_name = f"{IMAGES_DIR}/item_{int(time.time()*1000)}_{idx}.jpg"
            download_image(img_url, img_name)

            # Название и ссылка
            title_elem = card.find_element(By.CSS_SELECTOR, ".product__item__title a")
            title = title_elem.text.strip()
            url = title_elem.get_attribute("href")

            # Цена текущая
            price_elem = card.find_element(By.CSS_SELECTOR, ".product__item__price-current span")
            price = price_elem.text.strip()

            # Старая цена
            try:
                old_price = card.find_element(By.CSS_SELECTOR, ".product__item__price-old").text.strip()
            except:
                old_price = None

            products.append({
                "img": img_name,
                "title": title,
                "price": price,
                "old_price": old_price,
                "url": url
            })
        except Exception as e:
            print("Ошибка карточки:", e)
    return products

def get_next_page(driver):
    wait = WebDriverWait(driver, 15)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".js-content.product__item")))

    # desktop + mobile
    next_btns = driver.find_elements(By.CSS_SELECTOR, '.paginator-pages-list a[title="Следующая страница"]') + \
                driver.find_elements(By.CSS_SELECTOR, '.paginator-pages-list-mobile a[title="Следующая страница"]')

    for btn in next_btns:
        if btn.is_displayed() and 'следующая' in btn.text.lower():
            href = btn.get_attribute("href")
            if href:
                return href
    return None

def save_xlsx(products, filename="maximum_data.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.append(["Изображение", "Название", "Цена", "Старая цена", "Ссылка"])

    for prod in products:
        # Вставляем изображение
        img = XLImage(prod["img"])
        img.width = 100
        img.height = 100
        ws.append([None, prod["title"], prod["price"], prod["old_price"], prod["url"]])
        ws.add_image(img, f"A{ws.max_row}")

    wb.save(filename)
    print(f"[OK] Сохранено в {filename}")

def main():
    driver = setup_driver()
    all_products = []

    url = BASE_URL
    while url:
        driver.get(url)
        products = parse_page(driver)
        all_products.extend(products)

        url = get_next_page(driver)
        if url:
            print(f"Переход на следующую страницу: {url}")
            time.sleep(2)  # ждём прогрузки JS

    save_xlsx(all_products)
    driver.quit()

if __name__ == "__main__":
    main()
