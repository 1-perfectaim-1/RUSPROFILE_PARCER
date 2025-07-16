import time
import csv
import os
import re
import math
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ИСПРАВЛЕНИЕ: Создаем точный "словарь-переводчик" с названий на сайте в названия колонок
KEY_MAP = {
    'Генеральный директор': 'Должность',
    'Директор': 'Должность',
    'Управляющая организация': 'Должность',
    'Исполняющий обязанности генерального директора': 'Должность',
    'Президент': 'Должность',
    'Временно исполняющий обязанности генерального директора': 'Должность',
    'И.О.генерального директора': 'Должность',
    'ИНН': 'ИНН',
    'ОГРН': 'ОГРН',
    'Дата регистрации': 'Дата регистрации',
    'Уставный капитал': 'Уставный капитал',
    'Выручка': 'Выручка',
    'Основной вид деятельности': 'Основной вид деятельности'
}

def parse_company_data(company_element, fieldnames):
    """
    ФИНАЛЬНАЯ ВЕРСИЯ.
    Извлекает всю необходимую информацию, используя жесткую структуру и словарь-переводчик.
    """
    data = {key: '' for key in fieldnames}

    try:
        name_element = company_element.find_element(By.CSS_SELECTOR, ".company-item__title a")
        data['Название'] = name_element.text
        relative_link = name_element.get_attribute('href')
        if relative_link:
             data['Ссылка на Rusprofile'] = "https://www.rusprofile.ru" + relative_link
    except NoSuchElementException: pass

    try:
        data['Адрес'] = company_element.find_element(By.CSS_SELECTOR, "address.company-item__text").text
    except NoSuchElementException: pass

    details = company_element.find_elements(By.CSS_SELECTOR, ".company-item-info dl")
    for detail in details:
        try:
            key_text = detail.find_element(By.TAG_NAME, 'dt').text.strip()
            value_text = detail.find_element(By.TAG_NAME, 'dd').text.strip()

            # ИСПРАВЛЕНИЕ: Проверяем, есть ли ключ с сайта в нашем "переводчике"
            if key_text in KEY_MAP:
                column_name = KEY_MAP[key_text]
                # Особая логика для должности и руководителя
                if column_name == 'Должность':
                    data['Должность'] = key_text
                    data['Руководитель'] = value_text
                # Особая логика для выручки
                elif column_name == 'Выручка':
                    data[column_name] = value_text.split('\n')[0].strip()
                # Все остальные поля
                else:
                    data[column_name] = value_text
        except NoSuchElementException:
            continue
            
    return data

def format_time(seconds):
    if seconds < 0: seconds = 0
    mins, secs = divmod(int(seconds), 60)
    return f"{mins} мин {secs} сек"

def main():
    output_filename_csv = "rusprofile_data.csv"
    output_filename_xlsx = "rusprofile_data.xlsx"
    fieldnames = [
        'Название', 'Ссылка на Rusprofile', 'Должность', 'Руководитель', 'ИНН', 'ОГРН', 
        'Дата регистрации', 'Уставный капитал', 'Выручка', 
        'Основной вид деятельности', 'Адрес'
    ]
    
    if not os.path.exists(output_filename_csv):
        with open(output_filename_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames); writer.writeheader()
        print(f"Создан новый файл '{output_filename_csv}'")
    else:
        print(f"Найден существующий файл '{output_filename_csv}'. Новые данные будут добавлены в него.")

    chrome_options = Options(); service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)

    print("--- ШАГ 1: АВТОРИЗАЦИЯ ---")
    driver.get("https://www.rusprofile.ru/")
    input(">>> Пожалуйста, войдите в свой профиль в открывшемся окне браузера, а затем вернитесь сюда и нажмите Enter...")

    print("\n--- ШАГ 2: НАСТРОЙКА ФИЛЬТРОВ ---")
    driver.get("https://www.rusprofile.ru/search-advanced")
    input(">>> Пожалуйста, на странице поиска установите все нужные фильтры и нажмите кнопку 'Найти'. "
          "После того, как результаты поиска загрузятся на странице, вернитесь сюда и нажмите Enter...")
    
    total_pages = 0
    try:
        pager_description_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pager-holder .description")))
        pager_text = pager_description_element.text
        numbers_str = re.findall(r'\d+', pager_text.replace('\xa0', '').replace(' ', ''))
        numbers = [int(n) for n in numbers_str]
        
        if len(numbers) >= 3:
            items_per_page = numbers[1]
            total_items = numbers[2]
            total_pages = math.ceil(total_items / items_per_page)
            print(f"\n--- Найдено {total_items} организаций. Всего будет обработано ~{total_pages} страниц. ---")
        else: raise ValueError()
    except Exception:
        print(f"--- Не удалось определить общее количество страниц. ETA будет недоступен. ---")

    total_companies_found = 0
    page_number = 1
    time_per_page_history = []
    first_company_on_page_text = "" # Переменная для "Метода Сталкера"

    while True:
        start_time = time.time()
        page_info = f"Страница {page_number}"
        if total_pages > 0: page_info += f" из {total_pages}"
        print(f"\n--- Парсинг | {page_info} ---")
        
        page_data = []
        try:
            wait.until(EC.presence_of_element_located((By.ID, "additional-results")))
            wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")) > 0)
            time.sleep(1) # Даем прогрузиться всем элементам

            # ИСПРАВЛЕНИЕ: "Метод Сталкера" - проверяем, не застряли ли мы
            current_cards_check = driver.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")
            if current_cards_check:
                new_first_company_text = current_cards_check[0].text
                if new_first_company_text == first_company_on_page_text:
                    print("\n[!!!] ВНИМАНИЕ: Содержимое страницы не изменилось после клика 'Далее'. Скрипт остановлен.")
                    user_decision = input(">>> Проверьте браузер (CAPTCHA?). Попробуйте перейти на следующую страницу вручную.\n"
                                          ">>> Нажмите Enter, чтобы продолжить, или введите 'q' и Enter, чтобы выйти: ")
                    if user_decision.lower() == 'q': break
                    else: 
                        first_company_on_page_text = "" # Сбрасываем "сталкера", чтобы он обновился на след. итерации
                        print("Пробую продолжить работу...")
                        continue
                first_company_on_page_text = new_first_company_text

            num_companies_on_page = len(current_cards_check)
            print(f"Найдено {num_companies_on_page} компаний. Начинаю сбор...")

            for i in range(num_companies_on_page):
                # ... (вложенный цикл с попытками, который хорошо себя показал)
                for attempt in range(2):
                    try:
                        all_cards = driver.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")
                        company_data = parse_company_data(all_cards[i], fieldnames)
                        page_data.append(company_data)
                        break
                    except (StaleElementReferenceException, IndexError):
                        time.sleep(1)
            
            if page_data:
                with open(output_filename_csv, 'a', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writerows(page_data)
                total_companies_found += len(page_data)
                print(f"Сохранено {len(page_data)} записей. Всего в файле: {total_companies_found}.")

        except TimeoutException:
            print("Не удалось дождаться загрузки списка компаний. Завершение работы.")
            break

        end_time = time.time()
        # ... (расчет и вывод ETA, который работал хорошо)
        elapsed_time = end_time - start_time
        time_per_page_history.append(elapsed_time)
        average_time = sum(time_per_page_history) / len(time_per_page_history)
        if total_pages > 0 and page_number < total_pages:
            pages_left = total_pages - page_number
            eta = average_time * pages_left
            print(f"Затрачено: {format_time(elapsed_time)}. | ETA: {format_time(eta)}")
        else:
            print(f"Затрачено: {format_time(elapsed_time)}.")

        try:
            next_button = driver.find_element(By.CSS_SELECTOR, ".paging-list .nav-next")
            if "disabled" in next_button.get_attribute("class"):
                print("\nКнопка 'Далее' неактивна. Завершаю сбор.")
                break
            driver.execute_script("arguments[0].click();", next_button)
            page_number += 1
        except Exception as e:
            print(f"Не удалось нажать 'Далее' ({type(e).__name__}). Завершаю сбор.")
            break
            
    driver.quit()

    if total_companies_found > 0:
        print(f"\n--- Сбор данных завершен. ---")
        try:
            print("Конвертирую в .xlsx с правильным форматом...")
            df = pd.read_csv(output_filename_csv, dtype=str)
            df.to_excel(output_filename_xlsx, index=False, engine='openpyxl')
            print(f"Создан файл Excel: {output_filename_xlsx}")
            print(f"Исходные данные также сохранены в файле: {output_filename_csv}")
        except Exception as e:
            print(f"Не удалось создать .xlsx файл. Ошибка: {e}. Данные сохранены в {output_filename_csv}")
    else:
        print("\nНе было собрано ни одной записи.")

if __name__ == "__main__":
    main()