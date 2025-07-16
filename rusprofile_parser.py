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

def parse_company_data(company_element):
    """
    Извлекает всю необходимую информацию из одного блока компании (div.company-item).
    """
    data = {}
    HEAD_PERSON_KEYS = [
        'Генеральный директор', 'Директор', 'Управляющая организация',
        'Исполняющий обязанности генерального директора', 'Президент',
        'Временно исполняющий обязанности генерального директора',
        'И.О.генерального директора'
    ]

    try:
        name_element = company_element.find_element(By.CSS_SELECTOR, ".company-item__title a")
        data['Название'] = name_element.text
        relative_link = name_element.get_attribute('href')
        if relative_link:
             data['Ссылка на Rusprofile'] = "https://www.rusprofile.ru" + relative_link
        else:
            data['Ссылка на Rusprofile'] = 'Не найдено'
    except NoSuchElementException:
        data['Название'] = 'Не найдено'
        data['Ссылка на Rusprofile'] = 'Не найдено'

    try:
        data['Адрес'] = company_element.find_element(By.CSS_SELECTOR, "address.company-item__text").text
    except NoSuchElementException:
        data['Адрес'] = 'Не найден'

    details = company_element.find_elements(By.CSS_SELECTOR, ".company-item-info dl")
    for detail in details:
        try:
            key = detail.find_element(By.TAG_NAME, 'dt').text.strip()
            value_element = detail.find_element(By.TAG_NAME, 'dd')
            value = value_element.text.strip()
            
            if key in HEAD_PERSON_KEYS:
                data['Должность'] = key
                data['Руководитель'] = value
            elif key == 'Выручка':
                data[key] = value.split('\n')[0].strip()
            else:
                data[key] = value
        except NoSuchElementException:
            continue
            
    return data

def format_time(seconds):
    """Форматирует секунды в минуты и секунды для красивого вывода."""
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
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            print(f"Создан новый файл '{output_filename_csv}'")
    else:
        print(f"Найден существующий файл '{output_filename_csv}'. Новые данные будут добавлены в него.")

    chrome_options = Options()
    service = Service(ChromeDriverManager().install())
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
        else: raise ValueError("Не удалось распарсить информацию о страницах")
    except Exception as e:
        print(f"--- Не удалось определить общее количество страниц. ETA будет недоступен. Ошибка: {e} ---")

    total_companies_found = 0
    page_number = 1
    time_per_page_history = []

    while True:
        start_time = time.time()
        page_info = f"Страница {page_number}"
        if total_pages > 0: page_info += f" из {total_pages}"
        print(f"\n--- Парсинг | {page_info} ---")
        
        # --- НОВАЯ СТАБИЛЬНАЯ ЛОГИКА ПАРСИНГА СТРАНИЦЫ ---
        page_data = []
        try:
            wait.until(EC.presence_of_element_located((By.ID, "additional-results")))
            # Ожидаем, что количество элементов станет больше 0
            wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")) > 0)
            
            num_companies_on_page = len(driver.find_elements(By.CSS_SELECTOR, "#additional-results .company-item"))
            print(f"Найдено {num_companies_on_page} компаний. Начинаю сбор...")

            for i in range(num_companies_on_page):
                parsed = False
                for attempt in range(3): # 3 попытки на парсинг одного элемента
                    try:
                        # Каждый раз находим список заново
                        all_cards = driver.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")
                        current_card = all_cards[i] # Берем нужный по индексу
                        
                        company_data = parse_company_data(current_card)
                        page_data.append(company_data)
                        parsed = True
                        break # Если успешно, выходим из цикла попыток
                    except (StaleElementReferenceException, IndexError):
                        print(f"  [!] Ошибка Stale/Index для элемента #{i+1}, попытка {attempt + 2}...")
                        time.sleep(1) # Ждем и пробуем снова
                if not parsed:
                    print(f"  [X] Не удалось обработать элемент #{i+1} после нескольких попыток. Пропускаю.")
            
            if page_data:
                with open(output_filename_csv, 'a', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
                    writer.writerows(page_data)
                total_companies_found += len(page_data)
                print(f"Сохранено {len(page_data)} записей. Всего в файле: {total_companies_found}.")

        except TimeoutException:
            print("Не удалось дождаться загрузки списка компаний. Завершение работы.")
            break
        # --- КОНЕЦ НОВОЙ ЛОГИКИ ---

        end_time = time.time()
        elapsed_time = end_time - start_time
        time_per_page_history.append(elapsed_time)
        average_time = sum(time_per_page_history) / len(time_per_page_history)
        
        if total_pages > 0 and page_number < total_pages:
            pages_left = total_pages - page_number
            eta = average_time * pages_left
            print(f"Затрачено: {format_time(elapsed_time)}. | ETA: {format_time(eta)}")
        else:
            print(f"Затрачено: {format_time(elapsed_time)}.")

        MAX_RETRIES = 3
        clicked_successfully = False
        for attempt in range(MAX_RETRIES):
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, ".paging-list .nav-next")
                if "disabled" in next_button.get_attribute("class"):
                    clicked_successfully = False; break
                
                driver.execute_script("arguments[0].click();", next_button)
                page_number += 1
                clicked_successfully = True
                break
            except (NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException) as e:
                print(f"  [Попытка {attempt + 1}/{MAX_RETRIES}] Не удалось нажать 'Далее' ({type(e).__name__}). Жду 5 секунд...")
                time.sleep(5)
        
        if not clicked_successfully:
            try:
                 if "disabled" in driver.find_element(By.CSS_SELECTOR, ".paging-list .nav-next").get_attribute("class"):
                     break 
            except NoSuchElementException: pass

            print("\n!!! КРИТИЧЕСКАЯ ОШИБКА: Не удалось перейти на следующую страницу.")
            user_decision = input(">>> Проверьте браузер (CAPTCHA?). Попробуйте перейти на следующую страницу вручную.\n"
                                  ">>> Нажмите Enter, чтобы продолжить, или введите 'q' и Enter, чтобы выйти: ")
            if user_decision.lower() == 'q': break
            else: print("Пробую продолжить работу..."); continue 
    
    driver.quit()

    if total_companies_found > 0:
        print(f"\n--- Сбор данных завершен. ---")
        try:
            print("Конвертирую в .xlsx с правильным форматом...")
            # ИСПРАВЛЕНО: Читаем CSV, принудительно считая ИНН/ОГРН текстом
            df = pd.read_csv(output_filename_csv, dtype={'ИНН': str, 'ОГРН': str})
            df.to_excel(output_filename_xlsx, index=False, engine='openpyxl')
            print(f"Создан файл Excel: {output_filename_xlsx}")
            # ИСПРАВЛЕНО: CSV файл больше не удаляется
            print(f"Итоговые данные также сохранены в файле: {output_filename_csv}")
        except Exception as e:
            print(f"Не удалось создать .xlsx файл. Ошибка: {e}. Данные сохранены в {output_filename_csv}")
    else:
        print("\nНе было собрано ни одной записи.")

if __name__ == "__main__":
    main()