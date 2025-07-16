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

def parse_company_data(company_element, fieldnames):
    """
    ИСПРАВЛЕННАЯ ФУНКЦИЯ
    Извлекает всю необходимую информацию, используя жесткую структуру данных для надежности.
    """
    # ШАГ 1: Создаем пустой "шаблон" со всеми нужными колонками
    data = {key: '' for key in fieldnames}

    # ШАГ 2: Заполняем его, находя каждый элемент целенаправленно
    try:
        name_element = company_element.find_element(By.CSS_SELECTOR, ".company-item__title a")
        data['Название'] = name_element.text
        relative_link = name_element.get_attribute('href')
        if relative_link:
             data['Ссылка на Rusprofile'] = "https://www.rusprofile.ru" + relative_link
    except NoSuchElementException:
        pass # Если не нашли, значение в data останется пустым

    try:
        data['Адрес'] = company_element.find_element(By.CSS_SELECTOR, "address.company-item__text").text
    except NoSuchElementException:
        pass

    # Ищем все элементы <dl> внутри карточки
    details = company_element.find_elements(By.CSS_SELECTOR, ".company-item-info dl")
    for detail in details:
        try:
            key = detail.find_element(By.TAG_NAME, 'dt').text.strip()
            value = detail.find_element(By.TAG_NAME, 'dd').text.strip()

            # ШАГ 3: Кладем найденное значение в нужную ячейку "шаблона"
            if 'директор' in key.lower() or 'управляющая' in key.lower() or 'президент' in key.lower():
                data['Должность'] = key
                data['Руководитель'] = value
            elif key == 'ИНН':
                data['ИНН'] = value
            elif key == 'ОГРН':
                data['ОГРН'] = value
            elif key == 'Дата регистрации':
                data['Дата регистрации'] = value
            elif key == 'Уставный капитал':
                data['Уставный капитал'] = value
            elif key == 'Выручка':
                data['Выручка'] = value.split('\n')[0].strip() # Убираем данные о % роста
            elif key == 'Основной вид деятельности':
                data['Основной вид деятельности'] = value
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
    # Этот список теперь - единый источник истины для структуры данных
    fieldnames = [
        'Название', 'Ссылка на Rusprofile', 'Должность', 'Руководитель', 'ИНН', 'ОГРН', 
        'Дата регистрации', 'Уставный капитал', 'Выручка', 
        'Основной вид деятельности', 'Адрес'
    ]
    
    if not os.path.exists(output_filename_csv):
        with open(output_filename_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
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
        
        page_data = []
        try:
            wait.until(EC.presence_of_element_located((By.ID, "additional-results")))
            wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")) > 0)
            
            num_companies_on_page = len(driver.find_elements(By.CSS_SELECTOR, "#additional-results .company-item"))
            print(f"Найдено {num_companies_on_page} компаний. Начинаю сбор...")

            for i in range(num_companies_on_page):
                parsed = False
                for attempt in range(3):
                    try:
                        all_cards = driver.find_elements(By.CSS_SELECTOR, "#additional-results .company-item")
                        current_card = all_cards[i]
                        
                        # Передаем fieldnames для создания шаблона
                        company_data = parse_company_data(current_card, fieldnames)
                        page_data.append(company_data)
                        parsed = True
                        break 
                    except (StaleElementReferenceException, IndexError):
                        print(f"  [!] Ошибка Stale/Index для элемента #{i+1}, попытка {attempt + 2}...")
                        time.sleep(1)
                if not parsed:
                    print(f"  [X] Не удалось обработать элемент #{i+1} после нескольких попыток. Пропускаю.")
            
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
                     print("\nКнопка 'Далее' неактивна. Это была последняя страница.")
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
            df = pd.read_csv(output_filename_csv, dtype=str) # Читаем ВСЕ как текст
            df.to_excel(output_filename_xlsx, index=False, engine='openpyxl')
            print(f"Создан файл Excel: {output_filename_xlsx}")
            print(f"Исходные данные также сохранены в файле: {output_filename_csv}")
        except Exception as e:
            print(f"Не удалось создать .xlsx файл. Ошибка: {e}. Данные сохранены в {output_filename_csv}")
    else:
        print("\nНе было собрано ни одной записи.")

if __name__ == "__main__":
    main()