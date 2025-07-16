import time
import csv
import os
import re
import math
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Ваш рабочий словарь-переводчик
KEY_MAP = {
    'Генеральный директор': 'Должность', 'Директор': 'Должность',
    'Управляющая организация': 'Должность', 'Президент': 'Должность',
    'Исполняющий обязанности генерального директора': 'Должность',
    'Временно исполняющий обязанности генерального директора': 'Должность',
    'И.О.генерального директора': 'Должность', 'Управляющий': 'Должность',
    'Генральный директор': 'Должность',
    'ИНН': 'ИНН', 'ОГРН': 'ОГРН', 'Дата регистрации': 'Дата регистрации',
    'Уставный капитал': 'Уставный капитал', 'Выручка': 'Выручка',
    'Основной вид деятельности': 'Основной вид деятельности'
}

# Ваша рабочая функция парсинга
def parse_company_data(company_element: WebElement, fieldnames: list) -> dict:
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
            
            if key_text in KEY_MAP:
                column_name = KEY_MAP[key_text]
                if column_name == 'Должность':
                    data['Должность'] = key_text
                    data['Руководитель'] = value_text
                elif column_name == 'Выручка':
                    data[column_name] = value_text.split('\n')[0].strip()
                else:
                    data[column_name] = value_text
        except NoSuchElementException:
            continue
    return data

# Ваша рабочая функция календаря
def set_dates_and_search_js(driver: webdriver.Chrome, wait: WebDriverWait, start_date_str: str, end_date_str: str):
    """Напрямую вписывает даты и инициирует поиск с помощью JS-событий."""
    try:
        js_script = f"""
            var start_input = document.getElementById('date-reg-from');
            var end_input = document.getElementById('date-reg-to');
            if (!start_input || !end_input) {{ return false; }}
            
            start_input.value = '{start_date_str}';
            end_input.value = '{end_date_str}';
            
            var event = new Event('change', {{ bubbles: true }});
            end_input.dispatchEvent(event);
            return true;
        """
        driver.execute_script(js_script)
        
        time.sleep(1)

    except TimeoutException:
        print("  [i] Индикатор загрузки не появился, вероятно, для этого периода нет результатов или они загрузились мгновенно.")
    except Exception as e:
        print(f"  [X] Не удалось автоматически установить фильтр по дате. Перезагружаю страницу.")
        driver.get(driver.current_url)
        time.sleep(3)
        raise e

def main():
    output_filename_csv = "rusprofile_data.csv"
    output_filename_xlsx = "rusprofile_data.xlsx"
    fieldnames = list(KEY_MAP.values()) + ['Название', 'Ссылка на Rusprofile', 'Адрес', 'Руководитель']
    fieldnames = sorted(list(set(fieldnames)), key=lambda x: (
        'Название' not in x, 'Ссылка' not in x, 'Должность' not in x, 'Руководитель' not in x, 'ИНН' not in x, 'ОГРН' not in x, x
    ))

    if not os.path.exists(output_filename_csv):
        with open(output_filename_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames); writer.writeheader()
        print(f"Создан новый файл '{output_filename_csv}'")
    else:
        print(f"Найден существующий файл '{output_filename_csv}'. Данные будут добавлены в него.")
    
    chrome_options = Options()
    profile_path = os.path.join(os.getcwd(), "chrome_profile")
    chrome_options.add_argument(f"user-data-dir={profile_path}")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)

    print("--- ШАГ 1: НАСТРОЙКА ПРОФИЛЯ ---")
    driver.get("https://www.rusprofile.ru/")
    input(">>> Если вы запускаете скрипт первый раз, войдите в свой профиль. Нажмите Enter для продолжения...")

    print("\n--- ШАГ 2: НАСТРОЙКА ОСНОВНЫХ ФИЛЬТРОВ ---")
    driver.get("https://www.rusprofile.ru/search-advanced")
    input(">>> Пожалуйста, установите ВСЕ нужные фильтры (ОКОПФ и т.д.), КРОМЕ ДАТЫ. Нажмите Enter...")
    
    total_companies_collected = 0
    start_year = 2023
    end_year = 1991
    
    for year in range(start_year, end_year - 1, -1):
        for month_start in [1, 4, 7, 10]:
            month_end = month_start + 2
            
            start_date = datetime.date(year, month_start, 1)
            end_day = (datetime.date(year, month_end + 1, 1) - datetime.timedelta(days=1)).day if month_end < 12 else 31
            end_date = datetime.date(year, month_end, end_day)
            
            start_date_str = start_date.strftime('%d.%m.%Y')
            end_date_str = end_date.strftime('%d.%m.%Y')

            print("\n" + "="*50)
            print(f"--- Начинаем сбор за период: {start_date_str} - {end_date_str} ---")
            
            try:
                set_dates_and_search_js(driver, wait, start_date_str, end_date_str)
            except Exception as e:
                print(f"Критическая ошибка при установке фильтра. Пропускаю период. Детали: {type(e).__name__}")
                continue
            
            page_number = 1
            
            while True:
                print(f"--- {year} Q{month_start//3+1} | Парсинг страницы {page_number} ---")
                try:
                    all_cards = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#additional-results .company-item")))
                    
                    # ИСПРАВЛЕНИЕ: Запоминаем ID первой компании
                    first_company_id = all_cards[0].find_element(By.CSS_SELECTOR, "a").get_attribute('href')

                    page_data = [parse_company_data(card, fieldnames) for card in all_cards]
                    
                    if page_data:
                        with open(output_filename_csv, 'a', newline='', encoding='utf-8-sig') as f:
                            writer = csv.DictWriter(f, fieldnames=fieldnames); writer.writerows(page_data)
                        total_companies_collected += len(page_data)
                        print(f"Сохранено {len(page_data)} записей. Всего в файле: {total_companies_collected}.")

                except TimeoutException:
                    print("Компаний для этого периода/страницы не найдено."); break

                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, ".paging-list .nav-next:not(.disabled)")
                    driver.execute_script("arguments[0].click();", next_button)
                    page_number += 1
                    
                    # ИСПРАВЛЕНИЕ: Ждем, пока ID первой компании не изменится
                    WebDriverWait(driver, 20).until(
                        lambda d: d.find_element(By.CSS_SELECTOR, "#additional-results .company-item:first-child a").get_attribute('href') != first_company_id
                    )
                except TimeoutException:
                    print("Страница не обновилась (лимит 20 страниц) или это последняя страница. Завершаю сбор для этого периода."); break
                except (NoSuchElementException, StaleElementReferenceException):
                    print("Кнопка 'Далее' не найдена. Завершаю сбор для этого периода."); break
    
    driver.quit()
    print("\n" + "="*50 + "\n--- СБОР ДАННЫХ ПОЛНОСТЬЮ ЗАВЕРШЕН ---\n" + "="*50)
    if total_companies_collected > 0:
        try:
            print("Конвертирую в .xlsx с правильным форматом...")
            df = pd.read_csv(output_filename_csv, dtype=str)
            df.to_excel(output_filename_xlsx, index=False, engine='openpyxl')
            print(f"Создан файл Excel: {output_filename_xlsx}")
            print(f"Исходные данные также сохранены в файле: {output_filename_csv}")
        except Exception as e:
            print(f"Не удалось создать .xlsx файл. Ошибка: {e}. Данные сохранены в {output_filename_csv}")
    else:
        print("Не было собрано ни одной записи.")

if __name__ == "__main__":
    main()