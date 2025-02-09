import pandas as pd
import functools
from utils import set_manufacturer_for_organization, combine_dataframes
import re
import numpy as np

@functools.lru_cache(maxsize=128)
def get_prices(organization):
    """Повертає словник цін у залежності від вибраного представника."""
    prices_dict = {
        'Stada': {
            "Тиоцетам 10мл №10": 1,
            "Тиоцетам 5мл №10": 1,
            "Тиоцетам 400мг/100мг №30 табл.": 1,
            "Уролесан 180мл": 1,
            "Уролесан 25мл": 1,
            "Элкоцин 100мг №30 табл.": 1,
            "Ларитилен №20 табл. мяты": 1,
            "Ларитилен №20 табл. мяты и малины": 1,
            "Ларитилен №20 табл.мяты и лимон": 1
        },
        'Без организации': {}
    }
    return prices_dict.get(organization, {})

def process_pharma_choice_excel(file_path, organization):
    """Обробляє завантажений Excel-файл та повертає оброблений DataFrame."""
    excel_file = pd.ExcelFile(file_path)

    # Визначаємо, які листи необхідно обробляти
    sheets_to_keep = ['продажи', 'Отчет по реализации']
    available_sheets = [sheet for sheet in excel_file.sheet_names if sheet in sheets_to_keep]

    if not available_sheets:
        raise ValueError(f"Не знайдено потрібних листів у файлі: {file_path}")

    final_df = pd.DataFrame()  # Порожній DataFrame для об'єднання всіх листів

    for sheet in available_sheets:
        # Читаємо перші 10 рядків для пошуку заголовків
        preview = pd.read_excel(excel_file, sheet_name=sheet, nrows=10, header=None)

        # Пошук рядка, що містить заголовки
        header_row = None
        for i, row in preview.iterrows():
            if 'Наименование' in row.values:
                header_row = i
                break

        if header_row is not None:
            # Завантажуємо дані з визначеним рядком заголовків
            df = pd.read_excel(excel_file, sheet_name=sheet, skiprows=header_row, header=0)

            # Перейменовуємо колонки для листа "Отчет по реализации"
            if sheet == 'Отчет по реализации':
                column_rename_map = {
                    "Откуда продано": "Филиал",
                    "Дата продажи": "Дата",
                    "№ документа": "№",
                    "Кол-во": "Количество",
                    "Город/Район": "Регион"
                }
                df.rename(columns={key: value for key, value in column_rename_map.items() if key in df.columns}, inplace=True)

            df.columns = df.columns.str.strip()  # Видаляємо зайві пробіли в заголовках колонок

            # Обробка колонки "ИНН"
            if 'ИНН' in df.columns:
                df['ИНН'] = df['ИНН'].apply(lambda x: re.sub(r'^\s+|\s+$', '', str(x)))
                
                def clean_inn(x):
                    """Очищує значення ІНН, видаляючи дробову частину."""
                    try:
                        return str(int(float(x)))
                    except (ValueError, TypeError):
                        return str(x).strip()
                
                df['ИНН'] = df['ИНН'].apply(clean_inn)
                df['ИНН'] = df['ИНН'].replace(["", " "], np.nan)
                df = df[pd.to_numeric(df['ИНН'], errors='coerce').notna()]  # Видаляємо некоректні значення
                df = df.dropna(subset=['ИНН'])  # Видаляємо рядки з NaN у колонці "ИНН"
            else:
                raise ValueError("Колонка 'ИНН' відсутня в даних.")

            # Перетворення формату дати
            df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', dayfirst=True, errors='coerce')

            # Перетворення колонки "Количество" в числовий формат
            df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').round()
            df = df[df['Количество'].notna()]

            # Отримання цін залежно від організації
            prices = get_prices(organization)
            if 'Цена' not in df.columns:
                df['Цена'] = None
            if prices:
                df['Цена'] = df['Наименование'].map(prices).fillna(1)

            # Додавання виробника
            df = set_manufacturer_for_organization(df, organization)
            
            # Обчислення суми в залежності від організації
            if organization == 'Stada':
                df['Сумма'] = 1
            else:
                df['Сумма'] = df['Количество'].astype(float) * df['Цена']
            
            # Додаємо нумерацію рядків
            df['Номер'] = range(1, len(df) + 1)
            columns = ['Номер'] + [col for col in df.columns if col != 'Номер']
            df = df[columns]

            # Фільтрація рядків з короткими значеннями у "Наименование"
            if 'Наименование' in df.columns:
                df = df[df['Наименование'].str.len() >= 3]

                final_df = pd.concat([final_df, df], ignore_index=True)

    if final_df.empty:
        raise ValueError("Не вдалося знайти відповідні дані в жодному листі.")
    
    combined_df = combine_dataframes([final_df])  # Об'єднуємо результати у фінальний DataFrame
    
    return final_df
