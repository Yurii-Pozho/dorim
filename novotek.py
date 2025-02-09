import pandas as pd
import functools
import json
from utils import set_manufacturer_for_organization, combine_dataframes

@functools.lru_cache(maxsize=128)
def get_prices(organization):
    """Повертає словник із цінами залежно від вибраної організації."""
    prices_dict = {
        'Arterium': {  # Ціни для організації Arterium
            "Тиоцетам 10мл №10": 94077,
            "Тиоцетам 5мл №10": 69196,
            "Тиоцетам 400мг/100мг №30 табл.": 50853,
            "Уролесан 180мл": 58823,
            "Уролесан 25мл": 55407,
            "Элкоцин 100мг №30 табл.": 72611,
            "Ларитилен №20 табл. мяты": 23403,
            "Ларитилен №20 табл. мяты и малины": 23403,
            "Ларитилен №20 табл.мяты и лимон": 23403,
            "Тиоцетам р-р 10мл №10": 94077
        },
        'Stada': {  # Заглушка для Stada (всі ціни = 1)
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
        'Без организации': {}  # Якщо немає організації, ціни відсутні
    }
    return prices_dict.get(organization, {})

def process_novotek_excel(file_path, organization):
    """Обробляє Excel-файл та повертає DataFrame з потрібними змінами."""
    
    # Читаємо файл, пропускаючи перші 4 рядки (заголовок у 5-му рядку)
    df = pd.read_excel(file_path, skiprows=4, header=0)
    
    # Витягуємо з колонки "Регистратор" номер документа, дату та час
    df[['Номер документа', 'Дата', 'Время']] = df['Регистратор'].str.extract(r'(\d{6,11})\s+от\s+(\d{2}\.\d{2}\.\d{4})\s+(\d+:\d+:\d+)')
    
    # Видаляємо колонку "Регистратор", оскільки вона вже не потрібна
    df = df.drop('Регистратор', axis=1)
    
    # Видаляємо рядки, де відсутня "Номенклатура"
    df = df.dropna(subset=['Номенклатура'])
    
    # Видаляємо колонки, де всі значення NaN
    df = df.dropna(axis=1, how='all')
    
    # Конвертуємо "Дата" у формат datetime
    df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', dayfirst=True, errors='coerce')
    
    # Конвертуємо "Количество" у числовий формат та округлюємо значення
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').round()
    
    # Видаляємо рядки, де "Количество" відсутнє
    df = df[df['Количество'].notna()]
    
    # Отримуємо ціни для відповідної організації
    prices = get_prices(organization)
    
    # Додаємо колонку "Цена", якщо її немає, або залишаємо стару
    df['Цена'] = df['Номенклатура'].map(prices).fillna(1) if 'Цена' not in df.columns else df['Цена']
    
    # Встановлюємо виробника для організації
    df = set_manufacturer_for_organization(df, organization)
    
    # Розрахунок суми залежно від організації
    if organization == 'Stada':
        df['Сумма'] = 1  # Для Stada всі значення = 1
    elif organization == 'Arterium':
        df['Сумма'] = df.apply(lambda row: row['Количество'] * row['Цена'] if row['Номенклатура'] in prices else 1, axis=1)
    else:
        df['Сумма'] = df['Количество'].astype(float) * df['Цена']
    
    # Якщо відсутня колонка "Производитель", додаємо її
    if 'Производитель' not in df.columns:
        df['Производитель'] = 'Без организации'
    
    # Додаємо колонку "Номер" для нумерації рядків
    df['Номер'] = range(1, len(df) + 1)
    
    # Переміщуємо колонку "Номер" на перше місце
    columns = ['Номер'] + [col for col in df.columns if col != 'Номер']
    df = df[columns]
    
    def clean_inn(inn_value):
        """Очищує значення ІПН (видаляє зайві символи та перетворює в int)."""
        if pd.isna(inn_value):
            return 0
        cleaned_inn = str(inn_value).replace(',', '').replace('.', '').split(r'[-–]')[0].strip()
        return int(cleaned_inn)
    
    # Перевіряємо, чи є колонка "ИНН", і очищуємо її
    if 'ИНН' in df.columns:
        df['ИНН'] = df['ИНН'].apply(clean_inn)
    else:
        # Якщо "ИНН" немає, пробуємо визначити ІПН за назвою контрагента
        if 'Контрагент' not in df.columns:
            raise ValueError("Не знайдено колонок 'ИНН' або 'Контрагент'")
        
        # Завантажуємо JSON-файл із відповідністю контрагентів та їхніх ІПН
        try:
            with open('_INN.json', 'r', encoding='utf-8') as f:
                inn_data = json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError("Файл _INN.json не знайдено")
        
        def get_inn_from_json(client_name):
            """Отримує ІПН з JSON за назвою контрагента або повертає дефолтне значення."""
            if pd.isna(client_name):
                return 111111111
            cleaned_name = (client_name.strip()
                .replace('"', '')
                .replace('`', '')
                .replace('mas`uliyati cheklangan jamiyati', '')
                .strip()
            )
            return inn_data.get(cleaned_name, 111111111)
        
        # Створюємо колонку "ИНН" на основі контрагента
        df['ИНН'] = df['Контрагент'].apply(get_inn_from_json)

    # Видаляємо рядки, які містять слово "итого" (сума по всьому документу)
    df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains("итого").any(), axis=1)] 
    
    # Об'єднуємо DataFrame
    combined_df = combine_dataframes([df])
    
    return df
