import pandas as pd
import functools
from utils import set_manufacturer_for_organization, combine_dataframes

@functools.lru_cache(maxsize=128)
def get_prices(organization):
    """Повертає ціни в залежності від вибраного представника."""
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

def process_pharmaxi_excel(file_path, organization):
    """Обробляє Excel-файл з даними, видаляє зайві колонки та рядки, додає цінову інформацію."""
    
    # Завантажуємо Excel-файл, пропускаючи перші 6 рядків, використовуючи перший рядок після них як заголовки
    df = pd.read_excel(file_path, skiprows=6, header=0)

    # Перетворюємо назви стовпців у рядки для уникнення помилок
    df.columns = df.columns.astype(str)
    
    # Видаляємо колонки з назвою "Unnamed"
    df = df.loc[:, ~df.columns.str.contains('Unnamed')]
    
    # Видаляємо рядки, де значення у стовпці "ИНН" відсутні
    df = df.dropna(subset=['ИНН'])
    
    # Видаляємо всі повністю порожні колонки
    df = df.dropna(axis=1, how="all")
    
    # Видаляємо всі повністю порожні рядки
    df = df.dropna(axis=0, how="all")
    
    # Видаляємо порожні рядки на основі важливих колонок "Номенклатура" та "ИНН"
    df = df.dropna(subset=['Номенклатура', 'ИНН'])

    # Перевіряємо, чи значення у "ИНН" є числом, і видаляємо ті, що не є
    df = df[pd.to_numeric(df['ИНН'], errors='coerce').notna()]
    
    # Обробка стовпця "Дата", якщо він є у файлі
    if 'Дата' in df.columns:
        df['Дата'] = df['Дата'].str.extract(r'(\d{2}\.\d{2}\.\d{4})')
        df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', dayfirst=True, errors='coerce')

    # Перетворення значень у стовпці "Количество" у числовий формат та округлення
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce')
    df['Количество'] = df['Количество'].round()
    df = df[df['Количество'].notna()]
    
    # Отримання цін для вибраної організації
    prices = get_prices(organization)
    if 'Цена' not in df.columns:
        df['Цена'] = None
    if prices:
        df['Цена'] = df['Номенклатура'].map(prices).fillna(1)
        
    # Встановлюємо виробника для організації
    df = set_manufacturer_for_organization(df, organization)    

    # Якщо організація Stada, встановлюємо значення "Сумма" як 1 для всіх рядків
    if organization == 'Stada':
        df['Сумма'] = 1
    else:
        df['Сумма'] = df['Количество'].astype(float) * df['Цена']
     
    # Додаємо унікальний номер для кожного рядка
    df['Номер'] = range(1, len(df) + 1)
    
    # Переміщуємо колонку "Номер" на перше місце
    columns = ['Номер'] + [col for col in df.columns if col != 'Номер']
    df = df[columns]    
        
    # Об'єднуємо дані у фінальний датафрейм
    combined_df = combine_dataframes([df])    
    
    return df
