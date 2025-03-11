import pandas as pd
import functools
from utils import set_manufacturer_for_organization, combine_dataframes

@functools.lru_cache(maxsize=128)
def get_prices(organization):
    """Повертає ціни в залежності від вибраного представника."""
    # Словник з цінами для кожної організації
    prices_dict = {
        'Arterium': {
            "Ларитилен №20 табл. мяты": 24430,
            "Ларитилен №20 табл. мяты и малины": 24430,
            "Ларитилен №20 табл.мяты и лимон": 24430,
            "Резистол 50мл капли оральные": 84290,
            "Резистол капли оральные 20мл": 49000,
            "Тиотриазолин Форте 50мг/мл 4мл №10 р-р д/и.": 173532,
            "Тиоцетам 5мл №10": 70275,
            "Тиоцетам 400мг/100мг №30 табл.": 52007,
            "Тиоцетам 10мл №10": 94077,
            "Элкоцин 100мг №30 табл.": 76028,
            "Формен Комби №40 капс.": 94280,
            #-------------------------------------------
            "Ларитилен №20 табл. мяты":	24430,
            "Ларитилен №20 табл. мяты и малины":	24430,
            "Ларитилен №20 табл.мяты и лимон":	24430,
            "Резистол 50мл капли оральные":	84290,
            "Резистол капли оральные 20мл":	49000,
            "Тиотриазолин Форте 50мг/мл 4мл №10 р-р д/и.":	173532,
            "Тиоцетам 10мл №10":	94077,
            "Тиоцетам 400мг/100мг №30 табл.":	52007,
            "Тиоцетам 5мл №10":	70275,
            "Формен Комби №40 капс.":	94280,
            "Элкоцин 100мг №30 табл.":	76028
        },
        'Stada': {},
        'Без организации': {}
    }
    # Повертаємо словник цін для заданої організації або порожній словник, якщо така організація відсутня
    return prices_dict.get(organization, {})

def process_asklepiy_excel(file_path, organization):
    # Зчитування всіх аркушів з Excel-файлу у словник DataFrame-ів
    df_all = pd.read_excel(file_path, sheet_name=None)
    
    # Перевірка наявності потрібного аркуша (шукаємо "Реализация " або "Реализация")
    sheet_name = next((name for name in ["Реализация ", "Реализация"] if name in df_all), None)
    if sheet_name is None:
        raise ValueError("Аркуш 'Развернутый' не знайдено")
    
    # Отримуємо DataFrame з вибраного аркуша
    df = df_all[sheet_name]
    
    # Очищення назв колонок від пробілів
    df.columns = df.columns.str.strip()
    
    # Видалення колонок, де всі значення NaN
    df = df.dropna(axis=1, how='all')
    
    # Видалення рядків, де відсутні значення в колонках 'Наименование' та 'Производитель'
    df = df.dropna(subset=['Наименование', 'Производитель'])
    
    # Перевірка наявності стовпця з ІНН (або 'ИНН', або 'ИНН клиента')
    if 'ИНН' in df.columns or 'ИНН клиента' in df.columns:
        # Визначаємо, яке ім'я стовпця є в DataFrame
        column_name = 'ИНН' if 'ИНН' in df.columns else 'ИНН клиента'
        
        # Виконуємо обробку для відповідного стовпця:
        # - перетворюємо значення в рядок,
        # - видаляємо коми,
        # - перетворюємо в числовий тип (з обробкою помилок),
        # - заповнюємо NaN нулями та конвертуємо в цілий тип з підтримкою пропущених значень.
        df[column_name] = (df[column_name].astype(str)
                                        .str.replace(',', '', regex=False)
                                        .apply(pd.to_numeric, errors='coerce')
                                        .fillna(0)
                                        .astype('Int64'))
    
    # Конвертація колонки 'Дата' в тип datetime з заданим форматом і обробкою помилок
    df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', dayfirst=True, errors='coerce')
    
    # Якщо існує стовпець 'Кол-во', перейменовуємо його в 'Количество'
    if 'Кол-во' in df.columns:
        df.rename(columns={'Кол-во': 'Количество'}, inplace=True)

    if 'Сумма по полю Кол-во' in df.columns:
        df.rename(columns={'Сумма по полю Кол-во': 'Количество'}, inplace=True)
        
    # Переконуємося, що значення в стовпці 'Количество' є числовими, та округлюємо їх
    if 'Количество' in df.columns:
        df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce')
        df['Количество'] = df['Количество'].round()
    
    # Фільтруємо DataFrame, залишаючи лише рядки, де 'Количество' не є NaN
    df = df[df['Количество'].notna()]     
    
    # Отримуємо ціни для заданої організації
    prices = get_prices(organization)
    
    # Призначення ціни для кожного запису:
    # - використовується метод map для зіставлення назви товару з цінами,
    # - якщо ціна не знайдена, встановлюємо значення за замовчуванням (1)
    df['Цена'] = df['Наименование'].map(prices).fillna(1)
    df['Цена'] = pd.to_numeric(df['Цена'], errors='coerce')
    
    # Оновлення інформації про виробника за допомогою зовнішньої функції
    df = set_manufacturer_for_organization(df, organization)
    
    # Розрахунок суми для кожного запису
    if organization == 'Stada':
        # Для організації "Stada" сума встановлюється рівною 1
        df['Сумма'] = 1
    elif organization == 'Arterium':
        # Для організації "Arterium" сума обчислюється як добуток "Количество" на "Цена",
        # якщо назва товару міститься у словнику цін, інакше значення суми встановлюється в 1.
        df['Сумма'] = df.apply(
            lambda row: row['Количество'] * row['Цена'] if row['Наименование'] in prices else 1,
            axis=1
        )
    else:
        # Для інших організацій сума обчислюється як добуток "Количество" на "Цена"
        df['Сумма'] = df['Количество'].astype(float) * df['Цена'] 
        
    # Додаємо стовпець "Номер" як послідовний номер кожного рядка
    df['Номер'] = range(1, len(df) + 1)
    
    # Видалення колонок, де всі значення рівні нулю
    df = df.loc[:, (df != 0).any()]
    # Видалення колонок, де всі значення порожні (NaN)
    df = df.loc[:, df.notnull().any()]
    
    # Перестановка колонок так, щоб стовпець "Номер" був першим
    columns = ['Номер'] + [col for col in df.columns if col != 'Номер']
    df = df[columns]      
        
    # Об'єднання DataFrame за допомогою функції combine_dataframes 
    # (у даному випадку об'єднується лише один DataFrame)
    combined_df = combine_dataframes([df])
    
    return df
