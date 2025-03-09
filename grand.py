import pandas as pd
import functools
from utils import set_manufacturer_for_organization, combine_dataframes

@functools.lru_cache(maxsize=128)
def get_prices(organization):
    """Повертає ціни в залежності від вибраного представника."""
    # Словник з цінами для кожної організації.
    # Кожна організація має свій вкладений словник, де ключ — назва товару, а значення — його ціна.
    prices_dict = {
        'Arterium': {
            "Тиоцетам 10мл №10": 94077,
            "Тиоцетам 5мл №10": 70275,
            "Тиоцетам 400мг/100мг №30 табл.": 52007,
            "Элкоцин 100мг №30 табл.": 76028,
            "Ларитилен №20 табл. мяты": 24430,
            "Ларитилен №20 табл. мяты и малины": 24430,
            "Ларитилен №20 табл.мяты и лимон": 24430,
            "Формен Комби №40 капс.": 94280,
            #--------------------------------------
            'Ларитилен таб. №20 (мята)': 24430,
            'Ларитилен таб. №20 (мята/лимон)': 24430,
            'Ларитилен таб. №20 (мята/малина)': 24430,
            'Тиоцетам р-р 10мл №10': 94077,
            'Тиоцетам р-р 5мл №10': 70275,
            'Тиоцетам таб. п/о 400мг/100мг N30': 52007,
            'Элкоцин таб. п/о 100мг №30': 76028,
            'Формэн комби капс. №40': 94280,
            #---------------------------------------
            "Ларитилен №20 табл. мяты": 24430,
            "Ларитилен №20 табл. мяты и малины": 24430,
            "Ларитилен №20 табл.мяты и лимон": 24430,
            "Резистол капли оральные 50мл": 84290,
            "Резистол капли оральные 20мл": 49000,
            "РЕЗИСТОЛ КАПЛИ 20 МЛ": 49000,
            "Тиотриазолин форте амп. 50мг/мл 4мл №10": 173532
        },
        'Astra Zeneca': {
            "Беталок Зок таб. 100мг №30": 113791,
            "Брилинта таб. 90мг №56": 984154,
            "Нексиум таб. 40мг №14": 115488,
            "Пульмикорт сусп. 0,5мг/2мл №20": 340220,
            "Симбикорт Турбухалер пор. д/ингал. 160/4,5 мкг/доза 60доз №1": 281134,
            "Симбикорт Турбухалер пор. д/ингал. 80/4,5мкг/доза 60доз №1": 253408,
            "Форсига таб. 10мг №28": 365507
        },
        'Egis': {
            "Алзепил таб. 10мг №28": 373580,
            "Алзепил таб. 5мг №28": 202663,
            "Алотендин таб. 5мг/10мг №30": 81965,
            "Алотендин таб. 5мг/5мг №30": 88915,
            "Алотендин таб.10мг/10мг №30": 108460,
            "Бетадин мазь для нар.прим. 100мг/г 20г": 48786,
            "Бетадин р-р 100мг/г 1000мл": 233927,
            "Бетадин р-р 100мг/г 120мл": 67467,
            "Бетадин р-р 100мг/г 30мл": 37700,
            "Бетадин супп. вагин. 200мг №14": 86308,
            "Велаксин капс. 37,5мг №28": 96468,
            "Велаксин капс. 75мг №28": 198829,
            "Грандаксин таб. 50мг N60": 167830,
            "Грандаксин таб. 50мг №20": 63172,
            "Кетилепт таб. 200мг №60": 578836,
            "Клостилбегит таб. 50мг №10": 143890,
            "Летирам таб. 1000мг №60": 607518,
            "Летирам таб. 500мг №60": 341586,
            "Розулип таб. 10мг №28": 99961,
            "Розулип таб. 20мг №28": 165831,
            "Розулип Плюс капс. 10мг/10мг №60": 363204,
            "Розулип Плюс капс. 20мг/10мг №60": 399343,
            "Сорбифер Дурулес таб. №30": 57203,
            "Сорбифер Дурулес таб. №50": 71509,
            "Стимулотон таб. 50мг №30": 108865,
            "Супрастин 20мг/мл р-р для ин. 1мл №5": 34877,
            "Супрастин таб. 25мг №20": 29687,
            "Супрастинекс капли д/пр. вн. 5мг/мл 20мл": 70701,
            "Супрастинекс таб. 5мг №30": 73364,
            "Таллитон таб. 12,5мг №28": 79338,
            "Таллитон таб. 25мг №28": 97455,
            "Таллитон таб. 6,25мг №28": 60005
        },
        'Stada': {
            # Для організації Stada ціни не визначені (закоментовано) або можуть встановлюватися окремо.
            # "Тиоцетам 10мл №10": 1,
            # "Тиоцетам 5мл №10": 1,
            # "Тиоцетам 400мг/100мг №30 табл.": 1,
            # "Уролесан 180мл": 1,
            # "Уролесан 25мл": 1,
            # "Элкоцин 100мг №30 табл.": 1,
            # "Ларитилен №20 табл. мяты": 1,
            # "Ларитилен №20 табл. мяты и малины": 1,
            # "Ларитилен №20 табл.мяты и лимон": 1
        },
        'Без организации': {}
    }
    # Повертаємо словник цін для заданої організації, або порожній словник, якщо організація не знайдена.
    return prices_dict.get(organization, {})

def process_grand_excel(file_path, organization):
    """Обробляє завантажений Excel файл і повертає оброблений DataFrame."""
    
    # =========================
    # Зчитування Excel файлу
    # =========================
    # Зчитуємо всі аркуші Excel файлу у словник, де ключем є назва аркуша, а значенням – DataFrame.
    df_all = pd.read_excel(file_path, sheet_name=None)
    
    # =========================
    # Перевірка наявності потрібного аркуша
    # =========================
    # Шукаємо аркуш з назвою "Развернутый". Якщо такий аркуш відсутній, піднімаємо ValueError.
    sheet_name = next((name for name in ["Развернутый"] if name in df_all), None)
    if sheet_name is None:
        raise ValueError("Аркуш 'Развернутый' не знайдено")
    
    # Отримуємо DataFrame з потрібного аркуша.
    df = df_all[sheet_name]
    
    # =========================
    # Перейменування стовпців
    # =========================
    # Якщо в DataFrame є стовпець "Кол-во", перейменовуємо його в "Количество".
    if 'Кол-во' in df.columns:
        df.rename(columns={'Кол-во': 'Количество'}, inplace=True)
    
    # =========================
    # Обробка дати
    # =========================
    # Конвертуємо стовпець "Дата" у формат datetime згідно з форматом "день.місяць.рік".
    df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', dayfirst=True, errors='coerce')
    
    # =========================
    # Обробка кількості
    # =========================
    # Якщо стовпець "Количество" існує, перетворюємо його значення у числовий тип з обробкою помилок,
    # та округлюємо значення.
    if 'Количество' in df.columns:
        df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce')
        df['Количество'] = df['Количество'].round()
    
    # Фільтруємо рядки, де "Количество" має валідне числове значення (не NaN).
    df = df[df['Количество'].notna()]
    
    # =========================
    # Призначення цін
    # =========================
    # Отримуємо словник з цінами для заданої організації.
    prices = get_prices(organization)
    # Використовуючи метод map, зіставляємо значення в колонці "Наименование" із словником цін.
    # Якщо для товару не знайдено відповідну ціну, встановлюємо значення за замовчуванням (1).
    df['Цена'] = df['Наименование'].map(prices).fillna(1)
    # Перетворюємо значення в колонці "Цена" у числовий тип.
    df['Цена'] = pd.to_numeric(df['Цена'], errors='coerce')
    
    # =========================
    # Оновлення інформації про виробника
    # =========================
    # Викликаємо зовнішню функцію для встановлення або оновлення даних про виробника згідно з обраною організацією.
    df = set_manufacturer_for_organization(df, organization)
    
    # =========================
    # Розрахунок суми
    # =========================
    # Обчислюємо суму для кожного рядка на основі організації та наявності ціни для товару.
    if organization == 'Stada':
        df['Сумма'] = 1
    elif organization == 'Arterium':
        df['Сумма'] = df.apply(
            lambda row: row['Количество'] * row['Цена'] if row['Наименование'] in prices else 1,
            axis=1
        )
    elif organization == 'Astra Zeneca':
        df['Сумма'] = df.apply(
            lambda row: row['Количество'] * row['Цена'] if row['Наименование'] in prices else 1,
            axis=1
        )
    elif organization == 'Binnopharm Group':
        df['Сумма'] = df.apply(
            lambda row: row['Количество'] * row['Цена'] if row['Наименование'] in prices else 1,
            axis=1
        )
    elif organization == 'Egis':
        df['Сумма'] = df.apply(
            lambda row: row['Количество'] * row['Цена'] if row['Наименование'] in prices else 1,
            axis=1
        )
    else:
        df['Сумма'] = df['Количество'].astype(float) * df['Цена']
    
    # =========================
    # Фільтрація рядків за значенням у "Покупатель"
    # =========================
    # Якщо в DataFrame є стовпець "Покупатель", видаляємо рядки, де значення рівне "Списание".
    if 'Покупатель' in df.columns:
        df = df[df['Покупатель'] != 'Списание']
    
    # =========================
    # Призначення ІНН для покупців
    # =========================
    # Визначаємо відповідність для ІНН за значенням стовпця "Покупатель" за допомогою мапінгу.
    inn_mapping = {
        'Аптека филиала': 111111111,
        'Расход на филиал': 111111125,
        'Ташкент GRAND BEST РЦ': 111111112,
        'Андижон GRAND BEST РЦ': 111111113,
        'Бухоро Гранд сеть РЦ': 111111114,
        'Самарканд GRAND BEST РЦ': 111111115,
        'Самарканд Гранд сеть РЦ': 111111115,
        'Фаргона GRAND BEST РЦ': 111111116,
        'Фаргона Гранд сеть РЦ': 111111116,
        'Жиззах GRAND BEST РЦ': 111111117,
        'Жиззах Гранд сеть РЦ': 111111117,
        'Карши GRAND BEST РЦ': 111111118,
        'Карши Гранд сеть РЦ': 111111118,
        'Наманган BEST РЦ': 111111119,
        'Термез GRAND BEST РЦ': 111111120,
        'Термез Гранд сеть РЦ': 111111120,
        'Термиз GRAND РЦ': 111111120,
        'Ургенч BEST РЦ': 111111121,
        'Нукус GRAND BEST РЦ': 111111122,
        'Нукус Гранд сеть РЦ': 111111122,
    }
    # Якщо стовпець "Покупатель" є, використовуючи мапінг, встановлюємо значення ІНН.
    if 'Покупатель' in df.columns:
        df['ИНН'] = df['Покупатель'].map(inn_mapping).fillna(df['ИНН'])
    
    # =========================
    # Додавання порядкового номера
    # =========================
    # Додаємо нову колонку "Номер" з послідовними номерами для кожного рядка.
    df['Номер'] = range(1, len(df) + 1)
    
    # Переставляємо колонки так, щоб "Номер" був першою, а всі інші наступали після нього.
    columns = ['Номер'] + [col for col in df.columns if col != 'Номер']
    df = df[columns]
    
    # =========================
    # Об'єднання даних
    # =========================
    # Об'єднуємо дані, викликаючи зовнішню функцію combine_dataframes.
    # Тут DataFrame обгортається в список, що дозволяє об'єднати декілька DataFrame, якщо це необхідно.
    combined_df = combine_dataframes([df])
    
    # Повертаємо остаточний оброблений DataFrame.
    return df
