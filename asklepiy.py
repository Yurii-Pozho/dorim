import streamlit as st
import pandas as pd
import openpyxl
import io

# Конфігурація сторінки
st.set_page_config(
    page_title="Excel File Processor",
    layout="wide"
)

# Словник цін
prices = {
    "Тиоцетам 10мл №10": 92598,
    "Тиоцетам 5мл №10": 69196,
    "Тиоцетам 400мг/100мг №30 табл.": 50853,
    "Уролесан 180мл": 58823,
    "Уролесан 25мл": 55407,
    "Элкоцин 100мг №30 табл.": 72611,
    "Ларитилен №20 табл. мяты": 23403,
    "Ларитилен №20 табл. мяты и малины": 23403,
    "Ларитилен №20 табл.мяты и лимон": 23403
}

st.title('Обробка Excel файлу')

# Завантаження файлу
uploaded_file = st.file_uploader("Оберіть Excel файл", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Читаємо файл
        df_all = pd.read_excel(uploaded_file, sheet_name=None)

        # Перевіряємо наявність потрібного аркуша
        sheet_name = next((name for name in ["Реализация", "Реализация "] if name in df_all), None)

        if sheet_name is None:
            st.error("Аркуш 'Реализация' не знайдено")
            st.stop()

        df = df_all[sheet_name]

        # Базова обробка
        df['Дата'] = pd.to_datetime(df['Дата']).dt.strftime('%d.%m.%Y')
        df['Срок годности'] = pd.to_datetime(df['Срок годности']).dt.strftime('%d.%m.%Y')
        df.insert(0, 'Номер', range(1, len(df) + 1))

        # Перейменування колонок
        rename_dict = {
            'Кол-во': 'Количество',
            'ИНН Клиента': 'ИНН',
            'Адрес Клиента': 'Адрес'
        }
        df = df.rename(columns=rename_dict)

        # Перетворення 'ИНН' на цілі числа
        df['ИНН'] = df['ИНН'].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

        # Додавання цін та розрахунок суми
        df['Цена'] = 1
        for name, price in prices.items():
            df.loc[df['Наименование'] == name, 'Цена'] = price

        df['Количество'] = df['Количество'].fillna(0)
        df['Сумма'] = df['Количество'] * df['Цена']

        # Перестановка колонок: переміщаємо 'Количество' перед 'Цена'
        df = df[['Номер', 'Дата', 'Срок годности','Наименование', 'ИНН', 'Адрес','Количество', 'Цена','Сумма']]

        # Видалення пустих колонок
        df = df.dropna(axis=1, how='all')

        # Виведення результатів
        st.write("### Статистика")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Количество строк", len(df))
        with col2:
            st.metric("Сумм количества", f"{df['Количество'].sum():,.2f}")
        with col3:
            total_sum = df[df['Цена'] > 1]['Сумма'].sum()
            st.metric("Сума (цена > 1)", f"{total_sum:,.2f}")

        # Показ таблиці
        st.write("### Оброблені дані")
        st.dataframe(df)

        # Кнопка завантаження
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Результат')
        processed_file = output.getvalue()

        st.download_button(
            label="Загрузить результат",
            data=processed_file,
            file_name='output_asklepiy.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except FileNotFoundError as e:
        st.error(f"Файл не знайдено: {str(e)}")
    except PermissionError as e:
        st.error(f"Недостатньо прав доступу: {str(e)}")
    except Exception as e:
        st.error(f"Невідома помилка: {str(e)}")
