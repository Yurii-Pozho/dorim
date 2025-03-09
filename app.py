import streamlit as st
import pandas as pd
import io
from asklepiy import process_asklepiy_excel, combine_dataframes
from astor import process_astor_excel, combine_dataframes
from biotek import process_biotek_excel, combine_dataframes
from curatio import process_curatio_excel, combine_dataframes
from farm_lyuks import process_farm_lyuks_excel, combine_dataframes
from grand import process_grand_excel, combine_dataframes
from hurshida import process_hurshida_excel, combine_dataframes
from memory import process_memory_excel, combine_dataframes
from meros import process_meros_excel, combine_dataframes
from novotek import process_novotek_excel, combine_dataframes
from pharma_choice import process_pharma_choice_excel, combine_dataframes
from pharmaxi import process_pharmaxi_excel, combine_dataframes

# Настройки страницы
st.set_page_config(
    page_title="Excel File Processor",
    layout="wide"
)

# Инициализация сессионных состояний
st.session_state.setdefault('distributor', 'Выбрать дистрибьютора')
st.session_state.setdefault('organization', 'Без организации')

# Заголовок приложения
st.markdown(
    """
    <h1 style="text-align: center;">Обработчик DORIM</h1>
    """,
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader("Выберите Excel файлы", type=['xlsx', 'xls'], accept_multiple_files=True)

# Создание колонок для выбора дистрибьютора и организации
col1, col2 = st.columns(2)

with col1:
    st.session_state.distributor = st.selectbox(
        'Дистрибьютор',
        ['Выбрать дистрибьютора', 'Asklepiy', 'Astor', 'Biotek', 'Curatio', 'Farm Lyuks', 'Grand', 'Hurshida', 'Memory', 'Meros', 'Novotek', 'Pharma Choice', 'Pharmaxi'],
        index=['Выбрать дистрибьютора', 'Asklepiy', 'Astor', 'Biotek', 'Curatio', 'Farm Lyuks', 'Grand', 'Hurshida', 'Memory', 'Meros', 'Novotek', 'Pharma Choice', 'Pharmaxi'].index(st.session_state.distributor)
    )

with col2:
    # Выбор организации
    organization_choice = st.selectbox(
        'Маркетирующая организация',
        ['Без организации', 'Arterium', 'Astra Zeneca', 'Binnopharm Group', 'Egis', 'Stada'],
        index=['Без организации', 'Arterium', 'Astra Zeneca', 'Binnopharm Group', 'Egis', 'Stada'].index(st.session_state.organization)
    )

    # Кнопка под выбором организации
    if st.button('Подтвердить выбор организации'):
        st.session_state.organization = organization_choice

# Загрузка файлов
if uploaded_files:
    if st.session_state.distributor == 'Выбрать дистрибьютора':
        st.warning("Пожалуйста, выберите дистрибьютора перед загрузкой файлов.")
    else:
        try:
            # Обработка в зависимости от выбора дистрибьютора
            if st.session_state.distributor == 'Asklepiy':
                df_list = [process_asklepiy_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Astor':
                df_list = [process_astor_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Biotek':
                df_list = [process_biotek_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)  
                
            elif st.session_state.distributor == 'Curatio':
                df_list = [process_curatio_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Farm Lyuks':  
                df_list = [process_farm_lyuks_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list) 
                
            elif st.session_state.distributor == 'Grand':  
                df_list = [process_grand_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Hurshida':  
                df_list = [process_hurshida_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Memory':
                df_list = [process_memory_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Meros': 
                df_list = [process_meros_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list) 
                
            elif st.session_state.distributor == 'Novotek':
                df_list = [process_novotek_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            elif st.session_state.distributor == 'Pharma Choice':  
                df_list = [process_pharma_choice_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)   
                
            elif st.session_state.distributor == 'Pharmaxi':  
                df_list = [process_pharmaxi_excel(file, st.session_state.organization) for file in uploaded_files]
                df = combine_dataframes(df_list)
                
            else:
                st.error(f"Дистрибьютор '{st.session_state.distributor}' не поддерживается.")
                df = pd.DataFrame()                     

            # Обработка дат
            if 'Дата' in df.columns:
                min_date, max_date = df['Дата'].min(), df['Дата'].max()
                date_range = f"{min_date} - {max_date}" if pd.notnull(min_date) and pd.notnull(max_date) else 'Нет данных'
            else:
                date_range = 'Нет данных'
            
            # Вывод статистики
            st.write("### Статистика")
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric("Период", date_range)

            with col2:
                st.metric("Количество строк", len(df))

            if 'Количество' in df.columns:
                with col3:
                    st.metric("Сумма количества", f"{df['Количество'].sum():,.2f}")

            if 'Цена' in df.columns and 'Сумма' in df.columns:
                total_sum = df[df['Цена'] > 1]['Сумма'].sum()
                with col4:
                    st.metric("Сумма (цена > 1)", f"{total_sum:,.2f}")

            # Вывод таблицы
            st.write("### Обработанные данные")
            st.write(df)

            # Формирование динамического имени файла на основе выбранных параметров
            distributor_name = st.session_state.distributor.strip().lower().replace(' ', '_')
            organization_name = st.session_state.organization.strip().lower().replace(' ', '_')
            file_name = f"{distributor_name}_{organization_name}.xlsx"

            # Кнопка для загрузки результата
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Результат')

            st.download_button(
                label="Загрузить результат",
                data=output.getvalue(),
                file_name=file_name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {str(e)}")
