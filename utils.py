import pandas as pd

def set_manufacturer_for_organization(df, organization):
    """
    Встановлює виробника тільки якщо колонка "Производитель" відсутня, пуста
    або має значення "Без организации".
    """
    manufacturer_dict = {
        'Arterium': 'Arterium',
        'Stada': 'Stada',
        'Binnopharm Group': 'Binnopharm Group',
        'Astra Zeneca':'Astra Zeneca',
        'Egis':'Egis',
    }

    # Перевіряємо, чи колонка існує
    if 'Производитель' not in df.columns:
        # Якщо колонки немає, створюємо її
        df['Производитель'] = manufacturer_dict.get(organization, 'Без организации')
    else:
        # Заповнюємо тільки рядки зі значенням "Без организации", NaN або пустими рядками
        mask = (df['Производитель'].isna() | (df['Производитель'] == '') | (df['Производитель'] == 'Без организации'))
        df.loc[mask, 'Производитель'] = manufacturer_dict.get(organization, 'Без организации')

    return df


def combine_dataframes(df_list):
    """Об'єднує список DataFrame в один і форматує дату, після чого присвоює порядкові номери."""    
    combined_df = pd.concat(df_list, ignore_index=True)
    
    combined_df['Номер'] = range(1, len(combined_df) + 1)
    
    if 'Дата' in combined_df.columns:
        combined_df['Дата'] = pd.to_datetime(combined_df['Дата'], format='%d.%m.%Y', errors='coerce').dt.date
    
    return combined_df
