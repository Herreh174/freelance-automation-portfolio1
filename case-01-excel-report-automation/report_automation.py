import pandas as pd
import os
from datetime import datetime

# Папка с исходными CSV-файлами
DATA_DIR = 'sample_data'

# Название выходного файла
OUTPUT_FILE = 'final_report.xlsx'

def load_all_csv_files(folder_path):
    """Загружает все CSV-файлы из папки и объединяет их в один DataFrame"""
    all_data = pd.DataFrame()
    for file in os.listdir(folder_path):
        if file.endswith('.csv'):
            file_path = os.path.join(folder_path, file)
            df = pd.read_csv(file_path)
            all_data = pd.concat([all_data, df], ignore_index=True)
    return all_data

def clean_data(df):
    """Удаляет дубли, пустые строки и форматирует данные"""
    # Убедимся, что нужные столбцы существуют
    if 'Дата' not in df.columns or 'Сумма' not in df.columns:
        raise ValueError("Ожидаемые столбцы 'Дата' и 'Сумма' не найдены в данных.")

    # Преобразуем даты
    df['Дата'] = pd.to_datetime(df['Дата'], dayfirst=True, errors='coerce')

    # Преобразуем сумму
    df['Сумма'] = pd.to_numeric(df['Сумма'], errors='coerce')

    # Удалим строки, где дата или сумма невалидны
    df = df.dropna(subset=["Дата", "Сумма"])

    # Удалим дубликаты
    df = df.drop_duplicates()

    return df

def generate_report(df):
    """Группирует данные по дням и считает сумму и средний чек"""
    grouped = df.groupby(df['Дата'].dt.date).agg(
        Всего_продаж=('Сумма', 'sum'),
        Средний_чек=('Сумма', 'mean'),
        Количество_транзакций=('Сумма', 'count')
    ).reset_index()
    grouped['Дата'] = pd.to_datetime(grouped['Дата'])
    return grouped

def save_to_excel(df_summary):
    """Сохраняет отчет в Excel"""
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Отчет', index=False)
    print(f"✅ Отчет успешно сохранён в {OUTPUT_FILE}")

def main():
    print("📥 Загрузка и обработка данных...")
    raw_data = load_all_csv_files(DATA_DIR)
    clean = clean_data(raw_data)
    summary = generate_report(clean)
    save_to_excel(summary)

if __name__ == '__main__':
    main()
