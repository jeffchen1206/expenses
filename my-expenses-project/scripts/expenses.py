import pandas as pd
import psycopg2
from datetime import datetime, timedelta
import os

# 步驟 1: 讀取 Excel 文件
file_path = r"C:\Users\jeff\Desktop\dailycost.xlsx"
df = pd.read_excel(file_path, sheet_name="工作表1")

# 步驟 2: 資料清洗
def excel_date_to_datetime(excel_date):
    if pd.isna(excel_date):
        return None
    # 如果 excel_date 是字串格式 (如 '2025/7/1')
    if isinstance(excel_date, str):
        try:
            return datetime.strptime(excel_date, '%Y/%m/%d').date()
        except ValueError:
            return None
    # 如果 excel_date 是 Excel 日期序列號 (如 45839)
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(excel_date))).date()
    except (ValueError, TypeError):
        return None

# 應用日期轉換並只保留日期部分
df['date'] = df['date'].apply(excel_date_to_datetime)

# 確保 amount 是數字類型
df['amount'] = pd.to_numeric(df['amount'], errors='coerce')

# 移除無效行 (date 或 amount 為 NaN)
df = df.dropna(subset=['date', 'amount'])

# 步驟 3: 連接到 PostgreSQL
conn = psycopg2.connect(
    dbname="test",
    user="postgres",
    password="123456",
    host="localhost",
    port="5432"
)

# 步驟 4: 創建 expenses 表格 (如果不存在)
create_table_query = """
CREATE TABLE IF NOT EXISTS expenses (
    id SERIAL PRIMARY KEY,
    category VARCHAR(50),
    amount DECIMAL(10, 2),
    date DATE,
    note VARCHAR(255)
);
"""
cursor = conn.cursor()
cursor.execute(create_table_query)
conn.commit()

# 步驟 5: 將數據匯入 PostgreSQL
for index, row in df.iterrows():
    insert_query = """
    INSERT INTO expenses (category, amount, date, note)
    VALUES (%s, %s, %s, %s)
    ON CONFLICT (id) DO NOTHING;
    """
    cursor.execute(insert_query, (row['category'], row['amount'], row['date'], row['note']))

conn.commit()
cursor.close()
conn.close()

print("數據已成功匯入 PostgreSQL!")