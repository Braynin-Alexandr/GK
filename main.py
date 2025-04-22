import pandas as pd
import re
from datetime import datetime
import numpy as np
import holidays

file_name = 'Выгрузка.xlsx'
df = pd.read_excel(file_name)

ru_holidays = holidays.Russia(years=[2023, 2024, 2025])
holiday_dates = list(ru_holidays.keys())


def count_weekdays(start_date, end_date):
    return int(np.busday_count(
        start_date.date().strftime('%Y-%m-%d'),
        end_date.date().strftime('%Y-%m-%d'),
        holidays=holiday_dates
    ))


def get_status_info(row):
    text = row['История']
    order_id = row['Номер закупки']

    re_pattern = r'(\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}) ([^\n]+)'
    matches = list(re.finditer(re_pattern, text))

    status_info = {}

    for i in range(len(matches)-1):
        status = matches[i].group(2).strip()
        start_date = matches[i].group(1)
        finish_date = matches[i+1].group(1)
        start_date_dt = datetime.strptime(start_date, '%d.%m.%Y %H:%M:%S')
        finish_date_dt = datetime.strptime(finish_date, '%d.%m.%Y %H:%M:%S')
        work_days = count_weekdays(start_date_dt, finish_date_dt)

        if status not in status_info:
            status_info[status] = {'Затрачено дней': 0}

        status_info[status]['Затрачено дней'] += work_days

    last_row = matches[-1].group(2)
    if last_row not in status_info:
        status_info[last_row] = {'Затрачено дней': 0}

    return [{'Номер закупки': order_id,
             'Статус': status,
             'Длительность, дни': date['Затрачено дней']} for status, date in status_info.items()]


all_orders = []

for _, row in df.iterrows():
    status_data = get_status_info(row)
    all_orders.extend(status_data)

try:
    new_df = pd.DataFrame(all_orders)
    new_df.to_excel('Report.xlsx', index=False)
except Exception as e:
    print(f'Something went wrong: {e}')
else:
    print('Report successfully created.')
