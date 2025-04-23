import pandas as pd
import re
from datetime import datetime
import numpy as np


file_name = 'Выгрузка.xlsx'
df = pd.read_excel(file_name)

target_stage = 'Анализ цены МТР'
target_statuses = ['Назначение исполнителя', 'Исполнитель назначен', 'Анализ проведен', 'Анализ завершен']

holidays = [
    "2023-01-01", "2023-01-02", "2023-01-03", "2023-01-04", "2023-01-05", "2023-01-06", "2023-01-07", "2023-01-08",
    "2023-02-23", "2023-02-24", "2023-02-25", "2023-02-26",
    "2023-03-08",
    "2023-04-29", "2023-04-30", "2023-05-01",
    "2023-05-06", "2023-05-07", "2023-05-08", "2023-05-09",
    "2023-06-10", "2023-06-11", "2023-06-12",
    "2023-11-04", "2023-11-05", "2023-11-06"
]


def convert_to_datetime(date: str) -> datetime:
    """Convert a date string to a datetime object"""
    return datetime.strptime(date, '%d.%m.%Y %H:%M:%S')


def count_weekdays(start_date: str, end_date: str, holiday_dates: list=None) -> int:
    """Count the number of weekdays between start_date and end_date, excluding holidays"""
    if not isinstance(start_date, datetime):
        start_date = convert_to_datetime(start_date)
    if not isinstance(end_date, datetime):
        end_date = convert_to_datetime(end_date)

    holiday_dates = holiday_dates or []

    return int(np.busday_count(
        start_date.date().strftime('%Y-%m-%d'),
        end_date.date().strftime('%Y-%m-%d'),
        holidays=holiday_dates
    ))


def update_order_info(order_info: dict, stage: str, status: str, work_days: int) -> None:
    """Updates the order_info dictionary with the number of workdays for a given stage and status"""
    if stage not in order_info:
        order_info[stage] = {}

    if status not in order_info[stage]:
        order_info[stage][status] = 0

    order_info[stage][status] += work_days


def get_order_info(order: pd.Series) -> dict[str, dict[str:int]]:
    """Extract and process information from an order's history"""
    text = order['История']
    re_pattern = r'(?P<date>\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}) (?P<stage>[^\n]+)(?:\n(?P<status>.*?))?(?=\n\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}|\Z)'
    matches = list(re.finditer(re_pattern, text))

    if not matches:
        return {}

    result = {}
    for i in range(len(matches)-1):
        current = matches[i]
        next_match = matches[i + 1]

        stage = current.group("stage").strip()
        status = current.group("status").strip() if current.group("status") else ''
        start_date = current.group("date")
        finish_date = next_match.group("date")
        work_days = count_weekdays(start_date, finish_date, holidays)

        update_order_info(result, stage, status, work_days)

    last_stage = matches[-1].group("stage").strip()
    last_status = matches[-1].group("status").strip() if matches[-1].group("status") else ''
    update_order_info(result, last_stage, last_status, 0)

    return result


def get_status_info(order: dict, group_name: str, status_group: list) -> dict:
    """Extracts and calculates the total workdays for each status in the given status_group for a specific order"""
    group_data = order.get(group_name, {})
    status_days = {status: 0 for status in status_group}

    for status_name, days in group_data.items():
        for target in status_group:
            if status_name.startswith(target):
                status_days[target] += days
                break

    return status_days


all_results = []
for _, row in df.iterrows():
    order_number = row['Номер закупки']
    order_info = get_order_info(row)
    statuses_info = get_status_info(order_info, target_stage, target_statuses)
    order_result = {'Номер закупки': order_number} | statuses_info
    all_results.append(order_result)


try:
    new_df = pd.DataFrame(all_results)
    new_df.to_excel('Report.xlsx', index=False)
except Exception as e:
    print(f'Something went wrong: {e}')
else:
    print('Report successfully created.')
