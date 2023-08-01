from datetime import datetime, timedelta
import pytz
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

def save_to_excel(df, file_name):
    # Create a workbook and worksheet
    wb = Workbook()
    ws = wb.active

    # Set the title and apply formatting
    timezone_utc_plus_9 = pytz.timezone('Asia/Tokyo')
    time_scraped = datetime.now(timezone_utc_plus_9)
    str_time_scraped = time_scraped.strftime('%m月%d日 %H:%M')
    title = f"埼玉県営テニスコート名義 {str_time_scraped}取得"
    
    ws['A1'] = title
    title_font = Font(size=14, bold=True)
    ws['A1'].font = title_font

    # Write the DataFrame to the worksheet
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    # Set column width for better readability
    column_widths = [15, 20, 20, 15, 20]
    for i, width in enumerate(column_widths, start=1):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = width
      
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=4):
        for cell in row:
            cell.alignment = Alignment(horizontal='left')

    # Write the total number of rows at the bottom
    total_rows = len(df) + 4  # Add 4 for the two additional rows
    ws.cell(row=total_rows-1, column=1, value="利用可名義数:")
    ws.cell(row=total_rows-1, column=2, value=len(df))

    ws.cell(row=total_rows, column=1, value="総票数:")
    ws.cell(row=total_rows, column=2, value=len(df)*4)

    # Save the Excel file
    wb.save(file_name)


def check_schedule_within_30_minutes():
    # 日本の時刻
    timezone_utc_plus_9 = pytz.timezone('Asia/Tokyo')
    current_time = datetime.now(timezone_utc_plus_9)

    # 毎週金曜日のスケジュール
    friday_schedule = {'day_of_week': 4, 'start_time': '03:00', 'end_time': '03:30'}

    # 毎月第2水曜日のスケジュール
    second_wednesday = current_time.replace(day=15)
    while second_wednesday.weekday() != 2:  # 水曜日を探す
        second_wednesday += timedelta(days=1)
    second_wednesday_schedule = {'day': second_wednesday.day, 'start_time': '22:30', 'end_time': '23:59'}
    next_day_schedule = {'day': (second_wednesday + timedelta(days=1)).day, 'start_time': '00:00', 'end_time': '08:00'}

    # 現在時刻から30分後までをチェック
    for _ in range(31):
        # Check the current time and two schedules
        if check_single_schedule(current_time, friday_schedule) or check_single_schedule(current_time, second_wednesday_schedule) or check_single_schedule(current_time, next_day_schedule):
            return 1
        # Increment current_time by 1 minute for the next iteration
        current_time += timedelta(minutes=1)

    return 0


def check_single_schedule(current_time, schedule):
    start_time = datetime.strptime(schedule['start_time'], '%H:%M').time()
    end_time = datetime.strptime(schedule['end_time'], '%H:%M').time()

    if not (day := schedule.get("day")):
        # 毎週金曜日について
        day_of_week = schedule['day_of_week']

        if current_time.weekday() == day_of_week and start_time <= current_time.time() <= end_time:
            return True
    else:
        # 第2水曜日の判定
        if current_time.day == day and start_time <= current_time.time() <= end_time:
            return True

    return False
