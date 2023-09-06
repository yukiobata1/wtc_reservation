from datetime import datetime, timedelta
import pytz
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side
import re
import os
import glob
import openpyxl
import pandas as pd

def extract_kakutei(string):
    # 予約番号を抽出
    print(f"{string=}")
    facility_pattern = r"施設名\n(.*?)\n"
    reservation_date_pattern = r"予約日\n(.*?)\n"
    usage_time_pattern = r"使用時間\n(.*?)\n"
    lottery_date_pattern = r"抽選日\n(.*?)\n"
    reservation_number_pattern = r"予約番号\n([0-9]*)$"
    
    facility_match = re.search(facility_pattern, string)
    reservation_date_match = re.search(reservation_date_pattern, string)
    usage_time_match = re.search(usage_time_pattern, string)
    lottery_date_match = re.search(lottery_date_pattern, string)
    reservation_number_match = re.search(reservation_number_pattern, string)
    
    if facility_match:
        facility_name = facility_match.group(1)
    else:
        facility_name = None
    
    if reservation_date_match:
        reservation_date = reservation_date_match.group(1)
    else:
        reservation_date = None
    
    if usage_time_match:
        usage_time = usage_time_match.group(1)
    else:
        usage_time = None
    
    if lottery_date_match:
        lottery_date = lottery_date_match.group(1)
    else:
        lottery_date = None
    
    if reservation_number_match:
        reservation_number = reservation_number_match.group(1)
    else:
        reservation_number = None

    print(f"施設名: {facility_name}")
    print(f"予約日: {reservation_date}")
    print(f"使用時間: {usage_time}")
    print(f"抽選日: {lottery_date}")
    print(f"予約番号: {reservation_number}")


    return {"court":facility_name,
            "date": reservation_date,
            "time_range": usage_time, 
            "reservation_number": int(reservation_number)
           }
    

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
    column_widths = [15, 20, 20, 15]
    for i, width in enumerate(column_widths, start=1):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = width
      
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=4):
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

import pandas as pd

def create_date_dict(data):
    rows_dict = dict()
    columns = ['08:30', '10:30', '12:30', '14:30', '16:30', '18:30']
    
    for court, value in data.items():
        for date, times in value.items():
            rows_dict[date] = []
        
    # Iterate through the data dictionary to populate the DataFrame rows
    for court, value in data.items():
        for date, times in value.items():
            rows_dict[date].append([int(court)+1] + [times[time] for time in columns])
    return rows_dict

def save_votes(data, file_path):
    date_dict = create_date_dict(data)
    thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

    # Create a workbook and worksheet
    wb = Workbook()
    ws = wb.active

    # Set the title and apply formatting
    title = "票数"
    ws['A1'] = title
    title_font = Font(size=14, bold=True)
    ws['A1'].font = title_font
    
    count = 0
    for i, (date, courts) in enumerate(date_dict.items()):
        # set border to each cell
        for row in ws[f'B{1+i*16}:I{1+i*16+12}']:
            for cell in row:
                cell.border = thin_border
        ws.cell(row=1+i*16, column=2, value= f"{date}")
        # create header
        start_times = ['08:30', '10:30', '12:30', '14:30', '16:30', '18:30'] 
        for j, value in enumerate(start_times):
            ws.cell(row=1+i*16, column=2+j+1, value= f"{value}")

        # write number of votes for each day
        for j, court in enumerate(courts):
            ws.cell(row=1+i*16+(j+1), column=2, value= f"{j+1}番")
            ws.cell(row=1+i*16+(j+1), column=2+len(court), value= f"{j+1}番")
            for k, vote in enumerate(court[1:]):
                if vote=="unavailable":
                    vote = "休"
                ws.cell(row=1+i*16+(j+1), column=2+(k+1), value= f"{vote}")

    # Save the Excel file
    wb.save(file_path)


def get_vote_dest():
    # ファイルを検索するディレクトリのパスを指定
    import os

    BASE_DIR = os.path.join(os.getcwd(), "data")
    
    # ファイルを検索
    files = glob.glob(os.path.join(BASE_DIR, '*最終形態*'))
    
    # ファイルが見つからなかった場合の処理
    if not files:
        print("最終形態と書かれたファイルが見つかりませんでした")
    else:
        # ファイルの更新日時でソートし、最新のファイルを取得
        latest_file = max(files, key=os.path.getmtime)
        print("最新の最終形態ファイルのパス:", latest_file)
    df_path = latest_file

    wb = openpyxl.load_workbook(df_path)
    ws = wb.worksheets[0]

    vote_dest = pd.DataFrame(columns= ["date","time","court"])
    row_idx = 1
    n =0
    while 1+(row_idx-1)*16 < valid_row_count:
      date = ws[f"A{1+(row_idx-1)*16}"].value
      print(date)
      if date == None:
        row_idx += 1
        continue
      # 時間帯を取得
      times = [time.value for time in ws[f'B{1+(row_idx-1)*16}:F{1+(row_idx-1)*16}'][0]]
      # 1-12番コートの票数
      for i, row in enumerate(ws[f'B{2+(row_idx-1)*16}:F{1+(row_idx-1)*16+12}']):
        court = i+1
        for time, cell in zip(times, row):
          if cell.value == "休" or (rep := int(cell.value)) == 0:
            continue
          to_add = pd.DataFrame({"date": [date]*rep,"time": [time]*rep,"court": [court]*rep})
          vote_dest = pd.concat([vote_dest, to_add])
          n += rep
      row_idx+=1
    vote_dest = vote_dest.reset_index()
    vote_dest = vote_dest.drop("index", axis=1)

    return vote_dest
