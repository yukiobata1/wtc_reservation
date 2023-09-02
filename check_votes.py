import pandas as pd
import numpy as np
import os
import sys
from tqdm import tqdm
import datetime
import utils
import time
import inspect
import json

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException


# デバッグ用
DEBUG = False

if utils.check_schedule_within_30_minutes() == 1:
  print("毎月第2水曜22:30～翌8:00,毎週金曜3:00～3:30は利用できません。")
  exit(0)

# 個人で利用する目的などで利用不可能な通し番号
unused = [31,]

service = Service()
options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=service, options=options)

# 各コートの抽選数を検索

# 空き状況
driver.get("https://www.pa-reserve.jp/eap-ri/rsv_ri/i/im-0.asp?KLCD=119999")

# result = defaultdict(dict)
result = dict()
with open("sample_file.json", "w") as file:
    json.dump(result, file)

court_list = [
'第１テニスコート第１クレーコート',
'第１テニスコート第２クレーコート',
'第１テニスコート第３人工芝コート',
'第１テニスコート第４人工芝コート',
'第１テニスコート第５人工芝コート',
'第１テニスコート第６人工芝コート',
'第１テニスコート第７人工芝コート',
'第１テニスコート第８人工芝コート',
'第１テニスコート第９人工芝コート',
'第１テニスコート第１０人工芝コート',
'第２テニスコート第１１人工芝コート',
'第２テニスコート第１２人工芝コート'
]

weekday_correspondence = {
  0: "月曜日",
  1: "火曜日",
  2: "水曜日",
  3: "木曜日",
  4: "金曜日",
  5: "土曜日",
  6: "日曜日",}

for i, court in enumerate(court_list):
  result[i] = dict()
  # 各コートについて抽選数確認
  vacancy_button = driver.find_element(By.XPATH, "//a[contains(text(),'施設の空き状況')]")
  vacancy_button.click()

  # 所在地
  select_by_place = driver.find_element(By.XPATH, "//a[text()='所在地から検索']")
  select_by_place.click()

  western_area = driver.find_element(By.XPATH, "//a[contains(text(),'西部エリア')]")
  western_area.click()

  proceed_button = driver.find_element(By.XPATH, "//input[contains(@value,'次へ')]")
  proceed_button.click()

  tokorozawa_park = driver.find_element(By.XPATH, "//a[contains(text(),'所沢航空記念公園')]")
  tokorozawa_park.click()

  # 来月1日のdatetimeオブジェクトを取得
  current_date = datetime.date.today()
  next_month_date = current_date.replace(day=1) + datetime.timedelta(days=32)
  next_month_date = next_month_date.replace(day=1)

  year_input = driver.find_element(By.NAME, 'selYear')
  month_input = driver.find_element(By.NAME, 'selMonth')
  day_input = driver.find_element(By.NAME, 'selDay')
  
  year_input.clear()
  year_input.send_keys(str(next_month_date.year))  # Set the year value
  
  month_input.clear()
  month_input.send_keys(str(next_month_date.month))  # Set the month value
  
  day_input.clear()
  day_input.send_keys(str(next_month_date.day))  # Set the day value
  
  # 各コート選択
  target_text = court
  xpath_expression = f"//input[@type='RADIO'][following-sibling::text()[1][contains(., '{target_text}')]]"
  radio_button = driver.find_element(By.XPATH, xpath_expression)
  radio_button.click()

  # 翌月1日のページに遷移
  submit_button = driver.find_element(By.XPATH, "//input[contains(@value, 'ＯＫ')]")
  submit_button.click()

  # 申請数表示
  show_votes = driver.find_element(By.XPATH, "//input[@value='申請数表示']")
  show_votes.click()
  
  # 1日-月末の各日についてループ
  next_month = next_month_date.month
  while (next_month_date.month == next_month):
    str_date = next_month_date.strftime("%m-%d")
    print(f"{str_date}")
    result[i][str_date] = dict()
    target_time_ranges = [
    "06:30-08:30", "08:30-10:30", "10:30-12:30",
    "12:30-14:30", "14:30-16:30", "16:30-18:30",
    "18:30-20:30", 
    ]

    for target_time_range in target_time_ranges:
      # 辞書のキー取得 
      start_time = target_time_range[:5]
      # 抽選可能の場合
      try:
        # 抽選数を取得
        xpath_expression = f'//font[@color="Blue"][preceding::text()[1][contains(., "{target_time_range}")]]'
        next_element = driver.find_element(By.XPATH, xpath_expression)

        lottery_count_text = next_element.text.strip()
        lottery_count = int(lottery_count_text.split('<')[-1].strip('>').split(';')[-1])        
        # 結果に抽選数を追加
        result[i][str_date][start_time] = lottery_count
      
      except NoSuchElementException:
        try:
          # 抽選数が見当たらないとき、その時間帯は使用不可
          xpath_expression = f'//font[@color="Red"][following::text()[1][contains(., "{target_time_range}")]]'
          next_element = driver.find_element(By.XPATH, xpath_expression)
          result[i][str_date][start_time] = "unavailable"
        except NoSuchElementException:
          # 18:30-20:30, court No.1~10は使用不可
          result[i][str_date][start_time] = "unavailable"

    
    # 次の日に
    next_day_button = driver.find_element(By.XPATH, '//input[contains(@value,"次日")]')
    next_day_button.click()

    next_month_date += datetime.timedelta(days=1)
     
  # メニューに戻って、別のコートの票数を取得
  to_menu = driver.find_element(By.XPATH, "//input[contains(@value,'メニュー')]")
  to_menu.click()

DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
from datetime import timedelta, timezone
tz = timezone(timedelta(hours=+9), 'Asia/Tokyo')
current_date = datetime.datetime.now(tz).strftime("_%m月%d日%H時")
file_path = os.path.join(DATA_BASE, f'投票数{current_date}.xlsx')
utils.save_votes(result, file_path)

# たたき台の作成
data_copy = result.copy()
for court, value in data_copy.items():
  for date, times in value.items():
    for time, number in times.items():
      if number == "unavailable":
        continue
      else:
        data_copy[court][date][time] = 1 if number == 0 else number*3 
  

file_path = os.path.join(DATA_BASE, f'たたき台{current_date}.xlsx')
utils.save_votes(data_copy, file_path)
