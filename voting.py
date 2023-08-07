import pandas as pd
import numpy as np
import os
import sys
from tqdm import tqdm
import datetime
import utils
from collections import defaultdict

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.common.exceptions import NoSuchElementException

# デバッグ用
DEBUG = True

if utils.check_schedule_within_30_minutes() == 1:
  print("毎月第2水曜22:30～翌8:00,毎週金曜3:00～3:30は利用できません。")
  exit(0)

DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
accounts = pd.read_excel(os.path.join(DATA_BASE, "埼玉県営テニスコート名義.xlsx"), usecols="A:D", header=1)
# 空の行を除去
accounts = accounts.dropna()

# 各コート選択
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

time_conversion = {
  "6:30": "06:30-08:30",
  "8:30": "08:30-10:30",
  "10:30": "10:30-12:30", 
  "12:30": "12:30-14:30", 
  "14:30": "14:30-16:30",
  "16:30": "16:30-18:30",
}

options = Options()
# options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(), options=options)

# 予約ボタンクリック
driver.get("https://www.pa-reserve.jp/eap-ri/rsv_ri/i/im-0.asp?KLCD=119999")

def single_vote(date, time, court, userid, password):
  # 与えられた予約を実行
  reservation_button = driver.find_element(By.XPATH, "//*[text()='施設の予約']")
  reservation_button.click()
  
  user_to_fill = driver.find_element(By.XPATH, "//input[@name='txtUserCD']")
  user_to_fill.send_keys(userid)
  
  pass_to_fill = driver.find_element(By.XPATH, "//input[@name='txtPassword']")
  pass_to_fill.send_keys(password)
  
  ok_button = driver.find_element(By.XPATH, "//input[contains(@value,'ＯＫ')]")
  ok_button.click()

  # 所在地
  select_by_place = driver.find_element(By.XPATH, "//a[text()='所在地から検索／予約']")
  select_by_place.click()
  
  western_area = driver.find_element(By.XPATH, "//a[contains(text(),'西部エリア')]")
  western_area.click()
  
  proceed_button = driver.find_element(By.XPATH, "//input[contains(@value,'次へ')]")
  proceed_button.click()
  
  tokorozawa_park = driver.find_element(By.XPATH, "//a[contains(text(),'所沢航空記念公園')]")
  tokorozawa_park.click()

  # 来月の年数取得
  current_date = datetime.date.today()
  next_month_date = current_date.replace(day=1) + datetime.timedelta(days=32)
  next_month_date = next_month_date.replace(day=1)

  # 日時指定
  date = datetime.datetime.strptime(date, "%m-%d")

  print(f"{driver.find_element(By.XPATH, '//title').text=}")
  year_input = driver.find_element(By.NAME, 'selYear')
  month_input = driver.find_element(By.NAME, 'selMonth')
  day_input = driver.find_element(By.NAME, 'selDay')
  
  year_input.clear()
  year_input.send_keys(str(next_month_date.year))  # Set the year value
  
  month_input.clear()
  month_input.send_keys(str(date.month))  # Set the month value
  
  day_input.clear()
  day_input.send_keys(str(date.day))  # Set the day value

  # コートを指定
  target_text = court_list[court]
  xpath_expression = f"//input[@type='RADIO'][following-sibling::text()[1][contains(., '{target_text}')]]"
  radio_button = driver.find_element(By.XPATH, xpath_expression)
  radio_button.click()

  submit_button = driver.find_element(By.XPATH, "//input[contains(@value, 'ＯＫ')]")
  submit_button.click()

  xpath_expression = f"//input[@name='chkComa'][following-sibling::text()[1][contains(., '{time_conversion[time]}')]][following-sibling::font[@color='Blue']]"
  time_check = driver.find_element(By.XPATH, xpath_expression)
  time_check.click()
  
  reservation_button = driver.find_element(By.XPATH, "//input[contains(@value,  '予約する')]")
  reservation_button.click()

  # 確認画面
  proceed_button = driver.find_element(By.XPATH, "//input[contains(@value,'次へ')]")
  proceed_button.click()

  # 正しく遷移できているか確認
  # 日付
  date_exp = date.strftime("%m/%d")
  WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{date_exp}')]"))
    )
  # 時間
  # start_time, end_time = time_conversion[time].split("-")
  # print(f"{start_time, end_time=}")
  # WebDriverWait(driver, 10).until(
  #      EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{start_time}')]"))
  #  )
  #WebDriverWait(driver, 10).until(
  #      EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{end_time}')]"))
  #  )
  # コート
  WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{court_list[court]}')]"))
    )
  # 予約実行
  if DEBUG==False:
    driver.find_element(By.XPATH, f"//input[contains(@value, '予約実行')]")

  # メニューに戻って、別のコートの票数を取得
  to_menu = driver.find_element(By.XPATH, "//input[contains(@value,'メニュー')]")
  to_menu.click()

if __name__ == "__main__":
  vote_dest = pd.read_csv(os.path.join(DATA_BASE, "vote_dest.csv"))

used_votes = defaultdict(lambda: 0)
for i, row in tqdm(vote_dest.iterrows()):
  date = row.date
  time = row.time
  court = row.court
  user = accounts[accounts["通し番号"] == row.account]
  userid = user["ID"].iloc[0]
  password = user["パスワード"].iloc[0]
  print(f"{date, time, court, user, userid, password=}")
  
  single_vote(date=date, time=time, court=court, userid=userid, password=password)
  used_votes[row.account] += 1
  remain_votes = pd.DataFrame({"通し番号":  list(accounts["通し番号"]), "残り票数": [4-a[idx] for idx in list(accounts["通し番号"])]})
  remain_votes.to_csv(os.path.join(DATA_BASE, "remain_votes.csv"))