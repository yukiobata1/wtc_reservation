import pandas as pd
import numpy as np
import os
import sys
from tqdm import tqdm
from datetime import date
import utils

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.common.exceptions import NoSuchElementException

if check_schedule() == 0:
  print("毎月第2水曜22:30～翌8:00,毎週金曜3:00～3:30は利用できません。")
  exit(0)

DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
original = pd.read_excel(os.path.join(DATA_BASE, "埼玉県営テニスコート名義.xlsx"), usecols="A:D", header=1)

# 個人で利用する目的などで利用不可能な通し番号
unused = [31, 259]

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def check_available(userid: str, password: str, driver) -> int:
  # 使用不可→0を返す、使用可→1を返す
  
  # 予約ボタンクリック
  driver.get("https://www.pa-reserve.jp/eap-ri/rsv_ri/i/im-0.asp?KLCD=119999")
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
  # ここから確認用
  any_court = driver.find_element(By.XPATH, "//input[contains(@name,'rdo_SHISETU')]")
  any_court.click()
  
  ok_button = driver.find_element(By.XPATH, "//input[contains(@value,'ＯＫ')]")
  ok_button.click()
  
  # 有効期限切れ
  try:
    driver.find_element(By.XPATH, "//form[contains(text(),  '利用者の有効期限が切れています')]")
    return 0
    sys.exit()
  except NoSuchElementException:
    pass
  
  while 1:
    try: 
      available_court = driver.find_element(By.XPATH, "//input[contains(@name, 'chkComa')]")
      available_court.click()
      
      #実際に予約するわけではなく、登録停止の確認のため
      reservation_button = driver.find_element(By.XPATH, "//input[contains(@value,  '予約する')]")
      reservation_button.click()
      break
    
    except NoSuchElementException:
      # 次の日に遷移
      next_day = driver.find_element(By.XPATH, "//input[contains(@value,  '次日')]")
      next_day.click()
  
  # 利用停止の確認
  try:
    driver.find_element(By.XPATH, "//form[contains(text(),  '利用停止のため、予約することができません。')]")
    # print("この番号は利用停止中です。")
    return 0
  except NoSuchElementException:
    # print("この番号は使用可能です。")
    return 1
    
# 空の行を除去
df = df.dropna()

for index, row in tqdm(df.iterrows()):
  # 各番号についての利用可否確認
  
  # 利用不可能な番号はスキップ
  if row["通し番号"] in unused:
    availability.append(0)
  
  # 番号が利用できるかチェック
  else:
    userid = row["ID"]
    password = row["パスワード"]
    availability.append(check_available(userid, password, driver))
  
  print(f"  userid: {userid}, password: {password}, 利用可否: {availability[-1]}")

df["利用可否(1:可, 0:不可)"] = availability

# 作成したdataframeの保存
current_date = date.today().strftime("%y-%m-%d-%h-%m-%s")

file_path = os.path.join(DATA_BASE, f'所沢名義_{current_date}.xlsx')
df.to_excel(file_path, index=False)
