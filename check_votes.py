import pandas as pd
import numpy as np
import os
import sys
from tqdm import tqdm
from datetime import date
import utils
import time

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

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 各コートの抽選数を検索

# 空き状況
driver.get("https://www.pa-reserve.jp/eap-ri/rsv_ri/i/im-0.asp?KLCD=119999")
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

for court in court_list:
  # 各コートについて抽選数確認　
  children = driver.find_elements(By.XPATH, '/html/body/form/*')
  for i, child in enumerate(children):
    print(f"{i=}")
    print(f"{child.textContent=}")


#   while 1:
#     try: 
#       available_court = driver.find_element(By.XPATH, "//input[contains(@name, 'chkComa')]")
#       available_court.click()

#       #実際に予約するわけではなく、登録停止の確認のため
#       reservation_button = driver.find_element(By.XPATH, "//input[contains(@value,  '予約する')]")
#       reservation_button.click()
#       break

#     except NoSuchElementException:
#       # 次の日に遷移
#       next_day = driver.find_element(By.XPATH, "//input[contains(@value,  '次日')]")
#       next_day.click()

#   # 利用停止の確認
#   try:
#     driver.find_element(By.XPATH, "//form[contains(text(),  '利用停止のため、予約することができません。')]")
#     # print("この番号は利用停止中です。")
#     return 0
#   except NoSuchElementException:
#     # print("この番号は使用可能です。")
#     return 1

# # 空の行を除去
# df = df.dropna()
# availability = []

# for index, row in tqdm(df.iterrows()):
#   # 各番号についての利用可否確認

#   # デバッグ時はdebug_list内の番号のみ
#   if DEBUG == True:
#     if row["通し番号"] not in debug_list:
#       availability.append(0)
#       continue
#   # 利用不可能な番号はスキップ
#   if row["通し番号"] in unused:
#     availability.append(0)
#   # 番号が利用できるかチェック
#   else:
#     userid = row["ID"]
#     password = row["パスワード"]
#     availability.append(check_available(userid, password, driver))
  
#   print(f"  userid: {userid}, password: {password}, 利用可否: {availability[-1]}")

# df["利用可否(1:可, 0:不可)"] = availability
# available_df = df[df["利用可否(1:可, 0:不可)"] == 1]

# # 作成したdataframeの保存
# current_date = date.today().strftime("[%m-%d]")

# file_path = os.path.join(DATA_BASE, f'埼玉県営利用可名義{current_date}.xlsx')
# utils.save_to_excel(available_df, file_path)

