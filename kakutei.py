import pandas as pd
import numpy as np
import os
import sys
import glob
from tqdm import tqdm
from datetime import date
import utils
import logging
logging.basicConfig(filename='check_votes.log', encoding='utf-8', level=logging.ERROR)

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.common.exceptions import NoSuchElementException
from multiprocessing import Pool


if utils.check_schedule_within_30_minutes() == 1:
  print("毎月第2水曜22:30～翌8:00,毎週金曜3:00～3:30は利用できません。")
  exit(0)

options = Options()
# options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(), options=options)

def single_kakutei(num, userid: str, password: str):
  # 各アカウントに対して、確定作業を行う
  
  driver.get("https://www.pa-reserve.jp/eap-ri/rsv_ri/i/im-0.asp?KLCD=119999")
  for_cancel = driver.find_element(By.XPATH, "//a[contains(text(), '予約の確認／取消')]")
  for_cancel.click()
  
  user_to_fill = driver.find_element(By.XPATH, "//input[@name='txtUserCD']")
  user_to_fill.send_keys(userid)
  
  pass_to_fill = driver.find_element(By.XPATH, "//input[@name='txtPassword']")
  pass_to_fill.send_keys(password)
  
  ok_button = driver.find_element(By.XPATH, "//input[contains(@value,'ＯＫ')]")
  ok_button.click()

  proceed_button = driver.find_element(By.XPATH, "//input[contains(@value,'次へ')]")
  proceed_button.click()

  # 未確定の抽選のみ確認ボタン
  # 4, 5は抽選
  exclude_list = [0, 1, 2, 3, 8]
  for exclude_n in exclude_list:
    el = driver.find_element(By.XPATH, f"//input[@name='chk_tSta{exclude_n}']")
    el.click()

  pref_park = driver.find_element(By.XPATH, f"//input[@name='chkKubun'][following-sibling::text()[1][contains(., '県営公園')]]")
  pref_park.click()
  
  proceed_button = driver.find_element(By.XPATH, "//input[contains(@value,'次へ')]")
  proceed_button.click()

  import time as t
  t.sleep(1000)

  before_confirm = driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO'][following-sibling::text()[1][contains(., '当選未確定')]]")
  # need to be modified
  after_confirm = driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO'][following-sibling::text()[1][contains(., '抽選前')]]")
  # select all the stuff
  if len(before_confirm) == 0:
    print(f"No entries before confirmation")

  if len(after_confirm) == 0:
    print(f"No entries after confirmation")

  # 確定後の当選票
  for i, el in enumerate(after_confirm):
    # すでにcsv内に入ってなければ保存
    driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO'][following-sibling::text()[1][contains(., '当選確定済')]]")[i].click()

    votes_won = pd.read_csv(os.path.join(DATA_BASE, "votes_won.csv"))[["date", "court", "time_range", "通し番号", "userid", "password", "予約申請番号"]]

    confirmation_button = driver.find_element(By.XPATH, "//input[contains(@value, '確認')]")
    confirmation_button.click()
    
    text = driver.find_element(By.XPATH, "//form").text
    data = extract_kakutei(text)

    driver.execute_script("window.history.go(-1)")

    print(data)
    print(f"{num=}, {userid=}, {password=}")
    
    to_add = pd.DataFrame({"date":[data["date"]], "court":[data["court"]], "time_range": [data["time_range"]], "通し番号": [num], "userid": [userid], "password": [password], "予約申請番号":[data["reservation_number"]]})
    votes_won = pd.concat([votes_won, to_add])
    votes_won.to_csv(os.path.join(DATA_BASE, "votes_won.csv"))

  
  # 確定作業前の当選票
  for i, el in enumerate(before_confirm):
    # 確定し、エントリーに保存
    driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO'][following-sibling::text()[1][contains(., '当選未確定')]]")[i].click()

    votes_won = pd.read_csv(os.path.join(DATA_BASE, "votes_won.csv"))[["date", "court", "time_range", "通し番号", "userid", "password", "予約申請番号"]]

    confirmation_button = driver.find_element(By.XPATH, "//input[contains(@value, '確認')]")
    confirmation_button.click()
    
    form = driver.find_element(By.XPATH, "//form")
    text = form.text
    text = text.split("\n")
    court = text[text.index("◇施設名")+1]
    date = text[text.index("◇予約日")+1]
    time_range = text[text.index("◇使用時間")+1]
    
    driver.execute_script("window.history.go(-1)")
    import time as t
    t.sleep(10)

    # 確定
    kakutei_button = driver.find_element(By.XPATH, "//input[contains(@value, '利用確定')]")
    kakutei_button.click()

    yes_button = driver.find_element(By.XPATH, "//input[contains(@value, 'はい')]")
    yes_button.click()

    form = driver.find_element(By.XPATH, "//form")
    text = form.text
    string = string.split("確定後の予約申請番号は以下のとおりです。")
    reservation_number = string[1].replace("\n", "")

    print(f"{date=}, {court=}, {time_range=}, {num=}, {userid=}, {password=}, {reservation_number=}")
    
    to_add = pd.DataFrame({"date":[date], "court":[court], "time_range": [time_range], "通し番号": [num], "userid": [userid], "password": [password], "予約申請番号":[reservation_number]})
    votes_won = pd.concat([votes_won, to_add])
    votes_won.to_csv(os.path.join(DATA_BASE, "votes_won.csv"))


    


if __name__ == "__main__":
  DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
  file_path = glob.glob(os.path.join(DATA_BASE, "埼玉県営利用可名義*.xlsx"))[0]
  accounts = pd.read_excel(file_path,usecols="A:D", header=1)
  # 空の行を除去
  accounts = accounts.dropna()

  try:
    votes_won = pd.read_csv(os.path.join(DATA_BASE, "votes_won.csv"))[["date", "court", "time_range", "通し番号", "userid", "password", "予約申請番号"]]
  except FileNotFoundError:
    print("create a new votes_won")
    votes_won = pd.DataFrame([], columns = ["date", "court", "time_range", "通し番号", "userid", "password", "予約申請番号"])
    votes_won.to_csv(os.path.join(DATA_BASE, "votes_won.csv"))
    
  for i, row in accounts.iterrows():
    print(f"{row=}")
    print(row["通し番号"], row["ID"], row["パスワード"])
    single_kakutei(row["通し番号"], row["ID"], row["パスワード"])
