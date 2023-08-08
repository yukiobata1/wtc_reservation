import pandas as pd
import numpy as np
import os
import sys
import glob
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


if utils.check_schedule_within_30_minutes() == 1:
  print("毎月第2水曜22:30～翌8:00,毎週金曜3:00～3:30は利用できません。")
  exit(0)


DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
file_path = glob.glob(os.path.join(DATA_BASE, "埼玉県営利用可名義*.xlsx"))[0]
accounts = pd.read_excel(file_path,usecols="A:D", header=1)
# 空の行を除去
accounts = accounts.dropna()

options = Options()
# options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(), options=options)

# def select_element(element):
#   if not element.isSelected():
#     element.click()

def single_cancel(userid: str, password: str):
  # 使用不可→0を返す、使用可→1を返す
  
  # 予約ボタンクリック
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
  exclude_list = [0, 1, 2, 3, 6, 7, 8]
  for exclude_n in exclude_list:
    el = driver.find_element(By.XPATH, f"//input[@name='chk_tSta{exclude_n}']")
    el.click()

  pref_park = driver.find_element(By.XPATH, f"//input[@name='chkKubun'][following-sibling::text()[1][contains(., '県営公園')]]")
  pref_park.click()
  
  proceed_button = driver.find_element(By.XPATH, "//input[contains(@value,'次へ')]")
  proceed_button.click()

  reservation = driver.find_element(By.XPATH, "//input[@name='rdoYoyakuNO']")
  # # select all the stuff
  # [select_element(el) for el in all_reservation] 

  cancel_button = driver.find_element(By.XPATH, "//input[contains(@value, '取消')]")
  cancel_button.click()

  page_source = driver.page_source
  # Print all the text in the page
  print(page_source)

  confirm_button = driver.find_element(By.XPATH, "//input[contains(@value, 'はい')]")
  confirm_button.click()

id = "waseda1102"
password = "1102"

single_cancel(id, password)
print("success")
