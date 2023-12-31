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

driver = utils.get_driver()

def single_check_vote(num, userid: str, password: str):
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

  reservation = driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO'][following-sibling::text()[1][contains(., '抽選前')]]")
  # select all the stuff
  if len(reservation) == 0:
    logging.error(f"No existing reservation for no.{userid}")
    
  for i, el in enumerate(reservation):
    driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO'][following-sibling::text()[1][contains(., '抽選前')]]")[i].click()

    exact_dest = pd.read_csv(os.path.join(DATA_BASE, "exact_dest.csv"))[["date", "court", "time_range", "通し番号", "userid", "password"]]

    confirmation_button = driver.find_element(By.XPATH, "//input[contains(@value, '確認')]")
    confirmation_button.click()
    
    form = driver.find_element(By.XPATH, "//form")
    text = form.text
    text = text.split("\n")
    court = text[text.index("◇施設名")+1]
    date = text[text.index("◇予約日")+1]
    time_range = text[text.index("◇使用時間")+1]
    print(f"{date=}, {court=}, {time_range=}, {num=}, {userid=}, {password=}")
    # print(exact_dest)

    to_add = pd.DataFrame({"date":[date], "court":[court], "time_range": [time_range], "通し番号": [num], "userid": [userid], "password": [password]})
    exact_dest = pd.concat([exact_dest, to_add])
    exact_dest.to_csv(os.path.join(DATA_BASE, "exact_dest.csv"))
    
    driver.execute_script("window.history.go(-1)")

if __name__ == "__main__":
  DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
  file_path = glob.glob(os.path.join(DATA_BASE, "埼玉県営利用可名義*.xlsx"))[0]
  accounts = pd.read_excel(file_path,usecols="A:D", header=1)
  # 空の行を除去
  accounts = accounts.dropna()

  import pickle

  try:
    exact_dest = pd.read_csv(os.path.join(DATA_BASE, "exact_dest.csv"))[["date", "court", "time_range", "通し番号", "userid", "password"]]
  except FileNotFoundError:
    print("create a new exact_dest")
    exact_dest = pd.DataFrame([], columns = ["date", "court", "time_range", "通し番号", "userid", "password"])
    exact_dest.to_csv(os.path.join(DATA_BASE, "exact_dest.csv"))
  
  print(f"{exact_dest=}")
  # try:
  #   exact_used_row = pd.read_csv(os.path.join(DATA_BASE, "exact_used_row.csv"))[["used_row"]]
  # except:
  #   print("create new exact_used_row")
  #   exact_used_row = pd.DataFrame({"used_row": []})
  #   exact_used_row.to_csv(os.path.join(DATA_BASE, "exact_used_row.csv"))
    
  for i, row in accounts.iterrows():
    print(f"{row=}")
    # print(f"{i=}")
    # print(f"{exact_used_row=}")
    # if i in exact_used_row["used_row"].values:
    #   print(f"{exact_used_row=}")
    #   print("skipped")
    #   continue
    print(row["通し番号"], row["ID"], row["パスワード"])
    single_check_vote(row["通し番号"], row["ID"], row["パスワード"])
    # exact_used_row.loc[len(exact_used_row)] = [i]
    # exact_used_row.to_csv(os.path.join(DATA_BASE, "exact_used_row.csv"), index=False)
