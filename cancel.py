import pandas as pd
import numpy as np
import os
import sys
import glob
from tqdm import tqdm
from datetime import date
import utils
import logging
logging.basicConfig(filename='cancel.log', encoding='utf-8', level=logging.ERROR)

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


DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
file_path = glob.glob(os.path.join(DATA_BASE, "埼玉県営利用可名義*.xlsx"))[0]
accounts = pd.read_excel(file_path,usecols="A:D", header=1)
# 空の行を除去
accounts = accounts.dropna()

# def select_element(element):
#   if not element.isSelected():
#     element.click()


def single_cancel(userid: str, password: str):
  # 使用不可→0を返す、使用可→1を返す
  options = Options()
  options.add_argument('--headless')
  options.add_argument('--no-sandbox')
  options.add_argument('--disable-dev-shm-usage')
  driver = webdriver.Chrome(service=Service(), options=options)
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

  reservation = driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO']")
  # select all the stuff
  if len(reservation) == 0:
    logging.error(f"No existing reservation for no.{userid}")
    
  for i, el in enumerate(reservation):
    import time
    driver.find_elements(By.XPATH, "//input[@name='rdoYoyakuNO']")[i].click()
    time.sleep(3)
    
    cancel_button = driver.find_element(By.XPATH, "//input[contains(@value, '取消')]")
    cancel_button.click()
    time.sleep(3)

    confirm_button = driver.find_element(By.XPATH, "//input[contains(@value, 'はい')]")
    confirm_button.click()
    time.sleep(3)

    driver.execute_script("window.history.go(-1)")
    time.sleep(3)
    driver.execute_script("window.history.go(-1)")
    time.sleep(3)

  # page_source = driver.page_source
  # Print all the text in the page
  # print(page_source)
  # print(f"{userid= }")
  # print(f"{password=}")

def multi_cancel(df):
  print(f"{df=}")
  for i, row in df.iterrows():
    id = row["ID"]
    password = row["パスワード"]

    userids = [
  "id='edamame012'\n",
 "id='waseda1102'\n",
 "id='0000016254'\n",
 "id='0000098257'\n",
 "id='0000098191'\n",
 "id='0000047470'\n",
 "id='0000063466'\n",
 "id='0000098568'\n",
 "id='0000074264'\n",
 "id='1J21E09400'\n",
 "id='tauff5kc11'\n",
 "id='05121wsnfm'\n",
 "id='0000026942'\n",
 "id='2611726117'\n",
 "id='0000066723'\n",
 "id='0000108447'\n",
 "id='RINOUE0513'\n",
 "id='0000098554'\n",
 "id='aiakanen24'\n",
 "id='0000098181'\n",
 "id='0000098190'\n",
 "id='0000047472'\n",
 "id='0000070759'\n",
 "id='hahahawaii'\n",
 'id=1113896841\n',
 "id='1J21E14900'\n",
 "id='0000026941'\n",
 "id='rikoko0107'\n",
 "id='0000095747'\n",
 "id='kota122525'\n",
 "id='0000088272'\n",
 "id='0000108444'\n",
 "id='0000066725'\n",
 "id='misso0202y'\n",
 "id='0000102825'\n",
 "id='0000062599'\n",
 "id='30294910Aa'\n",
 "id='2444666668'\n",
 "id='0000026349'\n",
 "id='fffffffff3'\n",
 "id='0000095746'\n",
 "id='1F19076806'\n",
 "id='0000100401'\n",
 "id='0000107855'\n",
 "id='0000066724'\n",
 "id='0000033574'\n",
 "id='0807032802'\n",
 "id='balcerona9'\n",
 "id='yasumodayo'\n",
 "id='0000098182'\n",
 "id='0000026104'\n",
 "id='amachan727'\n",
 "id='0000090652'\n",
 "id='0000087299'\n",
 "id='0000108452'\n",
 "id='0000033568'\n",
 "id='0000063364'\n",
 'id=1271271271\n',
 "id='watanaaabe'\n",
 "id='tagurikoyo'\n",
 'id=5070000209\n',
 "id='0000026528'\n",
 "id='0000093704'\n",
 "id='0000098187'\n",
 "id='0000120822'\n",
 "id='0000033628'\n",
 "id='0000088256'\n",
 "id='0000060386'\n",
 "id='0000085596'\n",
 "id='miyusariko'\n",
 "id='0000063464'\n",
 "id='kuro336699'\n",
 "id='atsuya0NNS'\n",
 "id='0000060204'\n",
 "id='0000108441'\n",
 "id='yutaito020'\n",
 "id='0000094712'\n",
 "id='0000114380'\n",
 "id='0000025307'\n",
 'id=1227272727\n',
 "id='andrew1341'\n",
 "id='1E22H11600'\n", 
 "ERROR:root:No existing reservation for no.waseda1102",
 "ERROR:root:No existing reservation for no.0000016776",
"ERROR:root:No existing reservation for no.edamame012",
"ERROR:root:No existing reservation for no.0000094711",
"ERROR:root:No existing reservation for no.misaki0522",
"ERROR:root:No existing reservation for no.misaki0522",
"ERROR:root:No existing reservation for no.0000016776",
"ERROR:root:No existing reservation for no.0000094711",
"ERROR:root:No existing reservation for no.0000099816",
"ERROR:root:No existing reservation for no.0000042117",
"ERROR:root:No existing reservation for no.0000066718",
"ERROR:root:No existing reservation for no.0000099797",
"ERROR:root:No existing reservation for no.0000066717",
"ERROR:root:No existing reservation for no.edamame012",
"ERROR:root:No existing reservation for no.aotaku0205",
"ERROR:root:No existing reservation for no.0456982453",
"ERROR:root:No existing reservation for no.0000016254",
"ERROR:root:No existing reservation for no.nkamatafai",
"ERROR:root:No existing reservation for no.waseda1102",
"ERROR:root:No existing reservation for no.1J21E09400",
"ERROR:root:No existing reservation for no.0000026942",
"ERROR:root:No existing reservation for no.1113896841",
"ERROR:root:No existing reservation for no.1J21E14900",
"ERROR:root:No existing reservation for no.0000026941",
"ERROR:root:No existing reservation for no.2444666668",
"ERROR:root:No existing reservation for no.0000026349",
"ERROR:root:No existing reservation for no.0000033574",
"ERROR:root:No existing reservation for no.0807032802",
"ERROR:root:No existing reservation for no.0000026104",
"ERROR:root:No existing reservation for no.0000033568",
"ERROR:root:No existing reservation for no.0000026528",
"ERROR:root:No existing reservation for no.1271271271",
"ERROR:root:No existing reservation for no.0000033628",
"ERROR:root:No existing reservation for no.0000060204",
]

    if any([(id in s) for s in userids]):
      print(f"{id} is already cleared")
      continue
    single_cancel(id, password)
    print(f"{id=}")

if __name__ == "__main__":
  p = Pool(16)
  import numpy as np
  ac_list = np.array_split(accounts, 16)
  p.map(multi_cancel, ac_list)
