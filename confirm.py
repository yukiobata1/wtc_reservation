import pandas as pd
import numpy as np
import os

DATA_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
original = pd.read_excel(os.path.join(DATA_BASE, "埼玉県営テニスコート名義.xlsx", usecols="A:D", header=1)

# 使用可(1)or不可(0)
is_valid = [] 
# 個人で利用する目的などで利用不可能な通し番号
unused = [31, 259]

def preprocess(df, unused):
  # dfのnanを除外、使ってはならない番号も除外
  unused_indices = np.where(df["通し番号"].isin(unused))
  return df.dropna().drop(unused_indices[0])

df = preprocess(original, unused)

#for index, row in df.iterrows():
#  # 通し番号i番の抽選　
#  id = row["ID"]
#  password = row["パスワード"]

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

 
options = Options()
#options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
 

# 予約ボタンクリック
driver.get("https://www.pa-reserve.jp/eap-ri/rsv_ri/i/im-0.asp?KLCD=119999")
reservation_button = driver.find_element(By.XPATH, "/html/body/a[5]")
reservation_button.click()

# ログイン
userid = "0000098557"
password = "2118"

user_to_fill = driver.find_element(By.XPATH, "/html/body/form/input[1]")
user_to_fill.send_keys(userid)

pass_to_fill = driver.find_element(By.XPATH, "/html/body/form/input[2]")
pass_to_fill.send_keys(password)

ok_button = driver.find_element(By.XPATH, "/html/body/form/input[3]")
ok_button.click()

# 所在地

select_by_place = driver.find_element(By.XPATH, "/html/body/a[3]")
select_by_place.click()

western_area = driver.find_element(By.XPATH, "/html/body/form/a[1]")
western_area.click()

proceed_button = driver.find_element(By.XPATH, "/html/body/form/input[3]")
proceed_button.click()

tokorozawa_park = driver.find_element(By.XPATH, "/html/body/form/a[2]")
tokorozawa_park.click()


time.sleep(100)
