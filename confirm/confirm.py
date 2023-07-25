import pandas as pd
import numpy as np

original = pd.read_excel("/content/drive/MyDrive/埼玉県営テニスコート名義.xlsx", usecols="A:D", header=1)

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

# テスト
id = "misaki0522"
password = "1214"


