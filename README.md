# 概要
所沢のコートの抽選、確定作業を自動化

最低限のことしか書いていないので、問題があったら気兼ねなくyukiobata1|at|gmail.comまで連絡をお願いします。(|at|を@に置き換え)

## 使い方の説明
作業を始める前に、
クラウドストレージのwtc_save内埼玉県営テニスコート名義.xlsx**以外**のすべてのファイルを消すこと！

### セットアップ
Linux上で作成したので、windowsではwslを起動して作業するのがわかりやすい。
以下のコードを実行してください。
```bash
git clone https://github.com/yukiobata1/wtc_reservation
docker run -d -p 4444:4444 --shm-size="2g" selenium/standalone-chrome:4.13.0-20230926`
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```
### 手順
実行する前に以下のコードを実行
```bash
source venv/bin/activate
cd wtc_reservation
```

ログは```nohup.out```に出力されるため、```cat nohup.out```を用いて適宜確認する。
動作が不安定で、一回の作業では投票しきれていない、ということが頻発するため確認は必須。

1. 使用可能アカウント、投票数を調べる
```bash
> nohup.out | nohup python check_accounts.py &
> nohup.out | nohup python check_votes.py &
```

2. (会計から投票数のxlsxファイルを受け取ったら)投票する
```bash
> nohup.out | nohup python voting.py &
```

3. 抽選後に確定する。
```bash
> nohup.out | nohup python kakutei.py &
> nohup.out | nohup python get_won_votes.py &

```

### file_name 一覧
1. check_accounts.py
利用停止になっていない名義を取得
2. check_votes.py
今各コートに入っている票数を取得、入れる票のたたき台を作成
3. voting.py
与えられたxlsxシートを読み取り、投票する
5. kakutei.py
各アカウントで、当選した票に対して確定作業をする
6. get_won_votes.py
確定作業された票を保存する。
=====以上は使う=======
7. check_votes_after(未完成)
票が適切に入っているかどうか確認
8. cancel.py
入れた抽選票をすべてキャンセルする。合計10時間程度かかり、不安定
9. utils.py
xlsxを作成するなど、ヘルパー関数が入っている(直接は使わない。)

# エラー
500 errorの場合、時間をおいて(1時間くらい)やってみるとうまくいくかも

# 注意点
DATAとgs://wtc_saveから前のファイルは消す。
