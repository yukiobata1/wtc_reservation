# 概要
所沢のコートの抽選、確定作業を自動化するコード

## 使い方の説明
作業を始める前に、
クラウドストレージのwtc_save内埼玉県営テニスコート名義.xlsx**以外**のすべてのファイルを消すこと！

### 使用可能アカウント、投票数を調べる
まず、
```python
python check_accounts.py
```
1. check_accounts.py
利用停止になっていない名義を取得
2. check_votes.py
今各コートに入っている票数を取得、入れる票のたたき台を作成
3. voting.py 与えられたxlsxシートを読み取り、投票する
4. check_votes_after(未完成)
票が適切に入っているかどうか確認
5. kakutei.py
各アカウントで、当選した票に対して確定作業をする
6. votes_won.py
確定作業された票を保存する。
7. cancel.py
入れた抽選票をすべてキャンセルする。合計10時間程度かかり、不安定
8. utils.py
xlsxを作成するなど、ヘルパー関数が入っている

# エラー
500 errorの場合、時間をおいて(1時間くらい)やってみるとうまくいくかも

# 注意点
DATAとgs://wtc_saveから前のファイルは消す。
