# 概要
所沢のコートの抽選、確定作業を自動化するコード

## 各コードの説明
(画像付き説明書がいる。)
作業を始める前に、1. クラウドストレージのwtc_save内, 2.data内に入っている、埼玉県営テニスコート名義.xlsx**以外**のすべてのファイルを消す(dataフォルダ外に退避する)こと！
1. check_accounts.py
利用停止になっていない名義を取得
2. check_votes.py
今各コートに入っている票数を取得、入れる票のたたき台を作成
3. voting.py

与えられたxlsxシートを読み取り、投票する
5. check_votes_after(未完成)
票が適切に入っているかどうか確認
6. kakutei.py
各アカウントで、当選した票に対して確定作業をする
7. votes_won.py
確定作業された票を保存する。
8. cancel.py
入れた抽選票をすべてキャンセルする。合計10時間程度かかり、不安定
10. utils.py
xlsxを作成するなど、ヘルパー関数が入っている
