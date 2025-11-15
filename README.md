# CSV-GAS
Google Apps Script を使って請求書CSVと作業報告書CSVを自動生成するアプリ。 得意先マスタ・タイムチャージ・ケースデータマスタなど複数DBを参照します。
# Ambitious 請求書・作業報告書 CSV Generator (GAS)

Google Apps Script を使って  
「請求書CSV」と「作業報告書CSV（タイムチャージ）」  
を自動生成するアプリケーションです。

本ツールは Google スプレッドシートをDBとして参照し、  
Web アプリ（デプロイ型）として動作します。

---

## 📌 構成

gas/ # Apps Script コード
customers.gs # 得意先マスタ → 請求書CSV
timecharge.gs # タイムチャージ → 作業報告書CSV
util.gs # 日付処理・CSV変換など共通処理

ui/ # HTML/JS UI
ui_customers.html # 請求書CSV画面
ui_timecharge.html# 作業報告書CSV画面

docs/ # 仕様書（Workspace が読む重要ファイル）
spec_customers.md # 請求書抽出ロジック
spec_timecharge.md# タイムチャージ抽出ロジック
mapping_customers.md
mapping_timecharge.md

yaml
コードをコピーする

---

## 📌 使用するデータベース（Spreadsheet）

| 名称 | シート名 | 説明 |
|-----|----------|------|
| 得意先マスタ | 得意先マスタ | 請求書CSVの元データ |
| タイムチャージ | タイムチャージ | 作業報告書CSVの元データ |
| ケースデータマスタ | ケースデータマスタ | 案件名などを参照 |

---

## 📌 請求書CSVの仕様（要約）

- 当月請求 → pivot = 請求締日 + 1ヶ月
- 翌月請求 → pivot = 請求締日
- 取引開始日 ≤ pivot ≤ 取引終了日 で抽出
- 発行日 = 実行日
- 支払期限 = 発行月の月末

詳細 → docs/spec_customers.md

---

## 📌 作業報告書CSV（タイムチャージ）の仕様（要約）

- 発生日が「実行日から過去3ヶ月以内」
- 請求データ未作成（FLG空）だけ対象
- 1行1明細で出力
- ケース名/案件名はケースデータマスタから紐付ける予定

詳細 → docs/spec_timecharge.md

---

## 📌 今後の実装予定
- 案件名の自動関連付け
- UI統合（請求書＋作業報告書）
- 処理ログのUI出力
- バリデーション強化
