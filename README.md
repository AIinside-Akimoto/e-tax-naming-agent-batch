# Naming Agent Batch Tool

AI エージェント API を利用して PDF 等のファイル名を自動生成し、
指定フォルダーへコピー・整理する PowerShell + BAT ツールです。

主に **請求書 / 見積書 / 契約書などの書類整理の自動化**を目的としています。

---

# 概要

このツールは以下を行います。

1. 指定フォルダー内のファイルを取得
2. AI エージェント API に送信
3. AI が判断したファイル名へ変更
4. 出力フォルダーへコピー
5. 元ファイルを Original フォルダーへ整理（デフォルト）

さらに

- CSVログによるトレーサビリティ確保
- SHA-256 による改ざん検知
- 並列 API 実行

に対応しています。

---

# ディレクトリ構成

```
naming-agent/
│
├─ Invoke-NamingAgentBatch.ps1
├─ ファイル命名アシスタント.bat
├─ README.md
│
├─ Input/
├─ Output/
├─ Original/
│   └─ yyyyMMddHHmmss/
└─ logs/
```

---

# 必要条件

- Windows
- PowerShell 5.1 以上
- Leapnet AI エージェント API

---

# PowerShell 実行方法

```
powershell -ExecutionPolicy Bypass -File Invoke-NamingAgentBatch.ps1 `
  -InputFolder "C:\work\Input" `
  -OutputFolder "C:\work\Output" `
  -ApiBaseUrl "YOUR_API_URL" `
  -ApiKey "YOUR_API_KEY"
```

---

# BAT 実行方法

通常は BAT を利用します。

```
ファイル命名アシスタント.bat
  "InputFolder"
  "OutputFolder"
  "ApiBaseUrl"
  "ApiKey"
```

例

```
ファイル命名アシスタント.bat ^
"C:\work\naming-agent\Input" ^
"C:\work\naming-agent\Output" ^
"YOUR_API_URL" ^
"YOUR_API_KEY"
```

---

# BAT 引数

|番号|内容|必須|説明|
|---|---|---|---|
|1|入力フォルダー|Yes|処理対象ファイルが格納されているフォルダー|
|2|出力フォルダー|Yes|リネーム後ファイルの出力先|
|3|API Base URL|Yes|Leapnet AI Agent のエンドポイント|
|4|API Key|Yes|認証用 API キー|
|5|Timeout 秒|No|API タイムアウト秒（既定: 600）|
|6|CopyOriginal|No|リネーム対象外ファイルを Output にコピーするか（true/false）|
|7|ログファイルパス|No|ログ出力先ファイル|
|8|Parallelism|No|並列実行数（未指定時は自動計算）|
|9|OrganizeSourceFilesAfterCopy|No|元ファイルを Original に移動するか（既定: true）|
|10|OriginalFolder|No|元ファイル保管先フォルダー（未指定時は Output 親配下の Original）|
|11|MappingCsvPath|No|mapping.csv の出力先（未指定時は Output 配下に自動生成）|

---

# Parallelism (並列実行)

AI API を並列実行します。

未指定の場合は **CPU から自動決定**されます。

```
Parallelism = CPUコア数 - 1
最低 = 2
最大 = 8
```

例

```
Parallelism = 5
```

---

# 実行例（フルオプション）

```
ファイル命名アシスタント.bat ^
"C:\work\naming-agent\Input" ^
"C:\work\naming-agent\Output" ^
"YOUR_API_URL" ^
"YOUR_API_KEY" ^
600 ^
false ^
"C:\work\naming-agent\logs\rename.log" ^
5 ^
true ^
"C:\work\naming-agent\Original" ^
"C:\work\naming-agent\logs\mapping.csv"
```

---

# ログ

ログファイルを指定すると処理ログが保存されます。

```
rename.log
```

例

```
[INFO] InputFolder=C:\work\naming-agent\Input
[INFO] Parallelism=4
[INFO] RENAMED file=invoice.pdf
```

---

# mapping.csv

```
処理日時,処理結果,元ファイルフルパス,元ファイル名,検索用ファイル名,検索用ファイルフルパス,保管用ファイルフルパス
```

---

# 主な機能

- AIによるファイル命名
- Original保管（改ざん防止）
- CSVログ出力
- SHA-256記録
- 並列処理
- 日本語対応

---

# ライセンス

Internal Tool / PoC Use
