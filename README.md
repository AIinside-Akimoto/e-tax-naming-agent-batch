# Naming Agent Batch Tool

AI エージェント API を利用して PDF 等のファイル名を自動生成し、
指定フォルダーへコピーする PowerShell + BAT ツールです。

主に **請求書 / 見積書 / 契約書などの書類整理の自動化**を目的としています。

---

# 概要

このツールは以下を行います。

1. 指定フォルダー内のファイルを取得
2. AI エージェント API に送信
3. AI が判断したファイル名へ変更
4. 出力フォルダーへコピー

オプションで

- 元ファイルもコピー
- 元ファイル整理
- 並列 API 実行

などを制御できます。

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
│   └─ *.pdf
│
└─ Output/
```

---

# 必要条件

- Windows
- PowerShell 5.1 以上
- Leapnet AI Agent API

---

# PowerShell 実行方法

```
powershell -ExecutionPolicy Bypass -File Invoke-NamingAgentBatch.ps1 `
  -InputFolder "C:\work\Input" `
  -OutputFolder "C:\work\Output" `
  -ApiBaseUrl "https://example.com" `
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
"https://stg-agent.leapnet.com/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" ^
"your-api-key"
```

---

# BAT 引数

|番号|内容|必須|
|---|---|---|
|1|入力フォルダー|Yes|
|2|出力フォルダー|Yes|
|3|API Base URL|Yes|
|4|API Key|Yes|
|5|Timeout 秒|No|
|6|CopyOriginal (true/false)|No|
|7|ログファイルパス|No|
|8|Parallelism|No|
|9|OrganizeSourceFilesAfterCopy|No|

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

# OrganizeSourceFilesAfterCopy

有効にすると以下の整理を行います。

### 1 コピー成功したファイル

```
Input フォルダーから削除
```

### 2 コピー対象外ファイル

```
対象外
```

フォルダーへ移動します。

例

```
Input/
  請求書.pdf

Output/
  2024-01-01_会社名_10000.pdf

対象外/
  不明書類.pdf
```

---

# 実行例（フルオプション）

```
ファイル命名アシスタント.bat ^
"C:\work\naming-agent\Input" ^
"C:\work\naming-agent\Output" ^
"https://stg-agent.leapnet.com/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" ^
"your-api-key" ^
600 ^
false ^
"C:\work\naming-agent\rename.log" ^
5 ^
true
```

---

# ログ

ログファイルを指定すると処理ログが保存されます。

```
rename.log
```

例

```
[INFO] InputFolder=C:\work\Input
[INFO] Parallelism=4
[INFO] RENAMED file=invoice.pdf
```

---

# 出力ファイル名例

```
2024-01-01_株式会社サンプル_10000_20260310123456.pdf
```

形式

```
取引日_取引先名_金額_処理日時.pdf
```

---

# 主な機能

- AIによるファイル名生成
- 並列 API 実行
- 元ファイルコピー
- 元ファイル整理
- 詳細ログ出力
- 日本語ファイル名対応

---

# ライセンス

Internal Tool / PoC Use
