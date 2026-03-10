# e-tax-naming-agent-batch

> 電子帳簿保存法向けの **ファイル名検索性** を高めるために、命名支援 API を使って  
> `取引日付-取引先-金額.拡張子` 形式へ自動リネームして保存する PowerShell / BAT ツールです。

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Windows](https://img.shields.io/badge/Windows-Supported-success)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

## ✨ Features

- 入力フォルダー配下を **再帰的** にスキャン
- `.pdf` / `.txt` / `.csv` を API に送信
- API の `documents[0].new_filename` を使って自動リネーム
- 出力先に **元フォルダー構造を維持** してコピー
- `rename=false` のファイルも必要に応じて元名コピー
- 元ファイルの **作成日時 / 更新日時 / アクセス日時** を保持
- API の生 JSON を含む **詳細ログ** を出力
- `-WhatIf` / `-Confirm` に対応

---

## 📦 Repository Structure

```text
.
├─ Invoke-NamingAgentBatch.ps1
├─ ファイル命名アシスタント.bat
├─ settings.example.json
├─ README.md
├─ LICENSE
├─ CHANGELOG.md
├─ CONTRIBUTING.md
├─ SECURITY.md
├─ .gitignore
├─ docs/
│  ├─ architecture.md
│  ├─ api-response-example.json
│  └─ benchmark-checklist.md
└─ examples/
   ├─ sample-run.ps1
   └─ sample-run.bat
```

---

## 🚀 Quick Start

### 1. BAT で実行

```bat
ファイル命名アシスタント.bat "C:\work\input" "C:\work\output" "https://example.contoso.com" "YOUR_API_KEY"
```

### 2. リネームできないファイルもコピーする

```bat
ファイル命名アシスタント.bat "C:\work\input" "C:\work\output" "https://example.contoso.com" "YOUR_API_KEY" 900 true "C:\work\logs\rename.log"
```

### 3. PowerShell から直接実行

```powershell
.\Invoke-NamingAgentBatch.ps1 `
  -InputFolder "C:\work\input" `
  -OutputFolder "C:\work\output" `
  -ApiBaseUrl "https://example.contoso.com" `
  -ApiKey "YOUR_API_KEY" `
  -Verbose
```

### 4. 事前確認だけしたい場合

```powershell
.\Invoke-NamingAgentBatch.ps1 `
  -InputFolder "C:\work\input" `
  -OutputFolder "C:\work\output" `
  -ApiBaseUrl "https://example.contoso.com" `
  -ApiKey "YOUR_API_KEY" `
  -WhatIf
```

---

## 🧠 Expected API Behavior

このツールは、以下の命名支援 API を想定しています。

- 認証: `x-api-key`
- エンドポイント: `/assist-naming`
- 送信形式: `multipart/form-data`
- フォーム項目: `file`

期待レスポンス例:

```json
{
  "summary": {
    "total": 1,
    "rename_true_count": 1,
    "rename_false_count": 0
  },
  "documents": [
    {
      "original_filename": "請求書_2026年2月分.pdf",
      "rename": true,
      "new_filename": "2026-02-28-株式会社サンプル商事-110000.pdf",
      "notes": "",
      "extracted": {
        "date": "2026-02-28",
        "vendor": "株式会社サンプル商事",
        "amount": "110000"
      }
    }
  ]
}
```

---

## ⚙️ Parameters

| Parameter | Required | Default | Description |
|---|---:|---:|---|
| `InputFolder` | Yes | - | 入力フォルダー |
| `OutputFolder` | Yes | - | 出力フォルダー |
| `ApiBaseUrl` | Yes | - | API ベース URL |
| `ApiKey` | Yes | - | `x-api-key` に設定するキー |
| `Timeout` | No | `600` | API 呼び出しタイムアウト秒 |
| `CopyNonRenamedFiles` | No | `false` | `rename=false` でも元名コピー |
| `LogFilePath` | No | auto | ログファイル出力先 |
| `PassThru` | No | `false` | サマリーをオブジェクト出力 |

---

## 🪵 Logging

ログには以下を記録します。

- 実行パラメーター
- 対象ファイル数
- API の生レスポンス JSON
- 抽出結果 `date / vendor / amount`
- 保存先パス
- スキップ理由
- エラー内容
- サマリー統計

---

## 🔒 Security Notes

- API キーはソースコードに直書きしない
- 公開リポジトリに実 URL / 実キーを含めない
- 必要に応じて `settings.example.json` をコピーしてローカル設定ファイルを作る
- 実運用前に `-WhatIf` で保存先を確認する

---

## 🧪 Benchmarking

電子帳簿保存法向けの評価観点は `docs/benchmark-checklist.md` にまとめています。

最低限確認したい項目:

- 取引日付の抽出精度
- 取引先の抽出精度
- 金額の抽出精度
- `rename=false` の判定妥当性
- サブフォルダー維持
- タイムスタンプ保持
- 同名衝突時の連番処理

---

## 🤝 Contributing

Issue / Pull Request 歓迎です。  
詳細は [CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

---

## 📄 License

MIT License. See [LICENSE](LICENSE).
