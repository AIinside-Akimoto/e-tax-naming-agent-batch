# Architecture

## Overview

```text
InputFolder
   │
   ├─ 再帰走査 (.pdf / .txt / .csv)
   │
   ├─ 1ファイルずつ /assist-naming に送信
   │      └─ x-api-key 認証
   │
   ├─ API レスポンス
   │      └─ documents[0].rename / new_filename / extracted
   │
   ├─ リネーム可
   │      └─ OutputFolder に同一フォルダー構造でコピー
   │
   └─ リネーム不可
          ├─ CopyNonRenamedFiles=true なら元名コピー
          └─ false ならスキップ
```

## Main Design Points

- PowerShell 単体で配布しやすい構成
- Windows 業務端末でそのまま使いやすい `run.bat` 同梱
- API 応答をログへ残し、監査・検証しやすくする
- `SupportsShouldProcess` により `-WhatIf` で安全確認できる
