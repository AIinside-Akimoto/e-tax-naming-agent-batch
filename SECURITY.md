# Security Policy

脆弱性や秘密情報の取り扱いに関する注意事項です。

## Report a Vulnerability

公開 Issue に機密情報を書かないでください。  
API キーや内部 URL が関係する場合は、公開前に伏せて報告してください。

## Secret Handling

- API キーは Git にコミットしない
- 実運用 URL は必要最小限の範囲で共有する
- `settings.example.json` はサンプル用のみとする
- ログ共有時は機密ファイル名や API 応答を必要に応じてマスクする
