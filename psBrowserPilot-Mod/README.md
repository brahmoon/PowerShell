# psBrowserPilot

psBrowserPilot は、NodeFlow Local をベースにした PowerShell Visual Builder のブラウザ／デスクトップ統合パッケージです。既存の Web 版と同じ UI を保ちながら、ローカルファイルの保存・読み込みを OS ネイティブのダイアログで行えるようになりました。

## 必要要件

- Node.js 18 以上

## セットアップと起動

```bash
cd psBrowserPilot
npm install
npm start
```

`npm start` を実行すると、Electron ウィンドウ内にアプリが起動し、Web 版と同じ操作でフローの作成・保存・読み込みが行えます。アプリは `app://` スキームで静的アセットを配信するため、`index.html` をブラウザで直接開くとモジュールが読み込めません。必ず `npm start` でデスクトップアプリとして起動してください。

## ブラウザでの動作（PowerShell サーバー経由）

Electron を使用できない環境では、同梱の `psBrowserPilotLocal.ps1` を実行してローカル HTTP サーバーを起動することでブラウザから利用できます。静的アセットを `http://127.0.0.1:8787/` で配信するため、`file://` 起動時に発生していた CORS エラーを回避できます。

1. Windows で PowerShell を開きます。
2. リポジトリのルートから `cd psBrowserPilot` を実行します。
3. `./psBrowserPilotLocal.ps1` を実行します。
4. ブラウザで `http://127.0.0.1:8787/` を開きます。

サーバーは `Ctrl + C` で終了できます。`psBrowserPilotLocal.ps1` は PowerShell スクリプト実行 API を提供する `/runscript` エンドポイントも維持しているため、既存の自動化ワークフローからも引き続き呼び出せます。

## PowerShell スクリプトの即時実行

NodeFlow Local から生成したスクリプトは、ヘッダーの **Run in PowerShell** ボタンを押すことで、ブラウザから直接 PowerShell サーバーに送信して実行できます。実行結果やエラー出力はダイアログに表示され、ワンクリックでコピーすることも可能です。

- 設定メニュー（⚙）から PowerShell サーバーの URL を指定できます。既定値は `http://127.0.0.1:8787` で、付属の `psBrowserPilotLocal.ps1` を起動すると同じポートで待ち受けます。
- 「接続テスト」ボタンで `/` エンドポイントへの疎通確認を行えます。サーバーが稼働していない場合は、ステータスが赤色で通知されます。
- 実行ダイアログでは `/runScript` エンドポイントを利用しており、PowerShell 側で `ok = $false` や `Errors` が返された場合はエラーとして強調表示されます。

## 主な変更点

- フローの保存・読み込み、スクリプトのダウンロードを Electron IPC 経由でネイティブファイルダイアログに対応
- ブラウザ単体でも動作するよう、Electron API が利用できない環境では従来のダウンロード処理にフォールバック
