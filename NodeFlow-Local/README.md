# NodeFlow Local

NodeFlow Local は、PowerShell Visual Builder を Electron でラップしたデスクトップ版です。既存の Web 版と同じ UI を保ちながら、ローカルファイルの保存・読み込みを OS ネイティブのダイアログで行えるようになりました。

## 必要要件

- Node.js 18 以上

## セットアップと起動

```bash
cd NodeFlow-Local
npm install
npm start
```

`npm start` を実行すると、Electron ウィンドウ内にアプリが起動し、Web 版と同じ操作でフローの作成・保存・読み込みが行えます。

## 主な変更点

- フローの保存・読み込み、スクリプトのダウンロードを Electron IPC 経由でネイティブファイルダイアログに対応
- ブラウザ単体でも動作するよう、Electron API が利用できない環境では従来のダウンロード処理にフォールバック
