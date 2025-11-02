# psBrowserPilot ノードカタログ

このディレクトリには、ブラウザから再利用できる NodeFlow 用 PowerShell ノード定義を配置します。

## 配置方法

1. `*.json` 形式でノード仕様を保存します。ファイルはブラウザで読み込むだけでなく、チーム内で共有できるよう Git 管理してください。
2. ブラウザ UI の「Custom Nodes」>「サンプルをインポート」でファイルの中身をコピーし、保存します。
3. ノード仕様は以下のプロパティを含みます。
   - `id`: 一意な識別子（英数字と `_` のみ）
   - `label`: NodeFlow 上で表示する名前
   - `category`: パレット表示カテゴリ
   - `inputs` / `outputs`: ポート名配列
   - `constants`: 定数設定の配列（`key` / `default`）
   - `script`: 実行する PowerShell スクリプト。`{{input.Name}}` や `{{config.Key}}` プレースホルダを利用できます。

## 例

`sample-get-process.json` を参考に、再利用したい PowerShell 操作をノード化してください。
