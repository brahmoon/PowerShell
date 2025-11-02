# psBrowserPilot ノードライブラリ

`psBrowserPilot` では NodeFlow のカスタムノード機能を利用して PowerShell スクリプトを部品化します。
このディレクトリには有識者が用意したノード定義 (`*.json`) を配置し、ブラウザで `psBrowserPilot-GUI.html`
を開いた際にパレットへ自動読み込みできるようにします。

## ノードの追加手順
1. NodeFlow の「カスタムノードビルダー」でノードを作成し、右上のメニューからエクスポートします。
2. 出力された JSON ファイルをこのディレクトリに保存します。
3. 必要に応じて JSON 内の `id` や `label` を調整し、バージョン管理下にコミットします。
4. 変更内容を `psBrowserPilot-GUI.html` の `SAMPLE_NODE_TEMPLATES` に追記すると、サンプルとして再利用できます。

## 命名規則
- ファイル名は `category-name.json` のように、カテゴリとノードの概要が分かるようにします。
- ノード ID (`id`) は小文字の英数字とアンダースコアのみを使用してください。

## 共有のポイント
- ノードの説明 (`description`) には期待する入力値や前提条件を必ず記載します。
- 実行に追加モジュールが必要な場合は `psBrowserPilot-GUI.html` の「ロードするモジュール」に記載できるよう、
  README やノードの `notes` フィールドに追記してください。
