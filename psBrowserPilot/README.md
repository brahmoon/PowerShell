# psBrowserPilot

`psBrowserPilot` は、NodeFlow ベースのブラウザ UI から PowerShell を安全に実行するためのフルスタック サンプルです。

## 構成

- `psBrowserPilot.ps1` — HTTP API を公開する PowerShell ブリッジ。`/health`、`/sessions`、`/commands` といったエンドポイントを提供し、ブラウザから任意のスクリプトを送信できます。
- `psBrowserPilot-GUI.html` — NodeFlow を組み込んだブラウザ UI。カスタムノードを作成し、フローを構築して PowerShell に送信できます。
- `nodes/` — 共有可能なカスタムノード定義 (`*.json`) を配置するディレクトリ。
- `examples/` — UI の「Load」機能から読み込める NodeFlow グラフ例。

## 使い方

1. PowerShell で `psBrowserPilot.ps1` を実行し、`http://127.0.0.1:8080/` で待ち受けます。
2. `psBrowserPilot-GUI.html` をブラウザ (Chromium ベース推奨) で開きます。
3. ヘッダーの「新規セッション」で Runspace を作成し、NodeFlow フローを組み立てて **Run Script** をクリックします。
4. 実行履歴は画面下部のログパネルと `/sessions/{id}/history` API で確認できます。

Excel 固有ロジックに依存しないため、既存の `psExcelPilot` ノードをテンプレートとして再利用しながら、任意の PowerShell 操作をブラウザから実行できるようになっています。
