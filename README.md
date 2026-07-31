# 東京都環境計画書 消費電力量集計ツール (energy-kwh-app)

Streamlit製の専用部・共用部PDF集計ツール。**Streamlit Community Cloud 版**。

- 公開URL: https://energy-kwh-app-tyhbqna9bsfkuauce6qwyk.streamlit.app/

## 機能
- 専用部PDF（住戸別の一次エネ計算書）から消費電力量を抽出
- 共用部PDF（非住宅版エネルギー消費性能計算書）から建物全体・太陽光削減量を抽出
- 住戸リストCSVと組み合わせて建物全体の消費電力量を集計
- Excel / PDF レポート出力

## 対応PDFフォーマット
- 共用部PDF: Ver.3.10 (2026.04) 以降の新形式（4ページ目に二次エネ、太陽光は正値）
- 旧形式（3ページ目に二次エネ、太陽光はマイナス符号）も後方互換

## 同じツールの別リポジトリ（要注意）

このツールは2つのリポジトリで公開されている。**`app.py` は常に両方へ同じ修正を当てること。**

| リポジトリ | ホスティング | 備考 |
|---|---|---|
| [tezukatomoo/energy-kwh-app](https://github.com/tezukatomoo/energy-kwh-app) | Streamlit Community Cloud | このリポジトリ |
| [tezukatomoo/energy-web](https://github.com/tezukatomoo/energy-web) | Render（無料枠切れで停止中） | ローカル起動用 `energy-web起動.bat` あり |

2026-07-31 に `app.py` を energy-web 側の最新版へ統一済み。

## ローカル実行
```bash
pip install -r requirements.txt
streamlit run app.py
```
ログインパスワード: `energy2026`
