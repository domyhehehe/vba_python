# VBAからPythonへの移行メモ

このディレクトリは、`PedigreeQuery` のレース結果取得と馬血統取得を、VBA から Python へ移行するための作業用です。

現在の主な実行入口は [`pedigree_complete.py`](/d:/AI/VBA移行/pedigree_complete.py) です。  
このスクリプトは、[`csv/csv3.py`](/d:/AI/VBA移行/csv/csv3.py) の CSV 変換・統合ロジックと、`scrape_pedigree.py` で確認済みの Selenium ベース取得を組み合わせています。

## 現在の役割

- `pedigree_complete.py`
  - 実行用の統合入口
  - `requests` が `403` の場合は Selenium/Chrome にフォールバック
  - レース URL、レース一覧 URL、馬 URL、既存 `stakes_horses.csv` を入口にできる

- `csv/csv3.py`
  - 血統 CSV とレース CSV の正規化・追記・重複排除
  - 馬ごとの再帰ループ本体
  - 最終完了時のみ `blood.csv` を年順ソート

- `scrape_pedigree.py`
  - 旧検証用スクリプト
  - Selenium 取得まわりの参考実装

- `元VBA.md` / `元VBA2.md`
  - 旧 VBA 実装のメモ

## 出力先

既定では `output` フォルダに出力します。

- 血統 CSV: [`output/blood.csv`](/d:/AI/VBA移行/output/blood.csv)
- レース CSV: [`output/stakes_horses.csv`](/d:/AI/VBA移行/output/stakes_horses.csv)

必要なら `--blood-out` と `--stakes-out` で変更できます。

## 依存関係

必要な主なパッケージ:

- `requests`
- `beautifulsoup4`
- `selenium`
- `webdriver-manager`

環境例:

```powershell
.venv\Scripts\Activate.ps1
```

## 使い方

### 1. 特定の馬から血統をたどる

```powershell
python pedigree_complete.py --horse-url https://www.pedigreequery.com/eclipse
```

### 2. 特定のレース結果ページを取り込む

```powershell
python pedigree_complete.py --race-url "https://www.pedigreequery.com/index.php?query_type=stakes&search_bar=stakes&id=640"
```

### 3. レース一覧ページから複数レースを拾う

```powershell
python pedigree_complete.py --race-list-url "https://www.pedigreequery.com/index.php?query_type=stakes&search_bar=stakes&field=country&h=japan"
```

### 4. 既存のレース CSV を入口にして馬血統をたどる

```powershell
python pedigree_complete.py --stakes-csv d:\AI\VBA移行\output\stakes_horses.csv
```

## 主な挙動

- `requests` で取得できればそのまま使う
- `403` などで失敗した場合は headless Chrome で取得する
- 馬は 1 頭ずつ `blood.csv` に追記する
- 最後にだけ `blood.csv` を年順で並べ替える
- 不正な URL や root URL (`https://www.pedigreequery.com/`) は再帰対象から除外する

## ログの見方

実行中は次のようなログが出ます。

- `[HORSE] ...`
  - 今処理中の馬 URL

- `[OK] horse=... rows=... next=... pending=...`
  - その馬の書き込み完了
  - `rows` は今回生成した行数
  - `next` は次にたどる URL 数
  - `pending` は未処理キュー数

- `[DONE] horses_processed=... failed=... skipped=...`
  - 馬処理全体の完了

- `[DONE] blood.csv sorted rows=...`
  - 最後の年順ソート完了

## 補足

- `git status` はこのディレクトリが Git 管理下でないと失敗します。
- `PedigreeQuery` 側の応答やブロック条件は変わることがあります。
- Selenium 起動にはローカルの Chrome と対応するドライバ取得が必要です。
