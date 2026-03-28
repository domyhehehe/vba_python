# VBAからPythonへの移行メモ

このディレクトリは、`PedigreeQuery` のレース結果取得と馬血統取得を、VBA から Python へ移行するための作業用です。

現在の主な実行入口は [`pedigree_complete.py`](/d:/AI/VBA移行/pedigree_complete.py) です。  
このスクリプトは、[`csv/csv3.py`](/d:/AI/VBA移行/csv/csv3.py) の CSV 変換・統合ロジックに対して、`requests` と Selenium の取得手段を差し替えて動かします。

## 全体像

- `pedigree_complete.py`
  - 実行入口
  - `csv/csv3.py` を読み込み、取得処理だけを上書きする
  - `requests` で取得できない場合に Selenium/Chrome へフォールバックする

- `csv/csv3.py`
  - 実処理の本体
  - レースページ解析、馬ページ解析、CSV 追記、重複排除、再帰展開を担当する
  - `blood.csv` と `stakes_horses.csv` の整形ルールもここにある

- `scrape_pedigree.py`
  - Selenium 取得の検証用スクリプト
  - 旧実験コードとして残している

- `元VBA.md` / `元VBA2.md`
  - 旧 VBA 実装のメモ

## 出力ファイル

既定では `output` フォルダに出力します。

- 血統 CSV: [`output/blood.csv`](/d:/AI/VBA移行/output/blood.csv)
- レース CSV: [`output/stakes_horses.csv`](/d:/AI/VBA移行/output/stakes_horses.csv)

必要なら次で変更できます。

- `--blood-out`
- `--stakes-out`

## 依存関係

主な依存パッケージ:

- `requests`
- `beautifulsoup4`
- `selenium`
- `webdriver-manager`

環境例:

```powershell
.venv\Scripts\Activate.ps1
```

## 実行入口

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

### 5. `--stakes-csv` 経由の既取得馬を事前 SKIP したくない場合

```powershell
python pedigree_complete.py --stakes-csv d:\AI\VBA移行\output\stakes_horses.csv --no-skip-loaded-stakes-csv
```

## 取得フロー

### フロー1: `--horse-url` から血統をたどる

1. 指定された馬 URL を開始キューに積む
2. その馬ページを取得する
3. 血統表から 1 代目から 5 代目までを解析する
4. 1 代目から 4 代目までを `blood.csv` に書き込む
5. ページ上部の主役馬情報を別途抽出し、その行は `LoadURL=True` で保存する
6. 血統表の 5 代目にある馬 URL を次の取得候補としてキューに追加する
7. キューが空になるまで繰り返す

### フロー2: `--race-url` から取り込む

1. レース結果ページを取得する
2. 出走馬ごとの URL とレース成績を抽出する
3. `stakes_horses.csv` に 1 頭 1 行でマージ保存する
4. この時点では `blood.csv` は更新しない
5. 続けて `--stakes-csv` を渡した場合は、その CSV を入口に血統取得へ進める

### フロー3: `--race-list-url` から複数レースを取り込む

1. 一覧ページから各レース URL を収集する
2. 各レース URL に対してフロー2を繰り返す
3. 収集された馬一覧を後段の `--stakes-csv` 入力として使える

### フロー4: `--stakes-csv` から馬 URL を流し込む

1. `stakes_horses.csv` の URL 列を読み込む
2. URL の重複を除外する
3. 既定では、`blood.csv` 側で `LoadURL=True` の馬だけを事前 SKIP する
4. SKIP されなかった URL を馬取得キューへ積む
5. 以後の馬ページ処理自体は `--horse-url` と同じ再帰ロジックで進む

## 五代目取得ロジック

この処理は「5代目まで解析し、4代目まで保存し、5代目を次の入口にする」という形です。

- `parse_depth=5`
  - 血統表の `data-g="1"` から `data-g="5"` まで解析する
- `register_depth=4`
  - `blood.csv` に登録するのは 4 代目まで
- 5代目の扱い
  - 5代目の URL はその場で詳細行として保存しない
  - 次にアクセスする馬 URL としてキューへ追加する

このため、1 回のページ取得で得られるものは次の2種類です。

- 現ページの主役馬の詳細情報
- その主役馬の 1 代目から 4 代目の血統行

さらに 5 代目の馬へ遷移することで、次のページでその馬が主役になり、同じ処理が繰り返されます。

## 主役馬ロジック

各馬ページには「今開いているページの主役の馬」が存在します。コード上ではこの行を別扱いしています。

- 血統表本体からは祖先行を組み立てる
- それとは別に、ページ上部ヘッダから主役馬の情報を抽出する
- 主役馬には以下の詳細を入れる
  - `Horse Name`
  - `URL`
  - `Sire`
  - `Dam`
  - `Sex`
  - `Color`
  - `Year`
  - `Country`
  - `Family`
  - `DP`
  - `DI`
  - `CD`
  - `Starts`
  - `Wins`
  - `Places`
  - `Shows`
  - `CareerEarnings`
  - `Owner`
  - `Breeder`
  - `StateBred`
  - `WinningsText`
  - `SubjectInfoText`
- 主役馬の行だけ `LoadURL=True` で保存する

つまり `LoadURL=True` は「その馬のページ本体に実際にアクセスして、主役として詳細取得した」という意味です。

## SKIP 条件

### 1. 同一実行中の URL 重複

次は SKIP されます。

- すでに処理済みの URL
- すでにキュー投入済みの URL
- 不正な URL
- `https://www.pedigreequery.com/` のような root URL

### 2. 再帰展開時の既取得馬 SKIP

馬ページを処理したあと、5代目から次の候補 URL を集めます。  
このとき `blood.csv` 側で「すでに取得済み」と判定された主役馬 PK は次キューへ積みません。

通常の取得済み判定は次です。

- `LoadURL=True`
- または詳細項目が入っていて取得済み相当とみなせる行

### 3. `--stakes-csv` 専用の事前 SKIP

`--stakes-csv` 経由は大量 URL を流し込むことが多いため、血統ページへアクセスする前に事前 SKIP を行います。

既定では次の条件だけで事前 SKIP します。

- `blood.csv` に同じ PK が存在する
- その行の `LoadURL` が明示的に `True`

逆に次の行は `--stakes-csv` 経由でも取得対象に残ります。

- `LoadURL` が空
- `LoadURL` が `False` 相当
- 一部項目だけ埋まっているが、主役馬としては未取得

### 4. `--horse-url` は事前 SKIP しない

直接指定した `--horse-url` は、既存データがあっても事前 SKIP しません。  
これは個別馬の再取得や、データ修正のための上書き更新を可能にするためです。

## CSV 更新ルール

### `blood.csv`

- `PrimaryKey` 単位でマージする
- 既存行があれば空欄を新規値で補完する
- ただし主役馬 `subject_pk` はそのページ取得結果で上書きする
- 最終完了時にだけ年順で並べ替える

### `stakes_horses.csv`

- 1 頭 1 行で保持する
- `PrimaryKey` ごとにレース JSON をマージする
- 同一レースは `race_page_id / year / placing` で重複排除する

## 取得手段

`pedigree_complete.py` は `csv/csv3.py` の `fetch_html()` を差し替えています。

取得手順は次です。

1. まず `requests` で取得する
2. HTML 内に必要な要素が見つかればそのまま使う
3. `403` などで失敗した場合や期待内容が取れない場合は Selenium に切り替える
4. Selenium では Chrome を起動し、ページ読込完了と対象要素出現を待ってから HTML を返す

## ログの見方

### 馬取得

- `[HORSE] ...`
  - 現在処理中の馬 URL

- `[OK] horse=... rows=... next=... pending=...`
  - 1 頭分の書き込み完了
  - `rows` は今回生成した CSV 行数
  - `next` は 5代目から追加された次候補数
  - `pending` は残りキュー数

- `[DONE] horses_processed=... failed=... skipped=...`
  - 馬処理全体の完了

- `[DONE] blood.csv sorted rows=...`
  - 最後のソート完了

### レース取得

- `[LIST] collecting race URLs from ...`
  - 一覧ページからレース URL を収集中

- `[RACE] i/n ...`
  - レースページを処理中

- `[OK] horses added/merged: ...`
  - `stakes_horses.csv` へ追加またはマージした頭数

- `[CSV] loading horse URLs from ...`
  - `--stakes-csv` を読み込み中

- `[CSV] pre-skip loaded horses from --stakes-csv: input=... skipped=... queued=...`
  - 事前 SKIP の結果

## 補足

- `git status` はこのディレクトリが Git 管理下でないと失敗します
- `PedigreeQuery` 側の HTML 構造やブロック条件は変わることがあります
- Selenium 起動にはローカルの Chrome と対応ドライバ取得が必要です
