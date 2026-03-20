# VBAからPythonへの移行スクリプト

このスクリプトは、VBAのScrapePedigreeQueryをPythonに移行したものです。

## 機能
- レース一覧ページからレースURLを収集
- 各レースの詳細ページから1-3着の馬名を収集
- 結果をCSVファイル(result.csv)に出力

## 実行方法
```bash
python scrape_pedigree.py [URL]
```

URLを指定しない場合、デフォルトのURLを使用します。

## 依存関係
- requests
- beautifulsoup4
- openpyxl
- pandas
- selenium
- webdriver-manager

## 注意
サイトが403エラーを返す場合、User-Agentや他のヘッダーを調整してください。
もしくは、seleniumを使用するよう変更してください。