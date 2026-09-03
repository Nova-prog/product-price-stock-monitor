# product-price-stock-monitor

商品の公開価格・在庫・納期を確認し、CSVとExcelに記録するPythonツールです。

モノタロウとアスクルでは同一品の比較を想定し、関東化学（Cica-Web）と富士フイルム和光純薬（Wako）では指定した薬品の公開価格の履歴確認を想定しています。

## 重要な方針

- 公開ページで見える情報だけを取得します。
- 会員価格、契約価格、ログイン後価格、認証が必要な情報は対象外です。
- 価格、在庫、納期が取得できない場合は、0円や在庫なしとして扱わず、取得不可として記録します。
- 薬品はCAS番号だけで同等品とは判断しません。グレード、純度、濃度、容量、メーカーを確認してください。
- サイトの構造変更や利用規約により取得できなくなる場合があります。

## 動作環境

- Windows
- Python 3.10以上
- インターネット接続

## 初期設定

プロジェクトのフォルダーでコマンドプロンプトを開き、順番に実行します。

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 文字化け対策

`products.example.csv`は文字化けしにくいように英数字中心で作っています。

`products.csv`をExcelで編集して保存した場合、PCによってはShift-JIS形式になることがあります。`monitor.py`はUTF-8とShift-JISの両方を読むようにしています。

コマンドプロンプトでCSVを`type`表示したとき、先頭に`・ｿ`のような文字が出ることがあります。これはUTF-8 BOMの表示で、Excel出力や取得処理の失敗ではありません。

## 実行方法
```bat
.venv\Scripts\python.exe monitor.py
```

`run_daily.bat`からも実行できます。

```bat
run_daily.bat
```

## 商品設定

確認対象はローカルのローカルの`products.csv`に1行ずつ書きます。

GitHubには実運用の`products.csv`は置かず、テンプレートとして`products.example.csv`だけを置きます。DL後、最初に次のコマンドでローカル用ファイルを作ってください。

```bat
copy products.example.csv products.csv
```

`products.csv`には自分が実際に監視したい商品URLを書きます。このファイルは`.gitignore`で除外しているため、GitHubへ上げない運用です。

主な列は次の通りです。

| 列 | 内容 |
| --- | --- |
| enabled | `true`なら取得、`false`なら無視 |
| group | 同一品比較のグループ名 |
| category | `office_supply`、`lab_consumable`、`reagent`など |
| site | `monotaro`、`askul`、`kanto`、`wako` |
| name | 管理用の商品名 |
| url | 商品ページURL |
| product_code | 販売サイトの商品コード |
| manufacturer_code | メーカー型番 |
| jan | JANコード |
| cas | 薬品のCAS番号 |
| grade | 薬品の規格・グレード |
| capacity | 容量 |
| unit | 購入単位 |
| quantity / quantity_unit | 単価計算用の数量と単位 |
| expected_price_type | 例：`selling_price`、`list_price`、`suggested_delivery_price`など |
| expected_tax_status | 例：`tax_excluded`、`tax_included` |

モノタロウとアスクルを比較したい場合は、同じ`group`を付けます。ただし、同じ`group`に入れる前に、型番、JAN、容量、入数が本当に同じか確認してください。

## 出力

実行すると`data`フォルダーに出力します。

| ファイル | 内容 |
| --- | --- |
| latest.csv | 最新の取得結果 |
| history.csv | 日次の価格履歴 |
| comparison.csv | 同じgroup内の比較結果 |
| errors.csv | 取得不可・部分取得の記録 |
| product_monitor.xlsx | Excel版。最新一覧、販売店比較、価格履歴、取得エラーのシート |

Excelを開いたまま実行した場合、保存できないことがあります。その場合は時刻付きの別名ファイルに保存します。

## 毎日の自動実行

Windowsのタスクスケジューラで、`run_daily.bat`を毎日1回実行するタスクを登録してください。

PCが電源オフまたはスリープ中の場合、その時刻の取得はできません。タスクスケジューラ側で「スケジュールされた時刻にタスクを開始できなかった場合、すぐにタスクを実行する」を有効にすると、次回起動時に実行しやすくなります。

## 免責事項

取得結果の正確性・完全性・最新性は保証しません。購入判断には販売サイトの最新情報を確認してください。

## ライセンス

MIT License。詳細は`LICENSE`を参照してください。
