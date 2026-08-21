# ECWeaver / ECWeaver2 コマンドツール機能仕様案

## 目的

`ECWeaver` と `ECWeaver2` を、Excel / CSV の変換・加工・検査をまとめて行える多機能コマンドツールにする。

既存の `ExcelAppTools`、`ExcelInteropTools`、`ExcelTools`、`CsvFileReader`、`CsvFileWriter` を土台にし、必要に応じて各 `*Tools` に機能を追加する。

## 方針

- `ECWeaver` は Excel アプリケーションを PowerShell COM 経由で扱う機能を中心にする。
- `ECWeaver2` は `Microsoft.Office.Interop.Excel` を直接使う機能を中心にする。
- CSV の読み書き、行列加工、文字コード変換など、Excel 本体が不要な処理は共通仕様として両方に入れる。
- 既存の C# / .NET Framework 4.8 / x86 構成を維持する。
- まずは単機能コマンドを増やし、あとから複数処理をつなぐバッチ的な機能を追加する。
- 複数の Excel / CSV ファイルを混在指定できる統合変換では、入力をいったんシート相当の中間データに正規化し、加工してから出力形式を選ぶ。

## 既存ツールの位置づけ

### ExcelAppTools

Excel アプリケーションを PowerShell COM 経由で操作する。

既存機能:

- Excel ファイルの全シート読み込み
- Excel から CSV への変換
- Excel から PDF への変換
- 印刷
- プリンタ一覧取得
- プレースホルダ文字列置換
- プレースホルダ位置への画像差し込み

追加候補:

- 指定シートのみ CSV 出力
- シート名指定 / シート番号指定の両対応
- Excel から TSV / SSV 出力
- CSV / TSV から Excel ブック作成
- 複数 CSV から複数シート Excel 作成
- 複数 Excel / CSV 混在入力から 1 ブックまたは CSV 群を作成
- セル範囲指定で CSV 出力
- UsedRange のトリム、空行・空列除去
- シートコピー、削除、リネーム、並び替え
- ブック内文字列検索
- ブック内文字列一括置換
- セル書式の簡易指定
- ページ設定、印刷範囲、倍率、用紙方向の変更
- PDF 出力時のシート指定
- PDF 出力時の結合 / 分割
- エラー内容を PowerShell 側からより詳細に返す仕組み

### ExcelInteropTools

`Microsoft.Office.Interop.Excel` を直接使って Excel を操作する。

既存機能:

- Excel から PDF への変換
- 印刷
- プリンタ一覧取得

追加候補:

- ExcelAppTools と同等の Excel -> CSV
- シート一覧取得
- セル値読み込み
- セル値書き込み
- 範囲読み込み
- 範囲書き込み
- 行列削除 / 挿入
- オートフィット
- フィルタ適用
- ソート
- シート追加 / 削除 / リネーム
- ブック保存形式変換
- COM オブジェクト解放処理の共通化と強化

### ExcelTools

`.xlsx` を ZIP として展開し、Open XML 内部ファイルを直接扱う。

既存機能:

- `.xlsx` 展開閲覧
- `.xlsx` 展開編集
- ブック内画像収集
- ブック内画像置換

追加候補:

- 画像一覧をファイルに書き出し
- 画像差し替えのマッピング指定
- `sharedStrings.xml` の文字列検索
- `sharedStrings.xml` の文字列置換
- ワークシート XML の直接検査
- ブック内 XML の簡易ダンプ
- `.xlsx` の破損検査
- Excel アプリケーション不要で可能なメタ情報取得
- シート構成、リレーション、メディアファイルの一覧化

### CsvFileReader / CsvFileWriter

CSV / TSV / SSV の読み書きに使う。

追加候補:

- UTF-8 BOM / UTF-8 no BOM / Shift_JIS / UTF-16LE の明示指定
- 区切り文字の自動判定
- 改行コードの指定
- ヘッダーあり / なしの扱い
- 列名による列抽出
- 列番号による列抽出
- 行フィルタ
- セル置換
- 正規表現置換
- 重複行削除
- ソート
- グループ集計
- 縦横変換
- 複数 CSV の結合
- CSV 差分比較
- CSV 妥当性検査

## コマンド体系案

基本形:

```txt
ECWeaver.exe <command> [options] <input> <output>
ECWeaver2.exe <command> [options] <input> <output>
```

共通オプション:

```txt
--help
--version
--overwrite
--encoding <sjis|utf8|utf8bom|utf16le>
--delimiter <comma|tab|space|char>
--sheet <name-or-index>
--range <A1:D20>
--input-list <file>
--log <file>
--silent
--verbose
```

## 変換系コマンド

### excel-to-csv

Excel ブックを CSV に変換する。

```txt
ECWeaver.exe excel-to-csv input.xlsx output-dir
ECWeaver.exe excel-to-csv --sheet Sheet1 input.xlsx output.csv
ECWeaver.exe excel-to-csv --encoding utf8bom input.xlsx output-dir
```

仕様:

- シート指定なしの場合、全シートを `0001.csv`、`0002.csv` のように出力する。
- `sheet-names.txt` も出力する。
- シート指定ありの場合、単一 CSV として出力できる。
- 出力先が存在する場合は `--overwrite` がない限りエラーにする。

### excel-to-tsv

Excel ブックを TSV に変換する。

```txt
ECWeaver.exe excel-to-tsv input.xlsx output-dir
```

仕様:

- `excel-to-csv` と同じだが、区切り文字をタブにする。

### csv-to-excel

CSV を Excel ブックに変換する。

```txt
ECWeaver.exe csv-to-excel input.csv output.xlsx
ECWeaver.exe csv-to-excel --sheet Data input.csv output.xlsx
```

仕様:

- 1 CSV を 1 シートの `.xlsx` に変換する。
- シート名未指定の場合は `Sheet1` とする。

### csvs-to-excel

複数 CSV を複数シートの Excel ブックに変換する。

```txt
ECWeaver.exe csvs-to-excel input-dir output.xlsx
```

仕様:

- 入力ディレクトリ配下の `.csv` をファイル名順で読み込む。
- 各 CSV を 1 シートにする。
- シート名はファイル名から生成する。

### excel-to-pdf

Excel ブックを PDF に変換する。

```txt
ECWeaver.exe excel-to-pdf input.xlsx output.pdf
ECWeaver2.exe excel-to-pdf input.xlsx output.pdf
```

仕様:

- `ECWeaver` は `ExcelAppTools.ToPDF` を使う。
- `ECWeaver2` は `ExcelInteropTools.ToPDF` を使う。
- シート指定出力は追加機能として扱う。

### weave

複数の Excel / CSV ファイルを混在入力として受け取り、順番に処理して統合変換する。

```txt
ECWeaver.exe weave input1.csv input2.xlsx input3.csv --to-excel output.xlsx
ECWeaver.exe weave input1.csv input2.xlsx input3.csv --to-csv-dir output-dir
ECWeaver.exe weave input1.csv input2.xlsx input3.csv --to-same-dir output-dir
ECWeaver.exe weave --input-list files.txt --to-excel output.xlsx
```

仕様:

- 入力ファイルは複数指定できる。
- 入力ファイルは `.csv`、`.tsv`、`.ssv`、`.xls`、`.xlsx`、`.xlsm` の混在を許可する。
- コマンドラインで指定された入力ファイル、または `--input-list` で指定されたリストファイルを、記載順に処理する。
- `--input-list` は 1 行 1 ファイルの簡易形式とし、空行と `#` コメント行を許可する。
- CSV / TSV / SSV 入力は、1 ファイルを 1 シート相当の中間データとして扱う。
- Excel 入力は、シート指定がなければ全シートを順番に読み込み、1 シートを 1 シート相当の中間データとして扱う。
- 入力処理の過程で、列抽出、行抽出、セル置換、トリム、ソート、シート名変更などの加工を適用できるようにする。
- 初期版では加工なしの統合変換を優先し、加工指定は段階的に追加する。
- 出力先が存在する場合は `--overwrite` がない限りエラーにする。

出力モード:

- `--to-excel output.xlsx`
  - 全ての中間データを、1 つの `.xlsx` ブックのシートとして出力する。
  - シート名は入力ファイル名と元シート名から生成する。
  - シート名が重複する場合は `_001`、`_002` のような連番を付ける。
- `--to-csv-dir output-dir`
  - 全ての中間データを、CSV ファイル群として指定フォルダへ出力する。
  - Excel 入力の各シートも、それぞれ 1 CSV として出力する。
  - 出力ファイル名は入力ファイル名と元シート名から生成する。
- `--to-same-dir output-dir`
  - 入力ファイル単位のまとまりを保ち、指定フォルダへ出力する。
  - CSV / TSV / SSV 入力は、加工後も原則として同種のテキスト形式で出力する。
  - Excel 入力は、加工後も原則として Excel ブックとして出力する。
  - ただし、加工内容が元形式で保持できない場合はエラーにするか、明示オプションで変換を許可する。

中間データの考え方:

- 内部では、入力順を保持した `WorkbookUnit` と `SheetUnit` のような構造に正規化する。
- `WorkbookUnit` は元入力ファイルの単位を表す。
- `SheetUnit` は CSV 1 ファイル、または Excel 1 シートに相当する表データを表す。
- 出力モードは、この中間データを `.xlsx`、CSV 群、または入力ファイル単位の出力に変換するだけにする。
- 将来の加工機能は、中間データに対する処理として追加する。

## CSV 加工系コマンド

### csv-select-columns

CSV の列を抽出する。

```txt
ECWeaver.exe csv-select-columns --columns 1,3,5 input.csv output.csv
ECWeaver.exe csv-select-columns --headers Code,Name,Price input.csv output.csv
```

仕様:

- `--columns` は 1 始まりの列番号。
- `--headers` は 1 行目をヘッダーとして扱う。
- 存在しない列はエラーにする。

### csv-filter-rows

CSV の行を条件で抽出する。

```txt
ECWeaver.exe csv-filter-rows --column 2 --equals Tokyo input.csv output.csv
ECWeaver.exe csv-filter-rows --column 3 --contains error input.csv output.csv
ECWeaver.exe csv-filter-rows --column 4 --regex "^[0-9]+$" input.csv output.csv
```

仕様:

- `--equals`、`--contains`、`--regex` のいずれかを指定する。
- `--invert` で条件に一致しない行を出力する。
- `--has-header` 指定時はヘッダー行を常に残す。

### csv-replace

CSV のセル文字列を置換する。

```txt
ECWeaver.exe csv-replace --from old --to new input.csv output.csv
ECWeaver.exe csv-replace --regex "\s+" --to " " input.csv output.csv
```

仕様:

- 全セルを対象にする。
- `--column` 指定時は対象列だけ置換する。
- 通常置換と正規表現置換を切り替える。

### csv-merge

複数 CSV を縦方向に結合する。

```txt
ECWeaver.exe csv-merge input-dir output.csv
ECWeaver.exe csv-merge --pattern "*.csv" --skip-header input-dir output.csv
```

仕様:

- ファイル名順に結合する。
- `--skip-header` 指定時、2 ファイル目以降の先頭行を除外する。

### csv-sort

CSV を指定列でソートする。

```txt
ECWeaver.exe csv-sort --column 1 input.csv output.csv
ECWeaver.exe csv-sort --column 3 --numeric --desc input.csv output.csv
```

仕様:

- 文字列ソートを標準とする。
- `--numeric` で数値ソートする。
- `--has-header` 指定時はヘッダー行を固定する。

### csv-unique

重複行を削除する。

```txt
ECWeaver.exe csv-unique input.csv output.csv
ECWeaver.exe csv-unique --columns 1,2 input.csv output.csv
```

仕様:

- `--columns` 未指定時は行全体で重複判定する。
- 最初に出現した行を残す。

## Excel 加工系コマンド

### excel-list-sheets

Excel ブックのシート一覧を出力する。

```txt
ECWeaver.exe excel-list-sheets input.xlsx
```

仕様:

- 標準出力にシート番号とシート名を出す。
- `--output` 指定時はテキストファイルに保存する。

### excel-replace-text

Excel ブック内の文字列を置換する。

```txt
ECWeaver.exe excel-replace-text --from "**NAME**" --to "山田太郎" template.xlsx output.xlsx
```

仕様:

- `ExcelAppTools.ReplacePlaceholder` を土台にする。
- プレースホルダ形式だけでなく通常文字列置換も追加候補にする。
- フォントや背景色の指定は追加オプションとして扱う。

### excel-replace-picture

Excel ブック内の画像を置換する。

```txt
ECWeaver.exe excel-replace-picture input.xlsx output.xlsx image.png
```

仕様:

- `ExcelTools.ReplacePicture` を土台にする。
- 初期版ではブック内の画像を順番に差し替える。
- 追加版では画像番号やファイル名マッピングで差し替える。

### excel-extract-pictures

Excel ブック内の画像を抽出する。

```txt
ECWeaver.exe excel-extract-pictures input.xlsx output-dir
```

仕様:

- `ExcelTools.CollectPicture` を土台にする。
- 抽出ファイル名は出現順と元拡張子から生成する。

### excel-trim

Excel / CSV の余分な空行・空列を削除する。

```txt
ECWeaver.exe excel-trim input.xlsx output.xlsx
ECWeaver.exe csv-trim input.csv output.csv
```

仕様:

- 行末の空セルを削る。
- 末尾の空行を削る。
- 完全空列の削除はオプションにする。

## 検査系コマンド

### csv-info

CSV の基本情報を出力する。

```txt
ECWeaver.exe csv-info input.csv
```

出力候補:

- 行数
- 最大列数
- 最小列数
- 空行数
- 文字コード推定
- 区切り文字推定

### csv-validate

CSV の形式を検査する。

```txt
ECWeaver.exe csv-validate input.csv
ECWeaver.exe csv-validate --columns 10 input.csv
```

検査候補:

- 行ごとの列数不一致
- 空ヘッダー
- 重複ヘッダー
- 不正な引用符
- 必須列不足

### excel-info

Excel ブックの基本情報を出力する。

```txt
ECWeaver.exe excel-info input.xlsx
```

出力候補:

- シート数
- シート名一覧
- 各シートの行数・列数
- 画像数
- ファイルサイズ
- 拡張子

### excel-validate

Excel ブックを検査する。

```txt
ECWeaver.exe excel-validate input.xlsx
```

検査候補:

- Excel アプリケーションで開けるか
- `.xlsx` として ZIP 展開できるか
- シート名が空でないか
- 想定シートが存在するか
- 画像やリレーションが壊れていないか

## 差分・比較系コマンド

### csv-diff

2 つの CSV を比較する。

```txt
ECWeaver.exe csv-diff before.csv after.csv diff.csv
```

仕様:

- セル単位の差分を出力する。
- 行番号、列番号、変更前、変更後を出力する。
- `--key-columns` 指定時はキー列で行を対応付ける。

### excel-diff

2 つの Excel ブックを比較する。

```txt
ECWeaver.exe excel-diff before.xlsx after.xlsx diff-dir
```

仕様:

- 初期版では各シートを CSV 化して比較する。
- シート名単位の追加 / 削除 / 変更を出す。

## バッチ処理系コマンド

### run-script

処理定義ファイルに従って複数処理を実行する。

```txt
ECWeaver.exe run-script recipe.txt
```

仕様候補:

- 1 行 1 コマンドの簡易形式から始める。
- 空行と `#` コメントを許可する。
- 後続版で JSON / YAML 形式を検討する。

### watch

指定フォルダを監視して変換する。

```txt
ECWeaver.exe watch input-dir output-dir --command excel-to-csv
```

仕様候補:

- 初期版では対象外。
- 将来、フォルダ投入型の自動変換に使う。

## 優先実装順

### Phase 1: 最小実用コマンド

- `--help`
- `--version`
- `excel-to-csv`
- `excel-to-pdf`
- `csv-info`
- `csv-select-columns`
- `csv-filter-rows`
- `csv-replace`

### Phase 2: よく使う変換・加工

- `csv-to-excel`
- `csvs-to-excel`
- `weave` の加工なし統合変換
- `excel-list-sheets`
- `excel-extract-pictures`
- `excel-replace-picture`
- `csv-merge`
- `csv-sort`
- `csv-unique`

### Phase 3: 検査・差分

- `csv-validate`
- `excel-info`
- `excel-validate`
- `csv-diff`
- `excel-diff`

### Phase 4: 高度な Excel 操作

- `excel-replace-text`
- セル範囲指定
- シート操作
- 書式操作
- PDF 分割 / 結合
- 印刷設定

### Phase 5: 自動化

- `weave` の加工パイプライン拡張
- `run-script`
- `watch`
- 処理ログの構造化
- 終了コード体系の整理

## 終了コード案

```txt
0  正常終了
1  引数エラー
2  入力ファイルなし
3  出力先あり
4  CSV 読み込みエラー
5  Excel 読み込みエラー
6  変換エラー
7  検査エラー
8  外部アプリケーションエラー
9  不明なエラー
```

## 実装メモ

- `Program.Main5()` はコマンドディスパッチャにする。
- コマンドごとの処理は `Commands` 名前空間または `Tools` 配下に分割する。
- `ArgsReader` は既存のままでも始められるが、オプション解析が増えたら補助クラスを追加する。
- ファイル上書きは原則禁止し、`--overwrite` 指定時だけ許可する。
- 出力ファイルは一時ファイルへ作成してから移動する。
- CSV は矩形化が必要な処理と、不揃い行を保持すべき処理を分ける。
- Excel を使う処理は Excel 未インストール環境で明確なエラーを返す。
- `ECWeaver` と `ECWeaver2` の同名コマンドは、できるだけ同じ引数仕様にする。
- ただし内部実装の違いにより使える機能が違う場合は `--help` に明記する。

## 未決事項

- `ECWeaver` と `ECWeaver2` を最終的に統合するか、2 系統として残すか。
- CSV 出力の標準文字コードを Shift_JIS のままにするか、UTF-8 BOM に寄せるか。
- `.xls`、`.xlsm`、`.xlsx` の対応範囲。
- Excel 未インストール環境でどこまで機能提供するか。
- GUI 版 `ECWeaverGUI` との機能共有方法。
- コマンド名を短縮形でも受け付けるか。
- 設定ファイルを導入するか。
