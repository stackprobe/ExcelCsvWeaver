# ECWeaver / ECWeaver2 コマンドライン仕様案

## 目的

`ECWeaver` / `ECWeaver2` を、コマンドラインと将来の GUI の両方から同じ処理を呼び出せる形にする。

この仕様では、ユーザーが直接使うコマンド引数の形と、GUI から内部的に呼び出しやすい処理単位を定義する。

## 基本方針

- コマンド名は動詞-目的語の kebab-case にする。
- オプション名も kebab-case にする。
- 入力ファイル、出力ファイルは原則として末尾に置く。
- 破壊的な上書きは `--overwrite` 指定時だけ許可する。
- 標準動作は安全側に倒す。
- Excel 操作エンジンは `--engine` で選択できる。
- `--engine` 未指定時は `auto` として扱う。
- CSV 系コマンドは Excel 操作エンジンを使わない。
- 複数の Excel / CSV ファイルを混在入力できる統合変換では、入力をいったんシート相当の中間データに正規化してから出力形式を選ぶ。
- GUI からも同じコマンド定義を利用できるよう、各コマンドは単一責務にする。
- コマンド引数として受け取ったファイル名・フォルダ名は、引数解析時点で `SCommon.MakeFullPath` を通してフルパス化する。

## 実行形式

```txt
ECWeaver.exe <command> [options] [arguments]
ECWeaver2.exe <command> [options] [arguments]
```

例:

```txt
ECWeaver.exe excel-to-csv input.xlsx output-dir
ECWeaver.exe excel-to-pdf --engine app input.xlsx output.pdf
ECWeaver2.exe excel-to-pdf --engine interop input.xlsx output.pdf
ECWeaver.exe csv-select-columns --columns 1,3,5 input.csv output.csv
```

## 現状実装範囲

この文書は将来仕様を含む。以下は現在の実装済み範囲である。

### ECWeaver

```txt
help
version
excel-to-csv
excel-to-tsv
excel-to-pdf
csv-info
csv-select-columns
csv-filter-rows
csv-replace
csv-merge
csv-sort
csv-unique
excel-list-sheets
excel-extract-pictures
excel-replace-picture
printers
print
```

未実装コマンドを指定した場合は、未実装エラーにする。

### ECWeaver2

```txt
help
version
csv-to-excel
csvs-to-excel
excel-to-pdf
weave
csv-info
csv-select-columns
csv-filter-rows
csv-replace
csv-merge
csv-sort
csv-unique
printers
print
```

`weave` は現在 `--to-excel` のみ実装済みで、入力は `.csv`、`.tsv`、`.ssv` に限る。
`--to-csv-dir`、`--to-same-dir`、Excel 入力の混在処理は未実装である。

### Excel インストール要否

- CSV 専用コマンドは Excel を必要としない。
- `ECWeaver` の `excel-extract-pictures`、`excel-replace-picture` は `ExcelTools` による ZIP / Open XML 直接操作のため Excel を必要としない。
- `ExcelAppTools` または `ExcelInteropTools` を使う Excel 読み込み、Excel 作成、PDF 出力、印刷、プリンタ一覧は Excel または該当環境を必要とする。

### 現在の制限

- `--newline`、`--log`、`--verbose`、`--no-dialog`、`--range`、`--password` はオプション名として予約されているが、現在の処理では実質未対応である。
- 例外発生時は `Program.Main4()` でログ出力とエラーダイアログ表示を行う。`--no-dialog` による抑制は未実装である。
- 終了コード体系は案であり、現在は例外種別ごとの終了コード返却までは整理されていない。

## ヘルプ

```txt
ECWeaver.exe
ECWeaver.exe --help
ECWeaver.exe help
ECWeaver.exe help <command>
ECWeaver.exe <command> --help
```

仕様:

- 引数なしは全体ヘルプを表示する。
- 未知のコマンドはエラーにする。現在の実装では全体ヘルプの短縮版表示までは行わない。
- `help <command>` は指定コマンドの詳細ヘルプを表示する。
- `--help` は他のオプションより優先する。

## バージョン

```txt
ECWeaver.exe --version
ECWeaver.exe version
```

出力候補:

```txt
ECWeaver 1.0.0
```

## 共通オプション

### --engine

Excel 操作に使う処理エンジンを選択する。

```txt
--engine auto
--engine app
--engine interop
--engine zip
```

値:

```txt
auto     コマンド内容と実行ファイルから自動選択する
app      ExcelAppTools を使う
interop  ExcelInteropTools を使う
zip      ExcelTools を使う
```

仕様:

- 未指定時は `auto` とする。
- CSV 専用コマンドでは指定不可にする。
- 指定した engine がコマンドに対応していない場合は引数エラーにする。
- `ECWeaver` の `auto` は `app` を優先する。
- `ECWeaver2` の `auto` は `interop` を優先する。
- `.xlsx` の内部構造だけで完結する処理は `zip` を優先できる。

### --overwrite

出力先が存在する場合に上書きを許可する。

```txt
--overwrite
```

仕様:

- 未指定時、既存の出力ファイルまたは出力ディレクトリがあればエラーにする。
- 上書き時も一時ファイルへ出力してから置き換える。

### --encoding

CSV / TXT の入出力文字コードを指定する。

```txt
--encoding auto
--encoding sjis
--encoding utf8
--encoding utf8bom
--encoding utf16le
```

仕様:

- 入力時の未指定は `auto` とする。
- 出力時の未指定は当面 `sjis` とする。
- 将来 UTF-8 BOM を標準に変更する場合は別途判断する。

### --delimiter

CSV 系ファイルの区切り文字を指定する。

```txt
--delimiter comma
--delimiter tab
--delimiter space
--delimiter "|"
```

仕様:

- 未指定時は拡張子から推定する。
- `.csv` は comma、`.tsv` は tab、`.ssv` は space とする。
- 拡張子で判定できない場合は comma とする。

### --newline

出力ファイルの改行コードを指定する。

```txt
--newline crlf
--newline lf
```

仕様:

- 未指定時は `crlf` とする。

### --input-list

複数入力ファイルをリストファイルで指定する。

```txt
--input-list files.txt
```

仕様:

- 1 行 1 ファイルの簡易形式とする。
- 空行を無視する。
- `#` で始まる行をコメントとして扱う。
- コマンドライン上の入力ファイル群と併用できるかは、各コマンド仕様で定義する。

### --response

レスポンスファイルを読み込み、ファイル内の各行をコマンドライン引数として扱う。

```txt
--response args.txt
--response=args.txt
```

仕様:

- レスポンスファイルは 1 行 1 引数の簡易形式とする。
- 読み込んだ行は、通常のコマンドライン引数と同じ規則で解析する。
- レスポンスファイル内に `--response` を記述した場合も展開する。
- 循環参照は引数エラーにする。
- 現在の実装では CP932 として読み込む。

### --log

ログ出力先を指定する。

```txt
--log process.log
```

仕様:

- 引数解析時点で `SCommon.MakeFullPath` によりフルパス化する。
- 未指定時は既存の `ProcMain` のログ仕様に従う。
- 指定時は処理ログをファイルにも出力する。

### --silent

標準出力への通常メッセージを抑制する。

```txt
--silent
```

仕様:

- エラーは標準エラーまたはエラーダイアログへ出す。
- GUI からの呼び出しでは標準で有効にしてよい。

### --verbose

詳細ログを出す。

```txt
--verbose
```

仕様:

- 処理対象ファイル、選択 engine、行数、シート数などを表示する。
- `--silent` と同時指定された場合は引数エラーにする。

### --no-dialog

エラーダイアログを表示しない。

```txt
--no-dialog
```

仕様:

- コマンドライン利用時の自動処理向け。
- エラーは標準エラーと終了コードで返す。
- GUI から直接処理層を呼ぶ場合は使わなくてよい。

## Excel 共通オプション

### --sheet

対象シートを指定する。

```txt
--sheet Sheet1
--sheet 1
```

仕様:

- 数値だけの場合は 1 始まりのシート番号として扱う。
- それ以外はシート名として扱う。
- シート名と数値名が衝突する可能性があるため、将来 `--sheet-name` / `--sheet-index` の追加を検討する。

### --sheets

複数シートを指定する。

```txt
--sheets Sheet1,Sheet2,Sheet3
--sheets 1,2,3
```

仕様:

- カンマ区切りで指定する。
- シート名にカンマを含む場合は対象外とする。
- 厳密指定が必要になった場合はリストファイル指定を追加する。

### --range

セル範囲を指定する。

```txt
--range A1:D20
```

仕様:

- A1 形式のみ対応する。
- 未指定時は UsedRange を対象にする。

### --password

Excel ブックのオープンパスワードを指定する。

```txt
--password pass
```

仕様:

- 初期実装では対象外にしてよい。
- 対応時はコマンド履歴に残る点をヘルプに明記する。

## CSV 共通オプション

### --has-header

CSV の 1 行目をヘッダーとして扱う。

```txt
--has-header
```

仕様:

- フィルタ、ソート、重複削除などでヘッダー行を固定する。
- 列名指定を使うコマンドでは自動的に有効扱いにする。

### --columns

対象列を列番号で指定する。

```txt
--columns 1,3,5
```

仕様:

- 1 始まりの列番号とする。
- 範囲指定は将来対応とする。

### --headers

対象列をヘッダー名で指定する。

```txt
--headers Code,Name,Price
```

仕様:

- 1 行目をヘッダーとして扱う。
- 存在しないヘッダーはエラーにする。

## コマンド一覧

## 変換系

### excel-to-csv

Excel ブックを CSV に変換する。

```txt
ECWeaver.exe excel-to-csv [options] <input-excel> <output-dir>
ECWeaver.exe excel-to-csv [options] --sheet <sheet> <input-excel> <output-csv>
```

対応 engine:

```txt
auto
app
```

オプション:

```txt
--sheet <name-or-index>
--sheets <names-or-indexes>
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- 現在の実装先は `ECWeaver` である。`ECWeaver2` では未実装。
- `--sheet` 未指定時は全シートを出力ディレクトリへ出力する。
- 全シート出力時は `0001.csv` のような連番ファイル名にする。
- 全シート出力時は `sheet-names.txt` を出力する。
- `--sheet` 指定時は単一 CSV ファイルへ出力する。
- `--sheets` 指定時は指定シート群を出力ディレクトリへ出力する。
- `--sheet` と `--sheets` の同時指定はエラーにする。

### excel-to-tsv

Excel ブックを TSV に変換する。

```txt
ECWeaver.exe excel-to-tsv [options] <input-excel> <output-dir>
ECWeaver.exe excel-to-tsv [options] --sheet <sheet> <input-excel> <output-tsv>
```

仕様:

- `excel-to-csv --delimiter tab` の別名として扱う。
- 現在の実装先は `ECWeaver` である。`ECWeaver2` では未実装。

### csv-to-excel

CSV を 1 シートの Excel ブックに変換する。

```txt
ECWeaver2.exe csv-to-excel [options] <input-csv> <output-excel>
```

対応 engine:

```txt
auto
interop
```

オプション:

```txt
--sheet <sheet-name>
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- 現在の実装先は `ECWeaver2` である。`ECWeaver` では未実装。
- `--sheet` 未指定時のシート名は `Sheet1` とする。

### csvs-to-excel

複数 CSV を複数シートの Excel ブックに変換する。

```txt
ECWeaver2.exe csvs-to-excel [options] <input-dir> <output-excel>
```

対応 engine:

```txt
auto
interop
```

オプション:

```txt
--pattern <file-pattern>
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- 現在の実装先は `ECWeaver2` である。`ECWeaver` では未実装。
- `--pattern` 未指定時は `*.csv` とする。
- ファイル名順に読み込む。
- シート名は拡張子なしファイル名から生成する。

### excel-to-pdf

Excel ブックを PDF に変換する。

```txt
ECWeaver.exe excel-to-pdf [options] <input-excel> <output-pdf>
```

対応 engine:

```txt
auto
app
interop
```

オプション:

```txt
--sheet <name-or-index>
--sheets <names-or-indexes>
--overwrite
```

仕様:

- 初期版ではブック全体の PDF 出力を必須対応とする。
- シート指定 PDF 出力は追加対応とする。
- 現在は `ECWeaver` と `ECWeaver2` の両方に実装済み。
- `ECWeaver` では `auto` または `app`、`ECWeaver2` では `auto` または `interop` を使用する。

### weave

複数の Excel / CSV ファイルを混在入力として受け取り、順番に処理して統合変換する。

```txt
ECWeaver.exe weave [options] <input-file>... --to-excel <output-excel>
ECWeaver.exe weave [options] <input-file>... --to-csv-dir <output-dir>
ECWeaver.exe weave [options] <input-file>... --to-same-dir <output-dir>
ECWeaver.exe weave [options] --input-list <list-file> --to-excel <output-excel>
```

対応 engine:

```txt
auto
app
interop
```

オプション:

```txt
--input-list <list-file>
--to-excel <output-excel>
--to-csv-dir <output-dir>
--to-same-dir <output-dir>
--sheet <name-or-index>
--sheets <names-or-indexes>
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- 現在の実装先は `ECWeaver2` の `--to-excel` のみである。`ECWeaver` では未実装。
- 現在の `ECWeaver2 weave --to-excel` は `.csv`、`.tsv`、`.ssv` の入力だけを受け付ける。
- 現在の `ECWeaver2 weave --to-excel` では、Excel 入力、`--to-csv-dir`、`--to-same-dir`、加工指定は未実装。
- 入力ファイルは複数指定できる。
- 入力ファイルは `.csv`、`.tsv`、`.ssv`、`.xls`、`.xlsx`、`.xlsm` の混在を許可する。
- コマンドラインで指定された入力ファイル、または `--input-list` で指定されたリストファイルを、記載順に処理する。
- `<input-file>...` と `--input-list` の同時指定は初期版ではエラーにする。
- `--input-list` は共通オプション仕様に従う。
- CSV / TSV / SSV 入力は、1 ファイルを 1 シート相当の中間データとして扱う。
- Excel 入力は、シート指定がなければ全シートを順番に読み込み、1 シートを 1 シート相当の中間データとして扱う。
- `--sheet` / `--sheets` は Excel 入力にだけ適用する。
- `--to-excel`、`--to-csv-dir`、`--to-same-dir` のいずれか 1 つを必須とする。
- 出力モードの複数同時指定はエラーにする。
- 出力先が存在する場合は `--overwrite` がない限りエラーにする。
- 初期版では加工なしの統合変換を優先し、加工指定は段階的に追加する。

出力モード:

- `--to-excel <output-excel>` は、全ての中間データを 1 つの `.xlsx` ブックのシートとして出力する。
- `--to-csv-dir <output-dir>` は、全ての中間データを CSV ファイル群として指定フォルダへ出力する。現在は未実装。
- `--to-same-dir <output-dir>` は、入力ファイル単位のまとまりを保ち、指定フォルダへ出力する。現在は未実装。

名前生成:

- `--to-excel` のシート名は入力ファイル名と元シート名から生成する。
- `--to-csv-dir` の出力ファイル名は入力ファイル名と元シート名から生成する。
- 名前が重複する場合は `_001`、`_002` のような連番を付ける。
- Excel のシート名として使えない文字や長すぎる名前は、安全な名前へ変換する。

中間データ:

- 内部では、入力順を保持した `WorkbookUnit` と `SheetUnit` のような構造に正規化する。
- `WorkbookUnit` は元入力ファイルの単位を表す。
- `SheetUnit` は CSV 1 ファイル、または Excel 1 シートに相当する表データを表す。
- 出力モードは、この中間データを `.xlsx`、CSV 群、または入力ファイル単位の出力に変換するだけにする。
- 将来の加工機能は、中間データに対する処理として追加する。

## CSV 加工系

### csv-info

CSV の基本情報を表示する。

```txt
ECWeaver.exe csv-info [options] <input-csv>
```

オプション:

```txt
--encoding <encoding>
--delimiter <delimiter>
```

出力:

```txt
Rows: 100
MinColumns: 3
MaxColumns: 5
EmptyRows: 0
```

### csv-select-columns

CSV の列を抽出する。

```txt
ECWeaver.exe csv-select-columns [options] <input-csv> <output-csv>
```

オプション:

```txt
--columns <column-indexes>
--headers <header-names>
--has-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- `--columns` または `--headers` のどちらかを必須とする。
- `--columns` と `--headers` の同時指定はエラーにする。

### csv-filter-rows

CSV の行を条件で抽出する。

```txt
ECWeaver.exe csv-filter-rows [options] <input-csv> <output-csv>
```

オプション:

```txt
--column <column-index>
--header <header-name>
--equals <text>
--contains <text>
--regex <pattern>
--invert
--has-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- `--column` または `--header` のどちらかを必須とする。
- `--equals`、`--contains`、`--regex` のいずれかを必須とする。
- 条件指定の複数同時指定はエラーにする。

### csv-replace

CSV のセル文字列を置換する。

```txt
ECWeaver.exe csv-replace [options] <input-csv> <output-csv>
```

オプション:

```txt
--from <text>
--to <text>
--regex <pattern>
--column <column-index>
--header <header-name>
--has-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- 通常置換では `--from` と `--to` を必須とする。
- 正規表現置換では `--regex` と `--to` を必須とする。
- `--column` / `--header` 未指定時は全セルを対象にする。

### csv-merge

複数 CSV を縦方向に結合する。

```txt
ECWeaver.exe csv-merge [options] <input-dir> <output-csv>
```

オプション:

```txt
--pattern <file-pattern>
--skip-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- `--pattern` 未指定時は `*.csv` とする。
- ファイル名順に結合する。
- `--skip-header` 指定時は 2 ファイル目以降の先頭行を捨てる。

### csv-sort

CSV を指定列でソートする。

```txt
ECWeaver.exe csv-sort [options] <input-csv> <output-csv>
```

オプション:

```txt
--column <column-index>
--header <header-name>
--numeric
--desc
--has-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- `--column` または `--header` のどちらかを必須とする。
- 未指定時は昇順文字列ソートとする。

### csv-unique

CSV の重複行を削除する。

```txt
ECWeaver.exe csv-unique [options] <input-csv> <output-csv>
```

オプション:

```txt
--columns <column-indexes>
--headers <header-names>
--has-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

仕様:

- `--columns` / `--headers` 未指定時は行全体で重複判定する。
- 最初に出現した行を残す。

## Excel 加工系

### excel-list-sheets

Excel ブックのシート一覧を表示する。

```txt
ECWeaver.exe excel-list-sheets [options] <input-excel>
```

対応 engine:

```txt
auto
app
```

オプション:

```txt
--output <text-file>
--overwrite
```

出力:

```txt
1	Sheet1
2	Sheet2
```

仕様:

- 現在の実装先は `ECWeaver` である。`ECWeaver2` では未実装。
- 現在は `ExcelAppTools.LoadSheets` を使うため、`auto` または `app` を使用する。

### excel-info

Excel ブックの基本情報を表示する。

```txt
ECWeaver.exe excel-info [options] <input-excel>
```

対応 engine:

```txt
auto
app
interop
zip
```

出力候補:

```txt
Sheets: 3
Pictures: 5
Size: 123456
```

仕様:

- 現在は未実装である。

### excel-extract-pictures

Excel ブック内の画像を抽出する。

```txt
ECWeaver.exe excel-extract-pictures [options] <input-excel> <output-dir>
```

対応 engine:

```txt
auto
zip
```

オプション:

```txt
--overwrite
```

仕様:

- 現在の実装先は `ECWeaver` である。`ECWeaver2` では未実装。
- `auto` は `zip` と同じ扱いにする。
- 出力ファイル名は `0001.png` のような連番にする。
- Excel アプリケーションを使わず、`.xlsx` を ZIP として展開して画像ファイルを収集する。

### excel-replace-picture

Excel ブック内の画像を置換する。

```txt
ECWeaver.exe excel-replace-picture [options] <input-excel> <output-excel> <picture-file>
```

対応 engine:

```txt
auto
zip
```

オプション:

```txt
--index <picture-index>
--overwrite
```

仕様:

- 現在の実装先は `ECWeaver` である。`ECWeaver2` では未実装。
- 初期版では `--index` 未指定時、すべての画像を指定画像で置換する。
- `--index` は 1 始まりとする。
- Excel アプリケーションを使わず、`.xlsx` を ZIP として展開して画像ファイルを置換する。

### excel-replace-text

Excel ブック内の文字列を置換する。

```txt
ECWeaver.exe excel-replace-text [options] <input-excel> <output-excel>
ECWeaver2.exe excel-replace-text [options] <input-excel> <output-excel>
```

対応 engine:

```txt
auto
app
interop
zip
```

オプション:

```txt
--from <text>
--to <text>
--regex <pattern>
--sheet <name-or-index>
--overwrite
```

仕様:

- 実装先は `ECWeaver` と `ECWeaver2` の両方とする。
- 通常置換では `--from` と `--to` を必須とする。
- 正規表現置換では `--regex` と `--to` を必須とする。
- `app` / `interop` は Excel アプリケーション経由でブックを開き、表示上のセル文字列を置換する。
- `zip` は `.xlsx` 内部 XML 文字列置換として扱うため、表示結果の完全保証はしない。
- 初期実装では `--sheet` 未指定時に全シートを対象とし、`--sheet` 指定時に指定シートだけを対象とする。

### excel-replace-placeholder

Excel テンプレート内の `**NAME**` 形式のプレースホルダを置換する。

```txt
ECWeaver.exe excel-replace-placeholder [options] <template-excel> <output-excel>
ECWeaver2.exe excel-replace-placeholder [options] <template-excel> <output-excel>
```

対応 engine:

```txt
auto
app
```

オプション:

```txt
--set <placeholder=text>
--set-file <mapping-csv>
--overwrite
```

例:

```txt
ECWeaver.exe excel-replace-placeholder --set "**NAME**=山田太郎" template.xlsx output.xlsx
```

仕様:

- 実装先は `ECWeaver` と `ECWeaver2` の両方とする。
- `--set` は複数指定可能にする。
- `--set-file` は 1 列目を置換元、2 列目を置換先として扱う。
- `--set` と `--set-file` の少なくとも一方を必須とする。
- `app` / `interop` は Excel アプリケーション経由でブックを開き、表示上のセル文字列を置換する。
- `ECWeaver` では既存の `ExcelAppTools.ReplacePlaceholder` を実装素材として利用できる。
- `ECWeaver2` では `ExcelInteropTools` 側へ同等処理を追加し、`ECWeaver` と同じ引数仕様にする。

## 検査・比較系

### csv-validate

CSV の形式を検査する。

```txt
ECWeaver.exe csv-validate [options] <input-csv>
```

オプション:

```txt
--columns <count>
--has-header
--encoding <encoding>
--delimiter <delimiter>
```

仕様:

- 検査エラーがある場合は終了コード `7` を返す。

### excel-validate

Excel ブックを検査する。

```txt
ECWeaver.exe excel-validate [options] <input-excel>
```

対応 engine:

```txt
auto
app
interop
zip
```

仕様:

- `app` / `interop` は Excel で開けるか検査する。
- `zip` は `.xlsx` として展開できるか検査する。

### csv-diff

2 つの CSV を比較する。

```txt
ECWeaver.exe csv-diff [options] <before-csv> <after-csv> <output-csv>
```

オプション:

```txt
--key-columns <column-indexes>
--has-header
--encoding <encoding>
--delimiter <delimiter>
--overwrite
```

出力列:

```txt
Row
Column
Before
After
```

### excel-diff

2 つの Excel ブックを比較する。

```txt
ECWeaver.exe excel-diff [options] <before-excel> <after-excel> <output-dir>
```

対応 engine:

```txt
auto
app
interop
```

仕様:

- 初期版では両方の Excel を CSV 化して比較する。
- シートごとの差分 CSV を出力する。

## 印刷系

### printers

利用可能なプリンタ一覧を表示する。

```txt
ECWeaver.exe printers [options]
```

対応 engine:

```txt
auto
app
interop
```

現在の実装:

- `ECWeaver` は `auto` または `app`。
- `ECWeaver2` は `auto` または `interop`。

### print

Excel ブックを印刷する。

```txt
ECWeaver.exe print [options] <input-excel>
```

対応 engine:

```txt
auto
app
interop
```

オプション:

```txt
--printer <printer-name>
--sheet <name-or-index>
```

仕様:

- `--printer` 未指定時は既定プリンタに出力する。
- 印刷は破壊的ではないが、実行前に GUI 側では確認を出す。
- 現在は `--sheet` 未対応で、ブック全体を印刷する。
- `ECWeaver` は `auto` または `app`、`ECWeaver2` は `auto` または `interop` を使用する。

## 自動化系

### run-script

処理定義ファイルに従って複数コマンドを実行する。

```txt
ECWeaver.exe run-script [options] <script-file>
```

オプション:

```txt
--stop-on-error
--continue-on-error
```

仕様:

- 初期版では 1 行 1 コマンドのテキスト形式とする。
- 空行を無視する。
- `#` で始まる行をコメントとして扱う。
- 未指定時は `--stop-on-error` とする。

## GUI 連携方針

GUI はコマンドライン文字列を組み立てるだけにしない。

内部構成案:

```txt
CommandDefinition
CommandOptions
CommandRunner
CommandResult
```

方針:

- CLI は引数を `CommandOptions` に変換して `CommandRunner` を呼ぶ。
- GUI は画面入力を `CommandOptions` に変換して `CommandRunner` を呼ぶ。
- `CommandRunner` は標準出力や MessageBox に直接依存しない。
- CLI 表示、GUI 表示、ログ出力は外側で担当する。
- コマンド定義には、必須引数、任意オプション、対応 engine、上書き可否を持たせる。

## エラー表示

CLI 標準:

```txt
Error: input file not found: input.xlsx
```

詳細表示:

```txt
Error: Excel open failed.
Detail: 指定されたエクセルファイルは破損しているか、対応していない形式です。
```

仕様:

- `--verbose` 指定時は例外詳細も表示する。
- `--no-dialog` 指定時は MessageBox を出さない。
- GUI 呼び出し時は `CommandResult` のエラー情報を画面に表示する。

## 終了コード

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

## 実装優先度

### Phase 1

```txt
help
version
excel-to-csv
excel-to-pdf
csv-info
csv-select-columns
csv-filter-rows
csv-replace
```

### Phase 2

```txt
csv-to-excel
csvs-to-excel
weave
excel-list-sheets
excel-info
excel-extract-pictures
excel-replace-picture
csv-merge
csv-sort
csv-unique
```

### Phase 3

```txt
csv-validate
excel-validate
csv-diff
excel-diff
printers
print
```

### Phase 4

```txt
excel-replace-text
excel-replace-placeholder
weave の加工パイプライン拡張
run-script
```

## 未決事項

- `ECWeaver` と `ECWeaver2` の両方で完全に同じコマンドセットを提供するか。
- `--engine app` と `--engine interop` の結果差異をどこまで吸収するか。
- CSV 出力の標準文字コードを Shift_JIS のままにするか。
- GUI から外部プロセスとして CLI を呼ぶモードも残すか。
- `--sheet 1` をシート番号として扱う仕様で、数値名シートをどう扱うか。
- 複雑な条件指定を CLI オプションで続けるか、設定ファイル形式へ逃がすか。
