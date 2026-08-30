=========
ECWeaver2
=========


Excel / CSV を変換・加工するコマンドラインツールです。

ECWeaver2 は Microsoft Office Interop による Excel 操作を中心にしています。
CSV から Excel ブックを作成する処理、複数 CSV の Excel 化、
CSV の基本的な加工を行えます。


----
コマンド

ECWeaver2.exe <コマンド> [オプション] [引数...]

例:

	ECWeaver2.exe csv-to-excel input.csv output.xlsx
	ECWeaver2.exe csvs-to-excel CsvDir output.xlsx
	ECWeaver2.exe weave input1.csv input2.tsv --to-excel output.xlsx


----
ヘルプ・バージョン

ECWeaver2.exe
ECWeaver2.exe --help
ECWeaver2.exe help

	全体ヘルプを表示する。

ECWeaver2.exe help <コマンド>
ECWeaver2.exe <コマンド> --help

	指定コマンドのヘルプを表示する。

ECWeaver2.exe version
ECWeaver2.exe --version

	バージョンを表示する。


----
オプションの指定方法

オプションは -- で始まる名前で指定する。

値を取るオプションは、以下のどちらでも指定できる。

	--encoding sjis
	--encoding=sjis

フラグオプションは名前だけを指定する。

	--overwrite
	--silent

パスを表す引数・オプション値は、実行時にフルパスへ変換される。


----
共通オプション

--overwrite

	出力先が既に存在する場合に上書きする。
	指定しない場合、出力ファイルまたは出力フォルダが存在するとエラーになる。

--encoding (auto | sjis | utf8 | utf8bom | utf16le)

	CSV / TSV / SSV の文字コードを指定する。

	入力時の既定値は auto。
	出力時の既定値は sjis。

	auto を出力に指定した場合は sjis として扱う。

--delimiter (comma | tab | space | 1文字)

	CSV 系ファイルの区切り文字を指定する。

	comma ... カンマ
	tab   ... タブ
	space ... 半角スペース
	1文字 ... 指定した1文字

	未指定時は拡張子で判定する。
	.csv は comma、.tsv は tab、.ssv は space になる。

--engine (auto | interop)

	Excel 操作方式を指定する。

	auto    ... Interop を使用する。
	interop ... Microsoft Office Interop を使用する。

	Excel 作成、PDF、印刷では auto または interop を指定できる。
	CSV 専用コマンドでは --engine は指定できない。

--silent

	通常メッセージをコンソールへ出力しない。


----
Excel 変換

ECWeaver2.exe csv-to-excel [--sheet シート名] [--overwrite] 入力CSV 出力Excel

	CSV / TSV / SSV を 1 シートの Excel ブックとして保存する。

	--sheet を指定しない場合、シート名は Sheet1 になる。
	入力ファイルの区切り文字は拡張子または --delimiter で決まる。

例:

	ECWeaver2.exe csv-to-excel input.csv output.xlsx
	ECWeaver2.exe csv-to-excel --sheet Data input.tsv output.xlsx


ECWeaver2.exe csvs-to-excel [--pattern ファイルパターン] [--overwrite] 入力フォルダ 出力Excel

	入力フォルダ内の CSV ファイルを、複数シートの Excel ブックとして保存する。

	--pattern の既定値は *.csv。
	ファイル名順に読み込み、各ファイルを 1 シートにする。
	シート名は拡張子を除いたファイル名から作成する。

例:

	ECWeaver2.exe csvs-to-excel CsvDir output.xlsx
	ECWeaver2.exe csvs-to-excel --pattern "*.tsv" CsvDir output.xlsx


ECWeaver2.exe excel-to-pdf [--overwrite] 入力Excel 出力PDF

	Excel ブックを PDF に変換する。
	Microsoft Office Interop を使用する。


ECWeaver2.exe weave 入力CSV... --to-excel 出力Excel [--overwrite]
ECWeaver2.exe weave --input-list 入力リスト --to-excel 出力Excel [--overwrite]

	複数の CSV / TSV / SSV を 1 つの Excel ブックにまとめる。
	各入力ファイルを 1 シートとして出力する。

	入力ファイルはコマンドライン上に複数指定できる。
	または --input-list でリストファイルを指定できる。
	入力ファイル指定と --input-list は同時に指定できない。

	--to-excel は必須。
	--to-csv-dir と --to-same-dir は名前だけ予約されているが、ECWeaver2 では未実装。

	入力リストは SJIS のテキストファイル。
	1 行 1 ファイルで指定する。
	空行と # で始まる行は無視する。
	相対パスは入力リストのあるフォルダからの相対パスとして扱う。

例:

	ECWeaver2.exe weave input1.csv input2.tsv input3.ssv --to-excel output.xlsx
	ECWeaver2.exe weave --input-list files.txt --to-excel output.xlsx


ECWeaver2.exe excel-replace-text (--from 文字列 | --regex 正規表現) --to 置換後 [--sheet シート] [--overwrite] 入力Excel 出力Excel

	Excel ブック内のセル文字列を置換する。

	--from は通常の文字列置換。
	--regex は正規表現置換。
	--to は置換後文字列。
	--sheet を指定すると、指定シートだけを対象にする。
	--sheet を省略すると、全シートを対象にする。
	シート指定は、シート名または 1 始まりのシート番号で指定する。
	このコマンドの --engine は auto または interop。

例:

	ECWeaver2.exe excel-replace-text --from old --to new Book.xlsx Out.xlsx
	ECWeaver2.exe excel-replace-text --regex ""Item[0-9]+"" --to ItemX Book.xlsx Out.xlsx


ECWeaver2.exe excel-replace-placeholder (--set プレースホルダ=置換後 | --set-file 置換CSV) [--overwrite] テンプレートExcel 出力Excel

	Excel テンプレート内のプレースホルダを置換する。

	--set は プレースホルダ=置換後 の形式で指定する。
	--set は複数指定できる。
	--set-file は 1 列目を置換元、2 列目を置換後として読み込む。
	このコマンドの --engine は auto または interop。

例:

	ECWeaver2.exe excel-replace-placeholder --set ""**NAME**=山田太郎"" Template.xlsx Out.xlsx
	ECWeaver2.exe excel-replace-placeholder --set-file mapping.csv Template.xlsx Out.xlsx

----
印刷

ECWeaver2.exe printers

	利用可能なプリンタ名を一覧表示する。


ECWeaver2.exe print [--printer プリンタ名] 入力Excel

	Excel ブックを印刷する。
	--printer を指定しない場合は既定プリンタへ印刷する。


----
CSV 情報・加工

以下の CSV 系コマンドは、.csv / .tsv / .ssv を扱える。
区切り文字は --delimiter で明示できる。
--engine は指定できない。


ECWeaver2.exe csv-info [--encoding 文字コード] [--delimiter 区切り] 入力CSV

	CSV の行数・列数などを表示する。

出力例:

	Rows: 100
	MinColumns: 3
	MaxColumns: 5
	EmptyRows: 0


ECWeaver2.exe csv-select-columns (--columns 列番号リスト | --headers ヘッダー名リスト) [--overwrite] 入力CSV 出力CSV

	指定した列だけを抽出する。

	--columns は 1 始まりの列番号をカンマ区切りで指定する。
	--headers は 1 行目をヘッダーとして、ヘッダー名をカンマ区切りで指定する。
	--columns と --headers は同時に指定できない。

例:

	ECWeaver2.exe csv-select-columns --columns 1,3,5 input.csv output.csv
	ECWeaver2.exe csv-select-columns --headers Code,Name,Price input.csv output.csv


ECWeaver2.exe csv-filter-rows (--column 列番号 | --header ヘッダー名) (--equals 文字列 | --contains 文字列 | --regex 正規表現) [--invert] [--has-header] [--overwrite] 入力CSV 出力CSV

	条件に一致する行だけを出力する。

	--column は 1 始まりの列番号を指定する。
	--header は 1 行目をヘッダーとして、対象列をヘッダー名で指定する。
	--equals は完全一致、--contains は部分一致、--regex は正規表現一致。
	条件指定は1つだけ指定する。
	--invert を指定すると、一致しない行を出力する。
	--has-header または --header 指定時は、先頭行をヘッダーとして常に残す。

例:

	ECWeaver2.exe csv-filter-rows --column 2 --equals Tokyo input.csv output.csv
	ECWeaver2.exe csv-filter-rows --header Status --contains error input.csv output.csv


ECWeaver2.exe csv-replace (--from 文字列 | --regex 正規表現) --to 置換後 [--column 列番号 | --header ヘッダー名] [--overwrite] 入力CSV 出力CSV

	CSV セル内の文字列を置換する。

	--from は通常の文字列置換。
	--regex は正規表現置換。
	--to は置換後文字列。
	--column または --header を指定すると、対象列だけを置換する。
	列指定を省略すると全セルを対象にする。

例:

	ECWeaver2.exe csv-replace --from old --to new input.csv output.csv
	ECWeaver2.exe csv-replace --regex " + " --to " " input.csv output.csv


ECWeaver2.exe csv-merge [--pattern ファイルパターン] [--skip-header] [--overwrite] 入力フォルダ 出力CSV

	入力フォルダ内の CSV をファイル名順で縦に結合する。

	--pattern の既定値は *.csv。
	--skip-header を指定すると、2 ファイル目以降の先頭行を除外する。


ECWeaver2.exe csv-sort (--column 列番号 | --header ヘッダー名) [--numeric] [--desc] [--has-header] [--overwrite] 入力CSV 出力CSV

	指定列でソートする。

	--numeric を指定すると数値として比較する。
	--desc を指定すると降順にする。
	--has-header または --header 指定時は、先頭行をヘッダーとして固定する。


ECWeaver2.exe csv-unique [--columns 列番号リスト | --headers ヘッダー名リスト] [--overwrite] 入力CSV 出力CSV

	重複行を削除する。

	--columns または --headers を指定した場合、指定列だけで重複判定する。
	省略した場合、行全体で重複判定する。
	最初に出現した行を残す。


----
未実装のコマンド

以下のコマンド名は予約されているが、ECWeaver2 では未実装。
実行すると未実装エラーになる。

	excel-to-csv
	excel-to-tsv
	excel-list-sheets
	excel-info
	excel-extract-pictures
	excel-replace-picture
	csv-validate
	excel-validate
	csv-diff
	excel-diff
	run-script

Excel から CSV / TSV への変換、Excel ブック内画像の抽出・置換は ECWeaver を使用する。


----
注意

Excel 操作を行うコマンドは、Excel が利用できる環境で実行すること。

出力先が既に存在する場合は、--overwrite を指定しない限りエラーになる。
上書き時は既存の出力先を削除してから作成する。
