========
ECWeaver
========


Excel / CSV を変換・加工するコマンドラインツールです。

ECWeaver は Excel アプリケーション操作を使う処理を中心にしています。
Excel から CSV / TSV / PDF への変換、Excel ブック内画像の抽出・置換、
CSV の基本的な加工を行えます。


----
コマンド

ECWeaver.exe <コマンド> [オプション] [引数...]

例:

	ECWeaver.exe excel-to-csv Book.xlsx OutDir
	ECWeaver.exe excel-to-csv --sheet Sheet1 Book.xlsx Sheet1.csv
	ECWeaver.exe csv-select-columns --columns 1,3,5 input.csv output.csv
	ECWeaver.exe --response args.txt


----
ヘルプ・バージョン

ECWeaver.exe
ECWeaver.exe --help
ECWeaver.exe help

	全体ヘルプを表示する。

ECWeaver.exe help <コマンド>
ECWeaver.exe <コマンド> --help

	指定コマンドのヘルプを表示する。

ECWeaver.exe version
ECWeaver.exe --version

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

--engine (auto | app | zip)

	Excel 操作方式を指定する。

	auto ... コマンドに応じて自動選択する。
	app  ... Excel アプリケーション操作を使う。
	zip  ... .xlsx を ZIP として直接処理する。

	Excel からの読み込み、PDF、印刷では auto または app を指定できる。
	画像抽出・画像置換では auto または zip を指定できる。
	CSV 専用コマンドでは --engine は指定できない。

--response レスポンスファイル

	レスポンスファイルを読み込み、各行を1つのコマンドライン引数として扱う。
	--response=レスポンスファイル 形式でも指定できる。
	レスポンスファイル内に --response を記述した場合も展開する。
	循環参照、値不足、ファイル不存在はエラーになる。
	レスポンスファイルは SJIS のテキストファイルとして読み込む。

--silent

	通常メッセージをコンソールへ出力しない。


----
Excel 変換

ECWeaver.exe excel-to-csv [--sheet シート] [--sheets シートリスト] [--overwrite] 入力Excel 出力先

	Excel ブックを CSV に変換する。

	--sheet を指定しない場合:
		出力先はフォルダとして扱う。
		対象シートを 0001.csv, 0002.csv, ... として出力する。
		sheet-names.txt にシート名一覧を出力する。

	--sheet を指定した場合:
		出力先は単一 CSV ファイルとして扱う。
		指定シートだけを出力する。

	--sheets を指定した場合:
		カンマ区切りで複数シートを指定する。
		出力先はフォルダとして扱う。

	シート指定は、シート名または 1 始まりのシート番号で指定する。

例:

	ECWeaver.exe excel-to-csv Book.xlsx OutCsv
	ECWeaver.exe excel-to-csv --sheet 1 Book.xlsx Sheet1.csv
	ECWeaver.exe excel-to-csv --sheets Sheet1,Sheet3 Book.xlsx OutCsv


ECWeaver.exe excel-to-tsv [--sheet シート] [--sheets シートリスト] [--overwrite] 入力Excel 出力先

	Excel ブックを TSV に変換する。
	基本動作は excel-to-csv と同じ。
	出力ファイル名は 0001.tsv, 0002.tsv, ... になる。


ECWeaver.exe excel-to-pdf [--overwrite] 入力Excel 出力PDF

	Excel ブックを PDF に変換する。
	Excel アプリケーション操作を使用する。


----
Excel 情報・画像・印刷

ECWeaver.exe excel-list-sheets [--output 出力テキスト] 入力Excel

	Excel ブックのシート一覧を出力する。
	--output を指定しない場合はコンソールへ出力する。
	--output を指定した場合は指定ファイルへ出力する。

出力例:

	1	Sheet1
	2	Sheet2


ECWeaver.exe excel-extract-pictures [--overwrite] 入力Excel 出力フォルダ

	Excel ブック内の画像を抽出する。
	出力フォルダに 0001.png, 0002.jpeg, ... のように連番で出力する。
	このコマンドの --engine は auto または zip。


ECWeaver.exe excel-replace-picture [--index 番号] [--overwrite] 入力Excel 出力Excel 画像ファイル

	Excel ブック内の画像を指定画像で置換する。

	--index を指定しない場合:
		ブック内の全画像を置換する。

	--index を指定した場合:
		1 始まりの画像番号に一致する画像だけを置換する。

	このコマンドの --engine は auto または zip。


ECWeaver.exe excel-replace-text (--from 文字列 | --regex 正規表現) --to 置換後 [--sheet シート] [--overwrite] 入力Excel 出力Excel

	Excel ブック内のセル文字列を置換する。

	--from は通常の文字列置換。
	--regex は正規表現置換。
	--to は置換後文字列。
	--sheet を指定すると、指定シートだけを対象にする。
	--sheet を省略すると、全シートを対象にする。
	シート指定は、シート名または 1 始まりのシート番号で指定する。
	このコマンドの --engine は auto または app。

例:

	ECWeaver.exe excel-replace-text --from old --to new Book.xlsx Out.xlsx
	ECWeaver.exe excel-replace-text --regex ""Item[0-9]+"" --to ItemX Book.xlsx Out.xlsx


ECWeaver.exe excel-replace-placeholder (--set プレースホルダ=置換後 | --set-file 置換CSV) [--overwrite] テンプレートExcel 出力Excel

	Excel テンプレート内のプレースホルダを置換する。

	--set は プレースホルダ=置換後 の形式で指定する。
	--set は複数指定できる。
	--set-file は 1 列目を置換元、2 列目を置換後として読み込む。
	このコマンドの --engine は auto または app。

例:

	ECWeaver.exe excel-replace-placeholder --set ""**NAME**=山田太郎"" Template.xlsx Out.xlsx
	ECWeaver.exe excel-replace-placeholder --set-file mapping.csv Template.xlsx Out.xlsx

ECWeaver.exe printers

	利用可能なプリンタ名を一覧表示する。


ECWeaver.exe print [--printer プリンタ名] 入力Excel

	Excel ブックを印刷する。
	--printer を指定しない場合は既定プリンタへ印刷する。


----
CSV 情報・加工

以下の CSV 系コマンドは、.csv / .tsv / .ssv を扱える。
区切り文字は --delimiter で明示できる。
--engine は指定できない。


ECWeaver.exe csv-info [--encoding 文字コード] [--delimiter 区切り] 入力CSV

	CSV の行数・列数などを表示する。

出力例:

	Rows: 100
	MinColumns: 3
	MaxColumns: 5
	EmptyRows: 0


ECWeaver.exe csv-select-columns (--columns 列番号リスト | --headers ヘッダー名リスト) [--overwrite] 入力CSV 出力CSV

	指定した列だけを抽出する。

	--columns は 1 始まりの列番号をカンマ区切りで指定する。
	--headers は 1 行目をヘッダーとして、ヘッダー名をカンマ区切りで指定する。
	--columns と --headers は同時に指定できない。

例:

	ECWeaver.exe csv-select-columns --columns 1,3,5 input.csv output.csv
	ECWeaver.exe --response args.txt
	ECWeaver.exe csv-select-columns --headers Code,Name,Price input.csv output.csv


ECWeaver.exe csv-filter-rows (--column 列番号 | --header ヘッダー名) (--equals 文字列 | --contains 文字列 | --regex 正規表現) [--invert] [--has-header] [--overwrite] 入力CSV 出力CSV

	条件に一致する行だけを出力する。

	--column は 1 始まりの列番号を指定する。
	--header は 1 行目をヘッダーとして、対象列をヘッダー名で指定する。
	--equals は完全一致、--contains は部分一致、--regex は正規表現一致。
	条件指定は1つだけ指定する。
	--invert を指定すると、一致しない行を出力する。
	--has-header または --header 指定時は、先頭行をヘッダーとして常に残す。

例:

	ECWeaver.exe csv-filter-rows --column 2 --equals Tokyo input.csv output.csv
	ECWeaver.exe csv-filter-rows --header Status --contains error input.csv output.csv


ECWeaver.exe csv-replace (--from 文字列 | --regex 正規表現) --to 置換後 [--column 列番号 | --header ヘッダー名] [--overwrite] 入力CSV 出力CSV

	CSV セル内の文字列を置換する。

	--from は通常の文字列置換。
	--regex は正規表現置換。
	--to は置換後文字列。
	--column または --header を指定すると、対象列だけを置換する。
	列指定を省略すると全セルを対象にする。

例:

	ECWeaver.exe csv-replace --from old --to new input.csv output.csv
	ECWeaver.exe csv-replace --regex " + " --to " " input.csv output.csv


ECWeaver.exe csv-merge [--pattern ファイルパターン] [--skip-header] [--overwrite] 入力フォルダ 出力CSV

	入力フォルダ内の CSV をファイル名順で縦に結合する。

	--pattern の既定値は *.csv。
	--skip-header を指定すると、2 ファイル目以降の先頭行を除外する。


ECWeaver.exe csv-sort (--column 列番号 | --header ヘッダー名) [--numeric] [--desc] [--has-header] [--overwrite] 入力CSV 出力CSV

	指定列でソートする。

	--numeric を指定すると数値として比較する。
	--desc を指定すると降順にする。
	--has-header または --header 指定時は、先頭行をヘッダーとして固定する。


ECWeaver.exe csv-unique [--columns 列番号リスト | --headers ヘッダー名リスト] [--overwrite] 入力CSV 出力CSV

	重複行を削除する。

	--columns または --headers を指定した場合、指定列だけで重複判定する。
	省略した場合、行全体で重複判定する。
	最初に出現した行を残す。


----
未実装のコマンド

以下のコマンド名は予約されているが、ECWeaver では未実装。
実行すると未実装エラーになる。

	csv-to-excel
	csvs-to-excel
	weave
	excel-info
	csv-validate
	excel-validate
	csv-diff
	excel-diff
	run-script

CSV から Excel ブックを作成する処理や weave --to-excel は ECWeaver2 を使用する。


----
注意

Excel 操作を行うコマンドは、Excel が利用できる環境で実行すること。

出力先が既に存在する場合は、--overwrite を指定しない限りエラーになる。
上書き時は既存の出力先を削除してから作成する。
