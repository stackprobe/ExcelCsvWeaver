==============
ExcelCsvWeaver
==============


ExcelCsvWeaver は、Excel / CSV を変換・加工するためのアプリケーション一式です。

この配布物には、インストーラ、GUI プログラム、2 種類のコマンドラインツールが
含まれています。

通常の処理は、主に以下のコマンドラインツールから実行します。

	ECWeaver.exe
	ECWeaver2.exe

ECWeaverGUI は、将来 GUI が必要になった場合に備えて残しているプログラムです。
現時点では、専用 GUI ワークフローは未計画・未実装です。


----
内容

この配布物には、以下のプログラムが含まれます。

Installer.exe

	ExcelCsvWeaver をインストール、再インストール、アンインストールするための
	インストーラです。

ECWeaverGUI.exe

	将来 GUI 用のプログラムです。
	現時点ではプレースホルダーであり、実際の処理にはコマンドラインツールを
	使用します。

ECWeaver.exe

	Excel アプリケーション操作や .xlsx の直接処理を中心にした
	コマンドラインツールです。

	Excel から CSV / TSV / PDF への変換、Excel ブック内画像の抽出・置換、
	CSV の基本的な加工を行えます。

ECWeaver2.exe

	Microsoft Office Interop による Excel 操作を中心にした
	コマンドラインツールです。

	CSV から Excel ブックを作成する処理、複数 CSV の Excel 化、
	CSV / TSV / SSV を 1 つの Excel ブックにまとめる処理、
	CSV の基本的な加工を行えます。


----
インストール

Installer.exe を実行すると、インストール先フォルダを指定できます。

既定のインストール先は、ユーザーの LocalAppData 配下です。

	%LOCALAPPDATA%\ExcelCsvWeaver

インストール時には、デスクトップにショートカットを作成できます。
ショートカットの作成が不要な場合は、インストーラ画面でチェックを外してください。

インストール先フォルダが既に存在する場合は、確認メッセージが表示されます。
再インストール時は、必要なプログラムファイルを再配置します。


----
アンインストール

既にインストール済みの場合、Installer.exe からアンインストールできます。

アンインストールを実行すると、インストール先フォルダにあるファイルが削除されます。
デスクトップ上のショートカットを削除するかどうかは、実行時に確認されます。


----
コマンドラインツール

インストール後、以下の実行ファイルを使用します。

	ECWeaver\ECWeaver.exe
	ECWeaver2\ECWeaver2.exe

基本形:

	ECWeaver.exe <コマンド> [オプション] [引数...]
	ECWeaver2.exe <コマンド> [オプション] [引数...]

例:

	ECWeaver.exe excel-to-csv Book.xlsx OutDir
	ECWeaver.exe excel-to-pdf Book.xlsx Book.pdf
	ECWeaver.exe csv-select-columns --columns 1,3,5 input.csv output.csv
	ECWeaver2.exe csv-to-excel input.csv output.xlsx
	ECWeaver2.exe csvs-to-excel CsvDir output.xlsx
	ECWeaver2.exe weave input1.csv input2.tsv --to-excel output.xlsx


----
ヘルプ・バージョン

利用可能なコマンドは、各ツールの help で確認できます。

	ECWeaver.exe help
	ECWeaver2.exe help

指定コマンドの詳細を確認する場合:

	ECWeaver.exe help <コマンド>
	ECWeaver2.exe help <コマンド>

バージョンを確認する場合:

	ECWeaver.exe version
	ECWeaver2.exe version


----
ECWeaver の主な用途

ECWeaver は、Excel アプリケーション操作を使う処理と、
.xlsx を ZIP / Open XML として直接扱う処理を中心にしています。

主なコマンド:

	excel-to-csv
	excel-to-tsv
	excel-to-pdf
	excel-list-sheets
	excel-extract-pictures
	excel-replace-picture
	excel-replace-text
	excel-replace-placeholder
	csv-info
	csv-select-columns
	csv-filter-rows
	csv-replace
	csv-merge
	csv-sort
	csv-unique
	printers
	print

Excel から CSV / TSV への変換、Excel ブック内画像の抽出・置換は、
ECWeaver を使用します。


----
ECWeaver2 の主な用途

ECWeaver2 は、Microsoft Office Interop による Excel 操作を中心にしています。

主なコマンド:

	csv-to-excel
	csvs-to-excel
	excel-to-pdf
	weave
	excel-replace-text
	excel-replace-placeholder
	csv-info
	csv-select-columns
	csv-filter-rows
	csv-replace
	csv-merge
	csv-sort
	csv-unique
	printers
	print

CSV から Excel ブックを作成する処理や、
複数の CSV / TSV / SSV を 1 つの Excel ブックにまとめる処理は、
ECWeaver2 を使用します。


----
共通オプション

--overwrite

	出力先が既に存在する場合に上書きします。
	指定しない場合、出力ファイルまたは出力フォルダが存在するとエラーになります。

--encoding (auto | sjis | utf8 | utf8bom | utf16le)

	CSV / TSV / SSV の文字コードを指定します。
	入力時の既定値は auto、出力時の既定値は sjis です。

--delimiter (comma | tab | space | 1文字)

	CSV 系ファイルの区切り文字を指定します。
	未指定時は拡張子で判定します。

--engine

	Excel 操作方式を指定します。
	指定できる値は、使用するツールとコマンドによって異なります。

--response レスポンスファイル

	レスポンスファイルを読み込み、各行を1つのコマンドライン引数として扱います。
	レスポンスファイルは SJIS のテキストファイルとして読み込みます。

--silent

	通常メッセージをコンソールへ出力しません。


----
GUI について

ECWeaverGUI は、将来 GUI が必要になった場合に備えたプレースホルダーです。

現時点では、変換、加工、検査、印刷などを GUI から操作する専用画面は
未計画・未実装です。

実際の処理には、ECWeaver.exe または ECWeaver2.exe を使用してください。


----
注意

Excel 操作を行うコマンドは、Excel が利用できる環境で実行してください。

出力先が既に存在する場合は、--overwrite を指定しない限りエラーになります。

ECWeaver と ECWeaver2 では、同じ名前のコマンドでも内部の Excel 操作方式が
異なる場合があります。
用途に応じて、各ツールの Readme.txt または help を確認してください。
