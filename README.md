# ExcelCsvWeaver

ExcelCsvWeaver は、Excel / CSV の変換・加工を行うための Windows 向けツール群です。

主な利用対象はコマンドラインツールです。GUI プログラムも含まれていますが、現時点では将来開発用のプレースホルダーであり、専用 GUI ワークフローは未完成です。

## 構成

- `ECWeaver`
  - Excel アプリケーション操作や `.xlsx` の直接処理を中心にしたコマンドラインツールです。
  - Excel から CSV / TSV / PDF への変換、Excel ブック内画像の抽出・置換、CSV 加工などを扱います。
- `ECWeaver2`
  - Microsoft Office Interop による Excel 操作を中心にしたコマンドラインツールです。
  - CSV から Excel への変換、複数 CSV の Excel 化、`weave` による統合変換、CSV 加工などを扱います。
- `ECWeaverGUI`
  - 将来 GUI が必要になった場合に備えたプログラムです。
  - 現時点では未計画・未実装の機能が多いため、通常は CUI ツールを使用してください。
- `Installer`
  - 配布用インストーラです。
- `_spec`
  - コマンド仕様、機能方針、フォルダ構成などの仕様メモです。

## 使い方

各コマンドラインツールのヘルプを確認してください。

```bat
ECWeaver.exe help
ECWeaver2.exe help
```

例:

```bat
ECWeaver.exe excel-to-csv Book.xlsx OutDir
ECWeaver.exe excel-to-pdf Book.xlsx Book.pdf
ECWeaver2.exe csv-to-excel input.csv output.xlsx
ECWeaver2.exe weave input1.csv input2.tsv --to-excel output.xlsx
```

## 詳細

詳細な仕様は `_spec` 配下の Markdown を参照してください。

- `_spec/Project_Folder_Structure.md`
- `_spec/ECWeaver_CommandLine.md`
- `_spec/ECWeaver_CommandTool_Features.md`
- `_spec/ECWeaver_GUI.md`

各ツールの利用説明は、それぞれの `doc/Readme.txt` も参照してください。

- `ECWeaver/doc/Readme.txt`
- `ECWeaver2/doc/Readme.txt`
- `ECWeaverGUI/doc/Readme.txt`
- `Installer/doc/Readme.txt`

実装の詳細は各プロジェクトのソースコードを参照してください。

## 注意

Excel 操作を行うコマンドは、Excel が利用できる環境で実行してください。

出力先が既に存在する場合、多くのコマンドでは `--overwrite` を指定しない限りエラーになります。
