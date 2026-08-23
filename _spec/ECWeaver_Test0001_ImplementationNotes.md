# ECWeaver Test0001 実装メモ

## 目的

この文書は、`ECWeaver` 側に実装した `Tests.Test0001.Test01()` の概要と、`ECWeaver2` 側へ同種のテストを実装するときの作業指針をまとめる。

別チャットで `ECWeaver2` 側のテスト実装を依頼するときは、まずこの文書を読ませること。

依頼例:

```txt
_spec/ECWeaver_Test0001_ImplementationNotes.md を読んで、ECWeaver2 に Test0001.Test01 のテストを実装してください。
```

## ECWeaver 側で実装済みの場所

対象ファイル:

```txt
ECWeaver/HLTConsole/HLTConsole/Tests/Test0001.cs
ECWeaver/HLTConsole/HLTConsole/HLTConsole.csproj
```

`HLTConsole.csproj` には、テスト内で `ZipArchive` を使うために以下の参照を追加している。

```xml
<Reference Include="System.IO.Compression" />
```

## テスト起動方法

Debug ビルドで引数なし実行すると、`Program.Main3()` から以下が呼び出される。

```csharp
new Test0001().Test01();
```

テストフレームワークは使わず、`Test01()` が目的別の private メソッドを順に呼び出す構造にしている。

テスト失敗時は例外を投げる。例外は既存の上位処理で捕捉され、ログ出力とエラー表示が行われる。

テスト成功時はコンソールへ大きめの成功バナーを出力する。

## ディレクトリ方針

ECWeaver 側では以下を使っている。

```txt
永続テストデータ置き場:
C:\home\res\ExcelCsvWeaver\ECWeaver\Test0001

作業ディレクトリ:
C:\temp\ECWeaver
```

`Test01()` の最初で必ず以下を実行する。

```csharp
SCommon.DeleteAndCreateDir(BaseDir);
```

ECWeaver 側の現在の定数は以下。

```csharp
private const string ResourceDir = @"C:\home\res\ExcelCsvWeaver\ECWeaver\Test0001";
private const string BaseDir = @"C:\temp\ECWeaver";
private const string InputDir = ResourceDir + @"\input";
private const string OutputDir = BaseDir + @"\output";
```

ECWeaver2 側へ移植するときは、原則として以下のように置き換える。

```csharp
private const string ResourceDir = @"C:\home\res\ExcelCsvWeaver\ECWeaver2\Test0001";
private const string BaseDir = @"C:\temp\ECWeaver2";
private const string InputDir = ResourceDir + @"\input";
private const string OutputDir = BaseDir + @"\output";
```

`C:\home\res\ExcelCsvWeaver\ECWeaver2\` は永続的に存在すると考えてよい。必要な CSV / XLSX などのテストデータは、この配下へ自動生成または配置し、テストではフルパスで参照する。

## 実装スタイル

重要な方針:

- `Tests.Test0001` は既存と同様に `#if DEBUG` で囲む。
- Release 版では空のクラスになるようにする。
- 失敗時は例外を投げる。
- 成功時は分かりやすい成功メッセージをコンソールに表示する。
- 目視確認が必要なテストにしない。
- 目的ごとに private メソッドを分ける。
- 各テストメソッドには XML コメントでテスト目的を書く。
- テストデータはテスト開始時に作り直す。
- 出力先は毎回 `SCommon.DeleteAndCreateDir(BaseDir)` でクリーンにする。
- 外部テストフレームワークは追加しない。
- 既存の `ECWeaverProcessor` / `ECWeaverArgs` を直接呼び出して処理を実行する。

ECWeaver 側では以下のようなヘルパーを `Test0001` 内に置いている。

- `Run(string[] args)`
- `CaptureOutput(string[] args)`
- `InputFile(string localPath)`
- `OutputFile(string localPath)`
- `ReadCsv(string file)`
- `WriteCsv(...)`
- `AssertRows(...)`
- `AssertBytes(...)`
- `AssertThrows(...)`
- `AssertContains(...)`
- `AssertEquals(...)`
- `AssertTrue(...)`
- `ShowSuccessBanner()`

ECWeaver2 側でも同じ構造を使ってよい。

## ECWeaver 側で網羅した内容

ECWeaver 側の `Test0001.Test01()` では、Excel アプリケーションに依存しない基本動作を中心にテストしている。

### 共通・引数系

- `help` の全体ヘルプ出力
- `help <command>` のコマンド別ヘルプ出力
- `version` の出力
- 未知コマンドのエラー
- ECWeaver 未実装コマンドのエラー
- 引数個数不正のエラー
- 入力ファイルなしのエラー
- `--silent` による標準出力抑制

### CSV 情報

- `csv-info`
- 行数
- 最小列数
- 最大列数
- 空行数

### CSV 列抽出

- `csv-select-columns --columns`
- `csv-select-columns --headers`
- `--columns` / `--headers` の不足
- `--columns` / `--headers` の同時指定
- 列番号 `0`
- 範囲外列
- 存在しないヘッダー

### CSV 行抽出

- `csv-filter-rows --equals`
- `csv-filter-rows --contains`
- `csv-filter-rows --regex`
- `--invert`
- `--has-header`
- `--header` 指定時のヘッダー行維持
- 列指定不足
- 条件不足
- 条件の複数指定
- CSV コマンドへの `--engine` 指定エラー

### CSV 置換

- `csv-replace --from --to`
- 全セル置換
- `--column` 指定置換
- `--header` 指定の正規表現置換
- `--to` 不足
- `--from` / `--regex` 不足
- `--from` / `--regex` 同時指定

### CSV 結合

- `csv-merge`
- ファイル名順の結合
- `--skip-header`
- `--pattern`
- 入力ディレクトリなしのエラー

### CSV ソート

- `csv-sort` の文字列昇順
- `--numeric`
- `--desc`
- `--has-header`
- `--header`

### CSV 重複削除

- `csv-unique` の行全体重複削除
- `--columns` によるキー指定
- `--headers` によるキー指定
- `--columns` / `--headers` 同時指定エラー

### 区切り文字・文字コード

- `.tsv` 拡張子によるタブ区切り自動判定
- `--delimiter space`
- `--encoding utf8bom`
- UTF-8 BOM 出力確認
- 不正な `--delimiter`
- 不正な `--encoding`

### 上書き制御

- 既存出力先に対する `--overwrite` なしのエラー
- `--overwrite` ありでの置き換え

### ECWeaver 固有の ZIP 直操作

- `excel-extract-pictures --engine zip`
- `.xlsx` 内の画像抽出
- 連番ファイル名確認
- 画像バイト列確認
- `excel-replace-picture --engine zip`
- 全画像置換
- `--index` 指定置換
- 不正 engine
- 不正 index
- 置換用画像なし
- 出力先存在

## ECWeaver2 側へ移植するときの注意

ECWeaver2 は実装済みコマンドセットが ECWeaver と異なる。

ECWeaver2 側では、ECWeaver 固有の以下はそのまま移植しない。

```txt
excel-to-csv
excel-to-tsv
excel-list-sheets
excel-extract-pictures
excel-replace-picture
```

特に `excel-extract-pictures` / `excel-replace-picture` は ECWeaver の ZIP 直操作コマンドなので、ECWeaver2 側のテストには不要である。

一方で、ECWeaver2 側では以下を基本正常系・異常系として追加するのが望ましい。

```txt
csv-to-excel
csvs-to-excel
weave --to-excel
```

ECWeaver2 の Excel 出力系は `ExcelInteropTools` を使うため、実行環境に Microsoft Excel が必要になる可能性がある。既存実装や実行環境を確認し、Excel 起動が必要なテストを入れる場合も自動判定・自動検証できる内容にする。

目視確認は入れない。`.xlsx` の中身検証は、可能であれば次のような方法で自動化する。

- 生成された `.xlsx` が存在することを確認する。
- `.xlsx` を ZIP として開けることを確認する。
- `xl/workbook.xml` や `xl/worksheets/sheet*.xml` などの存在を確認する。
- 必要に応じて `sharedStrings.xml` などから期待文字列の存在を確認する。
- ECWeaver2 側の既存ヘルパーで Excel を読み戻せるなら、それを使って表データを検証する。

ただし、テストデータ作成や自動検証が困難な場合は、無理に目視テストにせずユーザーへ依頼する。

## ECWeaver2 側で推奨するテスト構成

ECWeaver2 側でも、まず CSV 共通コマンドのテストは ECWeaver 側とほぼ同じ内容で実装する。

共通化しやすいテスト:

- `help`
- `version`
- 未知コマンド
- 未実装コマンド
- `csv-info`
- `csv-select-columns`
- `csv-filter-rows`
- `csv-replace`
- `csv-merge`
- `csv-sort`
- `csv-unique`
- `--delimiter`
- `--encoding`
- `--overwrite`
- 入力なし、引数個数不正、オプション不正

ECWeaver2 固有として追加したいテスト:

- `csv-to-excel` の正常系
- `csv-to-excel --sheet <name>` の正常系
- `csv-to-excel` の出力先存在エラー
- `csvs-to-excel` の正常系
- `csvs-to-excel --pattern` の正常系
- `csvs-to-excel` の入力ディレクトリなしエラー
- `weave --to-excel` の CSV / TSV / SSV 混在入力
- `weave --to-excel` の出力先存在エラー
- `weave` の出力モード不足、複数出力モード指定、未対応入力などの異常系
- ECWeaver2 で未実装の ECWeaver 専用コマンドがエラーになること

Excel / プリンタに依存する以下は、基本テストに含める場合は環境依存を考慮する。

```txt
excel-to-pdf
printers
print
```

負荷テストや大規模データテストは不要。

## テストデータ例

ECWeaver 側では、テスト開始時に以下のような小さなデータを自動生成している。

`basic.csv`:

```csv
ID,Name,Price,Category
1,Apple,100,Fruit
2,Carrot,200,Vegetable
3,Banana,150,Fruit
4,Apple,200,Fruit
```

`duplicates.csv`:

```csv
ID,Name,Price,Category
1,Apple,100,Fruit
1,Apple,100,Fruit
2,Carrot,200,Vegetable
3,Banana,150,Fruit
4,Apple,200,Fruit
```

`info.csv`:

```csv
A,B,C
1,2,3,4

x,y
```

その他:

- `tab.tsv`
- `space.ssv`
- `utf8bom.csv`
- `merge/001.csv`
- `merge/002.csv`
- `merge/only.part`

ECWeaver2 側でも、この程度の小さなデータで十分である。

## ビルド確認

ECWeaver 側のビルド確認は以下で行った。

```txt
cd ECWeaver\HLTConsole
TestBuild.bat
```

ECWeaver2 側では以下を使用する。

```txt
cd ECWeaver2\HLTConsole
TestBuild.bat
```

`dotnet build` や `MSBuild` を直接使わず、対象フォルダの `AGENTS_HowToTestBuild.md` に従って `TestBuild.bat` を使う。

## 成功メッセージ

ECWeaver 側では、成功時に以下のようなバナーを出している。

```txt
============================================================
============================================================
====                                                    ====
====              ECWeaver Test0001 SUCCESS            ====
====                                                    ====
====              ALL TESTS PASSED                     ====
====                                                    ====
============================================================
============================================================
```

ECWeaver2 側ではツール名を `ECWeaver2` に置き換える。

## 文字コードと改行

本リポジトリのルールに従う。

- `.cs` は UTF-8 with BOM + CRLF
- `.md` は UTF-8 with BOM + CRLF
- `.bat` / `.txt` は CP932 + CRLF

無関係な自動整形や一括整形はしない。
