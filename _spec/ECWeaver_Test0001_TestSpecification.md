# ECWeaver / ECWeaver2 Test0001 テスト実装仕様

## 目的

この文書は、`ECWeaver` と `ECWeaver2` に実装済みの `Tests.Test0001.Test01()` のテスト仕様をまとめる。

両者のテストは、共通 CSV コマンドの基本的な正常系・異常系を同じ考え方で検証し、各ツール固有の Excel 関連コマンドだけを個別に検証する。

負荷テスト、大規模データテスト、目視確認が必要なテストは対象外とする。

## 対象ファイル

テスト本体:

```txt
ECWeaver/HLTConsole/HLTConsole/Tests/Test0001.cs
ECWeaver2/HLTConsole/HLTConsole/Tests/Test0001.cs
```

ZIP / XLSX 検証のための参照追加:

```txt
ECWeaver/HLTConsole/HLTConsole/HLTConsole.csproj
ECWeaver2/HLTConsole/HLTConsole/HLTConsole.csproj
```

`Test0001.cs` は `#if DEBUG` 内に実装し、Release 版では空のクラスとして扱われる。

## 起動仕様

Debug ビルドで引数なし実行した場合、`Program.Main3()` から次を呼び出す。

```csharp
new Test0001().Test01();
```

テストフレームワークは導入しない。`Test01()` が目的別の private メソッドを順番に呼び出す。

`Test01()` はテスト失敗やその他のエラーで発生した例外を内部で捕捉し、呼び出し側へ例外を投げない。失敗時はコンソールに失敗バナーと例外詳細を出力し、コンソール出力から原因を調査できるようにする。

成功時はコンソールに成功バナーを表示する。

## ディレクトリ仕様

テストデータは永続領域に置く。

```txt
ECWeaver:
C:\home\res\ExcelCsvWeaver\ECWeaver\Test0001

ECWeaver2:
C:\home\res\ExcelCsvWeaver\ECWeaver2\Test0001
```

作業ディレクトリは毎回削除して作り直す。

```txt
ECWeaver:
C:\temp\ECWeaver

ECWeaver2:
C:\temp\ECWeaver2
```

`Test01()` の先頭では、以下を実行する。

```csharp
SCommon.DeleteAndCreateDir(BaseDir);
```

`C:\home\res` 配下のテストデータは永続管理するため、現在の実装では `PrepareTestData()` 呼び出しをコメントアウトしている。テストデータを再生成したい場合は、一時的にコメントアウトを解除して実行する。

`SCommon.Pause_WaitSeconds = 0;` もコメントアウトしている。F5 実行時にコンソールがすぐ閉じると結果を確認しづらいためである。

## 共通実装方針

- `Run(string[] args)` で `ECWeaverArgs.Read(new ArgsReader(args))` と `ECWeaverProcessor.Run(...)` を直接呼び出す。
- 標準出力検証が必要なコマンドは `CaptureOutput(string[] args)` で `Console.Out` を差し替えて捕捉する。
- CSV 出力は `CsvFileReader.ReadToEnd(...)` で読み戻し、期待する行列と完全一致で比較する。
- XLSX は ZIP として開き、必要なエントリ、シート数、XML 内の期待文字列を検証する。
- 例外系は `AssertThrows(...)` で例外メッセージの一部一致を確認する。
- 目的ごとに private メソッドを分け、各メソッドに XML コメントでテスト目的を書く。
- 外部テストフレームワークは追加しない。
- テスト失敗時やその他のエラー発生時は、`Test01()` 内で例外を捕捉し、失敗バナーと例外詳細をコンソールに出力する。
- 成功時のみ成功バナーを出力する。

## 共通テストデータ

両テストで使う基本データは、小さく固定的な CSV / TSV / SSV とする。

代表例:

```csv
ID,Name,Price,Category
1,Apple,100,Fruit
2,Carrot,200,Vegetable
3,Banana,150,Fruit
4,Apple,200,Fruit
```

その他に以下を使用する。

- `duplicates.csv`: 重複削除テスト用
- `info.csv`: 行数、列数、空行数テスト用
- `tab.tsv`: タブ区切り自動判定用
- `space.ssv`: スペース区切り指定用
- `utf8bom.csv`: UTF-8 BOM 入出力用
- `merge/001.csv`, `merge/002.csv`: ファイル名順結合用
- `merge/only.part`: `--pattern` 指定用

ECWeaver 固有テストでは、`.xlsx` を ZIP として作成し、`xl/media/image1.png` などの画像エントリを持つ最小ワークブックを使用する。

ECWeaver2 固有テストでは、`excel-sheets/alpha.csv`, `excel-sheets/beta.csv`, `excel-sheets/only.part` と `unsupported.txt` を使用する。

## 共通テスト項目

### ヘルプ・バージョン・コマンドエラー

- `help` が全体ヘルプを出力すること。
- `help <command>` がコマンド別ヘルプを出力すること。
- `version` がツール名付きのバージョンを出力すること。
- 未知コマンドが `Unknown command` を含む例外になること。
- 各ツールで未実装のコマンドが未実装エラーになること。

### csv-info

- 行数、最小列数、最大列数、空行数を出力すること。
- `--silent` 指定時に標準出力が空になること。
- 入力ファイルなしをエラーにすること。
- 引数個数不正をエラーにすること。

### csv-select-columns

- `--columns` で列番号指定の列抽出ができること。
- `--headers` でヘッダー名指定の列抽出ができること。
- `--columns` / `--headers` 不足をエラーにすること。
- `--columns` / `--headers` 同時指定をエラーにすること。
- 列番号 `0`、範囲外列、存在しないヘッダーをエラーにすること。

### csv-filter-rows

- `--equals` で一致行を抽出できること。
- `--contains` で部分一致行を抽出できること。
- `--regex` で正規表現一致行を抽出できること。
- `--invert` で条件を反転できること。
- `--has-header` または `--header` 指定時にヘッダー行を維持できること。
- 列指定不足、条件不足、条件の複数指定をエラーにすること。
- CSV 系コマンドへの `--engine` 指定をエラーにすること。

### csv-replace

- `--from --to` で全セル置換できること。
- `--column` 指定で対象列だけ置換できること。
- `--header` と `--regex` でヘッダー名指定列の正規表現置換ができること。
- `--to` 不足、置換元指定不足、`--from` / `--regex` 同時指定をエラーにすること。

### csv-merge

- ディレクトリ内 CSV をファイル名順に結合できること。
- `--skip-header` で 2 ファイル目以降のヘッダーを除外できること。
- `--pattern` で対象ファイルを絞り込めること。
- 入力ディレクトリなしをエラーにすること。

### csv-sort

- 文字列昇順でソートできること。
- `--numeric` で数値昇順にソートできること。
- `--desc` で降順にソートできること。
- `--has-header` / `--header` 指定時にヘッダー行を維持できること。

### csv-unique

- 行全体の重複を削除できること。
- `--columns` でキー列を指定できること。
- `--headers` でキー列を指定できること。
- `--columns` / `--headers` 同時指定をエラーにすること。

### 区切り文字・文字コード

- `.tsv` 拡張子からタブ区切りを自動判定できること。
- `--delimiter space` でスペース区切りを読み書きできること。
- `--encoding utf8bom` で UTF-8 BOM 付きファイルを出力できること。
- UTF-8 BOM のバイト列を確認すること。
- 不正な `--delimiter` と不正な `--encoding` をエラーにすること。

### 上書き制御

- 出力先が存在する場合、`--overwrite` なしでは失敗すること。
- `--overwrite` ありでは既存出力先を置き換えられること。

## ECWeaver 固有テスト項目

ECWeaver は、Excel 画像操作系の ZIP 直操作コマンドを検証する。

### excel-extract-pictures

- `excel-extract-pictures --engine zip` で `.xlsx` 内の画像だけを抽出できること。
- 抽出ファイルが `0001.png`, `0002.jpg` のような連番名になること。
- 抽出画像のバイト列が期待値と一致すること。
- 画像以外の ZIP エントリを抽出しないこと。

### excel-replace-picture

- `excel-replace-picture --engine zip` で全画像を置換できること。
- `--index` 指定で対象画像だけ置換できること。
- 不正 engine、不正 index、置換用画像なし、出力先存在をエラーにすること。

ECWeaver では `csv-to-excel` を未実装コマンドとして扱い、未実装エラーを確認する。

## ECWeaver2 固有テスト項目

ECWeaver2 は、CSV から Excel へ出力するコマンドを検証する。

### csv-to-excel

- CSV から XLSX を生成できること。
- `--sheet` で指定したシート名が `xl/workbook.xml` に反映されること。
- 生成 XLSX を ZIP として開けること。
- workbook、worksheet、期待文字列を確認できること。

### csvs-to-excel

- ディレクトリ内 CSV をファイル名順の複数シートとして XLSX に出力できること。
- `--pattern` で対象ファイルを絞り込めること。
- シート名とセル文字列を ZIP 内 XML から確認できること。

### weave --to-excel

- CSV / TSV / SSV 混在入力を複数シートの XLSX に出力できること。
- 入力ファイル名に基づくシート名を確認できること。
- 各入力ファイル由来の文字列を ZIP 内 XML から確認できること。

### Excel 出力系の異常系

- `csv-to-excel --engine zip` を不正 engine としてエラーにすること。
- `csv-to-excel` の出力先存在をエラーにすること。
- `csvs-to-excel` の入力ディレクトリなしをエラーにすること。
- `csvs-to-excel --pattern` で対象なしの場合をエラーにすること。
- `weave` の出力モード不足をエラーにすること。
- `weave` の複数出力モード指定をエラーにすること。
- `weave` の入力ファイルなしをエラーにすること。
- `weave` の未対応拡張子をエラーにすること。
- `weave --to-excel` の出力先存在をエラーにすること。
- ECWeaver2 で未実装の `weave --to-same-dir` を未実装エラーにすること。

ECWeaver2 では `excel-to-csv` と `excel-extract-pictures` を未実装コマンドとして扱い、未実装エラーを確認する。

## 実行環境上の注意

`C:\temp` と `C:\home\res` 配下への書き込み権限が必要である。

ECWeaver2 の Excel 出力系は既存実装が Microsoft Excel Interop を使用するため、実行環境に Microsoft Excel が必要になる可能性がある。ただし、テスト検証自体は目視ではなく、生成された `.xlsx` を ZIP として開く自動検証で行う。

## ビルド確認

ビルド確認は各プロジェクト配下の `AGENTS_HowToTestBuild.md` に従い、`TestBuild.bat` を使用する。

```txt
cd ECWeaver\HLTConsole
TestBuild.bat

cd ECWeaver2\HLTConsole
TestBuild.bat
```

`dotnet build` や `MSBuild` を直接使うのではなく、既存のバッチを使う。

## 成功条件

すべてのテストメソッドが例外を投げずに完了し、最後に成功バナーが出力されること。

成功バナーのツール名はそれぞれ以下とする。

```txt
ECWeaver Test0001 SUCCESS
ECWeaver2 Test0001 SUCCESS
```

## 文字コードと改行

本リポジトリのルールに従う。

- `.cs` は UTF-8 with BOM + CRLF
- `.md` は UTF-8 with BOM + CRLF
- `.txt` / `.bat` は CP932 + CRLF

無関係な自動整形や一括整形は行わない。
