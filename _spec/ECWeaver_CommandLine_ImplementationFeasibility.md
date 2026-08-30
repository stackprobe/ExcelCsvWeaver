# ECWeaver / ECWeaver2 コマンドライン機能の実装可否メモ

## 目的

`ECWeaver_CommandLine.md` に記載されたコマンドライン機能について、`ECWeaver` と `ECWeaver2` のどちらで実装するのが自然かを整理する。

このメモでは、既存の `Tools\*.cs` を実装素材として扱う。
また、現在実装済みのコマンドについては実装先と制限を記録する。

## 判定方針

- すべての機能を `ECWeaver` と `ECWeaver2` の両方へ無理に実装する必要はない。
- どちらか一方で簡単または自然に実装できるなら、それを採用してよい。
- `ECWeaver` は `ExcelAppTools` と `ExcelTools` を持つため、Excel COM を使った読み取り・PDF・印刷・プレースホルダ処理、および `.xlsx` ZIP 内部操作に向いている。
- `ECWeaver2` は `ExcelInteropTools` と `Microsoft.Office.Interop.Excel` 参照を持つため、Excel Interop による新規ブック作成・複数シート作成・セル書き込みに向いている。
- CSV 系は両方に `CsvFileReader` / `CsvFileWriter` があるため、どちらでも実装可能と考える。
- Excel 未インストール環境で動かせるのは、CSV 系と `ExcelTools` による ZIP / Open XML 直接操作系である。

## 推奨実装先

| コマンド | 推奨実装先 | 理由 |
|---|---|---|
| `help` | 共通 | Tools 依存なし。 |
| `version` | 共通 | Tools 依存なし。 |
| `excel-to-csv` | ECWeaver | `ExcelAppTools.LoadSheets` が既に近い。 |
| `excel-to-tsv` | ECWeaver | `excel-to-csv` の delimiter 変更で実装しやすい。 |
| `csv-to-excel` | ECWeaver2 | Interop の `Workbooks.Add`、シート作成、セル書き込み、`SaveAs` が素直。 |
| `csvs-to-excel` | ECWeaver2 | 複数 CSV を複数シートへ書き出す処理は Interop が自然。 |
| `excel-to-pdf` | 共通 | 両方に PDF 出力素材がある。 |
| `weave --to-excel` | ECWeaver2 | 最終出力で `.xlsx` 新規作成が必要になるため。 |
| `weave --to-csv-dir` | ECWeaver | Excel 読み込み済みの `LoadSheets` を使いやすい。 |
| `weave --to-same-dir` | ECWeaver | Excel から CSV 群への出力と CSV コピー・変換が中心になるため。 |
| `csv-info` | 共通 | CSV Reader だけで実装可能。 |
| `csv-select-columns` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `csv-filter-rows` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `csv-replace` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `csv-merge` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `csv-sort` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `csv-unique` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `excel-list-sheets` | ECWeaver | `LoadSheets` でシート名取得が可能。 |
| `excel-info` | ECWeaver | `LoadSheets` と `ExcelTools` を組み合わせやすい。 |
| `excel-extract-pictures` | ECWeaver | `ExcelTools.CollectPicture` が最適。 |
| `excel-replace-picture` | ECWeaver | `ExcelTools.ReplacePicture` が最適。 |
| `excel-replace-text` | 共通 | `ECWeaver` は `ExcelAppTools`、`ECWeaver2` は `ExcelInteropTools` 経由で実装可能。必要に応じて `.xlsx` ZIP 内部 XML 置換も選択肢にできる。 |
| `excel-replace-placeholder` | 共通 | `ECWeaver` には `ExcelAppTools.ReplacePlaceholder` が既にあり、`ECWeaver2` も `ExcelInteropTools` 側へ同等処理を追加すれば実装可能。 |
| `csv-validate` | 共通 | CSV Reader だけで実装可能。 |
| `excel-validate` | ECWeaver | app / zip の両方式で検査しやすい。 |
| `csv-diff` | 共通 | CSV Reader/Writer だけで実装可能。 |
| `excel-diff` | ECWeaver | Excel を CSV 化して比較する流れを作りやすい。 |
| `printers` | 共通 | 両方にプリンタ一覧取得素材がある。 |
| `print` | 共通 | 両方に印刷素材がある。 |
| `run-script` | 共通 | コマンド層で実装する機能で、Tools 依存が薄い。 |

## 現在の実装状況

### ECWeaver

実装済み:

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

未実装:

```txt
csv-to-excel
csvs-to-excel
weave
excel-info
excel-replace-text
excel-replace-placeholder
csv-validate
excel-validate
csv-diff
excel-diff
run-script
```

`excel-replace-text` と `excel-replace-placeholder` は未実装だが、実装方針は `ECWeaver` / `ECWeaver2` の両対応とする。
`excel-extract-pictures` と `excel-replace-picture` は `ExcelTools` を使うため Excel は不要。
それ以外の Excel 読み込み、PDF、印刷、プリンタ一覧は `ExcelAppTools` を使う。

### ECWeaver2

実装済み:

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

未実装:

```txt
excel-to-csv
excel-to-tsv
excel-list-sheets
excel-info
excel-extract-pictures
excel-replace-picture
excel-replace-text
excel-replace-placeholder
csv-validate
excel-validate
csv-diff
excel-diff
run-script
```

`excel-replace-text` と `excel-replace-placeholder` は未実装だが、実装方針は `ECWeaver` / `ECWeaver2` の両対応とする。
`weave` は `--to-excel` のみ実装済み。
入力は `.csv`、`.tsv`、`.ssv` に限り、Excel 入力混在、`--to-csv-dir`、`--to-same-dir` は未実装。

## 結論

`ECWeaver_CommandLine.md` のコマンド群は、既存の `Tools\*.cs` を実装素材としてメソッド追加していく前提なら、明確な実装不可項目はない。

ただし、両方の実行ファイルに同じ機能を完全実装する必要はない。原則として、片方で自然に実装できる機能はその片方へ寄せる。

実装分担の大枠は次の通りとする。

- CSV 加工系は共通実装でよい。
- Excel 読み取り、CSV 化、画像抽出・置換、プレースホルダ置換、ZIP 内部操作は `ECWeaver` を主軸にする。
- `.xlsx` 新規作成、複数シート作成、CSV 群から Excel ブックを作る処理は `ECWeaver2` を主軸にする。
- PDF 出力、プリンタ一覧、印刷は両方に素材があるため、必要に応じて共通または個別に対応する。

この方針により、`--engine auto` では各機能の得意な実装へ寄せる設計が現実的になる。
