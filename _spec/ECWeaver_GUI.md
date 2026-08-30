# ECWeaverGUI 仕様メモ

## 目的

この文書は、`ECWeaverGUI` の現在の位置づけと、将来 GUI を実装する場合の基本方針をまとめる。

現時点では、GUI の具体的な画面仕様、操作フロー、実装計画は未定である。

## 現在の位置づけ

`ECWeaverGUI` は、将来 GUI が必要になった場合に備えて残しているプレースホルダーである。

現時点では、Excel / CSV 関連処理の主な利用手段は以下の CUI ツールとする。

```txt
ECWeaver.exe
ECWeaver2.exe
```

GUI から操作する専用ワークフローは、現在未計画・未実装である。

## 現在 GUI で提供しないもの

現時点では、以下の GUI 機能は提供しない。

- Excel から CSV / TSV への変換画面
- CSV から Excel への変換画面
- Excel から PDF への変換画面
- CSV 加工画面
- Excel ブック内画像の抽出・置換画面
- 印刷画面
- 実行履歴画面
- 設定画面

これらの機能は、必要になった時点で改めて仕様を決める。

## 当面の案内方針

GUI プログラムを配布物に含める場合、GUI 側では以下の内容をユーザーに案内する。

- GUI は将来開発用のプレースホルダーであること。
- 現時点では専用 GUI ワークフローが未計画・未実装であること。
- 実際の処理には `ECWeaver.exe` と `ECWeaver2.exe` を使用すること。
- 利用可能なコマンドは `help` で確認できること。

案内文の例:

```txt
ExcelCsvWeaver GUI is currently reserved for future development.

At this time, no dedicated GUI workflow has been planned or implemented.
Please use ECWeaver.exe or ECWeaver2.exe from the command line.
```

## 将来 GUI を実装する場合の方針

将来 GUI を実装する場合は、`ECWeaver_CommandLine.md` の「GUI 連携方針」に従う。

基本方針:

- GUI は、単にコマンドライン文字列を組み立てるだけの薄いラッパーにしない。
- CLI と GUI で同じ処理定義を共有できるようにする。
- GUI の画面入力は `CommandOptions` 相当の構造へ変換する。
- 実処理は `CommandRunner` 相当の処理層で実行する。
- 処理結果は `CommandResult` 相当の構造で受け取り、GUI 側で表示する。
- エラー、警告、ログ、実行結果の表示は GUI 側の責務とする。

将来の画面構成を検討する場合は、次のような単位で整理する。

- 入力ファイルまたは入力フォルダの選択
- 出力ファイルまたは出力フォルダの選択
- 処理種類の選択
- 処理オプションの設定
- 実行前チェック
- 実行結果とログの表示

## 実装優先度

現時点では GUI 実装の優先度は低い。

コマンドライン機能の整備、仕様確定、処理層の共通化を優先する。
GUI は、ユーザー操作で明確な需要が出た段階で改めて計画する。

## 関連文書

- `Project_Folder_Structure.md`
- `ECWeaver_CommandLine.md`
- `ECWeaver_CommandTool_Features.md`
