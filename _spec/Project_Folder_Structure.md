# ExcelCsvWeaver フォルダ構成

## 目的

この文書は、本プロジェクトのトップレベルフォルダ構成と、それぞれの役割をまとめる。

本プロジェクトについて不明な点がある場合は、まず `_spec` 配下の仕様書を読む。

## トップレベル構成

```txt
ExcelCsvWeaver/
  _spec/
  ECWeaver/
  ECWeaver2/
  ECWeaverGUI/
  Installer/
  MakeDistribution.bat
```

## _spec

本プロジェクトの仕様書をまとめるフォルダ。

コマンドライン仕様、機能方針、実装可否、フォルダ構成など、プロジェクト全体を理解するための情報をここに置く。
本プロジェクトについて不明な点は、ここを読めば分かる状態にする。

## ECWeaver

CUI ツール。

Excel / CSV 関連処理をコマンドラインから実行するためのプログラムである。
現状では、Excel アプリケーション操作系や、`.xlsx` を ZIP / Open XML として直接扱う処理を主に担当する。

代表的な処理:

- Excel から CSV / TSV への変換
- Excel から PDF への変換
- CSV 加工
- Excel シート一覧取得
- Excel ブック内画像の抽出・置換
- プリンタ一覧取得・印刷

## ECWeaver2

CUI ツール。

Excel / CSV 関連処理をコマンドラインから実行するためのプログラムである。
現状では、Microsoft Office Interop を使った Excel ブック作成系の処理を主に担当する。

代表的な処理:

- CSV から Excel への変換
- 複数 CSV から複数シート Excel への変換
- CSV / TSV / SSV 群を 1 つの Excel ブックへ統合する `weave --to-excel`
- Excel から PDF への変換
- CSV 加工
- プリンタ一覧取得・印刷

## ECWeaverGUI

GUI ラッパーアプリケーション。

内部的には `ECWeaver` / `ECWeaver2` などの CUI ツールを呼び出す。
GUI を使うユーザーから見ると、このアプリケーションがメインプログラムになる。

GUI 側は、ユーザー入力を受け取り、対応する CUI コマンドライン呼び出しを組み立てて実行する役割を持つ。

## Installer

インストーラプログラム。

インストール処理を担当する。現状では完成済みとみなし、通常の機能追加や仕様変更では基本的に触らない想定である。

## MakeDistribution.bat

リリースパッケージを作成するためのバッチファイル。

このバッチはユーザーが実行する。Codex は実行しない。

## 関連仕様

- `ECWeaver_CommandLine.md`
  - `ECWeaver` / `ECWeaver2` のコマンドライン仕様。
- `ECWeaver_CommandTool_Features.md`
  - コマンドツールとして提供する機能の方針。
- `ECWeaver_CommandLine_ImplementationFeasibility.md`
  - 各コマンドを `ECWeaver` / `ECWeaver2` のどちらで実装するのが自然かを整理したメモ。
