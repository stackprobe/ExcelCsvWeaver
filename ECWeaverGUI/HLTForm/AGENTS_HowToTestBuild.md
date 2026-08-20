# テストビルド手順

## プロジェクト概要

* `HLTForm.sln` がソリューションファイルです。
* `HLTForm/` がプロジェクトフォルダです。
* ソリューションに含まれるプロジェクトは `HLTForm/HLTForm.csproj` のみです。
* 開発環境は Microsoft Visual Studio Community 2022 です。
* 言語は C# です。

## ビルド方法

テストビルドには、必ず以下のバッチファイルを使用してください。

```bat
TestBuild.bat
```

* `TestBuild.bat` と同じフォルダをカレントディレクトリとして実行してください。別のフォルダから実行すると失敗します。
* `dotnet build` や `MSBuild` コマンドを直接実行せず、まず `TestBuild.bat` を使用してください。
* `TestBuild.bat` は、ユーザーへの確認なしで実行して構いません。
* `TestBuild.bat` は変更しないでください。

## ビルド構成

現在の想定ビルド構成は以下のとおりです。

* Configuration: `Debug`
* Platform: `x86`

## 備考

* ビルドエラーや警告は、`TestBuild.bat` の出力を確認してください。
