# Excel → PDF（Windows / Visual Studio 2022）

C# / Windows Forms / .NET Framework 4.8 のデスクトップアプリです。
ルートの既存ブラウザ版とは別に使用できます。

## 必要な環境
- Windows 10 / 11
- Visual Studio 2022：「.NET デスクトップ開発」ワークロードと .NET Framework 4.8 ターゲットパック
- 実行環境：.NET Framework 4.8、デスクトップ版 Microsoft Excel（COM登録済み）

NuGetパッケージやOffice参照の追加は不要です。ExcelをインストールしていないPCでは変換できません。

## 起動
1. `desktop/ExcelPdfConverter.sln` をVS2022で開きます。
2. 構成をReleaseにしてビルド、またはF5で実行します。
3. 「Excelを追加」またはドラッグ＆ドロップでブックを追加します。
4. 保存先を選び、「PDFに変換」を押します。

ビルド後の実行ファイル：`ExcelPdfConverter/bin/Release/ExcelPdfConverter.exe`

```powershell
MSBuild.exe desktop/ExcelPdfConverter.sln /p:Configuration=Release
```

## 動作
- xlsx / xls / xlsm / xlsb を複数まとめて処理します。
- ブック全体を1つのPDFへ出力します。ページサイズ・向き・改ページなどはExcelブックの印刷設定に従います。
- 印刷範囲を無視するオプションがあります。非表示シートなどの扱いはExcel標準のPDF出力に従います。
- PDFが存在する場合は `名前 (1).pdf` のように別名で保存します。既存PDFを上書きしません。
- 元ブックは読み取り専用で開き、保存せず閉じます。マクロ・イベントと外部リンクの更新を無効にします。
- 変換は専用STAスレッドで実行し、各ファイルの成否を表示します。
- 中止は現在のファイルの処理後に適用されます。Excel内部の処理中は強制中止しません。
- パスワード付きブックには対応しません。必要ならExcelで解除してから使用してください。
- ネットワークへのファイルアップロードは行いません。信頼できるブックを使用してください。

## 検証
VS2022 MSBuildによるReleaseビルドが成功しました。出力ファイル名・対応拡張子の6項目のテストも成功しました。
検証環境にExcelが登録されていないため、実際のPDF変換は未検証です。

```powershell
MSBuild.exe desktop/tests/OutputPathsTests.csproj /p:Configuration=Release
desktop/tests/bin/Release/OutputPathsTests.exe
```
実際のPDF変換確認にはExcelが必要です。Excel未導入環境でのビルド成功だけではPDF変換の成功は保証できません。

手動確認：複数シートと印刷範囲のあるブックを変換し、PDFのページ・表示をExcelの印刷プレビューと比較します。
同名ブック2つの変換、既存PDFがある保存先、パスワード付きブック、途中の失敗後の継続、中止も確認してください。

Excel公式API：
https://learn.microsoft.com/office/vba/api/excel.workbook.exportasfixedformat
