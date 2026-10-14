# Excel PDF Converter

## Windows版（Visual Studio 2022）

`desktop/ExcelPdfConverter.sln` をVisual Studio 2022で開いてください。
C# / Windows Forms / .NET Framework 4.8で、複数のExcelブックをPDFに変換できます。
**変換にはデスクトップ版Microsoft Excelが必要です。**

必要な環境、ビルドと使用方法は [desktop/README.md](desktop/README.md) にあります。

## ブラウザ版（Sheetform）

Excel / CSV ファイルをブラウザ内で読み込み、用紙サイズ・向き・余白などを調整して提出用PDFとして保存するシングルページアプリです。

```bash
python3 -m http.server 4173
```

`http://localhost:4173` を開いてください。ファイル処理はブラウザ内で完結し、PDF出力にはブラウザの印刷機能を利用します。
