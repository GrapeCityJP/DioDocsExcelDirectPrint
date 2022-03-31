// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("Excelファイルをプリンタダイアログ表示なしで直接印刷します");

Workbook.SetLicenseKey("");

// Excelファイルを読み込み
var workbook = new Workbook();
workbook.Open(@"InvoiceJP.xlsx");

// 印刷オプションを作成
PrintOutOptions options = new PrintOutOptions();

// 印刷するプリンターを設定
options.ActivePrinter = "Microsoft Print to PDF";

// 3部印刷 
options.Copies = 3;

// このワークブックを「Microsoft Print to PDF」で印刷
workbook.PrintOut(options);
