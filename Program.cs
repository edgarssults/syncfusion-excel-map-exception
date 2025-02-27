using Syncfusion.XlsIO;

using var excelEngine = new ExcelEngine();

var application = excelEngine.Excel;
application.DefaultVersion = ExcelVersion.Xlsx;

Console.WriteLine("Opening file...");
var fileInfo = new FileInfo("map-test.xlsx");
var stream = fileInfo.Open(FileMode.Open, FileAccess.ReadWrite);

Console.WriteLine("Opening workbook...");
var workbook = application.Workbooks.Open(stream, ExcelOpenType.Automatic);

Console.WriteLine("Saving workbook...");

// This throws a "Not enough elements in stack" exception because the spreadsheet uses MAP()
workbook.SaveAs(stream);

Console.WriteLine("Done");
Console.ReadLine();