# Overview

Contains a .NET console application that can be used to reproduce the exception that is thrown when trying to save an Excel spreadsheet that contains a MAP() function with the Syncfusion Excel library.

# Steps

Open the solution and run the app.

# Exception

````
System.ArgumentException
  HResult=0x80070057
  Message=Not enough elements in stack
  Source=Syncfusion.XlsIO.Portable
  StackTrace:
   at Syncfusion.XlsIO.Parser.Biff_Records.Formula.FunctionVarPtg.PushResultToStack(FormulaUtil formulaUtil, Stack`1 operands, Boolean isForSerialization)
   at Syncfusion.XlsIO.Implementation.FormulaUtil.ParsePtgArray(Ptg[] ptgs, Int32 row, Int32 col, Boolean bR1C1, NumberFormatInfo numberInfo, Boolean bRemoveSheetNames, Boolean isForSerialization, IWorksheet sheet)
   at Syncfusion.XlsIO.Implementation.FormulaUtil.ParsePtgArray(Ptg[] ptgs, Int32 row, Int32 col, Boolean bR1C1, NumberFormatInfo numberInfo, Boolean isForSerialization)
   at Syncfusion.XlsIO.Implementation.FormulaUtil.ParsePtgArray(Ptg[] ptgs, Int32 row, Int32 col, Boolean bR1C1, Boolean isForSerialization)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.Excel2007Serializator.SerializeArrayFormula(XmlWriter writer, ArrayRecord arrayRecord, String cellName)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.Excel2007Serializator.SerializeCell(XmlWriter writer, BiffRecordRaw record, RowStorageEnumerator rowStorageEnumerator, CellRecordCollection cells, Dictionary`2 hashNewParentIndexes, String cellTag)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.Excel2007Serializator.SerializeCells(XmlWriter writer, RowStorage row, CellRecordCollection cells, Dictionary`2 hashNewParentIndexes, String cellTag)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.Excel2007Serializator.SerializeRow(XmlWriter writer, RowStorage row, CellRecordCollection cells, Int32 iRowIndex, Dictionary`2 hashNewParentIndexes, String cellTag, Boolean isSpansNeeded)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.Excel2007Serializator.SerializeSheetData(XmlWriter writer, CellRecordCollection cells, Dictionary`2 hashNewParentIndexes, String cellTag, Dictionary`2 additionalAttributes, Boolean isSpansNeeded)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.Excel2007Serializator.SerializeWorksheet(XmlWriter writer, WorksheetImpl sheet, Stream streamStart, Stream streamConFormats, Dictionary`2 hashXFIndexes, Stream streamExtCondFormats)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.WorksheetDataHolder.SerializeWorksheetPart(WorksheetImpl sheet, Dictionary`2 hashNewXFIndexes)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.WorksheetDataHolder.SerializeWorksheet(WorksheetImpl sheet, Dictionary`2 hashNewXFIndexes, Dictionary`2 cacheFiles)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveWorksheet(WorksheetImpl sheet, String itemName, Dictionary`2 hashNewXFIndexes, Dictionary`2 cacheFiles)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveSheet(WorksheetBaseImpl sheet, String itemName, RelationCollection relations, String workbookPath, Dictionary`2 hashNewXFIndexes, Dictionary`2 cacheFiles)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveSheets(RelationCollection relations, String workbookItemName, Dictionary`2 hashNewXFIndexes, Dictionary`2 cacheFiles)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveWorkbookPart(Dictionary`2 hashNewXFIndexes, Dictionary`2 cacheFiles)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveWorkbook(ExcelSaveType saveAsType)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveDocument(ExcelSaveType saveType)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.SaveDocument(Stream stream, ExcelSaveType saveType)
   at Syncfusion.XlsIO.Implementation.XmlSerialization.FileDataHolder.Serialize(Stream stream, WorkbookImpl book, ExcelSaveType saveType)
   at Syncfusion.XlsIO.Implementation.WorkbookImpl.SaveAsInternal(Stream stream, ExcelSaveType saveType)
   at Syncfusion.XlsIO.Implementation.WorkbookImpl.SaveAs(Stream stream, ExcelSaveType saveType)
   at Syncfusion.XlsIO.Implementation.WorkbookImpl.SaveAs(Stream stream)
   at Program.<Main>$(String[] args) in D:\Experiments\syncfusion-excel-map-exception\Program.cs:line 18
````