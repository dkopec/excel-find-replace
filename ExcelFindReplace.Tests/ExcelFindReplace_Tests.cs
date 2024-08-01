using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace ExcelFindReplace.Tests
{
  public class ExcelFindReplaceTests : IDisposable
  {
    private readonly string _testFilePath;
    private readonly string _testFindValue;
    private readonly string _testCellReference;
    private readonly string _testSheetName;

    public ExcelFindReplaceTests()
    {
      // Create a temporary Excel file for testing
      _testFilePath = Path.Combine(Path.GetTempPath(), "test.xlsx");
      _testFindValue = "findValue";
      _testCellReference = "A1";
      _testSheetName = "Sheet1";
      CreateTestExcelFile(_testFilePath);
    }

    [Fact]
    public void Given_Spreadsheet_With_Value_Replace_Value()
    {
      //Arrange
      var filePath = new FileInfo(_testFilePath);
      var replaceText = "replaceText";

      //Act
      var result = ExcelFindReplace.ReplaceTextInExcel(filePath, _testSheetName, _testFindValue, replaceText);

      //Assert
      Assert.NotNull(result);
      Assert.Equivalent(replaceText, result);
    }

    private void CreateTestExcelFile(string filePath)
    {
      using var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
      var workbookPart = document.AddWorkbookPart();
      workbookPart.Workbook = new Workbook();
      var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
      worksheetPart.Worksheet = new Worksheet(new SheetData());

      var sheets = document.WorkbookPart?.Workbook.AppendChild(new Sheets());
      var sheet = new Sheet() { Id = document.WorkbookPart?.GetIdOfPart(worksheetPart), SheetId = 1, Name = _testSheetName };
      sheets?.Append(sheet);

      var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
      var row = new Row();
      var cellValue = InsertSharedStringItem(_testFindValue, workbookPart);
      var cell = new Cell() { CellReference = _testCellReference, DataType = CellValues.SharedString, CellValue = new CellValue(cellValue) };
      row.Append(cell);
      sheetData?.Append(row);
    }

    private static int InsertSharedStringItem(string text, WorkbookPart workbookPart)
    {
      SharedStringTablePart? shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
      // If the part does not contain a SharedStringTable, create one.
      if (shareStringPart?.SharedStringTable is null)
      {
        shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        shareStringPart.SharedStringTable = new SharedStringTable();
      }
      else if (shareStringPart.SharedStringTable is null)
      {
        shareStringPart.SharedStringTable = new SharedStringTable();
      }

      int i = 0;

      // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
      foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
      {
        if (item.InnerText == text)
        {
          return i;
        }

        i++;
      }

      // The text does not exist in the part. Create the SharedStringItem and return its index.
      shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
      shareStringPart.SharedStringTable.Save();

      return i;
    }

    public void Dispose()
    {
      if (File.Exists(_testFilePath))
      {
        File.Delete(_testFilePath);
      }
    }
  }
}
