using System.CommandLine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelFindReplace;

/// <summary>
/// This class is the cli interface.
/// </summary>
public partial class ExcelFindReplace
{
  static async Task<int> Main(string[] args)
  {
    var fileOption = new Argument<FileInfo>(
      name: "file",
      description: "The file path to use for the find and replace.");

    var sheetNameOption = new Option<string>(
      name: "--sheet",
      description: "The sheet to look in the spreadsheet.");

    var findOption = new Option<string>(
      name: "--find",
      description: "The value to look for in the spreadsheet.");

    var replaceOption = new Option<string>(
      name: "--replace",
      description: "The value to replace with the cell with in the spreadsheet.");

    var rowOffsetOption = new Option<int>(
      name: "--row_offset",
      description: "The value to replace with the cell with in the spreadsheet.",
      getDefaultValue: () => 0);

    var colOffsetOption = new Option<int>(
      name: "--column_offset",
      description: "The value to replace with the cell with in the spreadsheet.",
      getDefaultValue: () => 0);

    var rootCommand = new RootCommand("Excel xlsx find find and replace.");

    var replaceCommand = new Command("replace", "find and replace string"){
      fileOption,
      sheetNameOption,
      findOption,
      replaceOption,
      rowOffsetOption,
      colOffsetOption
    };

    replaceCommand.SetHandler((filePath, sheetName, findText, replaceText, rowOffset, columnOffset) =>
        {
          ReplaceTextInExcel(filePath!, sheetName, findText, replaceText, rowOffset, columnOffset);
        },
        fileOption, sheetNameOption, findOption, replaceOption, rowOffsetOption, colOffsetOption);

    rootCommand.Add(replaceCommand);

    return await rootCommand.InvokeAsync(args);
  }

  public static string? ReplaceTextInExcel(FileInfo filePath, string sheetName, string findText, string replaceText, int rowOffset = 0, int columnOffset = 0)
  {
    string? value = null;

    // Open the spreadsheet document for read-only access.
    using SpreadsheetDocument document = SpreadsheetDocument.Open(filePath.FullName, true);

    // Retrieve a reference to the workbook part.
    WorkbookPart? wbPart = document.WorkbookPart;
    // Find the sheet with the supplied name, and then use that
    // Sheet object to retrieve a reference to the first worksheet.
    Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

    // Throw an exception if there is no sheet.
    if (theSheet is null || theSheet.Id is null)
    {
      throw new ArgumentException($"A sheet with the name {sheetName} could not be found in file: {filePath.FullName}");
    }
    WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);

    Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellValue != null && GetCellValue(wbPart, c) == findText).FirstOrDefault();

    if (theCell == null)
    {
      Console.WriteLine($"Did not find value: {findText} in sheet {sheetName} in file: {filePath.FullName}");
      var allCells = wsPart?.Worksheet?.Descendants<Cell>()
                      .Where(c => c.CellValue != null);
      if (allCells != null)
      {
        foreach (var cell in allCells)
        {
          Console.WriteLine($"Cell: {cell.CellReference} contains: {GetCellValue(wbPart, cell)}");
        }
      }
    }
    else
    {
      Console.WriteLine($"Found value: {findText} in cell: {theCell?.CellReference}");
      // Get the cell reference (e.g., "A1")
      string? cellReference = theCell?.CellReference?.Value;
      // Determine the column and row of the cell
      string column = new string(cellReference?.Where(char.IsLetter).ToArray());
      string row = new string(cellReference?.Where(char.IsDigit).ToArray());

      // Get the next column (e.g., "B" if the current column is "A")
      string nextColumn = ((char)(column[0] + columnOffset)).ToString();
      int nextRow = int.Parse(row) + rowOffset;

      // Construct the reference for the adjacent cell (e.g., "B1")
      string adjacentCellReference = nextColumn + nextRow;

      // Find the adjacent cell
      Cell? adjacentCell = wsPart?.Worksheet?.Descendants<Cell>()
                          .FirstOrDefault(c => c.CellReference?.Value == adjacentCellReference);

      if (adjacentCell != null && adjacentCell.CellValue != null)
      {
        Console.WriteLine($"Offset cell {adjacentCellReference} contains: {GetCellValue(wbPart, adjacentCell)}");
        SetCellValue(wbPart, adjacentCell, replaceText);
        value = GetCellValue(wbPart, adjacentCell);
        Console.WriteLine($"Offset cell {adjacentCellReference} replaced with: {value}");
      }
      else
      {
        Console.WriteLine($"Adjacent cell {adjacentCellReference} is empty or not found.");
      }
    }
    return value;
  }
  private static string GetCellValue(WorkbookPart wbPart, Cell cell)
  {
    string? value = null;
    // If the cell does not exist, return an empty string.
    if (cell is null || cell.InnerText.Length < 0)
    {
      return string.Empty;
    }
    value = cell.InnerText;
    // If the cell represents an integer number, you are done.
    // For dates, this code returns the serialized value that
    // represents the date. The code handles strings and
    // Booleans individually. For shared strings, the code
    // looks up the corresponding value in the shared string
    // table. For Booleans, the code converts the value into
    // the words TRUE or FALSE.
    if (cell.DataType is not null)
    {
      if (cell.DataType.Value == CellValues.SharedString)
      {
        // For shared strings, look up the value in the
        // shared strings table.
        var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        // If the shared string table is missing, something
        // is wrong. Return the index that is in
        // the cell. Otherwise, look up the correct text in
        // the table.
        if (stringTable is not null)
        {
          value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
        }
      }
      else if (cell.DataType.Value == CellValues.Boolean)
      {
        switch (value)
        {
          case "0":
            value = "FALSE";
            break;
          default:
            value = "TRUE";
            break;
        }
      }
    }

    return value;
  }

  private static void SetCellValue(WorkbookPart wbPart, Cell cell, string value)
  {
    // If the cell represents an integer number, you are done.
    // For dates, this code returns the serialized value that
    // represents the date. The code handles strings and
    // Booleans individually. For shared strings, the code
    // looks up the corresponding value in the shared string
    // table. For Booleans, the code converts the value into
    // the words TRUE or FALSE.
    if (cell.DataType is not null)
    {
      if (cell.DataType.Value == CellValues.SharedString)
      {
        // For shared strings, look up the value in the
        // shared strings table.
        var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        // If the shared string table is missing, something
        // is wrong. Return the index that is in
        // the cell. Otherwise, look up the correct text in
        // the table.
        if (stringTable is not null)
        {
          OpenXmlElement sharedString = stringTable.SharedStringTable.ElementAt(int.Parse(cell.InnerText));
          sharedString.ReplaceChild(new Text(value), sharedString.FirstChild);
          stringTable.SharedStringTable.Save();
        }
      }
      else if (cell.DataType.Value == CellValues.Boolean)
      {
        throw new NotImplementedException();
        // switch (value)
        // {
        //   case "TRUE":
        //     value = 1;
        //     break;
        //   default:
        //     value = 0;
        //     break;
        // }
      }
      else if (cell.DataType.Value == CellValues.Number)
      {
        throw new NotImplementedException();
      }
    }

  }
  // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
  // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
  static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
  {
    // If the part does not contain a SharedStringTable, create one.
    if (shareStringPart.SharedStringTable is null)
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
}