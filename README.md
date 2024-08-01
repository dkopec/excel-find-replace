# excel_find_replace

A simple cli program to find and replace values in excel xlsx files.

Download from the Github Releases, run binary and pass in the appropriate arguments and options as seen below:

```text
Description:
  Excel xlsx find find and replace.

Usage:
  ExcelFindReplace [command] [options]

Options:
  --version       Show version information
  -?, -h, --help  Show help and usage information

Commands:
  replace <file>  find and replace string

Description:
  find and replace string

Usage:
  ExcelFindReplace replace <file> [options]

Arguments:
  <file>  The file path to use for the find and replace.

Options:
  --sheet <sheet>                  The sheet to look in the spreadsheet.
  --find <find>                    The value to look for in the spreadsheet.
  --replace <replace>              The value to replace with the cell with in the spreadsheet.
  --row_offset <row_offset>        The value to replace with the cell with in the spreadsheet. [default: 0]
  --column_offset <column_offset>  The value to replace with the cell with in the spreadsheet. [default: 0]
  -?, -h, --help                   Show help and usage information
```
