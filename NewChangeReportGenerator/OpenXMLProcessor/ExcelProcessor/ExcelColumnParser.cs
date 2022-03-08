using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NewChangeReportGenerator.OpenXMLProcessor.ExcelProcessor; 

public class ExcelColumnParser {
    private WorkbookPart _workbookPart;
    private WorksheetPart _worksheetPart;
    private SheetData _sheetData;

    public List<string> TextColumnToStringList(string columnName) {
        ExcelCellParser excelCellParser = new ExcelCellParser(_workbookPart);
        var stringList = new List<string>();

        foreach (var row in _sheetData.Elements<Row>()) {
            string currentRow = row.RowIndex;
            string cellTextValue = excelCellParser.GetTextFromCell(columnName + currentRow);
            stringList.Add(cellTextValue);
        }
        return stringList;
    }

    public List<string> UrlColumnToStringList(string columnName) {
        ExcelCellParser excelCellParser = new ExcelCellParser(_workbookPart);
        var stringList = new List<string>();

        foreach (var row in _sheetData.Elements<Row>()) {
            string currentRow = row.RowIndex;
            string cellTextValue = excelCellParser.GetUrlFromCell(columnName + currentRow);
            stringList.Add(cellTextValue);
        }
        return stringList;
    }

    public ExcelColumnParser(WorkbookPart workbookPart) {
        _workbookPart = workbookPart;
        _worksheetPart = workbookPart.WorksheetParts.First();
        _sheetData = _worksheetPart.Worksheet.Elements<SheetData>().First();
    }
}