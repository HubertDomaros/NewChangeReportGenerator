using System;
using System.Collections;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NewChangeReportGenerator.OpenXMLProcessor.ExcelProcessor; 

internal class ExcelCellParser {
    private readonly WorkbookPart _workbookPart;
    private readonly WorksheetPart _worksheetPart;

    public string GetTextFromCell(string cellCoordinates) {
        var returnedValue = "";

        var cellValuesList = _worksheetPart.Worksheet.Descendants<CellValue>().ToList();
        var cell =
            _worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellCoordinates) ??
            throw new NullReferenceException(
                "NullReferenceException cell object in ExcelCellParser class (possible null pointer exception)");

        var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
        var row = sheetData.Elements<Row>();


        var cellId = cell.InnerText;
        //For SharedStringTable, see explanation on StackOverflow
        //https://stackoverflow.com/questions/5115257/openxml-sdk-returning-a-number-for-cellvalue-instead-of-cells-text

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
            var stringTable = _workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            returnedValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellId)).InnerText;
        }

        return returnedValue;
    }

    public string GetUrlFromCell(string cellCoordinates) {
        var returnedValue = "";

        var cell = _worksheetPart.RootElement.Descendants<Cell>()
            .FirstOrDefault(c => c.CellReference == cellCoordinates) ?? throw new InvalidOperationException();

        var hyperlinkEnumerable = _worksheetPart.RootElement.Descendants<Hyperlinks>().First().Cast<Hyperlink>();
        var hyperlink = hyperlinkEnumerable.SingleOrDefault(i => i.Reference.Value == cell.CellReference.Value);
        var hyperlinksRelation = _worksheetPart.HyperlinkRelationships.SingleOrDefault(i => i.Id == hyperlink.Id);
        if (hyperlinksRelation != null) returnedValue = hyperlinksRelation.Uri.ToString();
        return returnedValue;
    }

    public ExcelCellParser(WorkbookPart workbookPart) {
        _workbookPart = workbookPart;
        _worksheetPart = workbookPart.WorksheetParts.First();
    }
}