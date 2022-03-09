using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using NewChangeReportGenerator.Core;

namespace NewChangeReportGenerator;

public class ChangeReportGeneratorController {
    private readonly string _filePath;

    public List<string> RowNumberList { get; private set; }
    public List<string> SapObjectList { get; private set; }
    public List<string> DefinedByList { get; private set; }
    public Dictionary<string, string>[] DefinedByDictionariesArray { get; private set; }

    public void ProcessExcelDocument() {
        using var spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        if (workbookPart == null) return;
        var changeReportDataService = new ChangeReportDataService(workbookPart);

        RowNumberList = changeReportDataService.RowNumberList;
        SapObjectList = changeReportDataService.SapObjectList;
        DefinedByList = changeReportDataService.DefinedByList;
        DefinedByDictionariesArray = changeReportDataService.DefinedByDictionariesArray;
    }



    public ChangeReportGeneratorController(string filePath) {
        _filePath = filePath;
        ProcessExcelDocument();
    }
}