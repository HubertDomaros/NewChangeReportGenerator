using System.Collections.Generic;
using ChangeNotificationGenerator.Core;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator;

public class ChangeNotificationGeneratorController {
    private readonly string _filePath;

    public List<string> RowNumberList { get; private set; }
    public List<string> SapObjectList { get; private set; }
    public List<string> DefinedByList { get; private set; }
    public Dictionary<string, string>[] DefinedByDictionariesArray { get; private set; }

    public void ProcessExcelDocument() {
        using var spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        if (workbookPart == null) return;
        var changeReportDataService = new ChangeNotificationDataService(workbookPart);

        RowNumberList = changeReportDataService.RowNumberList;
        SapObjectList = changeReportDataService.SapObjectList;
        DefinedByList = changeReportDataService.DefinedByList;
        DefinedByDictionariesArray = changeReportDataService.DefinedByDictionariesArray;
    }



    public ChangeNotificationGeneratorController(string filePath) {
        _filePath = filePath;
        ProcessExcelDocument();
    }
}