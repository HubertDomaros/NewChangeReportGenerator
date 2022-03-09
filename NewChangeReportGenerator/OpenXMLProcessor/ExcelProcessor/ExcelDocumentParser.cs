using System.Diagnostics;
using ChangeNotificationGenerator.Core;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor; 

public class ExcelDocumentParser {
    private readonly string _filePath;
    public ChangeNotificationDataModel ChangeNotificationDataModel { get; }

    private ChangeNotificationDataModel ProcessChangeNotificationData() {
        using var spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false);
        var dataModel = new ChangeNotificationDataModel(spreadsheetDocument.WorkbookPart);

        return dataModel;
    }

    public ExcelDocumentParser(string filePath) {
        _filePath = filePath;
        ChangeNotificationDataModel = ProcessChangeNotificationData();
    }
}