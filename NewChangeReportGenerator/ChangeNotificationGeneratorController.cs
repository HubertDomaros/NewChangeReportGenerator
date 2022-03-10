using System;
using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator;

public class ChangeNotificationGeneratorController {

    private ChangeNotificationDataModel _changeNotificationDataModel;
    

    public void ProcessExcelDocument(string excelFilePath) {
        using var spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        if (workbookPart == null) throw new ArgumentNullException("Workbook is null or empty");

        var changeNotificationDataModel = new ChangeNotificationDataModel(workbookPart);
        _changeNotificationDataModel = changeNotificationDataModel;
    }

    public void GenerateChangeNotificationDocument(string wordFilePath, CheckboxesConfig checkboxesConfig) {
        var changeNotificationDocument = new ChangeNotificationDocument(wordFilePath, _changeNotificationDataModel, checkboxesConfig);
        changeNotificationDocument.CreateChangeNotificationDocument();
    }
}