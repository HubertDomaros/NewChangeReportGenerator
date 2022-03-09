using System;
using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator;

public class ChangeNotificationGeneratorController {

    private readonly CheckboxesConfig _checkboxesConfig;
    private ChangeNotificationDataModel _changeNotificationDataModel;
    

    public void ProcessExcelDocument(string excelFilePath) {
        using var spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        if (workbookPart == null) throw new ArgumentNullException("Workbook is null or empty");

        var changeNotificationDataModel = new ChangeNotificationDataModel(workbookPart);
        _changeNotificationDataModel = changeNotificationDataModel;
    }

    public void GenerateChangeNotificationDocument(string wordFilePath) {
        ChangeNotificationDocument changeNotificationDocument = new ChangeNotificationDocument(wordFilePath, _changeNotificationDataModel, _checkboxesConfig);
    }

    public ChangeNotificationGeneratorController(CheckboxesConfig checkboxesConfig) {
        _checkboxesConfig = checkboxesConfig;
    }
}