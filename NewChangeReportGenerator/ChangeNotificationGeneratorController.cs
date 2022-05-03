using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor;
using ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor;

namespace ChangeNotificationGenerator;

public class ChangeNotificationGeneratorController {

    private ChangeNotificationDataModel _changeNotificationDataModel;
    

    public void ProcessExcelDocument(string excelFilePath) {
        var excelDocumentParser = new ExcelDocumentParser(excelFilePath);
        _changeNotificationDataModel = excelDocumentParser.ProcessChangeNotificationData();
    }

    public void GenerateChangeNotificationDocument(string wordFilePath, CheckboxesConfig checkboxesConfig) {
        var changeNotificationDocument = new ChangeNotificationDocument(wordFilePath, _changeNotificationDataModel, checkboxesConfig);
        changeNotificationDocument.CreateChangeNotificationDocument();
    }
}