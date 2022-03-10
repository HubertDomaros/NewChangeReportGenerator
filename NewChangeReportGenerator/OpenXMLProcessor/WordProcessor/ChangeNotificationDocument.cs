using ChangeNotificationGenerator.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor; 

public class ChangeNotificationDocument {
    private readonly string _filePath;
    private readonly ChangeNotificationDataModel _changeNotificationDataModel;
    private readonly CheckboxesConfig _checkboxesConfig;

    public void CreateChangeNotificationDocument() {
        using var wordDocument = WordprocessingDocument.Create(_filePath, WordprocessingDocumentType.Document);
        var mainDocumentPart = wordDocument.AddMainDocumentPart();

        //Creating document structure
        mainDocumentPart.Document = new Document();
        var document = mainDocumentPart.Document;
        var body = mainDocumentPart.Document.AppendChild(new Body());

        //Adding Change Report table
        var changeNotificationTable = new ChangeNotificationTable(mainDocumentPart, _changeNotificationDataModel, _checkboxesConfig);
        body.AppendChild(changeNotificationTable.InsertTable());

        //Saving file
        document.Save();
    }

    public ChangeNotificationDocument(string filePath, ChangeNotificationDataModel changeNotificationDataModel, CheckboxesConfig checkboxesConfig) {
        _filePath = filePath;
        _changeNotificationDataModel = changeNotificationDataModel;
        _checkboxesConfig = checkboxesConfig;
    }
}