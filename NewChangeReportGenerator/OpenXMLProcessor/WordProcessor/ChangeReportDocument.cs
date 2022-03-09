using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.Core;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor; 

public class ChangeReportDocument {
    private readonly string _filePath;
    private readonly ChangeReportDataService _sortedData;
    private readonly CheckboxesConfig _checkboxesConfig;

    public void CreateWordprocessingDocument() {
        using (var wordDocument = WordprocessingDocument.Create(_filePath, WordprocessingDocumentType.Document)) {
            var mainDocumentPart = wordDocument.AddMainDocumentPart();

            //Creating document structure
            var document = mainDocumentPart.Document.AppendChild(new Document());
            var body = document.AppendChild(new Body());

            //Adding Change Report table
            var changeReportTable = new ChangeReportTable(mainDocumentPart, _sortedData, _checkboxesConfig);
            body.AppendChild(changeReportTable.InsertTable());

            //Saving file
            document.Save();
        }
    }

    public ChangeReportDocument(string filePath, ChangeReportDataService sortedData, CheckboxesConfig checkboxesConfig) {
        _filePath = filePath;
        _sortedData = sortedData;
        _checkboxesConfig = checkboxesConfig;
    }
}