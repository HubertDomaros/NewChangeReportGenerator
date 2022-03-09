using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor; 

public class ExcelDocumentParser {
    private readonly string _filePath;
    private readonly SpreadsheetDocument _spreadsheetDocument;
    public WorkbookPart CurrentWorkbookPart { get; private set; }



    public ExcelDocumentParser(string filePath) {
        _filePath = filePath;
        var spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false);
        _spreadsheetDocument = spreadsheetDocument;
        if (spreadsheetDocument.WorkbookPart != null) {
            CurrentWorkbookPart = spreadsheetDocument.WorkbookPart;
        } else Debug.Print("WorkbookPart is null");
    }

     ~ExcelDocumentParser() {
        _spreadsheetDocument.Close();
        Debug.Print("Excel file closed successfully");
    }
}