using System;
using System.Collections.Generic;
using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.Warnings;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor; 

public class ExcelDocumentParser {
    private readonly string _filePath;
    public List<string> RowNumberList { get; private set; }
    public List<string> SapObjectList { get; private set; }
    public List<string> DefinedByList { get; private set; }

    public ChangeNotificationDataModel ProcessChangeNotificationData() {
        using var spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false);
        if (spreadsheetDocument.WorkbookPart == null) throw new ArgumentNullException("Workbook is null or empty");
        
        SetClassProprieties(spreadsheetDocument.WorkbookPart);
        CheckIfAllDefinedByItemsAreInSapObjectList();


        var dataModel = new ChangeNotificationDataModel(spreadsheetDocument.WorkbookPart);
        return dataModel;
    }

    private void SetClassProprieties(WorkbookPart workbookPart) {
        ExcelColumnParser columnParser = new ExcelColumnParser(workbookPart);
        RowNumberList = columnParser.UrlColumnToStringList("A");
        SapObjectList = columnParser.TextColumnToStringList("B");
        DefinedByList = columnParser.TextColumnToStringList("C");
    }

    private void CheckIfAllDefinedByItemsAreInSapObjectList() {
        List<string> notFoundItemsList = new List<string>();
        
        foreach (var definedByItem in DefinedByList) {
            if (!IsDefinedByItem(definedByItem)) notFoundItemsList.Add(definedByItem);
        }

        if (notFoundItemsList.Count > 0) new ElementsNotFoundWarning(notFoundItemsList);
    }

    private bool IsDefinedByItem(string definedByItem) {
        foreach (var sapObjectItem in SapObjectList) {
            if (definedByItem.Contains(sapObjectItem)) return true;
        }
        return false;
    }

    public ExcelDocumentParser(string filePath) {
        _filePath = filePath;
    }
}