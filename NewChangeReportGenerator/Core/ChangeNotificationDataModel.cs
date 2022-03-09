using System.Collections.Generic;
using System.Diagnostics;
using ChangeNotificationGenerator.DebugUtilities;
using ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator.Core;

public class ChangeNotificationDataModel {
    private readonly WorkbookPart _workbookPart;
    
    public List<string> RowNumberList { get; private set; }
    public List<string> SapObjectList { get; private set; }
    public List<string> DefinedByList { get; private set; }
    public List<Dictionary<string, string>> DefinedByItemsWithUrls { get; private set; }

    private void SetClassProprieties() {
        ExcelColumnParser columnParser = new ExcelColumnParser(_workbookPart);
        
        RowNumberList = columnParser.UrlColumnToStringList("A");
        SapObjectList = columnParser.TextColumnToStringList("B");
        DefinedByList = columnParser.TextColumnToStringList("C");

        MainSortingAlgorithm mainSortingAlgorithm =
            new MainSortingAlgorithm(RowNumberList, SapObjectList, DefinedByList);
        DefinedByItemsWithUrls = mainSortingAlgorithm.DefinedByItemsWithUrls();
    }

    private void PrintDebugProperties() {
        DebugUtils.PrintDebuggedList(RowNumberList, "Printing RowNumberList");
        DebugUtils.PrintDebuggedList(SapObjectList, "Printing SapObjectList");
        DebugUtils.PrintDebuggedList(DefinedByList, "Printing DefinedByList");
        DebugUtils.PrintDebuggedDictionariesList(DefinedByItemsWithUrls, "Printing DefinedByItemsWithUrls");
    }

    public ChangeNotificationDataModel(WorkbookPart workbookPart) {
        _workbookPart = workbookPart;
        SetClassProprieties();
#if DEBUG
        PrintDebugProperties();
#endif
    }
}