using System.Collections.Generic;
using System.Diagnostics;
using ChangeNotificationGenerator.DebugUtilities;
using ChangeNotificationGenerator.OpenXMLProcessor.ExcelProcessor;
using DocumentFormat.OpenXml.Packaging;

namespace ChangeNotificationGenerator.Core;

public class ChangeNotificationDataService {
    private readonly WorkbookPart _workbookPart;
    
    public List<string> RowNumberList { get; private set; }
    public List<string> SapObjectList { get; private set; }
    public List<string> DefinedByList { get; private set; }
    public Dictionary<string, string>[] DefinedByDictionariesArray { get; private set; }

    private void SetClassProprieties() {
        ExcelColumnParser columnParser = new ExcelColumnParser(_workbookPart);
        

        
        RowNumberList = columnParser.UrlColumnToStringList("A");
        SapObjectList = columnParser.TextColumnToStringList("B");
        DefinedByList = columnParser.TextColumnToStringList("C");

        PrintDebugProperties();

        MainSortingAlgorithm mainSortingAlgorithm =
            new MainSortingAlgorithm(RowNumberList, SapObjectList, DefinedByList);
        DefinedByDictionariesArray = mainSortingAlgorithm.GetDefinedByItemsWithUrls();
    }

    private void PrintListWhileDebugging(List<string> debuggedList) {
        foreach (var item in debuggedList) {
            Debug.Print(item);
        }
        Debug.Print("");
    }

    private void PrintDebugProperties() {
        DebugUtils.PrintDebuggedList(RowNumberList, "Printing RowNumberList");
        DebugUtils.PrintDebuggedList(SapObjectList, "Printing SapObjectList");
        DebugUtils.PrintDebuggedList(DefinedByList, "Printing DefinedByList");
        DebugUtils.PrintDebuggedDictionariesArray(DefinedByDictionariesArray, "Printing DefinedByDictionariesArray");
    }

    public ChangeNotificationDataService(WorkbookPart workbookPart) {
        _workbookPart = workbookPart;
        SetClassProprieties();
#if DEBUG
        PrintDebugProperties();
#endif
    }
}