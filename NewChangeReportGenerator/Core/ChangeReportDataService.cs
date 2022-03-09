using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Navigation;
using DocumentFormat.OpenXml.Packaging;
using NewChangeReportGenerator.OpenXMLProcessor.ExcelProcessor;

namespace NewChangeReportGenerator.Core;

public class ChangeReportDataService {
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
        Debug.Print("Printing RowNumberList");
        PrintListWhileDebugging(RowNumberList);
        Debug.Print("Printing SapObjectList");
        PrintListWhileDebugging(SapObjectList);
        Debug.Print("Printing DefinedByList");
        PrintListWhileDebugging(DefinedByList);
    }

    public ChangeReportDataService(WorkbookPart workbookPart) {
        _workbookPart = workbookPart;
        SetClassProprieties();
    }

    public ChangeReportDataService(WorkbookPart workbookPart, bool isDebugMode) {
        _workbookPart = workbookPart;
        SetClassProprieties();
        if (isDebugMode) {
            //PrintDebugProperties();
        }
    }
}