using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using NewChangeReportGenerator.OpenXMLProcessor.ExcelProcessor;

namespace NewChangeReportGenerator.Core;

public class SortedData {
    private WorkbookPart _workbookPart;
    
    public List<string> RowNumberList { get; private set; }
    public List<string> SapObjectList { get; private set; }
    public List<string> DefinedByList { get; private set; }
    public Dictionary<string, string>[] DefinedByDictionariesArray { get; private set; }

    private void SetClassProprieties() {
        ExcelColumnParser columnParser = new ExcelColumnParser(_workbookPart);
        MainSortingAlgorithm mainSortingAlgorithm =
            new MainSortingAlgorithm(RowNumberList, SapObjectList, DefinedByList);

        RowNumberList = columnParser.UrlColumnToStringList("A");
        SapObjectList = columnParser.TextColumnToStringList("B");
        DefinedByList = columnParser.TextColumnToStringList("C");
        DefinedByDictionariesArray = mainSortingAlgorithm.GetDefinedByItemsWithUrls();
    }

    public SortedData(WorkbookPart workbookPart) {
        _workbookPart = workbookPart;
        SetClassProprieties();
    }
}