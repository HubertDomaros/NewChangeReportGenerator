using System.Collections.Generic;

namespace NewChangeReportGenerator.Core;

public class SortedData {
    public string[] RowNumberArray { get; set; }
    public string[] SapObjectArray { get; set; }
    public Dictionary<string, string>[] DefinedByDictionariesArray { get; set; }

    public SortedData(string[] rowNumberArray, string[] sapObjectArray, Dictionary<string, string>[] definedByDictionariesArray) {
        RowNumberArray = rowNumberArray;
        SapObjectArray = sapObjectArray;
        DefinedByDictionariesArray = definedByDictionariesArray;
    }
}