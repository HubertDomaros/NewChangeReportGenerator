using System.Collections.Generic;
using NewChangeReportGenerator.OpenXMLProcessor.ExcelProcessor;

namespace NewChangeReportGenerator.Core;

public class SortedData {

    private ExcelDocumentParser _excelDocumentParser;
    public List<string> RowNumberArray { get; private set; }
    public List<string> SapObjectArray { get; private set; }
    public Dictionary<string, string>[] DefinedByDictionariesArray { get; private set; }

    public SortedData(ExcelDocumentParser excelDocumentParser) {
        _excelDocumentParser = excelDocumentParser;
    }
}