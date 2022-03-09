using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using NewChangeReportGenerator.Core;


namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor;

public class ChangeReportTable {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly ChangeReportDataService _sortedData;
    private readonly Dictionary<string, string>[] _definedByDictionariesArray;
    private readonly CheckboxesConfig _checkboxesConfig;

    public Table InsertTable() {
        Table table = new Table();
        ChangeReportRow changeReportRow = new ChangeReportRow(_mainDocumentPart, _sortedData, _checkboxesConfig);

        table.Append(new ChangeReportTableStyling().SetTableBorderProperties());

        for (var i = 0; i < _definedByDictionariesArray.Length; i++) {
            table.Append(changeReportRow.InsertDataRow(i));
        }

        return table;
    }

    private TableHeader InsertTableHeader() {
        return new TableHeader();
    }

    public ChangeReportTable(MainDocumentPart mainDocumentPart, ChangeReportDataService sortedData, CheckboxesConfig checkboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _sortedData = sortedData;
        _definedByDictionariesArray = sortedData.DefinedByDictionariesArray;
        _checkboxesConfig = checkboxesConfig;
    }
}