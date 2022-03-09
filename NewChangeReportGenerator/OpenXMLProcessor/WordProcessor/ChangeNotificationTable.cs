using System.Collections.Generic;
using ChangeNotificationGenerator.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor;

public class ChangeNotificationTable {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly ChangeNotificationDataService _sortedData;
    private readonly Dictionary<string, string>[] _definedByDictionariesArray;
    private readonly CheckboxesConfig _checkboxesConfig;

    public Table InsertTable() {
        Table table = new Table();
        ChangeNotificationRow changeReportRow = new ChangeNotificationRow(_mainDocumentPart, _sortedData, _checkboxesConfig);

        table.Append(new ChangeNotificationTableStyling().SetTableBorderProperties());

        for (var i = 0; i < _definedByDictionariesArray.Length; i++) {
            table.Append(changeReportRow.InsertDataRow(i));
        }

        return table;
    }

    private TableHeader InsertTableHeader() {
        return new TableHeader();
    }

    public ChangeNotificationTable(MainDocumentPart mainDocumentPart, ChangeNotificationDataService sortedData, CheckboxesConfig checkboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _sortedData = sortedData;
        _definedByDictionariesArray = sortedData.DefinedByDictionariesArray;
        _checkboxesConfig = checkboxesConfig;
    }
}