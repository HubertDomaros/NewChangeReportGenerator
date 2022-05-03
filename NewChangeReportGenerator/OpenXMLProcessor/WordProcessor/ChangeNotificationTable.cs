using System.Collections.Generic;
using ChangeNotificationGenerator.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor;

public class ChangeNotificationTable {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly ChangeNotificationDataModel _changeNotificationDataModel;
    private readonly List<Dictionary<string, string>> _definedByDictionariesList;
    private readonly CheckboxesConfig _checkboxesConfig;

    public Table InsertTable() {
        Table table = new Table();
        ChangeNotificationRow changeNotificationRow = new ChangeNotificationRow(_mainDocumentPart, _changeNotificationDataModel, _checkboxesConfig);
        
        table.Append(new ChangeNotificationTableStyling().SetTableBorderProperties());

        
        table.Append(changeNotificationRow.InsertHeaderRow());
        table.Append(InsertTableHeader());

        for (int i = 0; i < _definedByDictionariesList.Count; i++) {
            table.Append(changeNotificationRow.InsertDataRow(i));
        }

        return table;
    }

    private TableProperties InsertTableHeader() {
        var tableProperties = new TableProperties();
        TableLook tableLook = new TableLook() { FirstRow = true };
        TableHeader tableHeader = new TableHeader();
        
        tableProperties.Append(tableLook);
        tableProperties.Append(tableHeader);

        return new TableProperties();
    }

    public ChangeNotificationTable(MainDocumentPart mainDocumentPart, ChangeNotificationDataModel changeNotificationDataModel, CheckboxesConfig checkboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _changeNotificationDataModel = changeNotificationDataModel;
        _definedByDictionariesList = changeNotificationDataModel.DefinedByItemsWithUrls;
        _checkboxesConfig = checkboxesConfig;
    }
}