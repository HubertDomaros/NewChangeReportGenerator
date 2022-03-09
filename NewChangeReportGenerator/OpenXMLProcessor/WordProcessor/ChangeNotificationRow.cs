using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor; 

public class ChangeNotificationRow {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly ChangeNotificationDataService _sortedData;
    private readonly CheckboxesConfig _checkboxesConfig;

    public TableRow InsertDataRow(int rowNumber) {
        var row = new TableRow();

        row.Append(new RowNumberCell(_mainDocumentPart, _sortedData, _checkboxesConfig).InsertCell(rowNumber));
        row.Append(new SapMaterialsAndDocumentsCell(_mainDocumentPart, _sortedData, _checkboxesConfig).InsertCell(rowNumber));
        row.Append(new ChangesEffectsCommentsCell().InsertCell(rowNumber));
        row.Append(new SwitchOverInformationCell().InsertCell(rowNumber));

        return row;
    }

    public TableRow InsertHeaderRow(Table table) {
        var row = table.GetFirstChild<TableRow>();

        if (row.TableRowProperties == null)
            row.TableRowProperties = new TableRowProperties();

        row.TableRowProperties.AppendChild(new TableHeader());

        return new TableRow();
    }

    public ChangeNotificationRow(MainDocumentPart mainDocumentPart, ChangeNotificationDataService sortedData, CheckboxesConfig changeReportCheckboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _sortedData = sortedData;
        _checkboxesConfig = changeReportCheckboxesConfig;
    }
}