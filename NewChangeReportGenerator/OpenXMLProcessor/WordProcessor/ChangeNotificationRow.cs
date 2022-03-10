using ChangeNotificationGenerator.Core;
using ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor; 

public class ChangeNotificationRow {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly ChangeNotificationDataModel _sortedData;
    private readonly CheckboxesConfig _checkboxesConfig;

    public TableRow InsertDataRow(int rowNumber) {
        var row = new TableRow();

        row.Append(new RowNumberCell(_mainDocumentPart, _sortedData, _checkboxesConfig).InsertCell(rowNumber));
        row.Append(new SapMaterialsAndDocumentsCell(_mainDocumentPart, _sortedData, _checkboxesConfig).InsertCell(rowNumber));
        row.Append(new ChangesEffectsCommentsCell().InsertCell(rowNumber)); 
        row.Append(new SwitchOverInformationCell().InsertCell(rowNumber));

        return row;
    }

    public TableRow InsertHeaderRow() {
        TableRow headerRow = new TableRow();

        TableRowProperties headerRowProperties = new TableRowProperties(
            new CantSplit() { Val = OnOffOnlyValues.On },
            new TableHeader() { Val = OnOffOnlyValues.On }
        );
        headerRow.AppendChild(headerRowProperties);

        for (int i = 0; i < 4; i++) {
            TableCell cell = new TableCell();
            cell.Append(new Paragraph(new Run(new Text("Header row"))));
            headerRow.Append(cell);
        }
        return headerRow;
    }

    public ChangeNotificationRow(MainDocumentPart mainDocumentPart, ChangeNotificationDataModel sortedData, CheckboxesConfig changeReportCheckboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _sortedData = sortedData;
        _checkboxesConfig = changeReportCheckboxesConfig;
    }
}