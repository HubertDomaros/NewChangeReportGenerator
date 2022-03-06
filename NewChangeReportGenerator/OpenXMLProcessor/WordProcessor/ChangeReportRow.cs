using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.Core;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor; 

public class ChangeReportRow {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly SortedData _sortedData;
    private readonly CheckboxesConfig _checkboxesConfig;

    public TableRow InsertRow(int rowNumber) {
        var row = new TableRow();

        bool rowNumberCheckbox = _checkboxesConfig.RowNumberCheckboxBool;

        row.Append(new RowNumberCell(_mainDocumentPart, _sortedData, _checkboxesConfig).InsertCell(rowNumber));
        row.Append(new SapMaterialsAndDocumentsCell(_mainDocumentPart, _sortedData, _checkboxesConfig).InsertCell(rowNumber));
        row.Append(new ChangesEffectsCommentsCell().InsertCell(rowNumber));
        row.Append(new SwitchOverInformationCell().InsertCell(rowNumber));

        return row;
    }

    public ChangeReportRow(MainDocumentPart mainDocumentPart, SortedData sortedData, CheckboxesConfig changeReportCheckboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _sortedData = sortedData;
        _checkboxesConfig = changeReportCheckboxesConfig;
    }
}