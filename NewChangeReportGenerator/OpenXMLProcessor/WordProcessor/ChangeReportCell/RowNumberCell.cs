using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.Core;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

internal class RowNumberCell : IChangeReportCell {
    private readonly MainDocumentPart _mainDocumentPart;
    private readonly bool _rowNumberCheckbox;
    private readonly List<string> _rowNumberList;

    public TableCell InsertCell(int rowNumber) {
        var cell = new TableCell();
        
        cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(_mainDocumentPart, _rowNumberCheckbox, rowNumber.ToString(), _rowNumberList[rowNumber]));
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        
        return cell;
    }

    public RowNumberCell(MainDocumentPart mainDocumentPart, SortedData sortedData, CheckboxesConfig checkboxesConfig) {
        _mainDocumentPart = mainDocumentPart;
        _rowNumberList = sortedData.RowNumberList;
        _rowNumberCheckbox = checkboxesConfig.RowNumberCheckboxBool;
    }
}