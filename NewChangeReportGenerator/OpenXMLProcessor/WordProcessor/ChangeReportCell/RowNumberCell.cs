using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

internal class RowNumberCell : BaseChangeReportCell {
    private readonly bool _rowNumberCheckbox;

    public TableCell InsertCell(string rowNumber, string url) {
        var cell = new TableCell();
        cell.Append(HyperlinkUtils.InjectParagraphWithOptionalHyperlink(DocumentPart, _rowNumberCheckbox, rowNumber, url));
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        return cell;
    }

    public RowNumberCell(MainDocumentPart documentPart, bool rowNumberCheckbox) {
        DocumentPart = documentPart;
        _rowNumberCheckbox = rowNumberCheckbox;
    }
}