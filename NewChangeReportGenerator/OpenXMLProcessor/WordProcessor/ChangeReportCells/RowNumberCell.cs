using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.WordProcessorUtils;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCells; 

internal class RowNumberCell : ChangeReportCell {
    private static Paragraph? _rowNumberParagraph;
    private readonly bool _hyperlinkCheckbox;

    public TableCell InsertCell(int rowNumber, string url) {
        TableCell cell = new TableCell();

        if (_hyperlinkCheckbox) {
            _rowNumberParagraph = HyperlinkUtils.InjectHyperlinkIntoTable(DocumentPart, url, rowNumber.ToString());
            cell.Append(_rowNumberParagraph);
            cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        }
        else {
            cell.Append(new Paragraph(new Run(new Text(rowNumber.ToString()))));
            cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        }
        return cell;
    }

    public RowNumberCell(MainDocumentPart documentPart, bool hyperlinkCheckbox) {
        DocumentPart = documentPart;
        _hyperlinkCheckbox = hyperlinkCheckbox;
    }
}