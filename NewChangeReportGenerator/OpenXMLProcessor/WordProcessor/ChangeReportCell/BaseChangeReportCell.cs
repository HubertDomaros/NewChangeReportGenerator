using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;

internal abstract class BaseChangeReportCell {
    protected MainDocumentPart DocumentPart = null!;

    TableCell InsertCell() {
        return new TableCell();
    }
}