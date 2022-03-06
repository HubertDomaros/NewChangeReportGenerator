using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;

internal interface IChangeReportCell {
    TableCell InsertCell(int rowNumber);
}