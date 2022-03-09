using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell;

internal interface IChangeReportCell {
    TableCell InsertCell(int rowNumber);
}