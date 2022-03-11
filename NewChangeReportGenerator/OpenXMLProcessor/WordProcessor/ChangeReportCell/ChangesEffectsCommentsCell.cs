using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

public class ChangesEffectsCommentsCell {
    public TableCell InsertCell(int rowNumber) {
        var cell = new TableCell();
        cell.Append(new Paragraph(new Run(new Text())));
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        return cell;
    }
}