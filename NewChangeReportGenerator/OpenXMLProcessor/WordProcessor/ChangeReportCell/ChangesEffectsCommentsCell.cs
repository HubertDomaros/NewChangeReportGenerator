using DocumentFormat.OpenXml.Wordprocessing;
using TableCell = DocumentFormat.OpenXml.Drawing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;

namespace NewChangeReportGenerator.OpenXMLProcessor.WordProcessor.ChangeReportCell; 

public class ChangesEffectsCommentsCell {
    public TableCell InsertCell(int rowNumber) {
        var cell = new TableCell();
        cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5" }));
        return cell;
    }
}