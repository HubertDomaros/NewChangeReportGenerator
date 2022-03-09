using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeNotificationGenerator.OpenXMLProcessor.WordProcessor; 

internal class ChangeNotificationTableStyling {
    private readonly TopBorder _topBorder = new() {
        Val = new(BorderValues.Single),
        Size = 6
    };
    private readonly BottomBorder _bottomBorder = new() {
        Val = new(BorderValues.Single),
        Size = 6
    };
    private readonly LeftBorder _leftBorder = new() {
        Val = new EnumValue<BorderValues>(BorderValues.Single),
        Size = 6
    };
    private readonly RightBorder _rightBorder = new() {
        Val = new EnumValue<BorderValues>(BorderValues.Single),
        Size = 6
    };
    private readonly InsideHorizontalBorder _insideHorizontalBorder = new() {
        Val = new EnumValue<BorderValues>(BorderValues.Single),
        Size = 6
    };
    private readonly InsideVerticalBorder _verticalBorder = new() {
        Val = new EnumValue<BorderValues>(BorderValues.Single),
        Size = 6
    };

    public TableProperties SetTableBorderProperties() {
        TableBorders tableBorders = new TableBorders(_topBorder, _bottomBorder, _leftBorder, _rightBorder, _insideHorizontalBorder, _verticalBorder);
        return new TableProperties(tableBorders);
    }
}