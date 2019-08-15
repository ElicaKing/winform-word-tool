using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    class GenerateTableCellBorders
    {
        public GenerateTableCellBorders()
        {
        }
        // Creates an TableCellBorders instance and adds its children.
        public TableCellBorders Create()
        {
            TableCellBorders tableCellBorders = new TableCellBorders();
            TopLeftToBottomRightCellBorder topLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder() { Val = BorderValues.Nil };
            TopRightToBottomLeftCellBorder topRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder() { Val = BorderValues.Nil };

            tableCellBorders.Append(topLeftToBottomRightCellBorder);
            tableCellBorders.Append(topRightToBottomLeftCellBorder);
            return tableCellBorders;
        }
    }
}
