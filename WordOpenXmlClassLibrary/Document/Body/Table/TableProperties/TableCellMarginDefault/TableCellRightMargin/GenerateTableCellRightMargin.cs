using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCellRightMargin
    {
        private Int16Value width;
        private EnumValue<TableWidthValues> type;
        public GenerateTableCellRightMargin()
        {
            this.width = 108;
            this.type = TableWidthValues.Dxa;
        }
        // Creates an TableCellRightMargin instance and adds its children.
        public TableCellRightMargin Create()
        {
            TableCellRightMargin tableCellRightMargin = new TableCellRightMargin()
            {
                Width = width,
                Type = type
            };
            return tableCellRightMargin;
        }

    }
}
