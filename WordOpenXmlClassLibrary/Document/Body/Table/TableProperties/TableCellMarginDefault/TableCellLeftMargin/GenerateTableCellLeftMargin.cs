using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCellLeftMargin
    {
        private Int16Value width;
        private EnumValue<TableWidthValues> type;
        public GenerateTableCellLeftMargin()
        {
            this.width = 108;
            this.type = TableWidthValues.Dxa;
        }
        // Creates an TableCellLeftMargin instance and adds its children.
        public TableCellLeftMargin Create()
        {
            TableCellLeftMargin tableCellLeftMargin = new TableCellLeftMargin()
            {
                Width = width,
                Type = type
            };
            return tableCellLeftMargin;
        }

    }
}
