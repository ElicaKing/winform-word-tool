using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCellWidth
    {
        private StringValue width;
        private EnumValue<TableWidthUnitValues> type;

        public GenerateTableCellWidth()
        {
            this.width = "10682";
            this.type = TableWidthUnitValues.Dxa;
        }

        public GenerateTableCellWidth(StringValue width)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
            this.type = TableWidthUnitValues.Dxa;
        }

        public GenerateTableCellWidth(Int32Value tableWidth)
        {
            this.width = tableWidth + "";
            this.type = TableWidthUnitValues.Dxa;
        }

        // Creates an TableCellWidth instance and adds its children.
        public TableCellWidth Create()
        {
            TableCellWidth tableCellWidth = new TableCellWidth()
            {
                Width = width,
                Type = type
            };
            return tableCellWidth;
        }
    }
}
