using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableWidth
    {
        private StringValue width;
        private EnumValue<TableWidthUnitValues> type;

        public GenerateTableWidth()
        {
            this.width = "10682";
            this.type = TableWidthUnitValues.Dxa;
        }

        public GenerateTableWidth(StringValue width)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
            this.type = TableWidthUnitValues.Dxa;
        }
        public GenerateTableWidth(int tableWidth)
        {
            this.width = tableWidth + "";
            this.type = TableWidthUnitValues.Dxa;
        }

        // Creates an TableWidth instance and adds its children.
        public TableWidth Create()
        {
            TableWidth tableWidth = new TableWidth()
            {
                Width = width,
                Type = type
            };
            return tableWidth;
        }
    }
}
