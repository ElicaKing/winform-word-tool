using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableIndentation
    {
        private Int32Value width;
        private EnumValue<TableWidthUnitValues> type;

        public GenerateTableIndentation()
        {
            this.width = 0;
            this.type = TableWidthUnitValues.Dxa;
        }

        // Creates an TableIndentation instance and adds its children.
        public TableIndentation Create()
        {
            TableIndentation tableIndentation = new TableIndentation()
            {
                Width = width,
                Type = type
            };
            return tableIndentation;
        }

    }
}
