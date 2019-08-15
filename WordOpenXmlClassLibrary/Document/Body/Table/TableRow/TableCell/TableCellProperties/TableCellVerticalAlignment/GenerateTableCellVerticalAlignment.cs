using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCellVerticalAlignment
    {
        private EnumValue<TableVerticalAlignmentValues> val;

        public GenerateTableCellVerticalAlignment()
        {
            val = TableVerticalAlignmentValues.Center;
        }

        // Creates an TableCellVerticalAlignment instance and adds its children.
        public TableCellVerticalAlignment Create()
        {
            TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment()
            {
                Val = val
            };
            return tableCellVerticalAlignment;
        }

    }
}
