using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableLayout
    {
        private EnumValue<TableLayoutValues> type;
        public GenerateTableLayout()
        {
            type = TableLayoutValues.Fixed;
        }

        // Creates an TableLayout instance and adds its children.
        public TableLayout Create()
        {
            TableLayout tableLayout = new TableLayout()
            {
                Type = type
            };
            return tableLayout;
        }
    }
}
