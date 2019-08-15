using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTopMargin
    {
        private StringValue width;
        private EnumValue<TableWidthUnitValues> type;
        public GenerateTopMargin()
        {
            this.width = "0";
            this.type = TableWidthUnitValues.Dxa;
        }

        // Creates an TopMargin instance and adds its children.
        public TopMargin Create()
        {
            TopMargin topMargin = new TopMargin()
            {
                Width = width,
                Type = type
            };
            return topMargin;
        }
    }
}
