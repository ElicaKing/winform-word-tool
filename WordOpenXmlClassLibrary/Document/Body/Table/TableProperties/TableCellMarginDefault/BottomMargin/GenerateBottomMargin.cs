using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateBottomMargin
    {
        private StringValue width;
        private EnumValue<TableWidthUnitValues> type;
        public GenerateBottomMargin()
        {
            this.width = "0";
            this.type = TableWidthUnitValues.Dxa;
        }

        // Creates an BottomMargin instance and adds its children.
        public BottomMargin Create()
        {
            BottomMargin bottomMargin = new BottomMargin()
            {
                Width = width,
                Type = type
            };
            return bottomMargin;
        }
    }
}
