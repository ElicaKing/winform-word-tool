using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateBottomBorder
    {
        private EnumValue<BorderValues> val;
        private StringValue color;
        private UInt32Value size;
        private UInt32Value space;

        public GenerateBottomBorder()
        {
            this.val = BorderValues.Single;
            this.color = "auto";
            this.size = (UInt32Value)4U;
            this.space = (UInt32Value)0U;
        }
        // Creates an BottomBorder instance and adds its children.
        public BottomBorder Create()
        {
            BottomBorder bottomBorder = new BottomBorder()
            {
                Val = val,
                Color = color,
                Size = size,
                Space = space
            };
            return bottomBorder;
        }

    }
}
