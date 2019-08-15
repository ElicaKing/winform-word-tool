using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateInsideHorizontalBorder
    {
        private EnumValue<BorderValues> val;
        private StringValue color;
        private UInt32Value size;
        private UInt32Value space;

        public GenerateInsideHorizontalBorder()
        {
            this.val = BorderValues.Single;
            this.color = "auto";
            this.size = (UInt32Value)4U;
            this.space = (UInt32Value)0U;
        }
        // Creates an InsideHorizontalBorder instance and adds its children.
        public InsideHorizontalBorder Create()
        {
            InsideHorizontalBorder insideHorizontalBorder = new InsideHorizontalBorder()
            {
                Val = val,
                Color = color,
                Size = size,
                Space = space
            };
            return insideHorizontalBorder;
        }
    }
}
