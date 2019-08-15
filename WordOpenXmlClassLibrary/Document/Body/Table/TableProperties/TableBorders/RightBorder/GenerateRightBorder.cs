using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateRightBorder
    {
        private EnumValue<BorderValues> val;
        private StringValue color;
        private UInt32Value size;
        private UInt32Value space;

        public GenerateRightBorder()
        {
            this.val = BorderValues.Single;
            this.color = "auto";
            this.size = (UInt32Value)4U;
            this.space = (UInt32Value)0U;
        }
        // Creates an RightBorder instance and adds its children.
        public RightBorder Create()
        {
            RightBorder rightBorder = new RightBorder()
            {
                Val = val,
                Color = color,
                Size = size,
                Space = space
            };
            return rightBorder;
        }

    }
}
