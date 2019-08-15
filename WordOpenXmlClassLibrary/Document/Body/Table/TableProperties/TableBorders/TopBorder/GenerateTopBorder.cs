using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTopBorder
    {
        private EnumValue<BorderValues> val;
        private StringValue color;
        private UInt32Value size;
        private UInt32Value space;

        public GenerateTopBorder()
        {
            this.val = BorderValues.Single;
            this.color = "auto";
            this.size = (UInt32Value)4U;
            this.space = (UInt32Value)0U;
        }
        // Creates an TopBorder instance and adds its children.
        public TopBorder Create()
        {
            TopBorder topBorder = new TopBorder()
            {
                Val = val,
                Color = color,
                Size = size,
                Space = space
            };
            return topBorder;
        }

    }
}
