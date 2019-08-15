using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateInsideVerticalBorder
    {
        private EnumValue<BorderValues> val;
        private StringValue color;
        private UInt32Value size;
        private UInt32Value space;

        public GenerateInsideVerticalBorder()
        {
            this.val = BorderValues.Single;
            this.color = "auto";
            this.size = (UInt32Value)4U;
            this.space = (UInt32Value)0U;
        }

        // Creates an InsideVerticalBorder instance and adds its children.
        public InsideVerticalBorder Create()
        {
            InsideVerticalBorder insideVerticalBorder = new InsideVerticalBorder()
            {

                Val = val,
                Color = color,
                Size = size,
                Space = space
            };
            return insideVerticalBorder;
        }

    }
}
