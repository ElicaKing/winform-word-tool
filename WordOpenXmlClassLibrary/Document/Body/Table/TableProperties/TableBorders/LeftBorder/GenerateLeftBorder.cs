using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateLeftBorder
    {
        private EnumValue<BorderValues> val;
        private StringValue color;
        private UInt32Value size;
        private UInt32Value space;

        public GenerateLeftBorder()
        {
            this.val = BorderValues.Single;
            this.color = "auto";
            this.size = (UInt32Value)4U;
            this.space = (UInt32Value)0U;
        }
        // Creates an LeftBorder instance and adds its children.
        public LeftBorder Create()
        {

            LeftBorder leftBorder = new LeftBorder()
            {
                Val = BorderValues.Single,
                Color = "auto",
                Size = (UInt32Value)4U,
                Space = (UInt32Value)0U
            };
            return leftBorder;

        }

    }
}
