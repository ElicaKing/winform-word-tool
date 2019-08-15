using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateFontSize
    {
        // Creates an FontSize instance and adds its children.
        public FontSize Create()
        {
            FontSize fontSize = new FontSize()
            {
                Val = "21"
            };
            return fontSize;
        }

        public FontSize Create(StringValue Val)
        {
            FontSize fontSize = new FontSize()
            {
                Val = Val
            };
            return fontSize;
        }
    }
}
