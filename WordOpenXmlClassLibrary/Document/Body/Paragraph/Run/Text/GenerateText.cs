using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateText
    {
        public GenerateText()
        {
        }

        public Text Create(string str)
        {
            Text text = new Text();
            text.Text = str;
            return text;
        }
        public Text CreateSpace()
        {
            Text text = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text.Text = " ";
            return text;
        }
    }
}
