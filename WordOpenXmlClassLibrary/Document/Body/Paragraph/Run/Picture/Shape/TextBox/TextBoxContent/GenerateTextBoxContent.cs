using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTextBoxContent
    {
        // Creates an TextBox instance and adds its children.
        public TextBoxContent Create(params OpenXmlElement[] newChildren)
        {
            TextBoxContent textBoxContent = new TextBoxContent();
            textBoxContent.Append(newChildren);
            return textBoxContent;
        }

    }
}
