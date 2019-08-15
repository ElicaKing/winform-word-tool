using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;


namespace WordOpenXmlClassLibrary
{
    public class GenerateTextBox
    {
        // Creates an TextBox instance and adds its children.
        public TextBox Create(params OpenXmlElement[] newChildren)
        {
            TextBox textBox = new TextBox();
            textBox.Append(newChildren);
            return textBox;
        }

    }
}
