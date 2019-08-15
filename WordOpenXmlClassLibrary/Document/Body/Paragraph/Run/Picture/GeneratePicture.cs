using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace WordOpenXmlClassLibrary
{
    // Creates an Picture instance and adds its children.
    public class GeneratePicture
    {
        public Picture Create(params OpenXmlElement[] newChildren)
        {
            Picture picture = new Picture();
            picture.Append(newChildren);
            return picture;
        }
    }
}
