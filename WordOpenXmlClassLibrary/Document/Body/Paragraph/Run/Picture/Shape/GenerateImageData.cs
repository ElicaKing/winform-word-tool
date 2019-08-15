using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Vml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateImageData
    {
        // Creates an ImageData instance and adds its children.
        public ImageData Create()
        {
            ImageData imageData = new ImageData() { Title = "" };
            imageData.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            return imageData;
        }
    }
}
