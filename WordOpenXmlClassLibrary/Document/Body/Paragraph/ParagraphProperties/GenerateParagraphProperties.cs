using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordOpenXmlClassLibrary
{
    public class GenerateParagraphProperties
    {
        public ParagraphProperties Create(params OpenXmlElement[] newChildren)
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(newChildren);
            return paragraphProperties;
        }
    }
}
