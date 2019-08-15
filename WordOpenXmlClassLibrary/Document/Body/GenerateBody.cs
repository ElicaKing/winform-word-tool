using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    public class GenerateBody
    {
        // Creates an Body instance and adds its children.
        public Body Create(params OpenXmlElement[] newChildren)
        {
            Body body = new Body();
            body.Append(newChildren);
            return body;
        }
    }
}
