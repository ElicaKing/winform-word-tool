using System;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GeneratePageMargin
    {

        // Creates an PageMargin instance and adds its children.
        public PageMargin Create()
        {
            PageMargin pageMargin = new PageMargin()
            {
                Top = 1080,
                Right = (UInt32Value)1080U,
                Bottom = 1080,
                Left = (UInt32Value)1080U,
                Header = (UInt32Value)720U,
                Footer = (UInt32Value)720U,
                Gutter = (UInt32Value)0U
            };
            return pageMargin;
        }
    }
}
