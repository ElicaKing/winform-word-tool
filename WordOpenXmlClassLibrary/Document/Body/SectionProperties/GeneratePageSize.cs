using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GeneratePageSize
    {
        private UInt32Value width;
        private UInt32Value height;
        public GeneratePageSize()
        {
            this.width = 23757U;
            this.height = 16783U;
        }

        public GeneratePageSize(UInt32Value width, UInt32Value height)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
            this.height = height ?? throw new ArgumentNullException(nameof(height));
        }


        // Creates an PageSize instance and adds its children.
        public PageSize Create()
        {
            PageSize pageSize = new PageSize()
            {
                Width = width,
                Height = height
            };
            return pageSize;
        }
    }
}
