using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateShape
    {
        private string width;
        private string height;

        public GenerateShape()
        {
            this.width = "475.3pt";
            this.height = "73pt";
        }

        /// <summary>
        /// 创建方框
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public GenerateShape(string width, string height)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
            this.height = height ?? throw new ArgumentNullException(nameof(height));
        }

        public Shape Create(params OpenXmlElement[] newChildren)
        {
            Shape shape = new Shape()
            {
                Style = "" +
                //"position:absolute;" +
                //"left:0pt;" +
                //"margin-left:7.9pt;" +
                //"margin-top:5.9pt;" +
                "height:" + height + ";" +// 73pt;
                "width:" + width + ";" +//475.3pt;
                "z-index:1;" +
                "mso-width-relative:page;" +
                "mso-height-relative:page;",
                Filled = true,
                FillColor = "#FFFFFF",
                Stroked = true
            };
            shape.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            shape.Append(newChildren);
            return shape;
        }
    }
}
