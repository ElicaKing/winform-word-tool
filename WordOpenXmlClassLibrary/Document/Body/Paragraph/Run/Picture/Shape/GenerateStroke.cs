using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Vml;

namespace WordOpenXmlClassLibrary
{
    // Creates an Stroke instance and adds its children.
    public class GenerateStroke
    {
        private string width;
        private string color;
        public GenerateStroke()
        {
            this.width = "2pt";
            this.color = "#000000";
        }

        public GenerateStroke(string width)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
            if (width.Equals("0pt"))
            {
                this.color = "#ffffff";
            }
        }

        // Creates an Stroke instance and adds its children.
        public Stroke Create()
        {
            Stroke stroke = new Stroke()
            {
                Weight = width,// 2pt
                Color = color // #000000
            };
            return stroke;
        }
    }
}
