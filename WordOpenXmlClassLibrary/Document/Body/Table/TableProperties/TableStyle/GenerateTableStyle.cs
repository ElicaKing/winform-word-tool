using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableStyle
    {
        private StringValue Val;
        public GenerateTableStyle()
        {
            this.Val = "3";
        }

        public GenerateTableStyle(StringValue val)
        {
            Val = val ?? throw new ArgumentNullException(nameof(val));
        }

        // Creates an TableStyle instance and adds its children.
        public TableStyle Create()
        {
            TableStyle tableStyle = new TableStyle()
            {
                Val = Val // "3"
            };
            return tableStyle;
        }

    }
}
