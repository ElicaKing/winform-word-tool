using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateGridSpan
    {
        private Int32Value val;
        public GenerateGridSpan()
        {
            val = 20;
        }

        public GenerateGridSpan(Int32Value val)
        {
            this.val = val ?? throw new ArgumentNullException(nameof(val));
        }

        // Creates an GridSpan instance and adds its children.
        public GridSpan Create()
        {
            GridSpan gridSpan = new GridSpan()
            {
                Val = val
            };
            return gridSpan;
        }

    }
}
