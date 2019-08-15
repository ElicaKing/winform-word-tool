using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateGridColumn
    {
        private StringValue width;
        public GenerateGridColumn()
        {
            this.width = "534";
        }

        public GenerateGridColumn(StringValue width)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
        }

        // Creates an GridColumn instance and adds its children.
        public GridColumn Create()
        {
            GridColumn gridColumn = new GridColumn()
            {
                Width = width
            };
            return gridColumn;
        }
    }
}
