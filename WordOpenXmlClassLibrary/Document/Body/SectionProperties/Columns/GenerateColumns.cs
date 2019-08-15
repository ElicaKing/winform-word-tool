using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateColumns
    {
        private OnOffValue equalWidth;
        private Int16Value columnCount;
        public GenerateColumns()
        {
        }

        public GenerateColumns(Int16Value columnCount)
        {
            this.columnCount = columnCount ?? throw new ArgumentNullException(nameof(columnCount));
        }

        public GenerateColumns(OnOffValue equalWidth, Int16Value columnCount)
        {
            this.equalWidth = equalWidth ?? throw new ArgumentNullException(nameof(equalWidth));
            this.columnCount = columnCount ?? throw new ArgumentNullException(nameof(columnCount));
        }


        // Creates an Columns instance and adds its children.
        public Columns Create(params OpenXmlElement[] newChildren)
        {
            Columns columns = new Columns()
            {
                EqualWidth = equalWidth, // false,
                ColumnCount = columnCount // 2
            };
            columns.Append(newChildren);
            return columns;
        }
    }
}
