using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateColumn
    {
        private StringValue width;
        private StringValue space;
        public GenerateColumn()
        {
            width = "10586";
            space = "425";
        }

        public GenerateColumn(StringValue width)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
        }

        public GenerateColumn(StringValue width, StringValue space)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
            this.space = space ?? throw new ArgumentNullException(nameof(space));
        }

        // Creates an Column instance and adds its children.
        public Column Create()
        {
            Column column = new Column()
            {
                Width = width, //"10586",
                Space = space //"425"
            };
            return column;
        }


    }
}
