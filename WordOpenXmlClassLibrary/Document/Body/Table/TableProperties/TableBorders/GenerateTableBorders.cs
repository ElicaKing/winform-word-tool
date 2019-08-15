using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableBorders
    {
        public GenerateTableBorders()
        {
        }

        public TableBorders Create(params OpenXmlElement[] newChildren)
        {
            TableBorders tableBorders = new TableBorders();
            tableBorders.Append(newChildren);
            return tableBorders;
        }
    }
}
