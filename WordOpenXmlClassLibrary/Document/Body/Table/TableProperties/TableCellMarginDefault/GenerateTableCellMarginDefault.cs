using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCellMarginDefault
    {
        public GenerateTableCellMarginDefault()
        {
        }

        // Creates an TableCellMarginDefault instance and adds its children.
        public TableCellMarginDefault Create(params OpenXmlElement[] newChildren)
        {
            TableCellMarginDefault tableCellMarginDefault = new TableCellMarginDefault();
            tableCellMarginDefault.Append(newChildren);
            return tableCellMarginDefault;
        }

    }
}
