using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCellProperties
    {
        public GenerateTableCellProperties()
        {

        }
        // Creates an TableCellProperties instance and adds its children.
        public TableCellProperties Create(params OpenXmlElement[] newChildren)
        {
            TableCellProperties tableCellProperties = new TableCellProperties();
            tableCellProperties.Append(newChildren);
            return tableCellProperties;
        }

    }
}
