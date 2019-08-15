using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableRowProperties
    {
        public GenerateTableRowProperties()
        {
        }

        // Creates an TableRowProperties instance and adds its children.
        public TableRowProperties Create(params OpenXmlElement[] newChildren)
        {
            TableRowProperties tableRowProperties = new TableRowProperties();
            tableRowProperties.Append(newChildren);
            return tableRowProperties;
        }

    }
}
