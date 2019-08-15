using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTablePropertyExceptions
    {
        public GenerateTablePropertyExceptions()
        {
        }

        // Creates an TablePropertyExceptions instance and adds its children.
        public TablePropertyExceptions Create(params OpenXmlElement[] newChildren)
        {
            TablePropertyExceptions tablePropertyExceptions = new TablePropertyExceptions();
            tablePropertyExceptions.Append(newChildren);
            return tablePropertyExceptions;
        }


    }
}
