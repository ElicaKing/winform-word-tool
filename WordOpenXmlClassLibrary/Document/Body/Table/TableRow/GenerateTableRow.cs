using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


namespace WordOpenXmlClassLibrary
{
    public class GenerateTableRow
    {
        public GenerateTableRow()
        {
        }

        // Creates an TableRow instance and adds its children.
        public TableRow Create(params OpenXmlElement[] newChildren)
        {
            TableRow tableRow = new TableRow();
            tableRow.Append(newChildren);
            return tableRow;
        }

    }
}
