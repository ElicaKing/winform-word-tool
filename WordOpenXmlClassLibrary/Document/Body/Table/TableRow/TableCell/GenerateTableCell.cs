using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableCell
    {
        public GenerateTableCell()
        {
        }

        // Creates an TableCell instance and adds its children.
        public TableCell Create(params OpenXmlElement[] newChildren)
        {
            TableCell tableCell = new TableCell();
            tableCell.Append(newChildren);
            return tableCell;
        }
    }
}
