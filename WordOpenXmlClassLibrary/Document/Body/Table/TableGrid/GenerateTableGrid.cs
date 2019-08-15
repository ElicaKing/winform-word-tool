using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableGrid
    {
        public GenerateTableGrid()
        {
        }

        // Creates an TableGrid instance and adds its children.
        public TableGrid Create(params OpenXmlElement[] newChildren)
        {
            TableGrid tableGrid = new TableGrid();
            tableGrid.Append(newChildren);
            return tableGrid;
        }

    }
}
