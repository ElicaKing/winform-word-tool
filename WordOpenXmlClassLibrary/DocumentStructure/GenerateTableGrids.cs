using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableGrids
    {
        public GenerateTableGrids()
        {
        }

        public TableGrid Create(int tableWidth, int columnNum)
        {
            int columnWidth = tableWidth / columnNum;
            StringValue cWidth = columnWidth + "";
            TableGrid tableGrid = new TableGrid();
            for (int i = 0; i < columnNum; i++)
            {
                tableGrid.Append(new GenerateGridColumn(cWidth).Create());
            }
            return tableGrid;
        }
    }
}
