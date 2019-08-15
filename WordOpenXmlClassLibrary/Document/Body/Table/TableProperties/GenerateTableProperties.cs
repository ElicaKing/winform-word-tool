using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableProperties
    {
        public GenerateTableProperties()
        {
        }

        //Creates an Table instance and adds its children.
        public TableProperties Create(params OpenXmlElement[] newChildren)
        {
            TableProperties tableProperties = new TableProperties();
            tableProperties.Append(newChildren);
            return tableProperties;
        }
        public TableProperties Create(int tableWidth)
        {
            TableProperties tableProperties = new GenerateTableProperties().Create(
                new GenerateTableStyle().Create(),
                new GenerateTableWidth(tableWidth).Create(),
                new GenerateTableJustification().Create(),// 默认表格居中
                new GenerateTableIndentation().Create(),
                new GenerateTableBorders().Create(
                    new GenerateTopBorder().Create(),
                    new GenerateLeftBorder().Create(),
                    new GenerateBottomBorder().Create(),
                    new GenerateRightBorder().Create(),
                    new GenerateInsideHorizontalBorder().Create(),
                    new GenerateInsideVerticalBorder().Create()
                    ),
                new GenerateTableLayout().Create(),
                new GenerateTableCellMarginDefault().Create(
                    new GenerateTopMargin().Create(),
                    new GenerateTableCellLeftMargin().Create(),
                    new GenerateBottomMargin().Create(),
                    new GenerateTableCellRightMargin().Create()
                    )
                );
            return tableProperties;
        }
    }
}
