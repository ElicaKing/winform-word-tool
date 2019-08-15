using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTablePaneRow
    {
        public GenerateTablePaneRow()
        {
        }

        /// <summary>
        /// 创建方框行
        /// </summary>
        /// <param name="tableWidth">表格宽度</param>
        /// <param name="columnNum">列数</param>
        /// <returns></returns>
        public TableRow Create(int tableWidth, int columnNum)
        {
            int columnWidth = tableWidth / columnNum;

            TableRow tableRow = new GenerateTableRow().Create(
                new GenerateTablePropertyExceptions().Create(
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
                    ),
                new GenerateTableRowProperties().Create(
                    new GenerateTableRowHeight(454U).Create()
                    )
                );
            for (int i = 0; i < columnNum; i++)
            {
                tableRow.Append(
                    new GenerateTableCell().Create(
                        new GenerateTableCellProperties().Create(
                            new GenerateTableCellWidth(columnWidth).Create(),
                            new GenerateTableCellVerticalAlignment().Create()
                            ),
                        new GenerateParagraph().CreateCellParagraph()
                        )
                    );
            }
            return tableRow;
        }
    }
}
