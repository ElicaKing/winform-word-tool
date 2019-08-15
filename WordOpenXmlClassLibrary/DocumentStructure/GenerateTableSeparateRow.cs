using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableSeparateRow
    {
        public GenerateTableSeparateRow()
        {
        }

        /// <summary>
        /// 创建分隔行
        /// </summary>
        /// <param name="tableWidth">表格宽度</param>
        /// <param name="columnNum">列数</param>
        /// <returns></returns>
        public TableRow Create(int tableWidth, int columnNum)
        {
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
                    new GenerateTableRowHeight().Create()
                    ),
                new GenerateTableCell().Create(
                    new GenerateTableCellProperties().Create(
                        new GenerateTableCellWidth(tableWidth).Create(),
                        new GenerateGridSpan(columnNum).Create(),
                        new GenerateTableCellVerticalAlignment().Create()
                        ),
                    new GenerateParagraph().CreateCellParagraph()
                    )
                );
            return tableRow;
        }
        public TableRow CreateSpanRow(int tableWidth, UInt32Value rowHeight)
        {
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
                    new GenerateTableRowHeight(rowHeight).Create()
                    ),
                new GenerateTableCell().Create(
                    new GenerateTableCellProperties().Create(
                        new GenerateTableCellWidth(tableWidth).Create(),
                        new GenerateTableCellVerticalAlignment().Create()
                        ),
                    new GenerateParagraph().CreateCellParagraph()
                    )
                );
            return tableRow;
        }
    }
}
