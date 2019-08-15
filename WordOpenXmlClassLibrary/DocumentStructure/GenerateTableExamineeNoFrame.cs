using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;

namespace WordOpenXmlClassLibrary
{
    class GenerateTableExamineeNoFrame
    {
        public GenerateTableExamineeNoFrame()
        {
        }
        public TableRow CreateRowHeader(StringValue title, Int32Value tableWidth, Int32Value columnNum, UInt32Value rowHeight, StringValue fontSize)
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
                       new GenerateTableCantSplit().Create(),
                       new GenerateTableRowHeight(rowHeight).Create()
                       ),
                   new GenerateTableCell().Create(
                       new GenerateTableCellProperties().Create(
                           new GenerateTableCellWidth(tableWidth).Create(),
                           new GenerateGridSpan(columnNum).Create(),
                           new GenerateTableCellBorders().Create(),
                           new GenerateTableCellVerticalAlignment().Create()
                           ),
                       new GenerateParagraph().Create(
                           new GenerateParagraphProperties().Create(
                               new GenerateJustification().Create()
                               ),
                           new GenerateRun().Create(
                               new GenerateRunProperties().Create(
                                   new GenerateRunFonts().Create(),
                                   new GenerateFontSize().Create(fontSize)
                                   ),
                               new GenerateText().Create(title)
                               )
                           )
                       )
                   );
            return tableRow;
        }

        public TableRow CreateRowItem(StringValue title, Int32Value tableWidth, Int32Value columnNum, UInt32Value rowHeight, StringValue fontSize, JustificationValues justification)
        {
            Int32Value cellWidth = tableWidth / columnNum;

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
                       new GenerateTableCantSplit().Create(),
                       new GenerateTableRowHeight(rowHeight).Create()
                       )
                   );
            for (int i = 0; i < columnNum; i++)
            {
                Run runLeftSpace = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create("12")
                                       ),
                                   new GenerateText().Create(" ")
                                   );
                Run runRightSpace = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create("12")
                                       ),
                                   new GenerateText().Create(" ")
                                   );
                Run runLeft = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create("[ ")
                                   );
                Run runRight = new GenerateRun().Create(
                                  new GenerateRunProperties().Create(
                                      new GenerateRunFonts().Create(),
                                      new GenerateFontSize().Create(fontSize)
                                      ),
                                  new GenerateText().Create(" ]")
                                  );
                Run runContent = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new Color() { Val = "C0C0C0" },
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create(title)
                                   );

                Paragraph paragraph = new GenerateParagraph().Create(
                           new GenerateParagraphProperties().Create(
                               new GenerateJustification().Create(justification) // 分散对齐
                               )
                           );
                if (title.Equals(""))
                {
                    paragraph.Append(runContent);
                }
                else
                {
                    paragraph.Append(runLeft);
                    paragraph.Append(runLeftSpace);
                    paragraph.Append(runContent);
                    paragraph.Append(runRightSpace);
                    paragraph.Append(runRight);
                }
                tableRow.Append(
                   new GenerateTableCell().Create(
                       new GenerateTableCellProperties().Create(
                           new GenerateTableCellWidth(cellWidth).Create(),
                           new GenerateTableCellBorders().Create(),
                           new GenerateTableCellVerticalAlignment().Create()
                           ),
                       paragraph
                       )
                   );
            }
            return tableRow;
        }

        internal TableRow CreateRowObjectiveLandscape(int itemNo, string[] itemList, Int32Value tableWidth, Int32Value columnNum, UInt32Value rowHeight, StringValue fontSize)
        {
            Int32Value cellWidth = tableWidth / columnNum;

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
                          new GenerateTableCantSplit().Create(),
                          new GenerateTableRowHeight(rowHeight).Create()
                          )
                      );
            Run itemNumber = new GenerateRun().Create(
                new GenerateRunProperties().Create(
                    new GenerateRunFonts().Create(),
                    new Color() { Val = "C0C0C0" },
                    new GenerateFontSize().Create(fontSize)
                    ),
                new GenerateText().Create(itemNo + "")
                );

            Paragraph paragraphNumber = new GenerateParagraph().Create(
                       new GenerateParagraphProperties().Create(
                           new GenerateJustification().Create(JustificationValues.Center) // 分散对齐
                           )
                       );
            paragraphNumber.Append(itemNumber);
            tableRow.Append(
               new GenerateTableCell().Create(
                   new GenerateTableCellProperties().Create(
                       new GenerateTableCellWidth(cellWidth).Create(),
                       new GenerateTableCellBorders().Create(),
                       new GenerateTableCellVerticalAlignment().Create()
                       ),
                   paragraphNumber
                   )
               );
            foreach (string item in itemList)
            {
                Run runLeftSpace = new GenerateRun().Create(
                                  new GenerateRunProperties().Create(
                                      new GenerateRunFonts().Create(),
                                      new GenerateFontSize().Create(fontSize)
                                      ),
                                  new GenerateText().CreateSpace()
                                  );
                Run runRightSpace = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().CreateSpace()
                                   );
                Run runLeft = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create("[")
                                   );
                Run runRight = new GenerateRun().Create(
                                  new GenerateRunProperties().Create(
                                      new GenerateRunFonts().Create(),
                                      new GenerateFontSize().Create(fontSize)
                                      ),
                                  new GenerateText().Create("]")
                                  );
                Run runContent = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new Color() { Val = "C0C0C0" },
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create(item)
                                   );

                Paragraph paragraph = new GenerateParagraph().Create(
                           new GenerateParagraphProperties().Create(
                               new GenerateJustification().Create(JustificationValues.Center) // 分散对齐
                               )
                           );
                paragraph.Append(runLeft);
                paragraph.Append(runLeftSpace);
                paragraph.Append(runContent);
                paragraph.Append(runRightSpace);
                paragraph.Append(runRight);
                tableRow.Append(
                   new GenerateTableCell().Create(
                       new GenerateTableCellProperties().Create(
                           new GenerateTableCellWidth(cellWidth).Create(),
                           new GenerateTableCellBorders().Create(),
                           new GenerateTableCellVerticalAlignment().Create()
                           ),
                       paragraph
                       )
                   );
            }
            return tableRow;
        }

        internal TableRow CreateRowItem(int i, int tableWidth, int columnNum, UInt32Value rowHeight, StringValue fontSize, JustificationValues justification)
        {
            StringValue title = i + "";
            return CreateRowItem(title, tableWidth, columnNum, rowHeight, fontSize, justification);
        }


        public TableRow CreateRowObjectiveTitle(int itemNo, Int32Value tableWidth, Int32Value columnNum, UInt32Value rowHeight, StringValue fontSize)
        {
            Int32Value cellWidth = tableWidth / columnNum;

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
                       new GenerateTableCantSplit().Create(),
                       new GenerateTableRowHeight(rowHeight).Create()
                       )
                   );
            for (int i = 0; i < columnNum; i++)
            {
                string title = itemNo++ + "";
                Run runContent = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create(title)
                                   );

                Paragraph paragraph = new GenerateParagraph().Create(
                           new GenerateParagraphProperties().Create(
                               new GenerateJustification().Create(JustificationValues.Center) // 分散对齐
                               )
                           );
                paragraph.Append(runContent);
                tableRow.Append(
                   new GenerateTableCell().Create(
                       new GenerateTableCellProperties().Create(
                           new GenerateTableCellWidth(cellWidth).Create(),
                           new GenerateTableCellBorders().Create(),
                           new GenerateTableCellVerticalAlignment().Create()
                           ),
                       paragraph
                       )
                   );
            }
            return tableRow;
        }
        public TableRow CreateRowObjectiveItem(StringValue title, Int32Value tableWidth, Int32Value columnNum, UInt32Value rowHeight, StringValue fontSize, JustificationValues justification)
        {
            Int32Value cellWidth = tableWidth / columnNum;

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
                       new GenerateTableCantSplit().Create(),
                       new GenerateTableRowHeight(rowHeight).Create()
                       )
                   );
            for (int i = 0; i < columnNum; i++)
            {
                Run runLeftSpace = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().CreateSpace()
                                   );
                Run runRightSpace = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().CreateSpace()
                                   );
                Run runLeft = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create("[")
                                   );
                Run runRight = new GenerateRun().Create(
                                  new GenerateRunProperties().Create(
                                      new GenerateRunFonts().Create(),
                                      new GenerateFontSize().Create(fontSize)
                                      ),
                                  new GenerateText().Create("]")
                                  );
                Run runContent = new GenerateRun().Create(
                                   new GenerateRunProperties().Create(
                                       new GenerateRunFonts().Create(),
                                       new Color() { Val = "C0C0C0" },
                                       new GenerateFontSize().Create(fontSize)
                                       ),
                                   new GenerateText().Create(title)
                                   );

                Paragraph paragraph = new GenerateParagraph().Create(
                           new GenerateParagraphProperties().Create(
                               new GenerateJustification().Create(justification) // 分散对齐
                               )
                           );
                if (title.Equals(""))
                {
                    paragraph.Append(runContent);
                }
                else
                {
                    paragraph.Append(runLeft);
                    paragraph.Append(runLeftSpace);
                    paragraph.Append(runContent);
                    paragraph.Append(runRightSpace);
                    paragraph.Append(runRight);
                }
                tableRow.Append(
                   new GenerateTableCell().Create(
                       new GenerateTableCellProperties().Create(
                           new GenerateTableCellWidth(cellWidth).Create(),
                           new GenerateTableCellBorders().Create(),
                           new GenerateTableCellVerticalAlignment().Create()
                           ),
                       paragraph
                       )
                   );
            }
            return tableRow;
        }
    }
}
