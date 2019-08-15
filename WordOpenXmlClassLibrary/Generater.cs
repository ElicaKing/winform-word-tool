using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    /// <summary>
    /// 主构造器
    /// </summary>
    public class Generater
    {
        private Body body;
        private WordprocessingDocument wordprocessingDocument;

        public Generater(WordprocessingDocument document)
        {
            this.wordprocessingDocument = document;
            this.Init();
        }

        /// <summary>
        ///  初始化文档流
        /// </summary>
        private void Init()
        {
            MainDocumentPart mainDocumentPart = wordprocessingDocument.AddMainDocumentPart();
            mainDocumentPart.Document = new Document(new Body());
            this.body = mainDocumentPart.Document.Body;
        }

        /// <summary>
        /// 考号填涂区域
        /// </summary>
        internal void ExaminationNo()
        {
            int tableWidth = 5733; // 表格宽度
            int columnNum = 9;
            int rowNum = 10;
            UInt32Value rowHeight = 315;
            StringValue fontSize = "18";

            int positionX = 5070;
            int positionY = 36;

            Table table = new GenerateTable().Create(
                new GenerateTableProperties().Create(
                    new GenerateTableStyle().Create(),
                    new GenerateTablePositionProperties(positionX, positionY).Create(),
                    new GenerateTableOverlap().Create(),
                    new GenerateTableWidth(tableWidth).Create(),
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
                    ),
                new GenerateTableGrids().Create(tableWidth, columnNum)
                );

            // 首行
            table.Append(new GenerateTableExamineeNoFrame().CreateRowHeader("考   号", tableWidth, columnNum, rowHeight, fontSize));
            table.Append(new GenerateTableExamineeNoFrame().CreateRowItem("", tableWidth, columnNum, rowHeight, fontSize, JustificationValues.Distribute));
            for (int i = 0; i < rowNum; i++)
            {
                table.Append(new GenerateTableExamineeNoFrame().CreateRowItem(i, tableWidth, columnNum, rowHeight, fontSize, JustificationValues.Distribute));
            }

            this.body.Append(table);
        }

        /// <summary>
        /// 客观题
        /// </summary>
        /// <param name="borderSize">边框大小</param>
        /// <param name="extendLine">附加行</param>
        public void ObjectiveItem(int borderSize, int extendLine)
        {
            StringValue border = borderSize + "pt";
            this.body.Append(new GenerateTextFrame(border, extendLine).Create(""));
            this.body.Append(new GenerateBreakLine().Create());
        }

        public void ObjectiveItem(string[] itemList, int columnNum, int itemNo, PageOrientationValues orientation)
        {
            //string[] itemList = { "A", "B", "C", "D", "E" };//答案选项
            //int columnNum = 6;//题目数
            int rowNum = columnNum % 9 > 0 ? (columnNum / 9) + 1 : columnNum / 9;
            //int itemNo = 1;//题号起始

            int columnWidth = 1020;//单元格宽度
            UInt32Value rowHeight = 340;//单元格高度
            StringValue fontSize = "18";



            Table table = new Table();
            int tableWidth = columnWidth * columnNum;

            switch (orientation)
            {
                case PageOrientationValues.Landscape://风景模式、横向

                    columnNum = itemList.Length + 1;
                    tableWidth = columnWidth * columnNum;
                    for (int i = 0; i < columnNum; i++)
                    {
                        table.Append(new GenerateTableExamineeNoFrame().CreateRowObjectiveLandscape(itemNo, itemList, tableWidth, columnNum, rowHeight, fontSize));
                        itemNo++;
                    }
                    break;
                case PageOrientationValues.Portrait://肖像模式、竖向
                    for (int i = 0; i < rowNum; i++)
                    {
                        int colNum = 0;
                        if (columnNum > 9)
                        {
                            colNum = 9;
                        }
                        else
                        {
                            colNum = columnNum;
                        }

                        table.Append(new GenerateTableExamineeNoFrame().CreateRowObjectiveTitle(itemNo, tableWidth, colNum, rowHeight, fontSize));

                        foreach (string item in itemList)
                        {
                            table.Append(new GenerateTableExamineeNoFrame().CreateRowObjectiveItem(item, tableWidth, colNum, rowHeight, fontSize, JustificationValues.Center));
                        }
                        itemNo += 9;
                        columnNum -= 9;
                    }
                    break;
            }

            table.Append(
                  new GenerateTableProperties().Create(tableWidth),
                  new GenerateTableGrids().Create(tableWidth, columnNum)
                  );

            this.body.Append(new GenerateBreakLine().Create());
            this.body.Append(table);
            this.body.Append(new GenerateBreakLine().Create());
        }

        internal void CreateEnglishWriting(int rowNum)
        {

            int tableWidth = 9629; // 表格宽度
            UInt32Value rowHeight = 454; // 表格宽度
            int columnNum = 1; // 列数
            //int rowNum = 30; // 行数
            //int wordCount = 600; // 字数

            Table table = new GenerateTable().Create(
                  new GenerateTableProperties().Create(tableWidth),
                  new GenerateTableGrids().Create(tableWidth, columnNum)
                  );

            for (int i = 0; i < rowNum; i++)
            {
                table.Append(
                    new GenerateTableSeparateRow().CreateSpanRow(tableWidth, rowHeight)
                    );
            }

            this.body.Append(new GenerateBreakLine().Create());
            this.body.Append(new GenerateTextSubTitle().Create("书面表达", "32"));
            this.body.Append(table);
        }

        /// <summary>
        /// 主观题
        /// </summary>
        /// <param name="content">内容</param>
        /// <param name="borderSize">边框大小</param>
        /// <param name="extendLine">附加行</param>
        public void SubjectiveItem(string content, int borderSize, int extendLine)
        {
            if (borderSize == 0)
            {
                string[] striparr = content.Split(new string[] { "\n" }, StringSplitOptions.None);
                foreach (string item in striparr)
                {
                    this.body.Append(new GenerateBreakLine().Create());
                    this.body.Append(new GenerateTextItem().Create(item));
                }
                for (int i = 0; i < extendLine; i++)
                {
                    this.body.Append(new GenerateBreakLine().Create());
                }
            }
            else
            {
                StringValue border = borderSize + "pt";
                this.body.Append(new GenerateTextFrame(border, extendLine).Create(content));
                this.body.Append(new GenerateBreakLine().Create());
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            this.wordprocessingDocument.Save();
        }
        /// <summary>
        /// 考试标题
        /// </summary>
        /// <param name="examName"></param>
        public void ExamTitle(StringValue examName)
        {
            this.body.Append(new GenerateTextMainTitle().Create(examName));
            this.body.Append(new GenerateBreakLine().Create());
        }
        /// <summary>
        /// 信息填涂区域
        /// </summary>
        public void StudentInfo()
        {
            this.body.Append(new GenerateTextFrame("0pt", "175.6pt", true).Create("" +
                    "学校：________________________\n" +
                    "姓名：________________________\n" +
                    "试室：________________________\n" +
                    "座位：________________________\n \n"));
            this.body.Append(new GenerateBreakLine().Create());
            this.body.Append(new GenerateBreakLine().Create());
            this.body.Append(new GenerateBreakLine().Create());
            this.body.Append(new GenerateTextTitle().Create("缺考标记 [   ]"));
            this.body.Append(new GenerateTextTitle().Create("缺考考生，由监考员用2B铅笔填涂"));
            this.body.Append(new GenerateBreakLine().Create());
            this.body.Append(new GenerateBreakLine().Create());
        }
        /// <summary>
        /// 模块标题
        /// </summary>
        /// <param name="title"></param>
        internal void ModuleTitle(string title)
        {
            this.body.Append(new GenerateTextSubTitle().Create(title));
            this.body.Append(new GenerateBreakLine().Create());
        }
        /// <summary>
        /// 大题标题
        /// </summary>
        /// <param name="title"></param>
        public void HeaderTitle(string title)
        {
            this.body.Append(new GenerateTextTitle().Create(title));
        }
        /// <summary>
        /// 创建作文题
        /// </summary>
        /// <param name="rowNum"></param>
        public void CreateWriting(int rowNum)
        {

            int TableWidth = 9629; // 表格宽度
            int ColumnNum = 20; // 列数
            //int rowNum = 30; // 行数
            //int wordCount = 600; // 字数

            Table table = new GenerateTable().Create(
                  new GenerateTableProperties().Create(TableWidth),
                  new GenerateTableGrids().Create(TableWidth, ColumnNum)
                  );

            for (int i = 0; i < rowNum; i++)
            {
                table.Append(
                    new GenerateTableSeparateRow().Create(TableWidth, ColumnNum),
                    new GenerateTablePaneRow().Create(TableWidth, ColumnNum)
                    );
            }
            this.body.Append(table);
        }

        /// <summary>
        /// 页面设置
        /// </summary>
        /// <param name="pageSizeValue"></param>
        /// <param name="pageOrientationValu"></param>
        public void SectionProperties(PageSizeValues pageSizeValue, PageOrientationValues pageOrientationValu)
        {
            this.body.Append(new GenerateSectionProperties().Create(pageSizeValue, pageOrientationValu));
        }
    }
}
