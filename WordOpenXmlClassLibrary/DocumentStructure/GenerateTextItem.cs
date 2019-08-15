using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTextItem
    {
        public GenerateTextItem()
        {
        }

        /// <summary>
        /// 创建文本项
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph Create(string text)
        {
            Paragraph paragraph = new GenerateParagraph().Create(
                new GenerateRun().Create(
                    new GenerateRunProperties().Create(
                        //new GenerateBold().Create(),
                        new GenerateRunFonts().Create(),
                        new GenerateFontSize().Create()),
                    new GenerateText().Create(text)
                    )
                );
            return paragraph;
        }

        public Paragraph CreateBold(string text)
        {
            Paragraph paragraphBold = new GenerateParagraph().Create(
                new GenerateRun().Create(
                    new GenerateRunProperties().Create(
                        new GenerateBold().Create(),
                        new GenerateRunFonts().Create(),
                        new GenerateFontSize().Create()),
                    new GenerateText().Create(text)
                    )
                );
            return paragraphBold;
        }
    }
}
