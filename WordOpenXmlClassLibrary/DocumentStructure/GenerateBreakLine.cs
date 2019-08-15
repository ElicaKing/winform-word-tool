using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    public class GenerateBreakLine
    {
        /// <summary>
        /// 创建空行
        /// </summary>
        public GenerateBreakLine()
        {
        }

        /// <summary>
        /// 创建空行
        /// </summary>
        /// <returns></returns>
        public Paragraph Create()
        {
            Paragraph paragraph = new GenerateParagraph().Create(
                new GenerateRun().Create(
                    new GenerateRunProperties().Create(
                        new GenerateBold().Create(),
                        new GenerateRunFonts().Create(),
                        new GenerateFontSize().Create()),
                    new GenerateText().Create(" ")
                    )
                );
            return paragraph;
        }
    }
}
