using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTextTitle
    {
        /// <summary>
        /// 创建文本标题
        /// </summary>
        public GenerateTextTitle()
        {
        }
        /// <summary>
        /// 创建文本标题
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph Create(string text)
        {
            Paragraph paragraph = new GenerateParagraph().Create(
                new GenerateRun().Create(
                    new GenerateRunProperties().Create(
                        new GenerateBold().Create(),
                        new GenerateRunFonts().Create(),
                        new GenerateFontSize().Create()),
                    new GenerateText().Create(text)
                    )
                );
            return paragraph;
        }
    }
}
