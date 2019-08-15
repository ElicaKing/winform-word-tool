using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTextMainTitle
    {
        public GenerateTextMainTitle()
        {

        }
        public Paragraph Create(string text)
        {
            Paragraph paragraph = new GenerateParagraph().Create(
                new GenerateParagraphProperties().Create(
                    new GenerateJustification().Create(JustificationValues.Center)),
                new GenerateRun().Create(
                    new GenerateRunProperties().Create(
                        new GenerateBold().Create(),
                        new GenerateRunFonts().Create(),
                        new GenerateFontSize().Create("28")),
                    new GenerateText().Create(text)
                    )
                );
            return paragraph;
        }
    }
}
