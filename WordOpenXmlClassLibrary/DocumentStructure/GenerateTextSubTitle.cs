using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTextSubTitle
    {
        public GenerateTextSubTitle()
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
                        new GenerateFontSize().Create("21")),
                    new GenerateText().Create(text)
                    )
                );
            return paragraph;
        }

        public Paragraph Create(string text, string fontsize)
        {
            Paragraph paragraph = new GenerateParagraph().Create(
                new GenerateParagraphProperties().Create(
                    new GenerateJustification().Create(JustificationValues.Center)),
                new GenerateRun().Create(
                    new GenerateRunProperties().Create(
                        new GenerateBold().Create(),
                        new GenerateRunFonts().Create(),
                        new GenerateFontSize().Create(fontsize)),
                    new GenerateText().Create(text)
                    )
                );
            return paragraph;
        }
    }
}
