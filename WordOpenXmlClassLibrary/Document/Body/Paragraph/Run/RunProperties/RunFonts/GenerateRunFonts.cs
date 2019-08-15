using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateRunFonts
    {
        public RunFonts Create()
        {
            RunFonts runFonts = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "宋体" };
            return runFonts;
        }
    }
}
