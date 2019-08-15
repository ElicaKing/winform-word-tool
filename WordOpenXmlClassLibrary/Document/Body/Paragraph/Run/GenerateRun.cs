using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateRun
    {
        // Creates an Paragraph instance and adds its children.
        public Run Create(params OpenXmlElement[] newChildren)
        {
            Run run = new Run();
            run.Append(newChildren);
            return run;
        }
    }
}
