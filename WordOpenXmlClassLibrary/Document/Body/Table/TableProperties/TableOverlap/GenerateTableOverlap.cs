using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    class GenerateTableOverlap
    {
        public GenerateTableOverlap()
        {
        }
        // Creates an TableOverlap instance and adds its children.
        public TableOverlap Create()
        {
            TableOverlap tableOverlap = new TableOverlap()
            {
                Val = TableOverlapValues.Never
            };
            return tableOverlap;
        }
    }
}
