using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableJustification
    {
        public GenerateTableJustification()
        {
        }

        // Creates an TableJustification instance and adds its children.
        public TableJustification Create()
        {
            TableJustification tableJustification = new TableJustification()
            {
                Val = TableRowAlignmentValues.Center
            };
            return tableJustification;
        }

    }
}
