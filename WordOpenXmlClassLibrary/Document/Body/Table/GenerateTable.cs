using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTable
    {
        public GenerateTable()
        {
        }
        //Creates an Table instance and adds its children.
        public Table Create(params OpenXmlElement[] newChildren)
        {
            Table table = new Table();
            table.Append(newChildren);
            return table;
        }
    }
}
