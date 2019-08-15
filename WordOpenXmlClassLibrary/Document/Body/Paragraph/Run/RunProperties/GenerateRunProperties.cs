using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateRunProperties
    {
        // Creates an RunProperties instance and adds its children.
        public RunProperties Create(params OpenXmlElement[] newChildren)
        {
            RunProperties runProperties = new RunProperties();
            runProperties.Append(newChildren);
            return runProperties;
        }
    }
}
