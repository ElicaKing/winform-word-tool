using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateJustification
    {
        public GenerateJustification()
        {
        }

        public Justification Create()
        {
            return this.Create(JustificationValues.Center);
        }
        // Creates an Justification instance and adds its children.
        public Justification Create(EnumValue<JustificationValues> JcVal)
        {
            Justification justification = new Justification()
            {
                Val = JcVal
            };
            return justification;
        }

    }
}
