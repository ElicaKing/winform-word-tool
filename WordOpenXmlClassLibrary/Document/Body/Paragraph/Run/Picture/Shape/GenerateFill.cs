using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Vml;

namespace WordOpenXmlClassLibrary
{
    public class GenerateFill
    {
        // Creates an Fill instance and adds its children.
        public Fill Create()
        {
            Fill fill = new Fill() { On = true, FocusSize = "0,0" };
            return fill;
        }
    }
}
