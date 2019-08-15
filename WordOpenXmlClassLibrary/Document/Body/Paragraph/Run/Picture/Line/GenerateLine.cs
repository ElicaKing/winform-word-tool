using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace WordOpenXmlClassLibrary
{
    public class GenerateLine
    {
        public GenerateLine()
        {
        }

        // Creates an Line instance and adds its children.
        public Line Create()
        {
            Line line = new Line()
            {
                Style = "" +
                "position:absolute;" +
                "left:0pt;" +
                "margin-left:-16.95pt;" +
                "margin-top:-16.4pt;" +
                "height:0.05pt;" +
                "width:526.5pt;" +
                "z-index:251658240;" +
                "mso-width-relative:page;" +
                "mso-height-relative:page;",
                Filled = false,
                Stroked = true,
                OptionalNumber = 20
            };
            Path path = new Path() { ShowArrowhead = true };
            Fill fill = new Fill() { On = false, FocusSize = "0,0" };
            Stroke stroke = new Stroke() { Color = "#000000" };
            ImageData imageData = new ImageData() { Title = "" };
            Ovml.Lock @lock = new Ovml.Lock() { Extension = ExtensionHandlingBehaviorValues.Edit, AspectRatio = false };

            line.Append(path);
            line.Append(fill);
            line.Append(stroke);
            line.Append(imageData);
            line.Append(@lock);
            return line;
        }
    }
}
