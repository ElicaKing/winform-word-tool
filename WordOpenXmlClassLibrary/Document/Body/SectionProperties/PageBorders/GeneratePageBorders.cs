using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
namespace WordOpenXmlClassLibrary
{
    public class GeneratePageBorders
    {
        public GeneratePageBorders()
        {
        }
        // Creates an PageBorders instance and adds its children.
        public PageBorders Create()
        {
            PageBorders pageBorders = new PageBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)10U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)10U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)10U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)10U };

            pageBorders.Append(topBorder);
            pageBorders.Append(leftBorder);
            pageBorders.Append(bottomBorder);
            pageBorders.Append(rightBorder);
            return pageBorders;
        }


    }
}
