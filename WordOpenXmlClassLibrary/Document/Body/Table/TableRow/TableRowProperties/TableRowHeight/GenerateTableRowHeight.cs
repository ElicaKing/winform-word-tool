using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTableRowHeight
    {
        private UInt32Value val;
        private EnumValue<HeightRuleValues> heightType;
        public GenerateTableRowHeight()
        {
            this.val = (UInt32Value)113U;
            this.heightType = HeightRuleValues.AtLeast;
        }

        public GenerateTableRowHeight(UInt32Value val)
        {
            this.val = val ?? throw new ArgumentNullException(nameof(val));
            this.heightType = HeightRuleValues.AtLeast;
        }

        // Creates an TableRowHeight instance and adds its children.
        public TableRowHeight Create()
        {
            TableRowHeight tableRowHeight = new TableRowHeight()
            {
                Val = val,
                HeightType = heightType
            };
            return tableRowHeight;
        }

    }
}
