using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

namespace WordOpenXmlClassLibrary
{
    class GenerateTablePositionProperties
    {
        private Int32Value tablePositionX;
        private Int32Value tablePositionY;

        public GenerateTablePositionProperties()
        {
            tablePositionX = 2520;
            tablePositionY = 93;
        }

        public GenerateTablePositionProperties(Int32Value tablePositionX, Int32Value tablePositionY)
        {
            this.tablePositionX = tablePositionX ?? throw new ArgumentNullException(nameof(tablePositionX));
            this.tablePositionY = tablePositionY ?? throw new ArgumentNullException(nameof(tablePositionY));
        }

        // Creates an TablePositionProperties instance and adds its children.
        public TablePositionProperties Create()
        {
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties()
            {
                LeftFromText = 181,
                RightFromText = 181,
                VerticalAnchor = VerticalAnchorValues.Text,
                HorizontalAnchor = HorizontalAnchorValues.Page,
                TablePositionX = tablePositionX,
                TablePositionY = tablePositionY
            };
            return tablePositionProperties1;
        }

    }
}
