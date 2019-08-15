

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    public class GenerateSectionProperties
    {
        private PageSize pageSize;
        private PageMargin pageMargin;
        private Columns columns; // 720
        private DocGrid docGrid; // 360

        public GenerateSectionProperties()
        {
            this.pageSize = new PageSize();
            this.pageMargin = new PageMargin();
            this.columns = new Columns() { Space = "220" }; // 720
            this.docGrid = new DocGrid() { LinePitch = 100 }; // 360
        }

        // Creates an SectionProperties instance and adds its children.
        public SectionProperties Create(PageSizeValues pageSizeValue, PageOrientationValues pageOrientationValue)
        {
            this.PageSetting(pageSizeValue, pageOrientationValue);
            this.PageColumnsSetting(pageSizeValue, pageOrientationValue);

            SectionProperties sectionProperties = new SectionProperties();
            sectionProperties.Append(
                pageSize,
                pageMargin,
                new GeneratePageBorders().Create(),
                columns,
                docGrid
                );

            return sectionProperties;
        }

        private void PageColumnsSetting(PageSizeValues pageSizeValue, PageOrientationValues pageOrientationValue)
        {
            // 当页面大小为A3、横向时，自动分栏
            if (pageSizeValue.Equals(PageSizeValues.A3) && pageOrientationValue.Equals(PageOrientationValues.Portrait))
            {
                this.columns = new GenerateColumns(false, 2).Create(
                new GenerateColumn("10220", "1292.5").Create(),//10586
                new GenerateColumn("10220").Create()
                );
            }
        }

        private void PageSetting(PageSizeValues pageSizeValue, PageOrientationValues pageOrientationValue)
        {
            UInt32Value width = 11906U;
            UInt32Value height = 16838U;

            UInt32Value left = 1080U;
            UInt32Value right = 1080U;

            int top = 1080;
            int bottom = 1080;


            switch (pageSizeValue)
            {
                case PageSizeValues.A4:
                    width = 11906U;
                    height = 16838U;
                    break;
                case PageSizeValues.A3:
                    width = 16783U;
                    height = 23757U;
                    break;
            }

            switch (pageOrientationValue)
            {
                case PageOrientationValues.Landscape:
                    break;
                case PageOrientationValues.Portrait:
                    UInt32Value sweep = width;
                    width = height;
                    height = sweep;

                    int top_sweep = top;
                    top = (int)left.Value;
                    left = (uint)top_sweep;
                    break;
            }

            pageSize.Width = width;
            pageSize.Height = height;
            pageSize.Orient = pageOrientationValue;

            pageMargin.Top = top;
            pageMargin.Bottom = bottom;
            pageMargin.Left = left;
            pageMargin.Right = right;
            pageMargin.Header = (UInt32Value)720U;
            pageMargin.Footer = (UInt32Value)720U;
            pageMargin.Gutter = (UInt32Value)0U;
        }
    }
}
