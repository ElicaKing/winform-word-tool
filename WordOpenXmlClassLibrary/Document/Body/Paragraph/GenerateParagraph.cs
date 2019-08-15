using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordOpenXmlClassLibrary
{
    public class GenerateParagraph
    {
        // Creates an Paragraph instance and adds its children.
        public Paragraph Create(params OpenXmlElement[] newChildren)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Append(newChildren);
            return paragraph;
        }

        // Creates an Paragraph instance and adds its children.
        public Paragraph CreateCellParagraph()
        {
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            KeepNext keepNext = new KeepNext() { Val = false };
            KeepLines keepLines = new KeepLines() { Val = false };
            PageBreakBefore pageBreakBefore = new PageBreakBefore() { Val = false };
            WidowControl widowControl = new WidowControl() { Val = false };
            Kinsoku kinsoku = new Kinsoku();
            WordWrap wordWrap = new WordWrap();
            OverflowPunctuation overflowPunctuation = new OverflowPunctuation();
            TopLinePunctuation topLinePunctuation = new TopLinePunctuation() { Val = false };
            AutoSpaceDE autoSpaceDE = new AutoSpaceDE();
            AutoSpaceDN autoSpaceDN = new AutoSpaceDN();
            BiDi biDi = new BiDi() { Val = false };
            AdjustRightIndent adjustRightIndent = new AdjustRightIndent();
            SnapToGrid snapToGrid = new SnapToGrid();
            SpacingBetweenLines spacingBetweenLines = new SpacingBetweenLines() { Line = "113", LineRule = LineSpacingRuleValues.Exact };
            TextAlignment textAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Auto };

            paragraphProperties.Append(keepNext);
            paragraphProperties.Append(keepLines);
            paragraphProperties.Append(pageBreakBefore);
            paragraphProperties.Append(widowControl);
            paragraphProperties.Append(kinsoku);
            paragraphProperties.Append(wordWrap);
            paragraphProperties.Append(overflowPunctuation);
            paragraphProperties.Append(topLinePunctuation);
            paragraphProperties.Append(autoSpaceDE);
            paragraphProperties.Append(autoSpaceDN);
            paragraphProperties.Append(biDi);
            paragraphProperties.Append(adjustRightIndent);
            paragraphProperties.Append(snapToGrid);
            paragraphProperties.Append(spacingBetweenLines);
            paragraphProperties.Append(textAlignment);

            paragraph.Append(paragraphProperties);
            return paragraph;
        }

    }
}
