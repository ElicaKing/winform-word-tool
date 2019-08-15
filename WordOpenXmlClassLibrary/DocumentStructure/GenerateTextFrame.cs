using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Utils;

namespace WordOpenXmlClassLibrary
{
    public class GenerateTextFrame
    {
        string stroke = "2pt";
        int height = 28;
        int line = 0;

        string width = "485.3pt";
        bool bold = false;

        /// <summary>
        /// 文本框
        /// </summary>
        public GenerateTextFrame()
        {

        }

        /// <summary>
        /// 文本框
        /// </summary>
        /// <param name="stroke">边框宽度</param>
        public GenerateTextFrame(string stroke)
        {
            this.stroke = stroke ?? throw new ArgumentNullException(nameof(stroke));
        }


        /// <summary>
        /// 文本框
        /// </summary>
        /// <param name="line">空行</param>
        public GenerateTextFrame(int line)
        {
            this.line = line;
        }

        /// <summary>
        /// 文本框
        /// </summary>
        /// <param name="stroke">边框宽度</param>
        /// <param name="line">空行数</param>
        public GenerateTextFrame(string stroke, int line) : this(stroke)
        {
            this.line = line;
        }

        public GenerateTextFrame(string stroke, string width) : this(stroke)
        {
            this.width = width ?? throw new ArgumentNullException(nameof(width));
        }

        public GenerateTextFrame(string stroke, string width, bool bold) : this(stroke, width)
        {
            this.bold = bold;
        }

        public Paragraph Create()
        {
            return this.Create("\n");
        }
        public Paragraph Create(string text)
        {
            string[] striparr = text.Split(new string[] { "\n" }, StringSplitOptions.None);
            striparr = striparr.Where(s => !string.IsNullOrEmpty(s)).ToArray();
            int i = striparr.Length + line;

            TextBoxContent textBoxContent = new TextBoxContent();
            textBoxContent.Append(new GenerateBreakLine().Create());
            foreach (string str in striparr)
            {
                int byteLen = WordLengthUtil.getByteLength(str);
                if (byteLen > 48)
                {
                    i++;
                }
                if (bold)
                {
                    textBoxContent.Append(new GenerateTextItem().CreateBold(str));
                }
                else
                {
                    textBoxContent.Append(new GenerateTextItem().Create(str));
                }
                textBoxContent.Append(new GenerateBreakLine().Create());
            }

            if (i <= 1)
            {
                i++;
            }
            height = i * height;
            string Height = height + "pt";


            Paragraph paragraph = new GenerateParagraph().Create(
                new GenerateRun().Create(
                    new GeneratePicture().Create(
                        new GenerateShape(width, Height).Create(
                            new GeneratePath().Create(),
                            new GenerateFill().Create(),
                            new GenerateStroke(stroke).Create(),
                            new GenerateImageData().Create(),
                            new GenerateLock().Create(),
                            new GenerateTextBox().Create(textBoxContent)
                            )
                        )
                    )
                );
            return paragraph;
        }
    }
}
