using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary.Enum;

namespace WordOpenXmlClassLibrary
{
    public class GeneratedClass
    {
        private string filePath;
        private string examName;
        private PageSizeValues pageSize;
        private PageOrientationValues pageOrientation;

        public string FilePath { get => filePath; set => filePath = value; }
        public string ExamName { get => examName; set => examName = value; }
        public PageSizeValues PageSize { get => pageSize; set => pageSize = value; }
        public PageOrientationValues PageOrientation { get => pageOrientation; set => pageOrientation = value; }

        public void Create()
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(FilePath, WordprocessingDocumentType.Document))
            {
                int writingRowNum = 30;

                // 创建基础文档结构
                Generater generater = new Generater(wordprocessingDocument);
                generater.ExamTitle(ExamName);
                generater.ExaminationNo();
                generater.StudentInfo();

                generater.ModuleTitle("[一]积累运用（46分）");
                generater.HeaderTitle("一、读拼音，把词语写端正。（每个字0.5分，共4分）");
                generater.ObjectiveItem(new string[] { "A", "B", "C" }, 6, 1, PageOrientationValues.Portrait);

                generater.HeaderTitle("二、查字填空。（每空0.5分，共3分）");
                generater.SubjectiveItem("1、____________  ____________  ____________ 2、____________  ____________  ____________\n", 0, 0);

                generater.HeaderTitle("三、根据多音字的读音，选择正确的组织。（每小题1分，共6分）");
                generater.ObjectiveItem(new string[] { "A", "B", "C" }, 6, 1, PageOrientationValues.Portrait);

                generater.HeaderTitle("四、下面的词语，正确的填“A”，错误的填“B”。（每小题0.5分，共4分）");
                generater.ObjectiveItem(new string[] { "A", "B" }, 8, 1, PageOrientationValues.Portrait);

                generater.HeaderTitle("五、选词填空。（每小题1分，共5分）");
                generater.ObjectiveItem(new string[] { "A", "B", "C", "D", "E" }, 5, 1, PageOrientationValues.Landscape);

                generater.HeaderTitle("六、下面的句子运用了什么描写方法？。（每小题1分，共4分）");
                generater.ObjectiveItem(new string[] { "A", "B", "C", "D" }, 4, 1, PageOrientationValues.Landscape);

                generater.HeaderTitle("七、先解释带点字，再写出句子的意思。（每小题3分，共6分）");
                generater.SubjectiveItem("" +
                    "1、潜：____________________\n" +
                    "句子意思：____________________________________________________________\n" +
                    "2、善哉：____________________\n" +
                    "句子意思：____________________________________________________________\n", 0, 1);

                generater.HeaderTitle("八、按要求完成句子。（每句2分，共8分）");
                generater.SubjectiveItem("" +
                    "1、冬天，寒风呼啸着拂面而来，吹得人瑟瑟发抖\n" +
                    "2、____________________________________________________________\n" +
                    "3、____________________________________________________________\n" +
                    "4、____________________________________________________________\n" +
                    "5、____________________________________________________________\n", 0, 1);

                generater.HeaderTitle("九、根据课文内容填空。（每空1分，共6分）");
                generater.SubjectiveItem("" +
                    "1、____________________ ____________________ 2、____________________ ____________________\n" +
                    "2、____________________ ____________________", 0, 1);

                generater.ModuleTitle("[二]阅读冲浪（24分）");

                generater.HeaderTitle("一、阅读文段，完成练习");
                generater.SubjectiveItem("" +
                    "1、（     ）（     ）（2分）     2、____________________ ____________________\n" +
                    "3、____________________ ____________________（3分）  4、____________________", 0, 1);

                generater.HeaderTitle("二、阅读短文，回答问题。（14分）");
                generater.SubjectiveItem("1、____________  ____________  ____________ 2、____________  ____________  ____________\n", 0, 0);


                generater.ModuleTitle("[三]快乐作文（30分）");
                generater.SubjectiveItem("" +
                    "题目：一次_________的合作\n" +
                    "提示：你有过和他人合作的经历吗？选择一个合作成功或者合作失败的示例写下来，要写清楚和谁合作，怎么合作，结果怎样。要求语句通顺连贯，感情真实自然。\n" +
                    "评分要求：能把合作的经过写清楚、写具体，语句通顺连贯，感情真实自然。分三等评分：一等24~30分；二等18~23分；三等18分以下。", 0, 0);
                generater.CreateWriting(writingRowNum);

                generater.CreateEnglishWriting(writingRowNum);

                //// 设置页面大小：默认A4
                //generater.SectionProperties(PageSizeValues.A4, PageOrientationValues.Landscape);
                //generater.SectionProperties(PageSizeValues.A3, PageOrientationValues.Portrait);
                generater.SectionProperties(PageSize, PageOrientation);
                // 保存文档
                generater.Save();
            }
        }
    }
}
