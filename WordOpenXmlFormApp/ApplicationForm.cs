using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpenXmlClassLibrary;
using WordOpenXmlClassLibrary.Enum;
using WordOpenXmlClassLibrary.Utils;

namespace WordOpenXmlFormApp
{
    public partial class ApplicationForm : Form
    {
        public ApplicationForm()
        {
            InitializeComponent();
            InitializeDefault();
        }

        private void ButtonExportFile_Click(object sender, EventArgs e)
        {
            string filePath = FilePathUtil.DocumentFilePath();
            string examName = textBoxExamName.Text;
            int pageSize = comboBoxPageSize.SelectedIndex;
            int pageOrientation = comboBoxPageOrientation.SelectedIndex;

            GeneratedClass generater = new GeneratedClass();

            generater.FilePath = filePath;
            generater.ExamName = examName;

            switch (pageSize)
            {
                case (int)PageSizeValues.A3:
                    generater.PageSize = PageSizeValues.A3;
                    break;
                case (int)PageSizeValues.A4:
                    generater.PageSize = PageSizeValues.A4;
                    break;
            }
            switch (pageOrientation)
            {
                case (int)PageOrientationValues.Landscape:
                    generater.PageOrientation = PageOrientationValues.Landscape;
                    break;
                case (int)PageOrientationValues.Portrait:
                    generater.PageOrientation = PageOrientationValues.Portrait;
                    break;
            }

            generater.Create();

            toolStripStatusLabelTip.Text = filePath;
            MessageBox.Show(filePath, "导出文件");
        }

        private void ButtonExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否确定退出?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                System.Environment.Exit(0);
            }
        }

        private void ComboBoxPageSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxPageSize.SelectedIndex)
            {
                case (int)PageSizeValues.A3:
                    comboBoxPageOrientation.SelectedIndex = (int)PageOrientationValues.Portrait;
                    break;
                case (int)PageSizeValues.A4:
                    comboBoxPageOrientation.SelectedIndex = (int)PageOrientationValues.Landscape;
                    break;
            }
        }
    }
}
