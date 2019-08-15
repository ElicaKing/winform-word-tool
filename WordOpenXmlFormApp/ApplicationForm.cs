using System;
using System.Diagnostics;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.DocViewer.Forms;
using WordOpenXmlClassLibrary;
using WordOpenXmlClassLibrary.Enum;
using WordOpenXmlClassLibrary.Utils;

namespace WordOpenXmlFormApp
{
    public partial class ApplicationForm : Form
    {
        private int ZoomNum = 100;
        public ApplicationForm()
        {
            InitializeComponent();
            InitializeDefault();
        }

        private void ButtonPreviewFile_Click(object sender, EventArgs e)
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

            this.toolStripStatusLabelTip.Text = filePath;
            string filepath = toolStripStatusLabelTip.Text;
            LoadFromViewFile(filepath);

            toolStripProgressBar.Maximum = 100;//设置最大长度值
            toolStripProgressBar.Value = 0;//设置当前值
            toolStripProgressBar.Step = 20;//设置没次增长多少
            toolStripProgressBar.Visible = true;
            for (int i = 0; i < 5; i++)//循环
            {
                System.Threading.Thread.Sleep(100);//暂停1秒
                toolStripProgressBar.Value += toolStripProgressBar.Step; //让进度条增加一次
            }

            if (MessageBox.Show("文件已生成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question) == DialogResult.OK)
            {
                this.tabControl.SelectedIndex = 1;
                toolStripProgressBar.Visible = false;
            }

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

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (this.tabControl.SelectedIndex)
            {
                case 0:
                    break;
                case 1:
                    //string filepath = toolStripStatusLabelTip.Text;
                    //LoadFromViewFile(filepath);
                    break;
            }

        }

        private void LoadFromViewFile(string filepath)
        {
            try
            {
                this.ZoomNum = 100;
                // Load doc document from file 
                this.docDocumentViewer.LoadFromFile(filepath);
                this.docDocumentViewer.ZoomTo(ZoomNum);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ButtonZoomToDown_Click(object sender, EventArgs e)
        {
            if (ZoomNum > 50)
            {
                ZoomNum = ZoomNum - 5;
                this.docDocumentViewer.ZoomTo(ZoomNum);
            }
            else
            {
                MessageBox.Show("已缩放到最小", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void ButtonZoomToUp_Click(object sender, EventArgs e)
        {
            if (ZoomNum <= 120)
            {
                ZoomNum = ZoomNum + 5;
                this.docDocumentViewer.ZoomTo(ZoomNum);
            }
            else
            {
                MessageBox.Show("已缩放到最大", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ButtonExportFile_Click(object sender, EventArgs e)
        {
            string filepath = toolStripStatusLabelTip.Text;
            Process.Start("explorer.exe", filepath);
        }
    }
}
