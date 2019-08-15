using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordOpenXmlClassLibrary.Utils
{
    /// <summary>
    /// 文件路径工具类
    /// </summary>
    public static class FilePathUtil
    {

        public static string DocumentFilePath()
        {
            string dateTime = DateTimeUtil.DateTimeToTimeStamp(DateTime.Now).ToString();
            string fileName = dateTime + ".docx";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string filePath = desktopPath + "\\" + fileName;
            return filePath;
        }
    }
}
