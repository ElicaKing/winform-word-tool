using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordOpenXmlClassLibrary.Utils
{
    public static class WordLengthUtil
    {
        /// <summary>
        /// 返回字节数
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int getByteLength(string s)
        {
            int lh = 0;
            char[] q = s.ToCharArray();
            for (int i = 0; i < q.Length; i++)
            {
                if ((int)q[i] >= 0x4E00 && (int)q[i] <= 0x9FA5) // 汉字
                {
                    lh += 2;
                }
                //else
                //{
                //    lh += 1;
                //}
            }
            return lh;
        }
        /// <summary>
        /// 返回字数
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static int getWordLength(string s)
        {
            if (s != null)
                return s.Length;
            else
                return 0;
        }
        /// <summary>
        /// 返回数字（0~9）字数数量
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static int getdigitalLength(string s)
        {
            int lx = 0;
            char[] q = s.ToCharArray();
            for (int i = 0; i < q.Length; i++)
            {
                if ((int)q[i] >= 48 && (int)q[i] <= 57)
                {
                    lx += 1;
                }
            }
            return lx;
        }
        /// <summary>
        /// 返回字母（A~Z-a~z）字数数量
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static int getcharLength(string s)
        {
            int lz = 0;
            char[] q = s.ToLower().ToCharArray();//大写字母转换成小写字母
            for (int i = 0; i < q.Length; i++)
            {
                if ((int)q[i] >= 97 && (int)q[i] <= 122)//小写字母
                {
                    lz += 1;
                }
            }
            return lz;
        }
    }
}
