using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelMyPlugin
{
    public class Utility
    {
        // 判断字符串是否包含中文
        public static bool IsContainChinese(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return false;
            }

            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]");
        }
    }
}
