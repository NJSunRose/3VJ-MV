using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SWJXMLToCSV
{
    public class RegexMatchHelper
    {
        /// <summary>
        /// 如果输入字符串不是科学计数法的数值，则直接返回输入字符串,否则，将输入字符串转化为Double类型的数值字符串
        /// </summary>
        /// <param name="InputText">输入字符串</param>
        /// <returns>输出结果</returns>
        public static string convert_e_to_double(string InputText)
        {
            if (string.IsNullOrEmpty(InputText))
                return "";
            if (Regex.IsMatch(InputText, @"^[+-]?(?!0\d)\d+(\.\d+)?(E-?\d+)?$"))
            {
                Regex reg = new Regex(@"^[+-]?(?!0\d)\d+(\.\d+)?(E-?\d+)?$", RegexOptions.IgnoreCase);
                string text_new = InputText;
                MatchCollection mac = reg.Matches(text_new);
                foreach (Match m in mac)
                { 
                    string v0 = m.Groups[0].Value;// 转化为非科学计数法
                    double vv = Convert.ToDouble(v0); // 格式转换为非科学计数法 0.00000000                 
                    text_new = vv.ToString();
                }            
                return text_new;
            }
            else
                return InputText;
        }

        /// <summary>
        /// 如果输入字符串不是科学计数法的数值，则直接返回输入字符串,否则，将输入字符串转化为小数点后4位有效数字的Double类型的数值字符串
        /// </summary>
        /// <param name="InputText">输入字符串</param>
        /// <returns>输出结果</returns>
        public static string convert_e_to_double_more(string InputText)
        {
            if (string.IsNullOrEmpty(InputText))
                return "";
            if (Regex.IsMatch(InputText, @"^[+-]?(?!0\d)\d+(\.\d+)?(E-?\d+)?$"))
            {
                Regex reg = new Regex(@"^[+-]?(?!0\d)\d+(\.\d+)?(E-?\d+)?$", RegexOptions.IgnoreCase);
                string text_new = InputText;
                MatchCollection mac = reg.Matches(text_new);
                foreach (Match m in mac)
                {
                    string v0 = m.Groups[0].Value;// 转化为非科学计数法
                    double vv = Convert.ToDouble(v0); // 格式转换为非科学计数法 0.00000000                 
                    text_new = text_new.Replace(v0, vv.ToString());
                    //text_new = string.Format("{0:f4}", vv);
                }
                return text_new;
            }
            else
                return InputText;
        }
    }
}
