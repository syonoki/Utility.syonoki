using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Utility.syonoki.ExtensionMethods
{
    public static class StringExtension
    {
        #region string conversion extension
        public static string replace(this string s, string[,] criteria)
        {

            if (criteria.GetLength(1) != 2)
            {
                throw new ArgumentException($"{criteria.GetLength(1)} sized column array provided \r\n" +
                                            "replacement criteria array must be 2 dimension column size");
            }

            for (int i = 0; i < criteria.GetLength(0); i++)
            {
                s= s.Replace(criteria[i, 0], criteria[i, 1]);
            }

            return s;
        }
        public static string removeGap(this string s)
        {
            return s.Replace(" ", "");
        }

        public static string remove(this string s, string removingTarget)
        {
            return s.Replace(removingTarget, "");
        }

        public static IEnumerable<string> splitCommaSeparatedString(this string s)
        {
            return s.removeGap().Split(',');
        }
        #endregion

        #region string extraction
        public static string between(this string txt, string s1, string s2)
        {
            string[] spt = txt.Split(new string[] { s1 }, StringSplitOptions.None);

            if (spt.Length == 1)
                return "해당형식없음";
            spt = spt[1].Split(new string[] { s2 }, StringSplitOptions.None);
            return spt[0];
        }

        public static string betweenR(this string txt, string s1, string s2)
        {
            string[] spt = txt.Split(new string[] { s1 }, StringSplitOptions.None);

            if (spt.Length == 1)
                return "해당형식없음";
            spt = spt[0].Split(new string[] { s2 }, StringSplitOptions.None);
            return spt.Last();

        }
        #endregion

        #region double conversion
        public static double toDoubleWithRemovingNonNumericValue(this string txt)
        {
            return Convert.ToDouble(Regex.Replace(txt, @"[^\d\.]", ""));
        }

        public static string zeroThanNull(this string txt)
        {
            return txt == "0" ? "" : txt;
        }
        #endregion

        #region save as txt
        public static void saveAsTxt(this string txt, string path)
        {
            if(File.Exists(path))
                File.Delete(path);
            File.WriteAllText(path, txt);    
        }
        #endregion

        public static string stringArrayFlating(this IEnumerable<string> stringArray)
        {
            string retVal = String.Empty;
            foreach (var s in stringArray)
                retVal += s + "\r\n";
            
            return retVal;
        }
    }
}
