using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlConverter
{
    /// <summary>
    /// 区域字体
    /// </summary>
    public class RegionFont
    {
        public string Ascii;
        public string HighAnsi;
        public string EastAsia;

        public static RegionFont Parse(string str)
        {
            if (str == null || str.Trim().Length == 0)
            {
                return null;
            }

            RegionFont font = new RegionFont();
            String[] names = str.Split(new[] {','}, StringSplitOptions.RemoveEmptyEntries);
            List<String> fontFamily = new List<String>();
            for (int i = 0; i < names.Length; i++)
            {
                String fontName = names[i].Trim();
                try
                {
                    if (fontName[0] == '\'' && fontName[fontName.Length - 1] == '\'')
                    {
                        fontName = fontName.Substring(1, fontName.Length - 2);
                    }

                    fontFamily.Add(fontName);
                    if (i > 2)
                    {
                        break;
                    }
                }
                catch (ArgumentException)
                {
                    // the name is not a TrueType font or is not a font installed on this computer
                }
            }

            if (fontFamily.Count == 1)
            {
                font.EastAsia = fontFamily[0];
                font.HighAnsi = fontFamily[0];
                font.Ascii = fontFamily[0];
            }
            else if (fontFamily.Count == 2)
            {
                font.EastAsia = fontFamily[0];
                font.HighAnsi = fontFamily[1];
                font.Ascii = fontFamily[1];
            }
            else if (fontFamily.Count == 3)
            {
                font.EastAsia = fontFamily[0];
                font.HighAnsi = fontFamily[1];
                font.Ascii = fontFamily[2];
            }

            return font;
        }
    }
}