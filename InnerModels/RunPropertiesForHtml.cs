using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLWordExtension.InnerModels
{
    public class RunPropertiesForHtml
    {
        public String Bold
        {
            get;
            private set;
        }

        public String FontSize
        {
            get;
            private set;
        }

        public String Highlight
        {
            get;
            private set;
        }

        public String FontFamily
        {
            get;
            private set;
        }

        public String Color
        {
            get;
            private set;
        }

        public Boolean Italic
        {
            get;
            private set;
        }

        public String Underline
        {
            get;
            private set;
        }

        private RunPropertiesForHtml() { }

        public static RunPropertiesForHtml CreateRunPropertiesForHtml(RunProperties rp)
        {
            rp = rp ?? new RunProperties();
            var instance = new RunPropertiesForHtml
            {
                Bold = rp.Bold.GetValue(),
                FontFamily = GetFontFamily(rp.RunFonts),
                FontSize = rp.FontSize.GetValue(),
                Color = rp.Color.GetValue(),
                Highlight = rp.Highlight.GetValue(),
            };
            return instance;
        }

        public String GenerateStyleString()
        {
            String style = Color.GetStyleString("color");
            style += FontSize.GetStyleString("font-size");
            style += FontFamily.GetStyleString("font-family");
            style += Bold.GetStyleString("font-weight");
            style += Highlight.GetStyleString("background-color");
            return style.ToStyleString();
        }

        public static String GetFontFamily(RunFonts rfs)
        {
            var fontfamily = "";
            if (rfs != null)
            {
                Func<StringValue, String> GetStringValueVal = (sv) => sv == null ? "" : ("'" + sv.Value + "', ");
                fontfamily += GetStringValueVal(rfs.Ascii) + GetStringValueVal(rfs.HighAnsi) + GetStringValueVal(rfs.EastAsia);
                fontfamily = fontfamily.Length > 0 ? fontfamily.Substring(0, fontfamily.Length - 2) : fontfamily;
            }
            return fontfamily;
        }
    }
}
