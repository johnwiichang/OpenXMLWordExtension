using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLWordExtension.InnerModels
{
    public class ParagraphPropertiesForHtml
    {
        public String TextAlign
        {
            get;
            private set;
        }

        public static ParagraphPropertiesForHtml CreateRunPropertiesForHtml(ParagraphProperties pp)
        {
            pp = pp ?? new ParagraphProperties();
            var instance = new ParagraphPropertiesForHtml
            {
                TextAlign = pp.Justification.GetValue(),
            };
            return instance;
        }

        public String GenerateStyleString()
        {
            String style = TextAlign == "" ? "" : $"text-align: {TextAlign}, ";
            return style.ToStyleString();
        }

        private ParagraphPropertiesForHtml() { }
    }
}
