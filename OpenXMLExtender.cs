using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLWordExtension.InnerModels;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXMLWordExtension
{
    public static class OpenXMLExtender
    {
        private static Int32 currentPicture = 0;

        public static IEnumerable<ImagePart> GetImages(this Document doc)
        {
            return doc.MainDocumentPart.ImageParts;
        }

        public static UInt32 GetImageId(this Drawing d)
        {

            return (d.SearchAll("docPr").FirstOrDefault() as DocProperties).Id.Value;
        }

        public static String GetImageBase64(this ImagePart img)
        {
            List<Byte> bs = new List<byte>();
            int current = 0;
            var stream = img.GetStream(System.IO.FileMode.Open);
            while (current != -1)
            {
                current = stream.ReadByte();
                bs.Add((Byte)current);
            }
            return $"data:{img.ContentType};base64,{Convert.ToBase64String(bs.ToArray())}";
        }

        public static String GetImageBase64(this Document doc, Int32 pictureIndex)
        {
            List<Byte> bs = new List<byte>();
            var img = doc.MainDocumentPart.ImageParts.ElementAt(pictureIndex);
            int current = 0;
            var stream = img.GetStream(System.IO.FileMode.Open);
            while (current != -1)
            {
                current = stream.ReadByte();
                bs.Add((Byte)current);
            }
            return $"data:{img.ContentType};base64,{Convert.ToBase64String(bs.ToArray())}";
        }

        public static IEnumerable<OpenXmlElement> GetContent(this Body b)
        {
            return b.ChildElements.Where(x => x.LocalName == "p" || x.LocalName == "tbl");
        }

        public static IEnumerable<TableRow> GetRows(this Table t)
        {
            return t.ChildElements.Where(x => x.LocalName == "tr").Select(x => x as TableRow);
        }

        public static IEnumerable<TableCell> GetCells(this TableRow t)
        {
            return t.ChildElements.Where(x => x.LocalName == "tc").Select(x => x as TableCell);
        }

        public static IEnumerable<Paragraph> GetParagraphs(this Body b)
        {
            return b.ChildElements.Where(x => x.LocalName == "p").Select(x => x as Paragraph);
        }

        public static IEnumerable<Paragraph> GetParagraphs(this TableCell tc)
        {
            return tc.ChildElements.Where(x => x.LocalName == "p").Select(x => x as Paragraph);
        }

        public static IEnumerable<Table> GetTables(this Body b)
        {
            return b.ChildElements.Where(x => x.LocalName == "tbl").Select(x => x as Table);
        }

        public static IEnumerable<Run> GetRuns(this Paragraph p)
        {
            return p.ChildElements.Where(x => x.LocalName == "r").Select(x => x as Run);
        }

        public static Drawing GetDrawing(this Run r)
        {
            return r.ChildElements.Where(x => x.LocalName == "drawing").Select(x => x as Drawing).FirstOrDefault();
        }

        public static String GetValue(this OpenXmlLeafElement oxle)
        {
            String val = "";
            if (oxle != null)
            {
                switch (oxle.GetType().Name)
                {
                    case nameof(FontSize):
                        val = (oxle as FontSize).Val.Value;
                        break;
                    case nameof(Color):
                        val = (oxle as Color).Val.Value;
                        break;
                    case nameof(Bold):
                        val = oxle == null ? "" : "bold";
                        break;
                    case nameof(Highlight):
                        val = (oxle as Highlight).Val.Value.ToString();
                        break;
                    case nameof(Justification):
                        val = (oxle as Justification).Val.Value.ToString();
                        break;
                    default:
                        throw new Exception("Method can not work with object provided.");
                }
            }
            return val;
        }

        public static String GetStyleString(this String str, String xName)
        {
            return str == "" ? "" : $"{xName}: {str}; ";
        }

        public static RunPropertiesForHtml ToHtmlVersion(this RunProperties rp)
        {
            return RunPropertiesForHtml.CreateRunPropertiesForHtml(rp);
        }

        public static ParagraphPropertiesForHtml ToHtmlVersion(this ParagraphProperties pp)
        {
            return ParagraphPropertiesForHtml.CreateRunPropertiesForHtml(pp);
        }

        public static String ToStyleString(this String style)
        {
            return style.Length == 0 ? "" : $" style=\"{style.Substring(0, style.Length - 2)}\"";
        }

        public static String ToHtml(this Run r)
        {
            var draw = r.GetDrawing();
            if (draw != null)
            {
                currentPicture++;
                return $"<img id=\"openXml2HtmlImage{currentPicture}\" src=\"\"/>";
            }
            return $"<span{r.RunProperties.ToHtmlVersion().GenerateStyleString()}>{r.InnerText}</span>";
        }

        public static String ToHtml(this Paragraph p)
        {
            var ptag = $"<p{p.ParagraphProperties.ToHtmlVersion().GenerateStyleString()}>";
            foreach (var r in p.GetRuns())
            {
                ptag += r.ToHtml();
            }
            ptag += "</p>";
            return ptag;
        }

        public static String ToHtml(this Table t)
        {
            var tbl = $"<table style=\"width:100%; border-collapse: collapse;\" border=\"1\">";
            foreach (var row in t.GetRows())
            {
                tbl += "<tr>";
                foreach (var cell in row.GetCells())
                {
                    tbl += "<td>";
                    foreach (var p in cell.GetParagraphs())
                    {
                        tbl += p.ToHtml();
                    }
                    tbl += "</td>";
                }
                tbl += "</tr>";
            }
            return tbl += "</table>";
        }

        public static String ToHtml(this Body b)
        {
            var body = "<body>";
            foreach (var item in b.GetContent())
            {
                if (item.LocalName == "p")
                {
                    body += (item as Paragraph).ToHtml();
                }
                else if (item.LocalName == "tbl")
                {
                    body += (item as Table).ToHtml();
                }
            }
            return body + "</body>";
        }

        public static String ToHtml(this Document d)
        {
            var html = "<html><head><meta charset=\"utf-8\"></head>";
            html += d.Body.ToHtml() + "<script>";
            var imgs = d.GetImages();
            for (int i = 0; i < imgs.Count(); i++)
            {
                html += $"document.getElementById('openXml2HtmlImage{i + 1}').src=\"{imgs.ElementAt(i).GetImageBase64()}\";";
            }
            return html + "</script></html>";
        }

        static IEnumerable<OpenXmlElement> SearchAll(this OpenXmlElement oxe, String localname)
        {
            List<OpenXmlElement> oxes = new List<OpenXmlElement>();
            foreach (var item in oxe)
            {
                if (item.LocalName == localname)
                {
                    oxes.Add(item);
                }
            }
            foreach (var item in oxe.Elements())
            {
                SearchAll(item, localname).ToList().ForEach(x => oxes.Add(x));
            }
            return oxes;
        }
    }
}
