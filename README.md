# Office OpenXML WordProcessing Extension
This library provides a series of methods to easily convert docx documents to html format and access Office OpenXML Word Processing files.

# Compatibility information
Written in .NET Core.

# Functions

 - For the paragraph section of the HTML conversion.
 - Get the picture in the file (Base64 encoded).
 - Convert tables to HTML.

The above function currently only supports part of the stylized effect. It is recommended only for the extraction of text material.
Does not support encrypted Office documents.

# How to use
Convert a docx into html:
```
using DocumentFormat.OpenXml.Packaging;
using OpenXMLWordExtension;

class Program
{
    static void Main(string[] args)
    {
        WordprocessingDocument docx = WordprocessingDocument.Open("testDocx.docx", false);
        System.IO.File.WriteAllText(@"/Users/johnwii/Desktop/out.html", docx.MainDocumentPart.Document.ToHtml());
    }
}
```
You still need [OpenXML SDK][1].

## Tips
Because .NET Core does not support the System.Xml.Xsl namespace temporarily, graphics **(not images)** are not currently visible. If you need this part of the feature, you can try to integrate the *[VectorConvertor][2]* project.


  [1]: https://msdn.microsoft.com/en-us/library/office/bb448854.aspx
  [2]: https://github.com/JohnwiiChang/VectorConvertor
