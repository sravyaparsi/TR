using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace sample
{
    class Program
    {
        static void Main(string[] args)
        {
            
            
            XmlDocument xDoc = new XmlDocument();

            xDoc.Load("sample.xml");
            XmlNodeList name = xDoc.GetElementsByTagName("para");
            XmlNodeList report = xDoc.GetElementsByTagName("report-template");

            var reportid = report[1].Attributes["section-name"];
            using (WordprocessingDocument doc = WordprocessingDocument.Create("document.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
              
                 Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text(name[0].InnerText));
             



            }
        }
    }
}
