using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFile
{
    internal class WordReading
    {
        public void ReadDocx(string filepath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, false))
            {
                var paragraphs = wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();
                List<string> lines = new List<string>();
                foreach (var par in paragraphs)
                {
                    lines.Add(par.InnerText.ToString());
                }
                Console.WriteLine();

                //string childStr = "Sub Total";
                //foreach (var paragraph in paragraphs)
                //{
                //    if (paragraph.InnerText.Contains(childStr))
                //    {
                //        Console.WriteLine(paragraph.InnerText);
                //        break;
                //    }
                //    else
                //    {
                //        Console.WriteLine(paragraph.InnerText);

                //    }
                //}
                Console.ReadKey();
                //var paragraphs = wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();
                //var tables = wordDocument.MainDocumentPart.RootElement.Descendants<Table>();
                //foreach (var paragraph in paragraphs)
                //{
                //    Console.WriteLine(paragraph.InnerText);
                //}
                //foreach (var table in tables)
                //{
                //    foreach(var line in table)
                //    {
                //        Console.WriteLine(line.InnerText);
                //    }
                //}
            }
        }
    }
}
