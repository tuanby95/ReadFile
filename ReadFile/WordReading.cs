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
                foreach (var paragraph in paragraphs)
                {
                    Console.WriteLine(paragraph.InnerText);
                }
                Console.ReadKey();
            }
        }
    }
}
