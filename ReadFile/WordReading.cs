using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFile
{
    internal class WordReading
    {
        public void ReadDocx()
        {
            Document document = new Document(@"D:\C#\Project\ReadFile\ReadFile\test.docx");
            document.
            Document _document = new Document();
            DocumentBuilder builder = new DocumentBuilder(_document);
            _document.Append(document);
            _document.Save("mytestdocx.docx");
        }
    }
}
