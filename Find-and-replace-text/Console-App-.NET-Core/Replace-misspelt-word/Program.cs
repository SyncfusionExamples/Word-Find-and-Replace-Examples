using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Find_and_replace_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Opens the input Word document
                Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../Template.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                //Finds all the occurrence of a misspelt word and replaces with desired word in the Word document.
                document.Replace("Cyles", "Cycles", true, true);
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
