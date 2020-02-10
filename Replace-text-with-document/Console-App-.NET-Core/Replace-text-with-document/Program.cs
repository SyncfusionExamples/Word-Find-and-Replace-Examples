using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_with_document
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
                //Finds all the content placeholder text in the Word document.
                TextSelection[] textSelections = document.FindAll(new Regex(@"\[(.*)\]"));
                for (int i = 0; i < textSelections.Length; i++)
                {
                    //Replaces the content placeholder text with desired Word document.
                    docStream = File.OpenRead(Path.GetFullPath(@"../../../" + textSelections[i].SelectedText.TrimStart('[').TrimEnd(']') + ".docx"));
                    WordDocument subDocument = new WordDocument(docStream, FormatType.Docx);
                    docStream.Dispose();
                    document.Replace(textSelections[i].SelectedText, subDocument, true, true);
                    subDocument.Dispose();
                }
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
