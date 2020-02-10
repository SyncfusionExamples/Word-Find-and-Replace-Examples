using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_and_highlight_text
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
                //Finds all occurrence of the text in the Word document.
                TextSelection[] textSelections = document.FindAll("Adventure", true, true);
                for (int i = 0; i < textSelections.Length; i++)
                {
                    //Sets the highlight color for the searched text as Yellow.
                    textSelections[i].GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
