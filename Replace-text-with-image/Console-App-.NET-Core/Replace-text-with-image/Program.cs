using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Text.RegularExpressions;

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
                //Finds all the image placeholder text in the Word document.
                TextSelection[] textSelections = document.FindAll(new Regex("^//(.*)"));
                for (int i = 0; i < textSelections.Length; i++)
                {
                    //Replaces the image placeholder text with desired image.
                    Stream imageStream = File.OpenRead(Path.GetFullPath(@"../../../" + textSelections[i].SelectedText + ".png"));
                    WParagraph paragraph = new WParagraph(document);
                    WPicture picture = paragraph.AppendPicture(imageStream) as WPicture;
                    imageStream.Dispose();
                    TextSelection newSelection = new TextSelection(paragraph, 0, 1);
                    TextBodyPart bodyPart = new TextBodyPart(document);
                    bodyPart.BodyItems.Add(paragraph);
                    document.Replace(textSelections[i].SelectedText, bodyPart, true, true);
                }
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
