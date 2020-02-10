using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_with_merge_field
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
                //Finds all the placeholder text enclosed within '«' and '»' in the Word document
                TextSelection[] textSelections = document.FindAll(new Regex("«([(?i)image(?-i)]*:*[a-zA-Z0-9 ]*:*[a-zA-Z0-9 ]+)»"));
                string[] searchedPlaceholders = new string[textSelections.Length];
                for (int i = 0; i < textSelections.Length; i++)
                {
                    searchedPlaceholders[i] = textSelections[i].SelectedText;
                }
                for (int i = 0; i < searchedPlaceholders.Length; i++)
                {
                    //Replaces the placeholder text enclosed within '«' and '»' with desired merge field
                    WParagraph paragraph = new WParagraph(document);
                    paragraph.AppendField(searchedPlaceholders[i].TrimStart('«').TrimEnd('»'), FieldType.FieldMergeField);
                    TextSelection newSelection = new TextSelection(paragraph, 0, paragraph.Items.Count);
                    TextBodyPart bodyPart = new TextBodyPart(document);
                    bodyPart.BodyItems.Add(paragraph);
                    document.Replace(searchedPlaceholders[i], bodyPart, true, true, true);
                }
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
