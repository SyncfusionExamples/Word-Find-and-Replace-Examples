using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Find_and_replace_several_paragraphs
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
                docStream = File.OpenRead(Path.GetFullPath(@"../../../Source.docx"));
                WordDocument subDocument = new WordDocument(docStream, FormatType.Docx);
                docStream.Dispose();
                //Gets the content from another Word document to replace
                TextBodyPart replacePart = new TextBodyPart(subDocument);
                foreach (TextBodyItem bodyItem in subDocument.LastSection.Body.ChildEntities)
                {
                    replacePart.BodyItems.Add(bodyItem.Clone());
                }
                string placeholderText = "Suppliers/Vendors of Northwind" + "Customers of Northwind"
                    + "Employee details of Northwind traders" + "The product information"
                    + "The inventory details" + "The shippers" + "Purchase Order transactions"
                    + "Sales Order transaction" + "Inventory transactions" + "Invoices" + "[end replace]";
                //Finds the text that extends to several paragraphs and replaces it with desired content.
                document.ReplaceSingleLine(placeholderText, replacePart, false, false);
                subDocument.Dispose();
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
