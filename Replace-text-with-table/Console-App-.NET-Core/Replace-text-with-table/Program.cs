using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Xml;

namespace Replace_text_with_table
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
                //Creates a new table
                WTable table = new WTable(document);
                table.ResetCells(1, 6);
                table[0, 0].Width = 52f;
                table[0, 0].AddParagraph().AppendText("Supplier ID");
                table[0, 1].Width = 128f;
                table[0, 1].AddParagraph().AppendText("Company Name");
                table[0, 2].Width = 70f;
                table[0, 2].AddParagraph().AppendText("Contact Name");
                table[0, 3].Width = 92f;
                table[0, 3].AddParagraph().AppendText("Address");
                table[0, 4].Width = 66.5f;
                table[0, 4].AddParagraph().AppendText("City");
                table[0, 5].Width = 56f;
                table[0, 5].AddParagraph().AppendText("Country");
                //Imports data to the table
                ImportDataToTable(table);
                //Applies the built-in table style (Medium Shading 1 Accent 1) to the table
                table.ApplyStyle(BuiltinTableStyle.MediumShading1Accent1);
                //Replaces the table placeholder text with a new table
                TextBodyPart bodyPart = new TextBodyPart(document);
                bodyPart.BodyItems.Add(table);
                document.Replace("[Suppliers table]", bodyPart, true, true, true);
                //Saves the resultant file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }

        /// <summary>
        /// Imports the data from XML file to the table.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="System.Exception">reader</exception>
        /// <exception cref="XmlException">Unexpected xml tag  + reader.LocalName</exception>
        private static void ImportDataToTable(WTable table)
        {
            FileStream fs = new FileStream(@"../../../Suppliers.xml", FileMode.Open, FileAccess.Read);
            XmlReader reader = XmlReader.Create(fs);
            if (reader == null)
                throw new Exception("reader");
            while (reader.NodeType != XmlNodeType.Element)
                reader.Read();
            if (reader.LocalName != "SuppliersList")
                throw new XmlException("Unexpected xml tag " + reader.LocalName);
            reader.Read();
            while (reader.NodeType == XmlNodeType.Whitespace)
                reader.Read();
            while (reader.LocalName != "SuppliersList")
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "Suppliers":
                            //Adds new row to the table for importing data from next record.
                            WTableRow tableRow = table.AddRow(true);
                            ImportDataToRow(reader, tableRow);
                            break;
                    }
                }
                else
                {
                    reader.Read();
                    if ((reader.LocalName == "SuppliersList") && reader.NodeType == XmlNodeType.EndElement)
                        break;
                }
            }
            reader.Dispose();
            fs.Dispose();
        }
        /// <summary>
        /// Imports the data to the table row.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">reader</exception>
        /// <exception cref="XmlException">Unexpected xml tag  + reader.LocalName</exception>
        private static void ImportDataToRow(XmlReader reader, WTableRow tableRow)
        {
            if (reader == null)
                throw new Exception("reader");
            while (reader.NodeType != XmlNodeType.Element)
                reader.Read();
            if (reader.LocalName != "Suppliers")
                throw new XmlException("Unexpected xml tag " + reader.LocalName);
            reader.Read();
            while (reader.NodeType == XmlNodeType.Whitespace)
                reader.Read();
            while (reader.LocalName != "Suppliers")
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "SupplierID":
                            tableRow.Cells[0].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "CompanyName":
                            tableRow.Cells[1].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "ContactName":
                            tableRow.Cells[2].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "Address":
                            tableRow.Cells[3].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "City":
                            tableRow.Cells[4].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "Country":
                            tableRow.Cells[5].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        default:
                            reader.Skip();
                            break;
                    }
                }
                else
                {
                    reader.Read();
                    if ((reader.LocalName == "Suppliers") && reader.NodeType == XmlNodeType.EndElement)
                        break;
                }
            }
        }
    }
}
