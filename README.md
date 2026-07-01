# Word Find and Replace examples

This repository contains examples that illustrates how to find and replace text in Word documents programmatically in C# and VB.NET using the [.NET Word library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO) without Microsoft Word or Office interop dependencies.

## Find and replace text in Word documents using C#

The [.NET Word library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO) provides comprehensive APIs to find and replace text in a Word document with any desired text, image, hyperlink, paragraph, table, and part of a document or an entire document. It provides options to find text by matching case and whole words. You can find each occurrence one by one or all the occurrences of a text in the document. It saves your efforts and time by helping you to automate find and replace a pattern of text in a Word document using [Regex](https://docs.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?utm_source=github&utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples). It allows you to save the resultant document as Word document (DOCX, WordML), PDF, image, HTML and more.

## Key Features

- [Find and highlight text in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-text-in-a-word-document-and-format) - Find text from a Word document and format or highlight it.
- [Find and replace text in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-and-replace-text-with-other-text) - Find text from a Word document and replace it with another text.
- [Replace text with image in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-and-replace-text-with-an-image) - Find a placeholder text from a Word document and replace it with any desired image.
- [Replace text with merge field in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-and-replace-a-pattern-of-text-with-a-merge-field) - Find a pattern of text from a Word document and replace it with merge fields.
- [Replace text with table in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-and-replace-text-with-a-table) - Find a placeholder text from a Word document and replace it with a table.
- [Replace text with document in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-and-replace-text-in-word-document-with-another-document) - Find text from a Word document and replace it with another Word document.
- [Find and replace several paragraphs in C#](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace#find-and-replace-text-extending-to-several-paragraphs) - Find text that extends to several paragraphs in a Word document and replace it with another Word document.

## Screenshots

**Find and highlight text**

<p align="center"> 
<img src="Find-and-highlight-text/Images/Find-and-highlight-text.png" alt="Find text from a Word document and format or highlight it in C#"/> 
</p>

**Find and replace text**

<p align="center"> 
<img src="Find-and-replace-text/Images/Replace-misspelt-word.png" alt="Find a misspelt word from a Word document and replace it with desired word in C#"/> 
</p>

**Replace text with table**

<p align="center"> 
<img src="Replace-text-with-table/Images/Replace-text-with-table.png" alt="Find text from a Word document and replace it with a table in C#"/> 
</p>

## .NET Word Library

The Syncfusion&reg; DocIO is a [.NET Word library](https://www.syncfusion.com/document-sdk/net-word-library) allows you to add advanced Word document processing functionalities to any .NET application and does not require Microsoft Word application to be installed in the machine. It is a non-UI component that provides a full-fledged document instance model similar to the Microsoft Office COM libraries to iterate with the document elements explicitly and perform necessary manipulation. 

*   Support to [create Word document](https://www.syncfusion.com/document-sdk/net-word-library/create-word-documents) from scratch.
*   Support to open (read), modify and save existing Word documents.
*   Advanced [Mail merge](https://www.syncfusion.com/document-sdk/net-word-library/mail-merge) support with different data sources.
*   Ability to create or edit Word 97-2003 and later version documents (DOCX), and convert them to commonly used file formats such as [RTF](https://help.syncfusion.com/document-processing/word/conversions/rtf-conversions), [WordML](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-processing-xml-xml), [TXT](https://www.syncfusion.com/document-sdk/net-word-library/text-conversions), [HTML](https://www.syncfusion.com/document-sdk/net-word-library/html-conversions) and vice versa.
*   Ability to export a Word document as an [Image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion) and [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversion)
*   Ability to [merge](https://www.syncfusion.com/document-sdk/net-word-library/merge-word-documents) and [split](https://www.syncfusion.com/document-sdk/net-word-library/split-word-documents) Word documents.
*   Support to [compare](https://www.syncfusion.com/document-sdk/net-word-library/compare-word-documents) two DOCX format documents.
*   Ability to create and manipulate [charts](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-charts), [Shapes](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-shapes), and [Group shape](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-shapes#grouping-shapes) in DOCX and WordML format documents.
*   Ability to read and write [Built-In and Custom Document Properties](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-word-document#working-with-word-document-properties).
*   Ability to [find and replace](https://www.syncfusion.com/document-sdk/net-word-library/find-and-replace) text with its original formatting.
*   Ability to insert [Bookmarks](https://www.syncfusion.com/document-sdk/net-word-library/bookmark-in-word-document) and navigate corresponding bookmarks to insert, replace, and delete content.
*   Support to insert and edit the [form fields](https://www.syncfusion.com/document-sdk/net-word-library/form-filling-in-word-document).
*   Support to protect the document to [restrict access](https://www.syncfusion.com/document-sdk/net-word-library/protect-word-documents) to the elements present within the document.
*   Ability to [encrypt and decrypt](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-security) Word documents.
*   Support to [insert and extract OLE](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-paragraph#working-with-ole-objects) objects.
*   Support to run the DocIO applications in multi-thread and its thread safe.

Compatible Microsoft Word Versions
----------------------------------

*   Microsoft Word 97-2003
*   Microsoft Word 2007
*   Microsoft Word 2010
*   Microsoft Word 2013
*   Microsoft Word 2016
*   Microsoft Word 2019
*   Microsoft 365

Supported File Formats
----------------------

*   Creates, reads, and edits popular text file formats like [DOC](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#doc-to-docx-and-docx-to-doc), DOT, [DOCM](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#macros-docm-dotm), DOTM, [DOCX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-document-docx), [DOTX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-template-dotx), [HTML](https://www.syncfusion.com/document-sdk/net-word-library/html-conversions), [RTF](https://help.syncfusion.com/document-processing/word/conversions/rtf-conversions), [TXT](https://www.syncfusion.com/document-sdk/net-word-library/text-conversions), [Markdown](https://help.syncfusion.com/document-processing/word/conversions/markdown-to-word-conversion) and [XML (WordML)](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-processing-xml-xml).
*   Converts Word documents also to [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversionhub-docio-examples), [Image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion), and [ODT](https://help.syncfusion.com/document-processing/word/conversions/word-to-odt-conversion) files.

## How to run the examples
- Download this project to a location in your disk.
- Open the solution file using Visual Studio.
- Rebuild the solution to install the required NuGet packages.
- Run the application.

## Resources

*   **Product page:** [Syncfusion&reg; Word Framework](https://www.syncfusion.com/document-sdk/net-word-library)
*   **Documentation:** [Find and Replace in Word document](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace)
*   **Online demo:** [Syncfusion&reg; Word library - Online demos](https://document.syncfusion.com/demos/word/salesinvoice#/bootstrap5)
*   **GitHub Examples:** [Syncfusion&reg; Word library examples](https://github.com/SyncfusionExamples/DocIO-Examples)
*   **Blog:** [Syncfusion&reg; Word library - Blog](https://www.syncfusion.com/blogs/category/docio?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)
*   **Knowledge Base:** [Syncfusion&reg; Word library - Knowledge Base](https://www.syncfusion.com/kb/aspnetcore/docio?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)
*   **Ebooks:** [Syncfusion&reg; Word library - Ebooks](https://www.syncfusion.com/succinctly-free-ebooks?utm_source=nuget&utm_medium=listing&utm_campaign=aspnetcore-docio-nuget)
*   **FAQ:** [Syncfusion&reg; Word library - FAQ](https://www.syncfusion.com/faq/?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)

## Support and feedback
- For any other queries, reach our [Syncfusion&reg; support team](https://support.syncfusion.com/agent/tickets/create?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).
- Request new feature through [Syncfusion&reg; feedback portal](https://www.syncfusion.com/feedback/home?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

## License
This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/22.2.5/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples). You can purchase a licnense [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).
