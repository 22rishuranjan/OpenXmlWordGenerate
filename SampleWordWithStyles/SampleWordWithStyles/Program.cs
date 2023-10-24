
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SampleWordWithStyles.Helper;

namespace SampleWordWithStyles;

class Program
{
    public static string filePath = @"/Volumes/badmonster disk/Development/MyGit/dotnet/OpenXmlWordGenerate/SampleWord/SampleWord/Files/Doc1.docx";


    static void Main(string[] args)
    {


        


        using (WordprocessingDocument doc =
            WordprocessingDocument.Open(filePath, true))
        {

            ApplyStyleToSpecificPara(doc);

            //// Get the first paragraph.
            //Paragraph p =
            //  doc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
            //  .ElementAtOrDefault(1);

            //// Check for a null reference. 
            //if (p == null)
            //{
            //    throw new ArgumentOutOfRangeException("p",
            //        "Paragraph was not found.");
            //}

            //WordHelper.ApplyStyleToParagraph(doc, "OverdueAmount", "Overdue Amount", p);
        }
        Console.WriteLine("Hello, World!");
    }


    public static void AddTextToWord(string text)
    {
        var wordprocessingDocument = WordHelper.GetAWordDocument(filePath);
        var body = wordprocessingDocument.MainDocumentPart?.Document?.Body;
        Paragraph para = WordHelper.AddTextToParagraph("Adding paragraph to the document");
        body?.AppendChild<Paragraph>(para);
        wordprocessingDocument.Dispose();
    }

    public static void AddTextToWord(string text, Paragraph para)
    {
        var wordprocessingDocument = WordHelper.GetAWordDocument(filePath);
        var body = wordprocessingDocument.MainDocumentPart?.Document?.Body;
        body?.AppendChild<Paragraph>(para);
        wordprocessingDocument.Dispose();
    }

    public static void ApplyStyleToSpecificPara(WordprocessingDocument doc)
    {

        // Get all paragraph.
        var paraList = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>();

        foreach(Paragraph p in paraList)
        {
            if (p.HasAttributes)
            {
                var attrs = p.GetAttributes();
                foreach(var attr in attrs)
                {
                    if (attr.LocalName == "Id" && attr.Value == "new para")
                    {
                        WordHelper.ApplyStyleToParagraph(doc, "OverdueAmount", "Overdue Amount", p);
                    }

                }

               
            }
        }
        
    }
}
