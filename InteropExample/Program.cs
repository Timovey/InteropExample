using Word = Microsoft.Office.Interop.Word;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
//Word.Application app = null;
//try
//{
//    app = new Word.Application();
//    Object missing = Type.Missing;

//    app.Documents.Open(Path.Combine(Directory.GetCurrentDirectory(), "ExampleTemplate.docx"));

//    Word.Find find = app.Selection.Find;
//    find.Text = "<TUB>";
//    find.Replacement.Text = "Зубенко Михаил Петрович, мафиозник";

//    Object wrap = Word.WdFindWrap.wdFindContinue;
//    Object replace = Word.WdReplace.wdReplaceAll;

//    find.Execute(FindText: missing,
//        MatchCase: false,
//        Wrap: wrap,
//        Replace: replace,
//        Format: false,
//        ReplaceWith: missing);

//    Object newFIle = Path.Combine(Directory.GetCurrentDirectory(), "1.docx");

//    app.ActiveDocument.SaveAs2(newFIle);
//    app.ActiveDocument.Close();
//}
//catch (Exception ex)
//{

//}
//finally
//{
//    if(app != null)
//    {
//        app.Quit();
//    }
//}

string initialPath = Path.Combine(Directory.GetCurrentDirectory(), "ExampleTemplate.docx");
string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
File.Copy(initialPath, resultPath, overwrite: true);

Dictionary<string, string> marks = new Dictionary<string, string>()
            {
                { "TUB","Зубенко"},
            };

using (WordprocessingDocument document = WordprocessingDocument.Open(resultPath, true))
{
    Body documentBody = document.MainDocumentPart.Document.Body;
    List<Paragraph> paragraphsWithMarks = documentBody.Descendants<Paragraph>().Where(x => Regex.IsMatch(x.InnerText, @".*\[\w+\].*")).ToList();
    foreach (Paragraph paragraph in paragraphsWithMarks)
    {
        foreach (Match markMatch in Regex.Matches(paragraph.InnerText, @"\[\w+\]", RegexOptions.Compiled))
        {

            string paragraphMarkValue = markMatch.Value.Trim(new[] { '[', ']' });
            string markValueFromCollection;
            if (marks.TryGetValue(paragraphMarkValue, out markValueFromCollection))
            {
                string editedParagraphText = paragraph.InnerText.Replace(markMatch.Value, markValueFromCollection);
                paragraph.RemoveAllChildren<Run>();
                paragraph.AppendChild<Run>(new Run(new Text(editedParagraphText)));
            }
        }
    }
}
Console.WriteLine("Hello, World!");
