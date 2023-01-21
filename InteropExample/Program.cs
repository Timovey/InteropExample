using Word = Microsoft.Office.Interop.Word;

Word.Application app = null;
try
{
    app = new Word.Application();
    Object missing = Type.Missing;

    app.Documents.Open(Path.Combine(Directory.GetCurrentDirectory(), "ExampleTemplate.docx"));

    Word.Find find = app.Selection.Find;
    find.Text = "<TUB>";
    find.Replacement.Text = "Зубенко Михаил Петрович, мафиозник";

    Object wrap = Word.WdFindWrap.wdFindContinue;
    Object replace = Word.WdReplace.wdReplaceAll;

    find.Execute(FindText: missing,
        MatchCase: false,
        Wrap: wrap,
        Replace: replace,
        Format: false,
        ReplaceWith: missing);

    Object newFIle = Path.Combine(Directory.GetCurrentDirectory(), "1.docx");
    
    app.ActiveDocument.SaveAs2(newFIle);
    app.ActiveDocument.Close();
}
catch (Exception ex)
{

}
finally
{
    if(app != null)
    {
        app.Quit();
    }
}

Console.WriteLine("Hello, World!");
