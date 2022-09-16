void SaveAsPDF(string filepath)
{
    if (!File.Exists(filepath.ToString()))
    {
        Console.WriteLine($"{filepath} not exists");
        return;
    }

    // Create netoffice application
    NetOffice.WordApi.Application word = new NetOffice.WordApi.Application();

    // Hide the application
    word.Visible = false;
    word.ScreenUpdating = false;

    string SourceFile = filepath ;
    object filename = SourceFile;

    // Open the document
    NetOffice.WordApi.Document doc = word.Documents.Open(filename);
    doc.Activate();

    // Change the file extension from docx to pdf
    object outputFileName = SourceFile.Replace(".docx", ".pdf");
    object fileFormat = NetOffice.WordApi.Enums.WdSaveFormat.wdFormatPDF;

    // Saving the file as pdf
    if (!File.Exists(outputFileName.ToString()))
    {
        doc.SaveAs(outputFileName, fileFormat);
        Console.WriteLine($"Save pdf to {outputFileName}");
    }
    else
        Console.WriteLine($"{outputFileName} already exists");

    // Close the document
    object saveChanges = NetOffice.WordApi.Enums.WdSaveOptions.wdDoNotSaveChanges;
    ((NetOffice.WordApi._Document)doc).Close(saveChanges);

    // Quit the application
    ((NetOffice.WordApi._Application)word).Quit();
}