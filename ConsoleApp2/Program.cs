using Microsoft.Office.Interop.Word;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Application();
            object missing = System.Reflection.Missing.Value;
            object newFilename1 = "C:\\tmp\\Doc1.rtf";
            Document doc1 = app.Documents.Open(ref newFilename1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            object newFilename2 = "C:\\tmp\\Doc2.rtf";
            Document doc2 = app.Documents.Open(ref newFilename2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            var compareReport = app.CompareDocuments(doc1, doc2);
            object compareReportName = "C:\\tmp\\DocRes.rtf";
            compareReport.SaveAs2(ref compareReportName);

            object saveChanges = true;
            app.Quit(ref saveChanges);
        }
    }
}
