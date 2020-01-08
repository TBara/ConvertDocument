using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ConvertDocFormat
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Convert Document Format";
            int choice;
            Console.WriteLine("Select your action:\n1.  Convert .doc to .docx\n2.  Convert .docx to.doc\n3.  Convert Word document to PDF");
            do
            {
                choice = int.Parse(Console.ReadLine());
                if ((choice < 1) || (choice > 3))
                {
                    Console.WriteLine("Wrong choice");
                }
            } while ((choice < 1) || (choice > 3));


            if (choice == 1)
            {
                var rootFolder = GetFileFolder();

                foreach (var file in Directory.EnumerateFiles(rootFolder, "*.doc"))
                {
                    ConvertDocToDocx(file);
                }
            }
            else if (choice == 2)
            {
                var rootFolder = GetFileFolder();

                foreach (var file in Directory.EnumerateFiles(rootFolder, "*.docx"))
                {
                    ConvertDocxToDoc(file);
                }

            } else if (choice == 3)
            {
                var rootFolder = GetFileFolder();
                if (rootFolder != null)
                {
                    foreach (var file in Directory.EnumerateFiles(rootFolder, "*.doc*"))
                    {
                        ConvertDocToPDF(file);
                    }
                }
                else
                {
                    return;
                }
            }
            else
            {
                Console.WriteLine(String.Format("Chose: {0}", choice));
            }

        }
        static void ConvertDocxToDoc(string path)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            if (path.ToLower().EndsWith(".docx"))
            {
                try
                {
                    var sourceFile = new FileInfo(path);
                    var document = word.Documents.Open(sourceFile.FullName);

                    string newFileName = sourceFile.FullName.Replace(".docx", ".doc");
                    document.SaveAs2(newFileName, WdSaveFormat.wdFormatDocument97);

                    word.ActiveDocument.Close();
                    word.Quit();
                    File.Delete(path);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        static void ConvertDocToDocx(string path)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            if (path.ToLower().EndsWith(".doc"))
            {
                var sourceFile = new FileInfo(path);
                var document = word.Documents.Open(sourceFile.FullName);

                string newFileName = sourceFile.FullName.Replace(".doc", ".docx");
                document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument,
                                 CompatibilityMode: WdCompatibilityMode.wdWord2013);

                word.ActiveDocument.Close();
                word.Quit();
                File.Delete(path);
            }
        }

        static void ConvertDocToPDF(string path)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            if ((path.ToLower().EndsWith(".docx")) || (path.ToLower().EndsWith(".doc")))
            {
                try
                {
                    var sourceFile = new FileInfo(path);
                    var document = word.Documents.Open(sourceFile.FullName);
                    string newFileName = "";
                    if (path.ToLower().EndsWith(".docx"))
                    {
                        newFileName = sourceFile.FullName.Replace(".docx", ".pdf");
                    }
                    else if (path.ToLower().EndsWith(".doc"))
                    {
                        newFileName = sourceFile.FullName.Replace(".doc", ".pdf");
                    }
                    
                    document.SaveAs2(newFileName, WdSaveFormat.wdFormatPDF);

                    word.ActiveDocument.Close();
                    word.Quit();
                    File.Delete(path);
                    CleanUpExcelObj();
                }
                catch (Exception e)
                {
                    Console.WriteLine(String.Format("Error Message: {0}\nError Stack Trace: {1}",e.Message, e.StackTrace));
                }
            }
        }

        private static string GetFileFolder()
        {
            Console.WriteLine("Select folder with files to convert.");
            try
            {
                string saveFolder = "";
                var t = new Thread((ThreadStart)(() => {
                    FolderBrowserDialog folderDlg = new FolderBrowserDialog();
                    folderDlg.ShowNewFolderButton = true;
                    // Show the FolderBrowserDialog.  
                    DialogResult result = folderDlg.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        saveFolder = folderDlg.SelectedPath;
                        Environment.SpecialFolder root = folderDlg.RootFolder;
                    }
                }));

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                return saveFolder;
            }
            catch (Exception)
            {

                throw;
            }

        }

        //Clean up the resources
        public static void CleanUpExcelObj()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
