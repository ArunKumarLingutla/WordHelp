using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordHelp
{
    public class Program
    {
        static void Main(string[] args)
        {
            var wordObj = new WordUtility();

            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            DirectoryInfo directoryInfo = new DirectoryInfo(exePath);
            DirectoryInfo sourceFolder = Directory.GetParent(exePath).Parent.Parent.Parent;
            string wordFilesFolder = Path.Combine(sourceFolder.FullName, "TestCases");

            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var filePath= System.IO.Path.Combine(wordFilesFolder, "TemplateLandScape.docx");
            //var nctoword= System.IO.Path.Combine(desktopPath, "nctoword.docx");
            //var dummyTextFile= System.IO.Path.Combine(desktopPath, "DummyNCFile.txt");
            var filesToMerge = new string[] {
                System.IO.Path.Combine(wordFilesFolder, "Merge1.docx"),
                System.IO.Path.Combine(wordFilesFolder, "Merge2.docx"),
            };
            var destinationFilePath = System.IO.Path.Combine(wordFilesFolder, "Final Merge.docx");
            var templateFilePath = System.IO.Path.Combine(wordFilesFolder, "templateFile.docx");

            //WordUtility.OpenWordDocument(wordObj, multipleImagesFile);
            //WordUtility.ReplaceText(wordObj.wordDoc, "Hi", "Hello");
            //WordUtility.ReplaceImage(wordObj.wordDoc,System.IO.Path.Combine(desktopPath, "Test.png"));
            //WordUtility.InsertAPicture(wordObj.wordDoc, System.IO.Path.Combine(desktopPath, "Test.png"));
            //WordUtility.SaveWordProcessDocument(wordObj.wordDoc);
            //WordUtility.CloseWordProcessDocument(wordObj.wordDoc);
            //WordUtility.MergeDocuments(templateFilePath, filesToMerge, destinationFilePath);
            ////WordUtility.MergeDocuments(filesToMerge, destinationFilePath);

            //DocxConverter.ConvertToHtml(destinationFilePath, System.IO.Path.Combine(desktopPath, "Final Merge.html"));

            //WordUtility.ConvertNcToWord(dummyTextFile,nctoword);

            var multipleImagesFile = Path.Combine(wordFilesFolder, "MultipleImages.docx");
            File.Copy(templateFilePath, multipleImagesFile, true);
            WordUtility.OpenWordDocument(wordObj, multipleImagesFile);

            List<string> ImageFiles = Directory.GetFiles(wordFilesFolder, "*.jpg").ToList();
            WordUtility.InsertImagesWithCaptions(wordObj.wordDoc, ImageFiles);
                WordUtility.SaveAs(wordObj.wordDoc, Path.Combine(desktopPath, "MultipleImagesWithCaptions.docx"));
                WordUtility.CloseWordProcessDocument(wordObj.wordDoc);
            //foreach (string imageFile in ImageFiles)
            //{
            //    WordUtility.InsertAPicture(wordObj.wordDoc, imageFile);
            //}
            //WordUtility.SaveAs(wordObj.wordDoc, Path.Combine(desktopPath, "1 Tool Sheet.docx"));
            //WordUtility.CloseWordProcessDocument(wordObj.wordDoc);

            //WordUtility.MergeDocuments(templateFilePath, filesToMerge, destinationFilePath);
            //DocxConverter.ConvertToHtml(destinationFilePath, System.IO.Path.Combine(wordFilesFolder, "Final Merge.html"));



        }
    }
}
