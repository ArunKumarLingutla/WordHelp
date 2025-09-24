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

            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var filePath= System.IO.Path.Combine(desktopPath, "TemplateLandScape.docx");
            var nctoword= System.IO.Path.Combine(desktopPath, "nctoword.docx");
            var dummyTextFile= System.IO.Path.Combine(desktopPath, "DummyNCFile.txt");
            var filesToMerge = new string[] {
                System.IO.Path.Combine(desktopPath, "Merge1.docx"),
                System.IO.Path.Combine(desktopPath, "Merge2.docx"),
            };
            var destinationFilePath = System.IO.Path.Combine(desktopPath, "Final Merge.docx");
            var templateFilePath = System.IO.Path.Combine(desktopPath, "templateFile.docx");

            WordUtility.OpenWordDocument(wordObj, filePath);
            //WordUtility.ReplaceText(wordObj.wordDoc, "Hi", "Hello");
            //WordUtility.ReplaceImage(wordObj.wordDoc,System.IO.Path.Combine(desktopPath, "Test.png"));
            //WordUtility.InsertAPicture(wordObj.wordDoc, System.IO.Path.Combine(desktopPath, "Test.png"));
            //WordUtility.SaveWordProcessDocument(wordObj.wordDoc);
            //WordUtility.CloseWordProcessDocument(wordObj.wordDoc);
            //WordUtility.MergeDocuments(templateFilePath,filesToMerge, destinationFilePath);
            ////WordUtility.MergeDocuments(filesToMerge, destinationFilePath);

            //DocxConverter.ConvertToHtml(destinationFilePath, System.IO.Path.Combine(desktopPath, "Final Merge.html"));

            //WordUtility.ConvertNcToWord(dummyTextFile,nctoword);

            List<string> ImageFiles = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "*.jpg").ToList();

            foreach (string imageFile in ImageFiles) 
            {
                WordUtility.InsertAPicture(wordObj.wordDoc, imageFile);
            }
            WordUtility.SaveAs(wordObj.wordDoc,Path.Combine(desktopPath,"1 Tool Sheet.docx"));
            WordUtility.CloseWordProcessDocument(wordObj.wordDoc);

            WordUtility.MergeDocuments(templateFilePath, filesToMerge, destinationFilePath);
            DocxConverter.ConvertToHtml(destinationFilePath, System.IO.Path.Combine(desktopPath, "Final Merge.html"));

        }
    }
}
