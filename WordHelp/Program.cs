using System;
using System.Collections.Generic;
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
            var filePath= System.IO.Path.Combine(desktopPath, "Merge2.docx");
            var filesToMerge=new string[] {
                System.IO.Path.Combine(desktopPath, "Merge1.docx"),
                System.IO.Path.Combine(desktopPath, "Merge2.docx"),
            };
            var destinationFilePath = System.IO.Path.Combine(desktopPath, "Final Merge.docx");
            var templateFilePath = System.IO.Path.Combine(desktopPath, "templateFile.docx");

            WordUtility.OpenWordDocument(wordObj, filePath);
            WordUtility.ReplaceText(wordObj.wordDoc, "Hi", "Hello");
            WordUtility.ReplaceImage(wordObj.wordDoc,System.IO.Path.Combine(desktopPath, "Test.png"));
            WordUtility.InsertAPicture(wordObj.wordDoc, System.IO.Path.Combine(desktopPath, "Test.png"));
            WordUtility.SaveWordProcessDocument(wordObj.wordDoc);
            WordUtility.CloseWordProcessDocument(wordObj.wordDoc);
            WordUtility.MergeDocuments(templateFilePath,filesToMerge, destinationFilePath);
            //WordUtility.MergeDocuments(filesToMerge, destinationFilePath);
        }
    }
}
