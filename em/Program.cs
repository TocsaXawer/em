using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

class Program
{
    static string DirectoryOpen(string path, string extension)
    {
        Console.WriteLine("kiterjesztes: " + extension);

        DirectoryInfo folder = new DirectoryInfo(path);
        FileInfo[] files = folder.GetFiles("*" + extension, SearchOption.AllDirectories);
        List<string> fileList = new List<string>();
        foreach (FileInfo file in files)
        {
            fileList.Add(file.FullName);
        }

        StringBuilder result = new StringBuilder();
        result.AppendLine("File kiterjesztese " + extension + ":");
        foreach (string file in fileList)
        {
            result.AppendLine(file);
        }

        return result.ToString();
    }

    static void SaveToPST(string[] filepaths, string outputPath)
    {
        Outlook.Application outlookApp = new Outlook.Application();
        Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
        Outlook.Folder rootFolder = outlookNamespace.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
        Outlook.Folder targetFolder = rootFolder.Folders.Add("Hianyzo mappa", Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;

        foreach (string filepath in filepaths)
        {
            Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mailItem.Subject = Path.GetFileName(filepath);
            mailItem.Body = "Ez az attach file: " + Path.GetFileName(filepath);
            mailItem.Attachments.Add(filepath);
            mailItem.Save();
            mailItem.Move(targetFolder);
        }

        outlookNamespace.Logoff();
        outlookApp.Quit();
    }

    static void Main()
    {
        string directoryPath = "C:\\Users\\Tóth Csaba János\\source\\repos\\em\\em\\test";
        string extension = ".em";
        string[] filepaths = Directory.GetFiles(directoryPath, "*" + extension, SearchOption.AllDirectories);

        Console.WriteLine(DirectoryOpen("C:\\Users\\Tóth Csaba János\\source\\repos\\em\\em\\test", ".em"));

        SaveToPST(filepaths, "C:\\PST\\Output.pst");
        Console.WriteLine("File mentve egy PST-be.");
       
    }
}