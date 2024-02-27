using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.CompilerServices;
using Newtonsoft.Json;


class ConfigData
{
    public string path { get; set; }
    public string extension { get; set; }
}

class Program
{
    static void DirectoryOpen(string path, string extension)
    {
        Console.WriteLine("kiterjesztes: " + extension);

        DirectoryInfo folder = new DirectoryInfo(path);
        FileInfo[] files = folder.GetFiles("*" + extension, SearchOption.AllDirectories);
        List<string> fileList = new List<string>();
        foreach (FileInfo file in files)
        {
            fileList.Add(file.FullName);
        }

        Console.WriteLine("File kiterjesztese " + extension + ":");
        foreach (string file in fileList)
        {
            Console.WriteLine(file);
        }
    }


    static void SaveToPST(string[] filepaths, string outputPath)
    {
        Outlook.Application outlookApp = new Outlook.Application();
        Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
        Outlook.Folder rootFolder = outlookNamespace.Session.DefaultStore.GetRootFolder() as Outlook.Folder;

        // Ellenőrizzük, hogy a célmappa létezik-e, ha nem, létrehozzuk
        Outlook.Folder targetFolder;
        try
        {
            targetFolder = rootFolder.Folders["em_altal_letrehozva"] as Outlook.Folder;
        }
        catch (System.Exception)
        {
            targetFolder = rootFolder.Folders.Add("em_altal_letrehozva", Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
        }

        // Fájlok feldolgozása és küldése
        for (int i = 0; i < filepaths.Length; i++)
        {
            Console.WriteLine("Processing file: " + filepaths[i]);
            Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mailItem.Subject = Path.GetFileName(filepaths[i]);
            mailItem.Body = "This is the attached file: " + Path.GetFileName(filepaths[i]);
            mailItem.Attachments.Add(filepaths[i]);
            mailItem.Save();
            mailItem.Move(targetFolder);
        }

        outlookNamespace.Logoff();
        outlookApp.Quit();
    }

    static void print_config()
    {
            string filePath = @"C:\em\.config.json";

            // Ellenőrizze, hogy a fájl létezik-e
            if (File.Exists(filePath))
            {
                // Olvassa be a JSON tartalmat a fájlból
                string jsonContent = File.ReadAllText(filePath);

                // Deszerializálja a JSON-t egy objektummá
                ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

                // Most már használhatod a configData objektumot, amely tartalmazza a JSON-ből kiolvasott adatokat
                Console.WriteLine("A config file tartalma:");
                Console.WriteLine("------------------------");
                Console.WriteLine($"path: {configData.path}");
                Console.WriteLine($"extension: {configData.extension}");
            }
            else
            {
                Console.WriteLine("A fájl nem létezik: " + filePath);
            }
    }
    static void go()
    {
        string filePath = @"C:\em\.config.json";
        if (File.Exists(filePath))
        {
            // Olvassa be a JSON tartalmat a fájlból
            string jsonContent = File.ReadAllText(filePath);

            // Deszerializálja a JSON-t egy objektummá
            ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

            // Most már használhatod a configData objektumot, amely tartalmazza a JSON-ből kiolvasott adatokat

            string directoryPath = configData.path;
            string extension = configData.extension;

            string[] filepaths = Directory.GetFiles(directoryPath, "*" + extension, SearchOption.AllDirectories);
            SaveToPST(filepaths, "C:\\Output.pst");
        }
        else
        {
            Console.WriteLine("A fájl nem létezik: " + filePath);
        }
    }
    static void add_extension()
    {
        string filePath = @"C:\em\.config.json";

        if (File.Exists(filePath))
        {
            // Olvassa be a JSON tartalmat a fájlból
            string jsonContent = File.ReadAllText(filePath);

            // Deszerializálja a JSON-t egy objektummá
            ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

            // Kiterjesztes módosítása
            Console.Write("Kiterjesztes:");
            string newExtension = Console.ReadLine();
            configData.extension = newExtension;

            // A fájlba írás előtt JSON formátumra alakítjuk a configData objektumot
            string updatedJsonContent = JsonConvert.SerializeObject(configData, Formatting.Indented);

            // A fájlba írás
            File.WriteAllText(filePath, updatedJsonContent);

            Console.WriteLine("Kiterjesztés módosítva.");
        }
        else
        {
            Console.WriteLine("A fájl nem létezik: " + filePath);
        }
    }


    static void add_spot()
    {
        string filePath = @"C:\em\.config.json";

        if (File.Exists(filePath))
        {
            // Olvassa be a JSON tartalmat a fájlból
            string jsonContent = File.ReadAllText(filePath);

            // Deszerializálja a JSON-t egy objektummá
            ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

            // Path módosítása
            Console.Write("Mappa Eleresi helye:");
            string newPath = Console.ReadLine();
            configData.path = newPath;

            // A fájlba írás előtt JSON formátumra alakítjuk a configData objektumot
            string updatedJsonContent = JsonConvert.SerializeObject(configData, Formatting.Indented);

            // A fájlba írás
            File.WriteAllText(filePath, updatedJsonContent);

            Console.WriteLine("Mappa sikeresen módosítva.");
        }
        else
        {
            Console.WriteLine("A fájl nem létezik: " + filePath);
        }
    }

    static void CheckAndCreateConfigFile()
    {
        string directoryPath = @"C:\em";
        string filePath = @"C:\em\.config.json";

        // Ellenőrizze, hogy létezik-e a könyvtár
         if (!Directory.Exists(directoryPath))
         {
            // Ha nem létezik, hozza létre
            Directory.CreateDirectory(directoryPath);
             Console.WriteLine("Status: Config mappa inicializálás.");
         }

        // Ellenőrizze, hogy létezik-e a fájl
        if (!File.Exists(filePath))
        {
        // Ha nem létezik, hozza létre
             using (FileStream fs = File.Create(filePath))
             Console.WriteLine("Status: Config file configurálás.");
        }
        else
        {
            Console.WriteLine("Status: Üdv újra minden rendben van.\n\n");
        }
    }

    static void Help()
    {
        Console.WriteLine("|=============================|\n|Mappa hely megadas [1]       |\n|Kiterjesztes megadas [2]     |\n|Config File megjelenitese [3]|\n|Help [4]                     |\n|=============================|\n|Go [0]                       |\n|Exit [9]                     |\n|_____________________________|");
    }


    static void Menu()
    {
        Help();
        while (true)
        {
            Console.Write(">>");
            int input = Convert.ToInt32(Console.ReadLine());
            switch (input)
            {
                case 0: go(); break;
                case 1: add_spot(); break;
                case 2: add_extension(); break;
                case 3: print_config(); break;
                case 4: Help(); break;
                case 9: Console.WriteLine("Visszlat"); System.Environment.Exit(0); break;
            }

        }
    }
    static void Main()
    {
            //init check
            CheckAndCreateConfigFile();

            //Menu
            Menu();
    }
}
