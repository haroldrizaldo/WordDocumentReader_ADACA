using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Directory Path containing folders, subfolders & MS Word file(s)
        string folderPath = @"C:\Users\JohnHaroldERizaldo\Desktop\WordDocumentReader_ADACA\SampleDirectory";

        try
        {
            //Scan folders & subfolders recursively
            ScanFolder(folderPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error scanning the diretory {folderPath}: {ex.Message}");
        }
    }

    static void ScanFolder(string folderPath)
    {
        // Scan for Microsoft Word Files in the Directory Path
        foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
        {
            PrintFileProperties(filePath);
        }

        // Recursively process subfolders
        foreach (string subfolderPath in Directory.GetDirectories(folderPath))
        {
            ScanFolder(subfolderPath);
        }
    }

    static void PrintFileProperties(string filePath)
    {
        Console.WriteLine($"File Properties for: {filePath}");

        try
        {
            // Get Word file properties
            FileInfo fileInfo = new FileInfo(filePath);

            // Print Word file properties
            Console.WriteLine($"File Name: {fileInfo.Name}");
            Console.WriteLine($"File Size: {fileInfo.Length} bytes");
            Console.WriteLine($"Creation Time: {fileInfo.CreationTime}");
            Console.WriteLine($"Last Access Time: {fileInfo.LastAccessTime}");
            Console.WriteLine($"Last Write Time: {fileInfo.LastWriteTime}");
            //Separator
            Console.WriteLine($"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
        }
    }
}
