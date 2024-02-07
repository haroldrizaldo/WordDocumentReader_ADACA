using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        string folderPath = @"C:\Users\JohnHaroldERizaldo\Desktop\WordDocumentReader_ADACA\SampleDirectory";

        try
        {
            // Scan folders & subfolders recursively
            ScanFolder(folderPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error scanning the directory {folderPath}: {ex.Message}");
        }
    }

    static void ScanFolder(string folderPath)
    {
        // Scan for Microsoft Word Files in the Directory Path
        foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
        {
            ProcessDocument(filePath);
        }

        // Recursively process subfolders
        foreach (string subfolderPath in Directory.GetDirectories(folderPath))
        {
            ScanFolder(subfolderPath);
        }
    }

    static void ProcessDocument(string filePath)
    {
        using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
        {
            //Get the document body of the Word file
            string fileName = Path.GetFileName(filePath);

            var bodyProperties = wordDocument.MainDocumentPart.Document.Body;

            Console.WriteLine(fileName);

            try
            {
                if (bodyProperties != null)
                {
                    Console.WriteLine($"List of Document properties:");
                    //Get list of elements inside the document body
                    foreach (var element in bodyProperties.Elements())
                    {
                        // Check if the paragraph contains a field with a cross-reference
                        if (element.InnerText.Contains("REF"))
                        {
                            Console.WriteLine($"{element} | Cross-reference: YES");
                        }
                        else
                        {
                            Console.WriteLine($"{element} | Cross-reference: NO");
                        }
                    }
                    //Separator
                    Console.WriteLine("------------------------------------");

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing document {filePath}: {ex.Message}");
            }
        }
    }

}
