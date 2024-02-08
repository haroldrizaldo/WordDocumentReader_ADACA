using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
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
            // Initialize dictionaries to store document properties and their usage
            Dictionary<string, List<string>> propertyToDocuments = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> documentProperties = new Dictionary<string, List<string>>();

            // Scan folders & subfolders recursively
            ScanFolder(folderPath, propertyToDocuments, documentProperties);

            // Print cross-reference
            PrintCrossReference(propertyToDocuments, documentProperties);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error scanning the directory {folderPath}: {ex.Message}");
        }
    }

    static void ScanFolder(string folderPath, Dictionary<string, List<string>> propertyToDocuments, Dictionary<string, List<string>> documentProperties)
    {
        // Scan for Microsoft Word Files in the Directory Path
        foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
        {
            ProcessDocument(filePath, propertyToDocuments, documentProperties);
        }

        // Recursively process subfolders
        foreach (string subfolderPath in Directory.GetDirectories(folderPath))
        {
            ScanFolder(subfolderPath, propertyToDocuments, documentProperties);
        }
    }

    static void ProcessDocument(string filePath, Dictionary<string, List<string>> propertyUsage, Dictionary<string, List<string>> documentProperties)
    {
        try
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {

                //Get the Document Body
                var documentBody = wordDoc.MainDocumentPart.Document.Body.InnerXml;
                if (documentBody != null)
                {
                    //Get the list of Document Properties
                    PropertyInfo[] properties = typeof(PackageProperties).GetProperties();
                    foreach (PropertyInfo property in properties)
                    {
                        string propertyName = property.Name;
                        //Check if the Document Property is in the Document Body
                        if (documentBody.Contains(propertyName))
                        {
                            //Add in the List
                            if (!propertyUsage.ContainsKey(propertyName))
                                propertyUsage[propertyName] = new List<string>();

                            if (!propertyUsage[propertyName].Contains(Path.GetFileName(filePath)))
                                propertyUsage[propertyName].Add(Path.GetFileName(filePath));

                            if (!documentProperties.ContainsKey(Path.GetFileName(filePath)))
                                documentProperties[Path.GetFileName(filePath)] = new List<string>();

                            if (!documentProperties[Path.GetFileName(filePath)].Contains(propertyName))
                                documentProperties[Path.GetFileName(filePath)].Add(propertyName);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing document {filePath}: {ex.Message}");
        }
    }

    static void PrintCrossReference(Dictionary<string, List<string>> propertyUsage, Dictionary<string, List<string>> documentProperties)
    {
        // Print the cross-reference
        foreach (var propUsage in propertyUsage)
        {
            Console.WriteLine($"Document Property: {propUsage.Key}");
            Console.WriteLine("   - Used in:");

            foreach (var docName in propUsage.Value)
            {
                Console.WriteLine($"       - {docName}");
            }
        }

        Console.WriteLine();
    }
}
