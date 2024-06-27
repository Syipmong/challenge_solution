using System;
using System.IO;
using System.Linq;

// Shehu Yipmong Said BU/22C/IT/7533
public class OfficeFileSummary
{
    public static void Main()
    {
        // Define the directory name where files are stored
        string directoryName = "FileCollection";
        // Define the results file name where the summary will be saved
        string resultsFileName = "results.txt";

        // Ensure the directory exists
        Directory.CreateDirectory(directoryName);

        // Initialize counters and variables
        int xlsxCount = 0, docxCount = 0, pptxCount = 0;
        long xlsxSize = 0, docxSize = 0, pptxSize = 0;

        // Create a DirectoryInfo object to access the specified directory
        DirectoryInfo dirInfo = new DirectoryInfo(directoryName);

        // Enumerate files in the directory
        foreach (FileInfo file in dirInfo.GetFiles())
        {
            // Check if the file is an Office file
            if (IsOfficeFile(file))
            {
                switch (file.Extension.ToLower())
                {
                    case ".xlsx":
                        xlsxCount++;
                        xlsxSize += file.Length;
                        break;
                    case ".docx":
                        docxCount++;
                        docxSize += file.Length;
                        break;
                    case ".pptx":
                        pptxCount++;
                        pptxSize += file.Length;
                        break;
                }
            }
        }

        // Write results to file
        using (StreamWriter writer = new StreamWriter(resultsFileName))
        {
            writer.WriteLine("Office File Summary:");
            writer.WriteLine($"Excel files (.xlsx): {xlsxCount}, Total size: {xlsxSize} bytes");
            writer.WriteLine($"Word files (.docx): {docxCount}, Total size: {docxSize} bytes");
            writer.WriteLine($"PowerPoint files (.pptx): {pptxCount}, Total size: {pptxSize} bytes");
        }

        Console.WriteLine($"Results written to {resultsFileName}");
    }

    private static bool IsOfficeFile(FileInfo file)
    {
        // Helper function to check if a file is an Office file based on its extension
        string[] officeExtensions = { ".xlsx", ".docx", ".pptx" };
        return officeExtensions.Contains(file.Extension.ToLower());
    }
}
