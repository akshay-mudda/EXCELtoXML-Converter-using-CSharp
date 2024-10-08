using System;
using System.Configuration;
using System.IO;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace EXCELtoXML_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Read source and destination paths from the app.config file
                string sourceDirectory = ConfigurationManager.AppSettings["SourceDirectory"];
                string destinationDirectory = ConfigurationManager.AppSettings["DestinationDirectory"];

                // Ensure the destination directory exists
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                // Get all Excel files from the source directory
                string[] excelFiles = Directory.GetFiles(sourceDirectory, "*.xlsx");

                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No Excel files found in the source directory.");
                    return;
                }

                foreach (string excelFilePath in excelFiles)
                {
                    try
                    {
                        // Get the file name without extension
                        string fileName = Path.GetFileNameWithoutExtension(excelFilePath);

                        // Set the output XML file path
                        string outputFilePath = Path.Combine(destinationDirectory, fileName + ".xml");

                        // Convert Excel to XML
                        ConvertExcelToXml(excelFilePath, outputFilePath);

                        Console.WriteLine($"Successfully converted '{excelFilePath}' to XML.");

                        // Optionally, delete the source Excel file after conversion
                        File.Delete(excelFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting file {excelFilePath}: {ex.Message}");
                    }
                }

                Console.WriteLine("Conversion process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Method to convert Excel to XML
        public static void ConvertExcelToXml(string excelFilePath, string outputFilePath)
        {
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);  // Get the first worksheet
                var rows = worksheet.RangeUsed().RowsUsed();

                var xmlDocument = new XDocument(new XElement("Root"));

                foreach (var row in rows)
                {
                    if (row.RowNumber() == 1)
                    {
                        continue;  // Skip header row
                    }

                    var rowElement = new XElement("Row");

                    for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
                    {
                        string header = worksheet.Cell(1, col).Value.ToString();
                        string value = row.Cell(col).Value.ToString();
                        rowElement.Add(new XElement(header, value));
                    }

                    xmlDocument.Root.Add(rowElement);
                }

                // Save the XML document
                xmlDocument.Save(outputFilePath);
            }
        }
    }
}