using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDataReader;
using AutoCAD;

namespace AutoCADBatchRename
{
    class Program
    {
        // Configuration variables
        private static readonly string ExcelFilePath = @"C:\Users\ChristianAkabueze\OneDrive - WP3\Desktop\Test\source_excel.xlsx";
        private static readonly string TargetFolder = @"C:\Users\ChristianAkabueze\OneDrive - WP3\Desktop\Test";

        // Excel column names
        private static readonly string SourceColumnName = "Source";
        private static readonly string DrawingNumberColumnName = "Drawing Number";

        // Operation settings
        private static readonly int MaxRetries = 3;
        private static readonly int InterOperationDelay = 500; // milliseconds
        private static readonly int PostSuccessDelay = 1000; // milliseconds
        private static readonly int UpdateTitleblockDelay = 3000; // milliseconds
        private static readonly int CommandStatusCheckInterval = 500; // milliseconds
        private static readonly bool CloseAutoCADWhenDone = false;
        private static readonly string UpdateTitleblockCommand = "Updatetitleblock";
        private static readonly string CommandSuccessMessage = "All attributes updated!";

        // Retry backoff settings
        private static readonly int InitialRetryDelay = 2000; // milliseconds

        static void Main(string[] args)
        {
            DataTable dt = ReadExcelIntoDataTable(ExcelFilePath);
            if (dt == null) return;

            int successCount = 0;
            AcadApplication acadApp = null;

            try
            {
                // Get or create AutoCAD instance once
                try
                {
                    acadApp = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application");
                }
                catch
                {
                    Type acadType = Type.GetTypeFromProgID("AutoCAD.Application");
                    acadApp = (AcadApplication)Activator.CreateInstance(acadType, true);
                }
                acadApp.Visible = true;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        string sourceFile = dt.Rows[i][SourceColumnName].ToString();
                        string rawName = dt.Rows[i][DrawingNumberColumnName].ToString();
                        string newFilePath = Path.Combine(TargetFolder, $"{rawName}.dwg");

                        Console.WriteLine($"\nProcessing row {i + 1}:");
                        Console.WriteLine($"Source: {sourceFile}");
                        Console.WriteLine($"New DWG: {newFilePath}");

                        bool success = DuplicateDrawing(acadApp, sourceFile, newFilePath);

                        if (success)
                        {
                            successCount++;
                            Console.WriteLine("Success: Drawing duplicated.");
                        }
                        else
                        {
                            Console.WriteLine("Failed to duplicate drawing after retries.");
                        }

                        System.Threading.Thread.Sleep(InterOperationDelay);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error on row {i + 1}: {ex.Message}");
                    }
                }
            }
            finally
            {
                if (acadApp != null)
                {
                    if (CloseAutoCADWhenDone)
                    {
                        acadApp.Quit();
                    }
                    Marshal.ReleaseComObject(acadApp);
                }
            }

            Console.WriteLine($"\nProcess complete. Successfully duplicated {successCount} of {dt.Rows.Count} drawings.");
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static DataTable ReadExcelIntoDataTable(string excelFilePath)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    }).Tables[0];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
                return null;
            }
        }

        private static bool DuplicateDrawing(AcadApplication acadApp, string sourceFile, string newFile)
        {
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return false;
            }

            int retryCount = 0;
            AcadDocument doc = null;

            while (retryCount < MaxRetries)
            {
                try
                {
                    Console.WriteLine($"Attempt {retryCount + 1}: Opening DWG: {sourceFile}");
                    doc = acadApp.Documents.Open(sourceFile, false);

                    Console.WriteLine($"Attempt {retryCount + 1}: Saving As: {newFile}");
                    doc.SaveAs(newFile);
                    File.SetAttributes(newFile, File.GetAttributes(newFile) & ~FileAttributes.ReadOnly);

                    // Execute the UpdateTitleblock command
                    Console.WriteLine($"Running {UpdateTitleblockCommand} command...");

                    try
                    {
                        acadApp.ActiveDocument = doc;

                        // Send the command
                        doc.SendCommand($"{UpdateTitleblockCommand} ");

                        // Wait for the command to complete
                        Console.WriteLine($"Waiting {UpdateTitleblockDelay / 1000} seconds for command to complete...");
                        System.Threading.Thread.Sleep(UpdateTitleblockDelay);

                        // Since we can't directly monitor the command line with COM, 
                        // we'll use the delay and assume success if no exceptions occur
                        Console.WriteLine($"{UpdateTitleblockCommand} command execution completed.");

                        // Save the document after the UpdateTitleblock command completes
                        Console.WriteLine("Saving document with updated titleblock...");
                        doc.Save();
                        Console.WriteLine("Document saved successfully.");
                    }
                    catch (Exception cmdEx)
                    {
                        Console.WriteLine($"Error executing {UpdateTitleblockCommand} command: {cmdEx.Message}");
                    }

                    doc.Close(false);
                    Marshal.ReleaseComObject(doc);
                    doc = null;

                    System.Threading.Thread.Sleep(PostSuccessDelay);
                    return true;
                }
                catch (COMException comEx) when (comEx.ErrorCode == -2147417846)
                {
                    Console.WriteLine($"Attempt {retryCount + 1} failed: Application busy");
                    CleanupDocument(ref doc);
                    int backoffDelay = InitialRetryDelay * (retryCount + 1); // Exponential backoff
                    System.Threading.Thread.Sleep(backoffDelay);
                    retryCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error in DuplicateDrawing: {ex.Message}");
                    CleanupDocument(ref doc);
                    return false;
                }
            }
            return false;
        }

        private static void CleanupDocument(ref AcadDocument doc)
        {
            try
            {
                if (doc != null)
                {
                    if (!doc.Equals(null))
                    {
                        doc.Close(false);
                    }
                    Marshal.ReleaseComObject(doc);
                    doc = null;
                }
            }
            catch { /* Suppress cleanup errors */ }
        }
    }
}