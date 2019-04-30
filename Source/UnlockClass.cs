using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Xml;

namespace ExcelUnlockerVisual
{

    class UnlockClass
    {
        const string spinCount = "100000";
        const string salt = "1jO0ukVnaQWZ/hBmt7hXxg==";

        public static int Lock(string outputPath, string filePath, bool overWrite, bool vba, bool veryHidden, string password, IProgress<int> progress, IProgress<int> consoleProg)
        {
            string homeDirectory = "";
            int currentProgress = 0;
            int worksheetProgress = 0;

            string fileName = Path.GetFileName(filePath);

            if (overWrite)
            {
                outputPath = fileName;
            }

            homeDirectory = Path.GetDirectoryName(filePath);
            string directoryID = Guid.NewGuid().ToString();
            string directoryPath = "C:\\Temp\\" + directoryID;
            currentProgress += 5;
            progress.Report(currentProgress);
            string fileExtension = Path.GetExtension(filePath);

            // Uses the C:\Temp\... path from earlier, and makes the working directory
            Directory.CreateDirectory(directoryPath);
            currentProgress += 5;
            progress.Report(currentProgress);

            // Copies the workbook to the working directory as a .zip
            File.Copy(filePath, directoryPath + "\\workBook.zip");
            currentProgress += 5;
            progress.Report(currentProgress);

            // Extracts the .zip contents
            ZipFile.ExtractToDirectory(directoryPath + "\\workBook.zip", directoryPath + "\\workBook");
            currentProgress += 5;
            progress.Report(currentProgress);

            if (veryHidden)
            {
                AddVeryHidden(directoryPath + "\\workBook\\xl\\workbook.xml");
            }

            if (vba)
            {
                if (File.Exists(directoryPath + "\\workBook\\xl\\vbaProject.bin"))
                {
                    try
                    {
                        // Reads vbaProject.bin
                        byte[] buf = File.ReadAllBytes(directoryPath + "\\workBook\\xl\\vbaProject.bin");

                        // Encodes the binary as hex, which allows us to edit the three hex couplets below
                        var str = new SoapHexBinary(buf).ToString();


                        // Find the VBA protection key ("DPB") and replaces it with a nothing key ("DBx")
                        // This causes the VBA editor to go through recovery and delete protection
                        str = str.Replace("444278", "445042");

                        // Writes the hex back to binary into the vbaProject.bin file
                        File.WriteAllBytes(directoryPath + "\\workBook\\xl\\vbaProject.bin", SoapHexBinary.Parse(str).Value);
                    }
                    catch
                    {
                        return 2;
                    }
                }
            }
            currentProgress += 10;
            progress.Report(currentProgress);
            currentProgress += 10;
            progress.Report(currentProgress);

            // Removes workbook-level structure protection
            AddWorkbookProtection(directoryPath + "\\workBook\\xl\\workbook.xml", password);

            // Searches in the decompiled workbook's xl directory for worksheets, and
            string[] worksheets = Directory.GetFiles(directoryPath + "\\workBook\\xl\\worksheets");

            worksheetProgress = 50 / worksheets.Length;

            // Calls the RemoveSheetProtection method for each of them
            foreach (string worksheet in worksheets)
            {
                AddSheetProtection(worksheet);
                currentProgress += worksheetProgress;
                progress.Report(currentProgress);
            }

            // Recompiles the workbook with the newly unprotected sheets
            ZipFile.CreateFromDirectory(directoryPath + "\\workBook", directoryPath + "\\Book1Mod.zip");
            currentProgress += 10;
            progress.Report(currentProgress);

            // Copies the .zip back as an Excel workbook to the home directory
            try
            {
                File.Copy(directoryPath + "\\Book1Mod.zip", outputPath, overWrite);
                progress.Report(100);
                consoleProg.Report(0);
            }
            catch
            {
                Directory.Delete(directoryPath, true);
                progress.Report(100);
                consoleProg.Report(1);
                return 1;
            }


            // Deletes the C:\Temp\... directory.
            Directory.Delete(directoryPath, true);

            return 0;
        }

        public static int Unlock(string outputPath, string filePath, bool overWrite, bool unlockVBA, bool removeVeryHidden, IProgress<int> progress, IProgress<int> consoleProg)
        {
            // Initial variable declarations
            string homeDirectory = "";
            int currentProgress = 0;
            int worksheetProgress = 0;

            // Gets the filename from the provided file path
            string fileName = Path.GetFileName(filePath);

            // Creates what we hope will be the new name of the file
            if (overWrite)
            {
                outputPath = fileName;
            }

            // Saves the home directory to a string
            homeDirectory = Path.GetDirectoryName(filePath);


            // Creates a GUID for the temporary directory, and creates C:\Temp\... filepath
            string directoryID = Guid.NewGuid().ToString();
            string directoryPath = "C:\\Temp\\" + directoryID;
            currentProgress += 5;
            progress.Report(currentProgress);


            // Gets the file extension from the filename
            string fileExtension = Path.GetExtension(filePath);


            // Uses the C:\Temp\... path from earlier, and makes the working directory
            Directory.CreateDirectory(directoryPath);
            currentProgress += 5;
            progress.Report(currentProgress);

            // Copies the workbook to the working directory as a .zip
            File.Copy(filePath, directoryPath + "\\workBook.zip");
            currentProgress += 5;
            progress.Report(currentProgress);

            // Extracts the .zip contents
            ZipFile.ExtractToDirectory(directoryPath + "\\workBook.zip", directoryPath + "\\workBook");
            currentProgress += 5;
            progress.Report(currentProgress);

            if (removeVeryHidden)
            {
                RemoveVeryHidden(directoryPath + "\\workBook\\xl\\workbook.xml");
            }

            if (unlockVBA)
            {
                if (File.Exists(directoryPath + "\\workBook\\xl\\vbaProject.bin"))
                {
                    try
                    {
                        // Reads vbaProject.bin
                        byte[] buf = File.ReadAllBytes(directoryPath + "\\workBook\\xl\\vbaProject.bin");

                        // Encodes the binary as hex, which allows us to edit the three hex couplets below
                        var str = new SoapHexBinary(buf).ToString();


                        // Find the VBA protection key ("DPB") and replaces it with a nothing key ("DBx")
                        // This causes the VBA editor to go through recovery and delete protection
                        str = str.Replace("445042", "444278");

                        // Writes the hex back to binary into the vbaProject.bin file
                        File.WriteAllBytes(directoryPath + "\\workBook\\xl\\vbaProject.bin", SoapHexBinary.Parse(str).Value);
                    }
                    catch
                    {
                        return 2;
                    }
                }
            }
            currentProgress += 10;
            progress.Report(currentProgress);

            // Removes workbook-level structure protection
            RemoveWorkbookProtection(directoryPath + "\\workBook\\xl\\workbook.xml");

            // Searches in the decompiled workbook's xl directory for worksheets, and
            string[] worksheets = Directory.GetFiles(directoryPath + "\\workBook\\xl\\worksheets");

            worksheetProgress = 50 / worksheets.Length;

            // Calls the RemoveSheetProtection method for each of them
            foreach (string worksheet in worksheets)
            {
                RemoveSheetProtection(worksheet);
                currentProgress += worksheetProgress;
                progress.Report(currentProgress);
            }

            // Recompiles the workbook with the newly unprotected sheets
            ZipFile.CreateFromDirectory(directoryPath + "\\workBook", directoryPath + "\\Book1Mod.zip");
            currentProgress += 10;
            progress.Report(currentProgress);

            // Copies the .zip back as an Excel workbook to the home directory
            try
            {
                File.Copy(directoryPath + "\\Book1Mod.zip", outputPath, overWrite);
                progress.Report(100);
                consoleProg.Report(0);
            }
            catch
            {
                Directory.Delete(directoryPath, true);
                progress.Report(100);
                consoleProg.Report(1);
                return 1;
            }


            // Deletes the C:\Temp\... directory.
            Directory.Delete(directoryPath, true);

            return 0;

        }

        static void RemoveSheetProtection(string fileName)
        {
            // Initializes our document, and loads it from the filename
            var doc = new System.Xml.XmlDocument();
            doc.Load(fileName);

            XmlNodeList protections = doc.GetElementsByTagName("sheetProtection");
            try
            {
                foreach (XmlNode element in protections)
                {
                    element.ParentNode.RemoveChild(element);
                }
            }
            catch
            {

            }
            doc.Save(fileName);
        }

        static string GenerateHash(string password)
        {
            var path = AppDomain.CurrentDomain.BaseDirectory;
            var outputBuilder = new StringBuilder();
            Process process = new Process();

            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.Arguments = $"{password} {salt}";
            process.StartInfo.FileName = Path.Combine(path, "generateHash.exe");

            process.EnableRaisingEvents = true;
            process.OutputDataReceived += new DataReceivedEventHandler
            (
                delegate (object sender, DataReceivedEventArgs e)
                {
                    outputBuilder.Append(e.Data);
                }
            );
            process.Start();
            process.BeginOutputReadLine();
            process.WaitForExit();
            process.CancelOutputRead();

            return outputBuilder.ToString();
        }

        static void AddSheetProtection(string fileName)
        {

        }

        static void AddVeryHidden(string fileName)
        {
            var docWorkbook = new System.Xml.XmlDocument();
            docWorkbook.Load(fileName);
            var sheets = docWorkbook.GetElementsByTagName("sheet");
            foreach (XmlElement sheet in sheets)
            {
                if (sheets.Item(0) == sheet) continue;
                sheet.SetAttribute("state", "veryHidden");
            }
            docWorkbook.Save(fileName);
        }

        static void AddWorkbookProtection(string fileName, string password)
        {
            var docWorkbook = new System.Xml.XmlDocument();
            docWorkbook.Load(fileName);
            XmlNode workbook = docWorkbook.DocumentElement;
            XmlElement workbookProtection = workbook.OwnerDocument.CreateElement("workbookProtection", docWorkbook.DocumentElement.NamespaceURI);
            workbookProtection.SetAttribute("workbookAlgorithmName", "SHA-512");
            workbookProtection.SetAttribute("workbookHashValue", GenerateHash(password));
            workbookProtection.SetAttribute("workbookSaltValue", salt);
            workbookProtection.SetAttribute("workbookSpinCount", spinCount);
            workbookProtection.SetAttribute("lockStructure", "1");
            workbook.InsertAfter(workbookProtection, docWorkbook.GetElementsByTagName("xr:revisionPtr")[0]);
            docWorkbook.Save(fileName);
        }

        static void RemoveWorkbookProtection(string fileName)
        {
            var docWorkbook = new System.Xml.XmlDocument();
            docWorkbook.Load(fileName);
            XmlNodeList protections = docWorkbook.GetElementsByTagName("workbookProtection");
            try
            {
                foreach (XmlNode element in protections)
                {
                    element.ParentNode.RemoveChild(element);
                }
            }
            catch
            {

            }
            docWorkbook.Save(fileName);
        }

        static void RemoveVeryHidden(string fileName)
        {
            var docWorkbook = new System.Xml.XmlDocument();
            docWorkbook.Load(fileName);
            var sheets = docWorkbook.GetElementsByTagName("sheet");
            foreach (XmlNode sheet in sheets)
            {
                if (sheet.Attributes["state"] != null)
                {
                    sheet.Attributes.Remove(sheet.Attributes["state"]);
                }
            }
            docWorkbook.Save(fileName);
        }

        static int IndexOfValidFileName(string[] args)
        {
            foreach (string argument in args)
            {
                if (Path.GetExtension(argument) == ".xlsx" || Path.GetExtension(argument) == ".xlsm" || Path.GetExtension(argument) == ".xlam")
                {
                    return Array.IndexOf(args, argument);
                }
            }
            return -1;
        }
    }

}
