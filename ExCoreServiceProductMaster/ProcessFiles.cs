using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using Camstar.WCF.ObjectStack;
using Camstar.WCF.Services;
using OpcenterWikLibrary;
using System.Configuration;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;

namespace ExCoreServiceProductMaster
{
    public class ProductFamilyInstance
    {
        public string Name { get; set; }
        public string Description { get; set; }
    }
    public class ProcessFiles
    {
        public ProcessFiles()
        {
            try
            {
                UNCFolderPath.Connect();
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        }
        public async void QueuingFiles(string sourceFolder, string queueFolder, string completedFolder, string errorFolder, string referenceFolder = "")
        {
            await Task.Run(() =>
            {
                try
                {
                    if (referenceFolder != "") if (Directory.GetFiles(referenceFolder, "*.csv").Length > 0) throw new ArgumentException($"There's file on folder {referenceFolder}. Please complete or clean this folder first before continue import Order BOM!");
                    if (Directory.GetFiles(sourceFolder, "*.csv").Length > 0)
                    {
                        string uniqueFolder = $"{queueFolder}\\{Guid.NewGuid()}";
                        Directory.CreateDirectory(uniqueFolder);
                        if (Directory.Exists(uniqueFolder))
                        {
                            foreach (string sFileName in Directory.GetFiles(sourceFolder, "*.csv"))
                            {
                                string sDestinationFileName = uniqueFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                                if (!File.Exists(sDestinationFileName)) File.Move(sFileName, sDestinationFileName);
                            }
                        }
                        ProcessingFile(uniqueFolder, completedFolder, errorFolder);
                    }
                }
                catch (Exception ex)
                {
                    EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
                }
            });
        }
        public void ProcessingFile(string uniqueQueueFolder, string completedFolder, string errorFolder)
        {
            try
            {
                // Retrieve file from Source Folder
                foreach (string sFileName in Directory.GetFiles(uniqueQueueFolder, "*.csv"))
                {
                    string Message = "";
                    bool bResult = false;
                    EventLogUtil.LogEvent("Processing" + sFileName, System.Diagnostics.EventLogEntryType.Information, 3);
                    bResult = ProcessingFileMasterProduct(sFileName, out Message);
                    EventLogUtil.LogEvent("Finish processing file:" + sFileName, System.Diagnostics.EventLogEntryType.Information, 3);

                    // Move the file to either the completed or error depending on result
                    string sDestinationFileName = "";
                    int iFileNo = 0;
                    string sFileExt = sFileName.Substring(sFileName.LastIndexOf(".")); //Get File Extension
                    while (true)
                    {
                        if (bResult) sDestinationFileName = completedFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        else sDestinationFileName = errorFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        if (iFileNo > 0) sDestinationFileName = sDestinationFileName.Substring(0, sDestinationFileName.Length - sFileExt.Length) + " (" + iFileNo.ToString() + ") " + sDestinationFileName.Substring(sDestinationFileName.Length - sFileExt.Length, sFileExt.Length);
                        if (!File.Exists(sDestinationFileName))
                        {
                            try
                            {
                                File.Move(sFileName, sDestinationFileName);
                                // Create an error log file with the last error event log
                                if (!bResult)
                                {
                                    StreamWriter oFile = null;
                                    try
                                    {
                                        string errorMessage = EventLogUtil.LastLogError;
                                        if (EventLogUtil.LastLogError == null) errorMessage = $"Something wrong when tried to processing File: {sFileName}. {Message}";
                                        oFile = new StreamWriter(sDestinationFileName + ".log");
                                        oFile.WriteLine(errorMessage);
                                        throw new ArgumentException($"{errorMessage}.\nMove {sFileName} to {sDestinationFileName}");
                                    }
                                    catch (Exception ex)
                                    {
                                        EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
                                    }
                                    finally
                                    {
                                        if (oFile != null) oFile.Close();
                                        if (oFile != null) oFile.Dispose();
                                    }
                                }
                                else
                                {
                                    EventLogUtil.LogEvent("Move " + sFileName + " to " + sDestinationFileName, System.Diagnostics.EventLogEntryType.Information, 3);
                                }
                            }
                            catch (Exception exFileMoveFailure)
                            {
                                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == exFileMoveFailure.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + exFileMoveFailure.Source, exFileMoveFailure);
                            }
                            break;
                        }
                        iFileNo = iFileNo + 1;
                    }
                }
                if (Directory.Exists(uniqueQueueFolder)) Directory.Delete(uniqueQueueFolder);
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        }
        private bool ConvertToDoubleCommaDecimal(string value, out double result)
        {
            CultureInfo provider = new CultureInfo("en-US");
            NumberStyles styles = NumberStyles.Integer | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;
            bool status = double.TryParse(value, styles, provider, out result);
            return status;
        }
        private string[] SmartSplit(string line, char separator = ',')
        {
            if (line.Contains(';')) separator = ';';
            var inQuotes = false;
            var token = "";
            var lines = new List<string>();
            for (var i = 0; i < line.Length; i++)
            {
                var ch = line[i];
                if (inQuotes) // process string in quotes, 
                {
                    if (ch == '"')
                    {
                        if (i < line.Length - 1 && line[i + 1] == '"')
                        {
                            i++;
                            token += '"';
                        }
                        else inQuotes = false;
                    }
                    else token += ch;
                }
                else
                {
                    if (ch == '"') inQuotes = true;
                    else if (ch == separator)
                    {
                        lines.Add(token);
                        token = "";
                    }
                    else token += ch;
                }
            }
            lines.Add(token);
            return lines.ToArray();
        }
        public bool ProcessingFileMasterProduct(string FileName, out string Message)
        {
            // Declare Variable
            Message = "";
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool result = false;
            var Product = new List<string>();
            var Description = new List<string>();
            var Procurement = new List<string>();
            var UOM = new List<string>();
            var ProductType = new List<string>();
            var ProductFamily = new List<string>();
            var ProductFamilyDescription = new List<string>();
            var ProductFamilyInstance = new List<ProductFamilyInstance>();

            try
            {

                //Read Csv line
                string[] lineCSV = System.IO.File.ReadAllLines(FileName);

                //Validation
                if (SmartSplit(lineCSV[0],',').Length - 1 != Convert.ToInt32(ConfigurationManager.AppSettings["LengthCSV"]))
                {
                    Message = $"The Column CSV have wrong number, make sure the number of column CSV is {ConfigurationManager.AppSettings["LengthCSV"]}";
                    return false;
                }

                //Read Csv line
                var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Encoding = Encoding.UTF8, // Our file uses UTF-8 encoding.
                    Delimiter = lineCSV[0].Contains(';') ? ";" : ","
                };
                using (var reader = new StreamReader(FileName))
                using (var csv = new CsvReader(reader, configuration))
                {
                    var records = csv.GetRecords<ProductFormat>();
                    foreach (var data in records)
                    {
                        //Console.WriteLine($"{data.Product} - {data.Description} - {data.Procurement} - {data.UOM} - {data.ProductType} - {data.ProductFamily} - {data.ProductFamilyDescription}");
                        Product.Add(data.Product.TrimStart('0'));
                        Description.Add(data.Description);
                        Procurement.Add(data.Procurement);
                        UOM.Add(data.UOM);
                        ProductType.Add(data.ProductType);
                        ProductFamily.Add(data.ProductFamily);
                        ProductFamilyDescription.Add(data.ProductType);
                        ProductFamilyInstance.Add(new ProductFamilyInstance { Name = data.ProductFamily, Description = data.ProductFamilyDescription });
                    }
                }

                string[] createUOM = UOM.Distinct().Where(x => !string.IsNullOrEmpty(x)).ToArray();
                foreach (var sUOM in createUOM)
                {
                    Console.WriteLine($"Create UOM {sUOM}");
                    oServiceUtil.SaveUOM(sUOM);
                }

                string[] createProductType = ProductType.Distinct().Where(x => !string.IsNullOrEmpty(x)).ToArray();
                foreach (var sProductType in createProductType)
                {
                    Console.WriteLine($"Create ProductType {sProductType}");
                    oServiceUtil.SaveProductType(sProductType);
                }

                ProductFamilyInstance[] createProductFamily = ProductFamilyInstance.GroupBy(p => p.Name).Select(g => g.First()).ToArray();
                foreach (var sProductFamily in createProductFamily)
                {
                    Console.WriteLine($"Create Product Family {sProductFamily.Name}");
                    oServiceUtil.SaveProductFamily(sProductFamily.Name, sProductFamily.Description);
                }

                for (int j = 0; j < lineCSV.Length - 1; j++)
                {
                    Console.WriteLine($"{Product[j]} - {Description[j]} - {Procurement[j]} - {UOM[j]} - {ProductType[j]} - {ProductFamily[j]}");
                    if (Product[j] != "")
                    {
                        result = oServiceUtil.SaveProduct(Product[j], "1", "", Description[j], "", ProductType[j], "", "", "", "", "", ProductFamily[j], Procurement[j], UOM[j], 1);
                        if (!result) break;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
                return false;
            }
        }
    }
}
