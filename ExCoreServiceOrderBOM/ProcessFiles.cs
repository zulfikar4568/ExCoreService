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
using CsvHelper.Configuration;
using CsvHelper;

namespace ExCoreServiceOrderBOM
{
    public class ProcessFiles
    {
        public ProcessFiles()
        {
            try
            {
                UNCFolderPath.Connect();
            } catch (Exception ex)
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
                    bResult = ProcessingFileOrderBOM(sFileName, out Message);
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
                                }else
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
        public bool ProcessingFileOrderBOM(string FileName, out string Message)
        {
            // Declare Variable
            Message = "";
            ServiceUtil oServiceUtil = new ServiceUtil();
            var ProductionOrder = new List<string>();
            var Scanning = new List<string>();
            var MaterialGroup = new List<string>();
            var Material = new List<string>();
            var Qty= new List<string>();
            List<MfgOrderChanges> oMfgOrders = new List<MfgOrderChanges>();

            string[] PPAFilter = ConfigurationManager.AppSettings["PPAFilter"].Split(',');
            string ScanningFilter = ConfigurationManager.AppSettings["ScanningFilter"].ToString();
            string DefaultOperationNumber = ConfigurationManager.AppSettings["DefaultOperationNumber"].ToString();
            string OperationNumber = DefaultOperationNumber;
            string DefaultProductDesc = ConfigurationManager.AppSettings["DefaultProductDesc"].ToString();
            string DefaultProductType = ConfigurationManager.AppSettings["DefaultProductType"].ToString();
            try
            {
                ProductTypeMaintService oServiceProduct = new ProductTypeMaintService(AppSettings.ExCoreUserProfile);
                bool ObjectExists = oServiceUtil.ObjectExists(oServiceProduct, new ProductTypeMaint(), DefaultProductType);
                if (!ObjectExists && DefaultProductType != "") oServiceUtil.SaveProductType(DefaultProductType);
            }
            catch (Exception ex)
            {
                EventLogUtil.LogEvent(ex.Message, System.Diagnostics.EventLogEntryType.Warning, 3);
            }

            try
            {
                //Read Csv line
                string[] lineCSV = System.IO.File.ReadAllLines(FileName);
                //Validation
                if (SmartSplit(lineCSV[0], ',').Length - 1 != Convert.ToInt32(ConfigurationManager.AppSettings["LengthCSV"]))
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
                    var records = csv.GetRecords<OrderBOMFormat>();
                    foreach (var data in records)
                    {
                        //Console.WriteLine($"{data.ProductionOrder.TrimStart('0')} - {data.Scanning} - {data.MaterialGroup} - {data.Material} - {data.Qty}");
                        ProductionOrder.Add(data.ProductionOrder.TrimStart('0'));
                        Scanning.Add(data.Scanning);
                        MaterialGroup.Add(data.MaterialGroup);
                        Material.Add(data.Material.TrimStart('0'));
                        Qty.Add(data.Qty);
                    }
                }

                NamedObjectRef[] getMfgOrderList = oServiceUtil.GetListMfgOrder();
                string[] sListtMfgOrder = getMfgOrderList.Where(x => x != null).Select(x => x.Name.ToString()).ToArray();
                string[] mergeMfgOrder = sListtMfgOrder.Intersect(ProductionOrder.Distinct().ToList()).ToArray();
                foreach (var oMfgOrder in mergeMfgOrder)
                {
                    MfgOrderChanges getMfgOrder = oServiceUtil.GetMfgOrder(oMfgOrder);
                    if (getMfgOrder != null)
                    {
                        //Console.WriteLine($"Ada {getMfgOrder.Name}");
                        oMfgOrders.Add(getMfgOrder);
                    }
                    else
                    {
                        EventLogUtil.LogEvent($"Production or Manufacturing Order: {oMfgOrder} is not found!", System.Diagnostics.EventLogEntryType.Warning, 3);
                    }
                }

                foreach (var oMfgOrder in oMfgOrders)
                {
                    //Console.WriteLine($"Still in Processs !!!!!! {oMfgOrder.Name}");
                    ERPRouteChanges oERPRoute = oServiceUtil.GetERPRouteFromMfgOrder(oMfgOrder);
                    if (oERPRoute == null) continue;
                    if (oMfgOrder.Qty == null && oMfgOrder.Containers != null)
                    {
                        EventLogUtil.LogEvent($"Production or Manufacturing Order: {oMfgOrder.Name.Value} can't be used, it might be Production Order have a Container or doesn't have Qty!.\nTry to remove this {oMfgOrder.Name.Value} data on material list.", System.Diagnostics.EventLogEntryType.Warning, 3);
                        continue;
                    }
                    List<dynamic> cMaterialList = new List<dynamic>();
                    for (int j = 0; j < lineCSV.Length - 1; j++)
                    {
                        double number;
                        if (oMfgOrder.Name.ToString() != ProductionOrder[j]) continue;
                        if (!ConvertToDoubleCommaDecimal(Qty[j], out number)) continue;
                        if (Scanning[j] != ScanningFilter) continue;
                        if (PPAFilter.Length > 0)
                        {
                            foreach (var item in PPAFilter)
                            {
                                if (MaterialGroup[j].Contains(item))
                                {
                                    OperationNumber = "2";
                                    break;
                                } else
                                {
                                    OperationNumber = DefaultOperationNumber;
                                }
                            }
                        }
                        ProductMaintService oServiceProduct = new ProductMaintService(AppSettings.ExCoreUserProfile);
                        bool ObjectExists = oServiceUtil.ObjectExists(oServiceProduct, new ProductMaint(), Material[j], "");
                        if (!ObjectExists && MaterialGroup[j] != "" && Material[j] != "")
                        {
                            oServiceUtil.SaveProductFamily(MaterialGroup[j]);
                            oServiceUtil.SaveProduct(Material[j], "1", "", DefaultProductDesc, "", DefaultProductType, "", "", "", "", "", MaterialGroup[j]);
                        }
                        if (oERPRoute.RouteSteps != null)
                        {
                            if (oERPRoute.RouteSteps.Length == 0) continue;
                            foreach (var routeStep in oERPRoute.RouteSteps)
                            {
                                if (routeStep.Sequence == null) continue;
                                if (routeStep.Sequence.Value == OperationNumber && routeStep.Name != null)
                                {
                                    cMaterialList.Add(new MfgOrderMaterialListItmChanges() { Product = new RevisionedObjectRef(Material[j]), QtyRequired = number / oMfgOrder.Qty.Value, IssueControl = IssueControlEnum.LotAndStockPoint, RouteStep = new NamedSubentityRef(routeStep.Name.Value), wikScanning = new Primitive<string>() { Value = Scanning[j] }});
                                    // Console.WriteLine($"{j} | {ProductionOrder[j]} | {OperationNumber} | {Material[j]} | {number / oMfgOrder.Qty.Value}");
                                }
                            }
                        }
                        else
                        {
                            EventLogUtil.LogEvent($"ERP Route: {oERPRoute.Name.Value} doesn't have routeSteps!. Material will included when updated!", System.Diagnostics.EventLogEntryType.Warning, 3);
                        }
                    }
                    bool resultMfgOrder = oServiceUtil.SaveMfgOrder(oMfgOrder.Name.ToString(), "", "", "", "", "", "", 0, cMaterialList, oERPRoute.Name != null ? oERPRoute.Name.Value : "");
                    if (!resultMfgOrder) throw new ArgumentException($"Something wrong when tried to update Manufacturing or Production Order: {oMfgOrder.Name.Value}.\nThe {oMfgOrder.Name.Value} data is the cause of error, try to remove this {oMfgOrder.Name.Value} data on order BOM list.");
                }
                return true;
            } catch(Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
                return false;
            }
        }
    }
}
