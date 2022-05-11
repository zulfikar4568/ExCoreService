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

namespace ExCoreServiceOrder
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
                    bResult = ProcessingFileMfgOrder(sFileName, out Message);
                    EventLogUtil.LogEvent("Finish processing file:" + sFileName, System.Diagnostics.EventLogEntryType.Information, 3);

                    // Move the file to either the completed or error depending on result
                    string sDestinationFileName = "";
                    int iFileNo = 0;
                    string sFileExt = sFileName.Substring(sFileName.LastIndexOf(".")); //Get File Extension
                    while (true)
                    {
                        if (bResult) sDestinationFileName = completedFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        else sDestinationFileName = errorFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        if (iFileNo > 0)sDestinationFileName = sDestinationFileName.Substring(0, sDestinationFileName.Length - sFileExt.Length) + " (" + iFileNo.ToString() + ") " + sDestinationFileName.Substring(sDestinationFileName.Length - sFileExt.Length, sFileExt.Length);
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
                                } else
                                {
                                    EventLogUtil.LogEvent(Message + "  Move " + sFileName + " to " + sDestinationFileName, System.Diagnostics.EventLogEntryType.Information, 3);
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
        public bool ProcessingFileMfgOrder(string FileName, out string Message)
        {
            Message = "";
            bool bStatus = true;
            ServiceUtil oServiceUtil = new ServiceUtil();
            var MfgLine = new List<string>();
            var MfgOrder = new List<string>();
            var Product = new List<string>();
            var OrderType = new List<string>();
            var Qty = new List<string>();
            var StartTime = new List<string>();
            var EndTime = new List<string>();
            var OrderStatus = new List<string>();

            string[] FilterMfgLines = ConfigurationManager.AppSettings["FilterWorkCenter"].Split(',');
            string[] FilterOrderTypes = ConfigurationManager.AppSettings["FilterOrderType"].Split(',');
            string[] FilterOrderStatus = ConfigurationManager.AppSettings["FilterSystemStatus"].Split(',');

            try
            {
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
                    var records = csv.GetRecords<OrderFormat>();
                    foreach (var data in records)
                    {
                        //Console.WriteLine($"{data.WorkCenter} - {data.Order} - {data.Material} - {data.OrderType} - {data.TargetQty} - {data.StartTime} - {data.EndTime}");
                        MfgLine.Add(data.WorkCenter);
                        MfgOrder.Add(data.Order.TrimStart('0'));
                        Product.Add(data.Material.TrimStart('0'));
                        OrderType.Add(data.OrderType);
                        Qty.Add(data.TargetQty);
                        StartTime.Add(data.StartTime);
                        EndTime.Add(data.EndTime);
                        OrderStatus.Add(data.SystemStatus);
                    }
                }

                ProductMaintService oServiceProduct = new ProductMaintService(AppSettings.ExCoreUserProfile);
                for (int i = 0; i < lineCSV.Length - 1; i++)
                {
                    double number;
                    if (ConvertToDoubleCommaDecimal(Qty[i], out number))
                    {
                        bool result = false;

                        // Validation or Filtering Data
                        string sMfgLineChecked = "";
                        string sOrderTypeChecked = "";
                        string sOrderStatusChecked = "";
                        int countOrderStatus = 0;
                        
                        if (FilterOrderTypes.Length > 0)
                        {
                            foreach (var item in FilterOrderTypes)
                            {
                                if (OrderType[i].Contains(item))
                                {
                                    sOrderTypeChecked = OrderType[i]; 
                                    break;
                                }
                            }
                        }
                        if (sOrderTypeChecked == "")
                        {
                            EventLogUtil.LogEvent($"{MfgOrder[i]} OrderType {OrderType[i]} skipped, cause there's not on the filter list!", System.Diagnostics.EventLogEntryType.Warning, 3);
                            continue;
                        }
                        
                        if (FilterMfgLines.Length > 0)
                        {
                            foreach (var item in FilterMfgLines)
                            {
                                if (MfgLine[i].Contains(item))
                                {
                                    sMfgLineChecked = MfgLine[i]; 
                                    break;
                                }
                            }
                        }
                        if (sMfgLineChecked == "")
                        {
                            EventLogUtil.LogEvent($"{MfgOrder[i]} MfgLine or WorkCenter {MfgLine[i]} skipped, cause there's not on the filter list!", System.Diagnostics.EventLogEntryType.Warning, 3);
                            continue;
                        }

                        
                        string[] OrderStatusArray = OrderStatus[i].Split(' ');
                        if (OrderStatusArray.Length > 0)
                        {
                            foreach (var itemActual in OrderStatusArray)
                            {
                                foreach (var itemConfig in FilterOrderStatus)
                                {
                                    if (itemActual == itemConfig)
                                    {
                                        sOrderStatusChecked = itemActual;
                                        countOrderStatus++;
                                    }
                                }
                            }
                        }
                        if (countOrderStatus != 1)
                        {
                            EventLogUtil.LogEvent($"{MfgOrder[i]} SystemStatus skipped, cause there's no on the filter list or container more than on the list!", System.Diagnostics.EventLogEntryType.Warning, 3);
                            continue;
                        }
                        Console.WriteLine($"{i}. - {MfgOrder[i]} - {Product[i]} - {number} - {StartTime[i]} - {EndTime[i]} - {sOrderStatusChecked} - {sOrderTypeChecked} - {sMfgLineChecked}");
                        if (sOrderStatusChecked != "") oServiceUtil.SaveOrderStatus(sOrderStatusChecked, isOrderStateEnum.Open);
                        if (sOrderTypeChecked != "") oServiceUtil.SaveOrderType(sOrderTypeChecked);
                        if (sMfgLineChecked != "") oServiceUtil.SaveMfgLine(sMfgLineChecked);
                        if (oServiceUtil.ObjectExists(oServiceProduct, new ProductMaint(), Product[i], "") && MfgOrder[i] != "")
                        {
                            Console.WriteLine($"{i}. - {MfgOrder[i]} - {Product[i]} - {number} - { StartTime[i]} - {EndTime[i]} - {OrderStatus[i]} - {OrderType[i]} - {MfgLine[i]}");
                            result = oServiceUtil.SaveMfgOrder(MfgOrder[i],
                                "",
                                "",
                                Product[i],
                                "",
                                "",
                                "",
                                number,
                                null,
                                "",
                                oServiceUtil.IsDate(StartTime[i]) == true ? StartTime[i] : "",
                                oServiceUtil.IsDate(EndTime[i]) == true ? EndTime[i] : "",
                                "",
                                sOrderStatusChecked,
                                null,
                                sOrderTypeChecked,
                                sMfgLineChecked,
                                false);
                            if (!result)
                            {
                                bStatus = false;
                                Message = $"Something Wrong when import the Order! {MfgOrder[i]}";
                                break;
                            }
                        }
                    }
                }
                Message += "Finish Looping All the rows!";
                return bStatus;
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
                return false;
            }
        }
        private bool ConvertToDoubleCommaDecimal(string value, out double result)
        {
            CultureInfo provider = new CultureInfo("en-US");
            NumberStyles styles = NumberStyles.Integer | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;
            bool status = double.TryParse(value, styles, provider, out result);
            return status;
        }
    }
}
