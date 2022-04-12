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
        //Services mode can be: MfgOrder, OrderBOM, MasterProduct
        public void ProcessingFile(string sourceFolder, string completedFolder, string errorFolder)
        {
            try
            {
                // Retrieve file from Source Folder
                foreach (string sFileName in Directory.GetFiles(sourceFolder, "*.csv"))
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
                        if (bResult)
                        {
                            sDestinationFileName = completedFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        }
                        else
                        {
                            sDestinationFileName = errorFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        }
                        if (iFileNo > 0)
                        {
                            sDestinationFileName = sDestinationFileName.Substring(0, sDestinationFileName.Length - sFileExt.Length) + " (" + iFileNo.ToString() + ") " + sDestinationFileName.Substring(sDestinationFileName.Length - sFileExt.Length, sFileExt.Length);
                        }
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
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        }
        private static string[] SmartSplit(string line, char separator = ',')
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

                for (int i = 1; i < lineCSV.Length; i++)
                {
                    string[] rowData = SmartSplit(lineCSV[i], ',');
                    MfgLine.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["WorkCenter"])]);
                    MfgOrder.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["Order"])]);
                    Product.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["Material"])]);
                    OrderType.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["OrderType"])]);
                    Qty.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["TargetQty"])]);
                    StartTime.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["StartTime"])]);
                    EndTime.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["EndTime"])]);
                    OrderStatus.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["SystemStatus"])]);

                }
                ProductMaintService oServiceProduct = new ProductMaintService(AppSettings.ExCoreUserProfile);
                for (int i = 0; i < lineCSV.Length - 1; i++)
                {
                    double number;
                    if (ConvertToDoubleCommaDecimal(Qty[i], out number))
                    {
                        bool result = false;
                        //Console.WriteLine($"{i}. - {MfgOrder[i]} - {Product[i]} - {number} - { StartTime[i]} - {EndTime[i]} - {OrderStatus[i]} - {OrderType[i]} - {MfgLine[i]}");

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
                        if (oServiceUtil.ObjectExists(oServiceProduct, new ProductMaint(), Product[i], ""))
                        {
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
                                true);
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
