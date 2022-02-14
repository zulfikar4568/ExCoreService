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
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        }
        public bool ProcessingFileMfgOrder(string FileName, out string Message)
        {
            Message = "";
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool result = false;
            var ProductionOrder = new List<string>();
            var Product = new List<string>();
            var Qty = new List<string>();
            var StartTime = new List<string>();
            var EndTime = new List<string>();

            try
            {
                string[] lineCSV = System.IO.File.ReadAllLines(FileName);
                //Validation
                if (lineCSV[0].Split(',').Length - 1 != Convert.ToInt32(ConfigurationManager.AppSettings["LengthCSV"]))
                {
                    Message = $"The Column CSV have wrong number, make sure the number of column CSV is {ConfigurationManager.AppSettings["LengthCSV"]}";
                    return false;
                }

                for (int i = 1; i < lineCSV.Length; i++)
                {
                    string[] rowData = lineCSV[i].Split(',');
                    ProductionOrder.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["PO"])]);
                    Product.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["Product"])]);
                    Qty.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["Qty"])]);
                    StartTime.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["StartTime"])]);
                    EndTime.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["EndTime"])]);

                }
                ProductMaintService oServiceProduct = new ProductMaintService(AppSettings.ExCoreUserProfile);
                OrderStatusMaintService oServiceOrderStatus = new OrderStatusMaintService(AppSettings.ExCoreUserProfile);
                string sOrderStatus = oServiceUtil.ObjectExists(oServiceOrderStatus, new OrderStatusMaint(), AppSettings.DefaultOrderStatus) == true ? AppSettings.DefaultOrderStatus : "";
                for (int i = 0; i < lineCSV.Length - 1; i++)
                {
                    if (oServiceUtil.ObjectExists(oServiceProduct, new ProductMaint(), Product[i], "") && oServiceUtil.CanCovertTo(Qty[i], "System.Double"))
                    {
                        result = oServiceUtil.SaveMfgOrder(ProductionOrder[i],
                            "",
                            "",
                            Product[i],
                            "",
                            "",
                            "",
                            Convert.ToDouble(Qty[i]),
                            null,
                            "",
                            oServiceUtil.IsDate(StartTime[i]) == true ? StartTime[i] : "",
                            oServiceUtil.IsDate(EndTime[i]) == true ? EndTime[i] : "",
                            "",
                            sOrderStatus,
                            true);
                    }

                    if (!result) break;
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
