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
                    bResult = ProcessingFileOrderBOM(sFileName, out Message);
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
        public bool ProcessingFileOrderBOM(string FileName, out string Message)
        {
            // Declare Variable
            Message = "";
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool resultMfgOrder = false;
            bool resultQueue = false;
            var ProductionOrder = new List<string>();
            var OperationNumber = new List<string>();
            var PartRequired = new List<string>();
            var Qty= new List<string>();
            List<MfgOrderChanges> oMfgOrders = new List<MfgOrderChanges>();
            try
            {
                //Read Csv line
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
                    OperationNumber.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["OperationNumber"])]);
                    PartRequired.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["PartRequired"])]);
                    Qty.Add(rowData[Convert.ToInt32(ConfigurationManager.AppSettings["Qty"])]);
                }

                foreach (var filteredMfgOrder in ProductionOrder.Distinct().ToList())
                {
                    MfgOrderChanges getMfgOrder = oServiceUtil.GetMfgOrder(filteredMfgOrder);
                    if (getMfgOrder != null)
                    {
                        oMfgOrders.Add(getMfgOrder);
                    }
                    else
                    {
                        EventLogUtil.LogEvent($"Production or Manufacturing Order: {filteredMfgOrder} is not found!", System.Diagnostics.EventLogEntryType.Warning, 3);
                    }
                }
                foreach (var oMfgOrder in oMfgOrders)
                {
                    isMaterialQueueMaintService oQueueService = new isMaterialQueueMaintService(AppSettings.ExCoreUserProfile);
                    bool bQueueExists = oServiceUtil.ObjectExists(oQueueService, new isMaterialQueueMaint(), oMfgOrder.Name.Value);
                    ERPRouteChanges oERPRoute = oServiceUtil.GetERPRouteFromMfgOrder(oMfgOrder);
                    if (oERPRoute != null)
                    {
                        if (bQueueExists)
                        {
                            if (oMfgOrder.Qty != null && oMfgOrder.Containers == null)
                            {
                                List<dynamic> cMaterialList = new List<dynamic>();
                                List<dynamic> cMaterialQueueDetails = new List<dynamic>();
                                int iSequence = 0;
                                for (int j = 0; j < lineCSV.Length - 1; j++)
                                {
                                    if (oMfgOrder.Name.ToString() == ProductionOrder[j])
                                    {
                                        ProductMaintService oServiceProduct = new ProductMaintService(AppSettings.ExCoreUserProfile);
                                        bool ObjectExists = oServiceUtil.ObjectExists(oServiceProduct, new ProductMaint(), PartRequired[j], "");
                                        if (ObjectExists)
                                        {
                                            iSequence++;
                                            if (oERPRoute.RouteSteps != null)
                                            {
                                                if (oERPRoute.RouteSteps.Length > 0)
                                                {
                                                    foreach (var routeStep in oERPRoute.RouteSteps)
                                                    {
                                                        if (routeStep.Sequence != null)
                                                        {
                                                            if (routeStep.Sequence.Value == OperationNumber[j] && routeStep.Name != null && oServiceUtil.CanCovertTo(Qty[j], "System.Double"))
                                                            {
                                                                cMaterialList.Add(new MfgOrderMaterialListItmChanges() { Product = new RevisionedObjectRef(PartRequired[j]), QtyRequired = Convert.ToDouble(Qty[j]) / oMfgOrder.Qty.Value, IssueControl = IssueControlEnum.NoTracking, RouteStep = new NamedSubentityRef(routeStep.Name.Value) });
                                                                cMaterialQueueDetails.Add(new isMaterialQueueDetailsChanges() { isProduct = new RevisionedObjectRef(PartRequired[j]), isQty = Convert.ToDouble(Qty[j]), isQtyAvailable = Convert.ToDouble(Qty[j]), isUOM = new NamedObjectRef(AppSettings.DefaultUOM), isRemovalStrategy = isRemovalStrategyEnum.FIFO, isSequence = iSequence, isConsumedQty = 0, isInventoryLocation = new NamedObjectRef(AppSettings.DefaultInventoryLocation) });
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                EventLogUtil.LogEvent($"ERP Route: {oERPRoute.Name.Value} doesn't have routeSteps!. Material will included when updated!", System.Diagnostics.EventLogEntryType.Warning, 3);
                                            }
                                        }
                                        //Console.WriteLine($"{j} | {ProductionOrder[j]} | {OperationNumber[j]} | {PartRequired[j]} | {Qty[j]}");
                                    }else
                                    {
                                        iSequence = 0;
                                    }
                                }
                                resultQueue = oServiceUtil.SaveManageQueue(oMfgOrder.Name.ToString(), oMfgOrder.Name.ToString(), cMaterialQueueDetails);
                                resultMfgOrder = oServiceUtil.SaveMfgOrder(oMfgOrder.Name.ToString(), "", "", "", "", "", "", 0, cMaterialList, oERPRoute.Name != null ? oERPRoute.Name.Value : "");
                                if (!resultMfgOrder) throw new ArgumentException($"Something wrong when tried to update Manufacturing or Production Order: {oMfgOrder.Name.Value}.\nThe {oMfgOrder.Name.Value} data is the cause of error, try to remove this {oMfgOrder.Name.Value} data on order BOM list.");
                                if (!resultQueue) throw new ArgumentException($"Something wrong when tried to update Queue: {oMfgOrder.Name.Value}.\nThe {oMfgOrder.Name.Value} data is the cause of error, try to remove this {oMfgOrder.Name.Value} data on order BOM list.");
                            }
                            else
                            {
                                EventLogUtil.LogEvent($"Production or Manufacturing Order: {oMfgOrder.Name.Value} can't be used, it might be Production Order have a Container or doesn't have Qty!.\nTry to remove this {oMfgOrder.Name.Value} data on material list.", System.Diagnostics.EventLogEntryType.Warning, 3);
                            }
                        }
                        else
                        {
                            EventLogUtil.LogEvent($"Production or Manufacturing Order: {oMfgOrder.Name.Value} can't be used, this {oMfgOrder.Name.Value} doesn't have Queue.\nTry to remove this {oMfgOrder.Name.Value} data on material list.", System.Diagnostics.EventLogEntryType.Warning, 3);
                        }
                    }
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
