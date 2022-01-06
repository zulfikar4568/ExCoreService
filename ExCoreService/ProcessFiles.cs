using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using Camstar.WCF.ObjectStack;
using Camstar.WCF.Services;

namespace ExCoreService
{
    public class ProcessFiles
    {
        public ProcessFiles()
        {
            try
            {
                NetUtil.Connect(AppSettings.SourceUNCPath, AppSettings.SourceUNCPathUsername, AppSettings.SourceUNCPathPassword);
                NetUtil.Connect(AppSettings.CompletedUNCPath, AppSettings.CompletedUNCPathUsername, AppSettings.CompletedUNCPathPassword);
                NetUtil.Connect(AppSettings.ErrorUNCPath, AppSettings.ErrorUNCPathUsername, AppSettings.SourceUNCPathPassword);
                if (!Directory.Exists(AppSettings.SourceFolder)) Directory.CreateDirectory(AppSettings.SourceFolder);
                if (!Directory.Exists(AppSettings.CompletedFolder)) Directory.CreateDirectory(AppSettings.CompletedFolder);
                if (!Directory.Exists(AppSettings.ErrorFolder)) Directory.CreateDirectory(AppSettings.ErrorFolder);
            } catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        } 
        public void ProcessingFile()
        {
            try
            {
                // Retrieve file from Source Folder
                foreach (string sFileName in Directory.GetFiles(AppSettings.SourceFolder, "*.csv"))
                {
                    EventLogUtil.LogEvent("Processing" + sFileName, System.Diagnostics.EventLogEntryType.Information, 3);
                    bool bResult = ProcessingFileOrderBOM(sFileName);
                    EventLogUtil.LogEvent("Completed" + sFileName, System.Diagnostics.EventLogEntryType.Information, 3);

                    // Move the file to either the completed or error depending on result
                    string sDestinationFileName = "";
                    int iFileNo = 0;
                    string sFileExt = sFileName.Substring(sFileName.LastIndexOf(".")); //Get File Extension
                    while (true)
                    {
                        if (bResult)
                        {
                            sDestinationFileName = AppSettings.CompletedFolder + "\\" + System.IO.Path.GetFileName(sFileName);
                        }
                        else
                        {
                            sDestinationFileName = AppSettings.ErrorFolder + "\\" + System.IO.Path.GetFileName(sFileName);
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
                                EventLogUtil.LogEvent("Move" + sFileName + "to" + sDestinationFileName, System.Diagnostics.EventLogEntryType.Information, 3);
                                // Create an error log file with the last error event log
                                if (!bResult)
                                {
                                    StreamWriter oFile = null;
                                    try
                                    {
                                        oFile = new StreamWriter(sDestinationFileName + ".log");
                                        oFile.WriteLine(EventLogUtil.LastLogError);
                                    }
                                    catch { }
                                    finally
                                    {
                                        if (oFile != null) oFile.Close();
                                        if (oFile != null) oFile.Dispose();
                                    }
                                }
                            }
                            catch (Exception exFileMoveFailure)
                            {
                                EventLogUtil.LogErrorEvent(typeof(Program).Assembly.GetName().Name == exFileMoveFailure.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + exFileMoveFailure.Source, exFileMoveFailure);
                            }
                            break;
                        }
                        iFileNo = iFileNo + 1;
                    }
                }
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        }
        public bool ProcessingFileOrderBOM(string FileName)
        {
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool result = false;
            string[] lineCSV = System.IO.File.ReadAllLines(FileName);
            var ProductionOrder = new List<string>();
            var OperationNumber = new List<string>();
            var PartRequired = new List<string>();
            var Qty = new List<string>();
            List<MfgOrderChanges> oMfgOrders = new List<MfgOrderChanges>();

            for (int i = 1; i < lineCSV.Length; i++)
            {
                string[] rowData = lineCSV[i].Split(',');
                ProductionOrder.Add(rowData[0]);
                OperationNumber.Add(rowData[1]);
                PartRequired.Add(rowData[2]);
                Qty.Add(rowData[3]);

            }
            var UniqueMfgOrder = ProductionOrder.Distinct().ToList();
            for (int i = 0; i < UniqueMfgOrder.Count; i++)
            {
                oMfgOrders.Add(oServiceUtil.GetMfgOrder(UniqueMfgOrder[i]));
            }
            foreach (var oMfgOrder in oMfgOrders)
            {
                if (oMfgOrder != null && oMfgOrder.Qty.Value != 0)
                {
                    List<dynamic> cMaterialList = new List<dynamic>();
                    for (int j = 0; j < lineCSV.Length - 1; j++)
                    {
                        if (oMfgOrder.Name.ToString() == ProductionOrder[j])
                        {
                            ProductMaintService oServiceProduct = new ProductMaintService(AppSettings.ExCoreUserProfile);
                            bool ObjectExists = oServiceUtil.ObjectExists(oServiceProduct, new ProductMaint(), PartRequired[j], "");
                            if (ObjectExists)
                            {
                                cMaterialList.Add(new MfgOrderMaterialListItmChanges() { Product = new RevisionedObjectRef(PartRequired[j]), QtyRequired = Convert.ToDouble(Qty[j]) / oMfgOrder.Qty.Value, IssueControl = IssueControlEnum.LotAndStockPoint });
                            }
                            Console.WriteLine($"{j} | {ProductionOrder[j]} | {OperationNumber[j]} | {PartRequired[j]} | {Qty[j]}");
                        }
                    }
                    result = oServiceUtil.SaveMfgOrder(oMfgOrder.Name.ToString(), "", "", "", "", "", "", 0, cMaterialList);
                    if (!result) break;
                }
            }
            return result;
        }
        public bool ProcessingFileMfgOrder(string FileName)
        {
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool result = false;
            string[] lineCSV = System.IO.File.ReadAllLines(FileName);
            var ProductionOrder = new List<string>();
            var Product = new List<string>();
            var Workflow = new List<string>();
            var Qty = new List<string>();
            var StartTime = new List<string>();
            var EndTime = new List<string>();


            for (int i = 1; i < lineCSV.Length; i++)
            {
                string[] rowData = lineCSV[i].Split(',');
                ProductionOrder.Add(rowData[0]);
                Product.Add(rowData[1]);
                Workflow.Add(rowData[5]);
                Qty.Add(rowData[6]);
                StartTime.Add(rowData[8]);
                EndTime.Add(rowData[9]);

            }
            Console.WriteLine("Production Order | Description | Product | Qty | ");
            for (int i = 0; i < lineCSV.Length - 1; i++)
            {
                Console.WriteLine($"{i} | {ProductionOrder[i]} | {Product[i]} | {Workflow[i]} | {Qty[i]} | {StartTime[i]} |{ EndTime[i]} |");
                result = oServiceUtil.SaveMfgOrder(ProductionOrder[i], "", "", Product[i], "", Workflow[i],"", Convert.ToDouble(Qty[i]), null, StartTime[i] , EndTime[i], "", "Released", true);
                if (!result) break;
            }
            return result;
        }
    }
}
