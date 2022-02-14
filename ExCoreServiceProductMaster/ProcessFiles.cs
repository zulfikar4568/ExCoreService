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

namespace ExCoreServiceProductMaster
{
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
        //Services mode can be: MfgOrder, OrderBOM, MasterProduct
        public void ProcessingFile(string sourceFolder, string completedFolder, string errorFolder)
        {
            try
            {
                // Retrieve file from Source Folder
                foreach (string sFileName in Directory.GetFiles(sourceFolder, "*.csv"))
                {
                    bool bResult = false;
                    EventLogUtil.LogEvent("Processing" + sFileName, System.Diagnostics.EventLogEntryType.Information, 3);
                    bResult = ProcessingFileMasterProduct(sFileName);
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
                                        if (EventLogUtil.LastLogError == null) errorMessage = $"Something wrong when tried to processing File: {sFileName}.";
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
            }
            catch (Exception ex)
            {
                EventLogUtil.LogErrorEvent(AppSettings.AssemblyName == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source, ex);
            }
        }
        public bool ProcessingFileMasterProduct(string FileName)
        {
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool result = false;
            string[] lineCSV = System.IO.File.ReadAllLines(FileName);
            var ProductNumber = new List<string>();
            var Description = new List<string>();
            var ProductType = new List<string>();

            for (int i = 1; i < lineCSV.Length; i++)
            {
                string[] rowData = lineCSV[i].Split(',');
                ProductNumber.Add(rowData[0]);
                Description.Add(rowData[1]);
                ProductType.Add(rowData[2]);
            }

            for (int j = 0; j < lineCSV.Length - 1; j++)
            {
                result = oServiceUtil.SaveProduct(ProductNumber[j], "1", "", Description[j], "", ProductType[j]);
                if (!result) break;
            }
            return result;
        }
    }
}
