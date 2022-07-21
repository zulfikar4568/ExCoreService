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

namespace ExCoreServiceRework
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
                    bResult = ProcessingFileRework(sFileName, out Message);
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
        public bool ProcessingFileRework(string FileName, out string Message)
        {
            // Declare Variable
            Message = "";
            ServiceUtil oServiceUtil = new ServiceUtil();
            bool result = false;
            var PORework = new List<string>();
            var Container = new List<string>();
            string LastPORework = "";
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
                    var records = csv.GetRecords<ReworkFormat>();
                    foreach (var data in records)
                    {
                        Console.WriteLine($"Rework: {data.ReworkPO} - Container: {data.Container}");
                        PORework.Add(data.ReworkPO.TrimStart('0'));
                        Container.Add(data.Container);
                    }
                }

                for (int j = 0; j < lineCSV.Length - 1; j++)
                {
                    if (LastPORework != PORework[j])
                    {
                        MfgOrderChanges bResult = oServiceUtil.GetMfgOrder(PORework[j]);
                        if (bResult == null)
                        {
                            continue;
                        }
                        else
                        {
                            LastPORework = PORework[j];
                        }
                    }
                    bool bContainerExists = oServiceUtil.ContainerExists(Container[j]);
                    if (!bContainerExists)
                    {
                        EventLogUtil.LogEvent($"Container: {Container[j]} doesn't exists!", System.Diagnostics.EventLogEntryType.Warning, 3);
                        continue;
                    }
                    ContainerMaintDetail oMaintDetail = new ContainerMaintDetail() { MfgOrder = new NamedObjectRef(LastPORework) };
                    result = oServiceUtil.ExecuteContainerMaintenance(Container[j], oMaintDetail);
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
