﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Camstar.WCF.Services;
using Camstar.WCF.ObjectStack;
using Camstar.Util;
using Camstar.Exceptions;

namespace ExCoreService
{
    class ServiceUtil
    {
        #region REUSABLE FUNCTION
        public string LastResultMessage = "";
        public bool ProcessResult(ResultStatus Result, ref string ResultMessage, bool IgnoreException = true)
        {
            try
            {
                ResultMessage = "";
                if (Result is null) { return false; }
                if (!Result.IsSuccess)
                {
                    ExceptionDataType oExceptionData = Result.ExceptionData;
                    if (oExceptionData is null)
                    {
                        ResultMessage = Result.Message;
                    }
                    else
                    {
                        ResultMessage = oExceptionData.Description;
                    }
                    throw new Exception(ResultMessage);
                }
                else
                {
                    ResultMessage = Result.Message;
                }
                return Result.IsSuccess;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                if (!IgnoreException) throw (ex);
                return false;
            }
            finally
            {
                LastResultMessage = ResultMessage;
            }
        }
        public bool ObjectExists(dynamic ServiceRef, dynamic ServiceObject, string Name)
        {
            string sMessage = "";
            ResultStatus oResultStatus;
            ServiceObject.ObjectToChange = new NamedObjectRef(Name);
            oResultStatus = ServiceRef.Load(ServiceObject);
            return ProcessResult(oResultStatus, ref sMessage, true);
        }
        public bool ObjectExists(dynamic ServiceRef, dynamic ServiceObject, string Name, string Revision)
        {
            ResultStatus oResultStatus;
            if (Revision != "")
            {
                ServiceObject.ObjectToChange = new RevisionedObjectRef(Name, Revision);
            }
            else
            {
                ServiceObject.ObjectToChange = new RevisionedObjectRef(Name);
            }
            oResultStatus = ServiceRef.Load(ServiceObject);
            string sMessage = "";
            return ProcessResult(oResultStatus, ref sMessage, true);
        }
        public DataPointSummary GetDataPointSummaryRef(dynamic Service, dynamic ServiceObject, dynamic ServiceObject_Request, dynamic ServiceObject_Info, ref string DataCollectionName, ref string DataCollectionRev)
        {
            string sMessage = "";
            MoveIn_Result oServiceObject_Result = null;
            ServiceObject_Request.Info = ServiceObject_Info;
            if (DataCollectionName != "")
            {
                ServiceObject.DataCollectionDef = new RevisionedObjectRef() { Name = DataCollectionName, Revision = DataCollectionRev, RevisionOfRecord = (DataCollectionRev == "") };
            }
            else
            {
                ServiceObject_Request.Info.DataCollectionDef = new Info(true);
            }
            ServiceObject_Request.Info.ParametricData = new ParametricData_Info();
            ServiceObject_Request.Info.ParametricData.RequestValue = true;
            ResultStatus oResultStatus = Service.GetDataPoints(ServiceObject, ServiceObject_Request, out oServiceObject_Result);
            if (ProcessResult(oResultStatus, ref sMessage, false))
            {
                if (oServiceObject_Result.Value != null)
                {
                    if ((oServiceObject_Result.Value.DataCollectionDef != null) && (oServiceObject_Result.Value.ParametricData != null))
                    {
                        var withBlock = (DataPointSummary)oServiceObject_Result.Value.ParametricData;
                        if (withBlock.DataPointDetails.Count() > 0)
                        {
                            dynamic oDataCollectionDef = withBlock.DataPointDetails[0].DataPoint.Parent;
                            DataCollectionName = oDataCollectionDef.Name;
                            DataCollectionRev = oDataCollectionDef.Revision;
                            return (DataPointSummary)oServiceObject_Result.Value.ParametricData;
                        }
                    }
                }
            }
            return null;
        }
        public DataPointSummary SetDataPointSummary(DataPointSummary DataPointSummaryRef, DataPointDetails[] DataPoints)
        {
            DataPointSummary oDataPointSummary = null;
            string sDataName = "";
            foreach (DataPointDetails oDataPointRef in DataPointSummaryRef.DataPointDetails)
            {
                foreach (DataPointDetails oDataPoint in DataPoints)
                {
                    if (!(oDataPoint is null))
                    {
                        sDataName = "";
                        if (oDataPoint.DataName != "")
                        {
                            sDataName = (string)oDataPoint.DataName;
                        }
                        else if (!(oDataPoint.DataPoint is null))
                        {
                            sDataName = oDataPoint.DataPoint.Name;
                        }
                        if ((sDataName != "") && (sDataName == oDataPointRef.DataName))
                        {
                            if (oDataPointSummary is null)
                            {
                                oDataPointSummary = new DataPointSummary() { FieldAction = Camstar.WCF.ObjectStack.Action.Create };
                            }
                            if (oDataPointSummary.DataPointDetails is null)
                            {
                                var objDataPointDetails = oDataPointSummary.DataPointDetails;
                                Array.Resize(ref objDataPointDetails, 1);
                                oDataPointSummary.DataPointDetails = objDataPointDetails;
                            }
                            else
                            {
                                var objDataPointDetails = oDataPointSummary.DataPointDetails;
                                Array.Resize(ref objDataPointDetails, oDataPointSummary.DataPointDetails.Count() + 1);
                                oDataPointSummary.DataPointDetails = objDataPointDetails;
                            }
                            oDataPointSummary.DataPointDetails[oDataPointSummary.DataPointDetails.Count() - 1] = new DataPointDetails()
                            {
                                CDOTypeName = "DataPointDetails",
                                ListItemAction = ListItemAction.Add,
                                DataPoint = new NamedSubentityRef(sDataName) { Parent = oDataPointRef.DataPoint.Parent },
                                DataType = oDataPointRef.DataType,
                                DataValue = oDataPoint.DataValue
                            };
                        }
                    }
                }
            }
            return oDataPointSummary;
        }
        public DataPointSummary SetDataPointSummary(object DataCollectionRef, DataPointDetails[] DataPoints)
        {
            DataPointSummary oDataPointSummary = null;
            string sDataName = "";
            foreach (DataPointDetails oDataPoint in DataPoints)
            {
                if (!(oDataPoint is null))
                {
                    sDataName = "";
                    if (oDataPoint.DataName != "")
                    {
                        sDataName = (string)oDataPoint.DataName;
                    }
                    else if (!(oDataPoint.DataPoint is null))
                    {
                        sDataName = oDataPoint.DataPoint.Name;
                    }
                    if (sDataName != "")
                    {
                        if (oDataPointSummary is null)
                        {
                            oDataPointSummary = new DataPointSummary() { FieldAction = Camstar.WCF.ObjectStack.Action.Create };
                        }
                        if (oDataPointSummary.DataPointDetails is null)
                        {
                            var objDataPointDetails = oDataPointSummary.DataPointDetails;
                            Array.Resize(ref objDataPointDetails, 1);
                            oDataPointSummary.DataPointDetails = objDataPointDetails;
                        }
                        else
                        {
                            var objDataPointDetails = oDataPointSummary.DataPointDetails;
                            Array.Resize(ref objDataPointDetails, oDataPointSummary.DataPointDetails.Count() + 1);
                            oDataPointSummary.DataPointDetails = objDataPointDetails;
                        }
                        oDataPointSummary.DataPointDetails[oDataPointSummary.DataPointDetails.Count() - 1] = new DataPointDetails()
                        {
                            CDOTypeName = "DataPointDetails",
                            ListItemAction = ListItemAction.Add,
                            DataPoint = new NamedSubentityRef(sDataName) { Parent = (BaseObjectRef)DataCollectionRef },
                            DataType = oDataPoint.DataType,
                            DataValue = oDataPoint.DataValue
                        };
                    }
                }
            }
            return oDataPointSummary;
        }
        public bool IsDate(string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                DateTime dt;
                return (DateTime.TryParse(input, out dt));
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region MAINTENANCE FUNCTION
        public bool SaveManageInventory(string NameMaterialQueue, string ManageInventory, string ProductNumber, string BatchNumber = "", double Qty = 0, string UOM = "", bool IgnoreException = true)
        {
            isManageInventoryService oService = null;
            try
            {
                isManageInventory oServiceObject = null;
                ResultStatus oResulstStatus = null;
                oService = new isManageInventoryService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent("Prepare Input Data", System.Diagnostics.EventLogEntryType.Information, 2);
                oServiceObject = new isManageInventory() { sswMaterialQueue = new NamedObjectRef(NameMaterialQueue), isProduct = new RevisionedObjectRef(ProductNumber), isInventoryLocation = new NamedObjectRef(ManageInventory) };
                if (BatchNumber != "") oServiceObject.isLot = BatchNumber;
                if (Qty != 0) oServiceObject.isQty = Qty;
                if (UOM != "") oServiceObject.isUOM = new NamedObjectRef(UOM);

                oResulstStatus = oService.ExecuteTransaction(oServiceObject);
                string sMessage = "";
                bool statusManageInventory = ProcessResult(oResulstStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 2);
                return statusManageInventory;
            }catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            } finally
            {
                if (oService != null) oService.Close();
            }
        }
        public bool SaveMfgOrder(string Name, string Description = "", string Notes = "", string ProductName = "", string ProductRevision = "", double Qty = 0, List<dynamic> MaterialList = null ,string PlannedStartDate = "", string PlannedCompletedDate = "", string ReleaseDate = "", bool AutoCreateQueue = false, bool IgnoreException = true)
        {
            MfgOrderMaintService oService = null;
            try
            {
                MfgOrderMaint oServiceObject = null;

                //check object exists
                oService = new MfgOrderMaintService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent("Checking Mfg Order" + Name, System.Diagnostics.EventLogEntryType.Information, 2);
                bool bObjectExists = ObjectExists(oService, new MfgOrderMaint(), Name);

                // Prepare Object
                EventLogUtil.LogEvent("Preparing Mfg Order" + Name, System.Diagnostics.EventLogEntryType.Information, 2);
                oServiceObject = new MfgOrderMaint();
                if (bObjectExists)
                {
                    oServiceObject.ObjectToChange = new NamedObjectRef(Name);
                    oService.BeginTransaction();
                    oService.Load(oServiceObject);
                }

                //Prepare input data
                oServiceObject = new MfgOrderMaint();
                oServiceObject.ObjectChanges = new MfgOrderChanges();
                oServiceObject.ObjectChanges.Name = new Primitive<string>() { Value = Name };
                if (AutoCreateQueue != false) oServiceObject.ObjectChanges.sswAutoCreateQueue = AutoCreateQueue;
                if (ProductName != "" && ProductRevision != "")
                {
                    oServiceObject.ObjectChanges.Product = new RevisionedObjectRef(ProductName, ProductRevision);
                } else if( ProductName != "")
                {
                    oServiceObject.ObjectChanges.Product = new RevisionedObjectRef(ProductName);
                }
                if (Qty > 0)  oServiceObject.ObjectChanges.Qty = new Primitive<double>() { Value = Qty };
                if (Description != "") oServiceObject.ObjectChanges.Description = new Primitive<string>() { Value = Description };
                if (Notes != "")  oServiceObject.ObjectChanges.Notes = new Primitive<string>() { Value = Notes };
                if (MaterialList != null)
                {
                    if (MaterialList.Count > 0)
                    {
                        oServiceObject.ObjectChanges.ReplaceDetails = true;
                        oServiceObject.ObjectChanges.ReplaceDetailsSubentityListNames = "MaterialList";
                        //MaterialList.ForEach(i => Console.Write("{0}\t", i));
                        var oMaterialList = oServiceObject.ObjectChanges.MaterialList;
                        Array.Resize(ref oMaterialList, MaterialList.Count);
                        Console.WriteLine(oMaterialList.Length);
                        oServiceObject.ObjectChanges.MaterialList = oMaterialList;
                        for (int index = 0; index < MaterialList.Count; index++)
                        {
                            Console.WriteLine(MaterialList[index].Product.Name);
                            oServiceObject.ObjectChanges.MaterialList[index] = MaterialList[index];
                        }
                    }
                }
                if (IsDate(PlannedStartDate)) oServiceObject.ObjectChanges.PlannedStartDate = new Primitive<DateTime>() { Value = Convert.ToDateTime(PlannedStartDate) };
                if (IsDate(PlannedCompletedDate)) oServiceObject.ObjectChanges.PlannedCompletionDate = new Primitive<DateTime>() { Value = Convert.ToDateTime(PlannedCompletedDate) };
                if (IsDate(ReleaseDate)) oServiceObject.ObjectChanges.ReleaseDate = new Primitive<DateTime>() { Value = Convert.ToDateTime(ReleaseDate) };

                // Save the Data
                if (bObjectExists)
                {
                    EventLogUtil.LogEvent("Updating Mfg Order" + Name, System.Diagnostics.EventLogEntryType.Information, 2);
                    oService.ExecuteTransaction(oServiceObject);
                }
                else
                {
                    EventLogUtil.LogEvent("Creating Mfg Order" + Name, System.Diagnostics.EventLogEntryType.Information, 2);
                    oService.BeginTransaction();
                    oService.New(oServiceObject);
                    oService.ExecuteTransaction();
                }
                string sMessage = "";
                bool statusMfgOrder = ProcessResult(oService.CommitTransaction(), ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 2);
                return statusMfgOrder;
            } catch(Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (oService != null) oService.Close();
            }
        }
        #endregion

        #region CONTAINER TXN FUNCTION
        public bool ExecuteStart(string ContainerName, string MfgOrder = "", string ProductName = "", string ProductRevision = "", string WorkflowName = "", string WorkflowRevision = "", string Level = "", string Owner = "", string StartReason = "", string PriorityCode = "", double Qty = 0, string UOM = "", string Comments = "", string EmployeeName = "", string TxnDateStr = "", bool IgnoreException = true)
        {
            StartService oService = null;
            try
            {
                // Setup Service User Profile
                oService = new StartService(AppSettings.ExCoreUserProfile);

                //Setup Input Data
                EventLogUtil.LogEvent("Setting Input Data for Start Container", System.Diagnostics.EventLogEntryType.Information, 3);
                Start oServiceObject = new Start();
                oServiceObject.Details = new StartDetails();
                oServiceObject.CurrentStatusDetails = new CurrentStatusStartDetails();
                oServiceObject.Details.ContainerName = new Primitive<string>() { Value = ContainerName };
                if (MfgOrder != "") oServiceObject.Details.MfgOrder = new NamedObjectRef(MfgOrder);
                if (ProductName != "" && ProductRevision != "")
                {
                    oServiceObject.Details.Product = new RevisionedObjectRef(ProductName, ProductRevision);
                } else if (ProductName != "")
                {
                    oServiceObject.Details.Product = new RevisionedObjectRef(ProductName);
                }
                if (WorkflowName != "" && WorkflowRevision != "")
                {
                    oServiceObject.CurrentStatusDetails.Workflow = new RevisionedObjectRef(WorkflowName, WorkflowRevision);
                } else if (WorkflowName != "")
                {
                    oServiceObject.CurrentStatusDetails.Workflow = new RevisionedObjectRef(WorkflowName);
                }
                if (Level != "") oServiceObject.Details.Level = new NamedObjectRef(Level);
                if (Owner != "") oServiceObject.Details.Owner = new NamedObjectRef(Owner);
                if (StartReason != "") oServiceObject.Details.StartReason = new NamedObjectRef(StartReason);
                if (PriorityCode != "") oServiceObject.Details.PriorityCode = new NamedObjectRef(PriorityCode);
                if (UOM != "") oServiceObject.Details.UOM = new NamedObjectRef(UOM);
                if (Comments != "") oServiceObject.Comments = Comments;
                if (EmployeeName != null) oServiceObject.Employee = new NamedObjectRef(EmployeeName);
                if (IsDate(TxnDateStr)) oServiceObject.TxnDate = Convert.ToDateTime(TxnDateStr);

                //Execute Transaction
                EventLogUtil.LogEvent("Executing Start Container", System.Diagnostics.EventLogEntryType.Information, 3);
                ResultStatus oResultStatus = oService.ExecuteTransaction(oServiceObject);

                //Process Result
                string sMessage = "";
                bool statusExecuteStart = ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 3);
                return statusExecuteStart;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        public bool ExecuteMoveIn(string ContainerName, string ResourceName, string DataCollectionName = "", string DataCollectionRev = "", DataPointDetails[] DataPoints = null, string Comments = "", string EmployeeName = "", string TxnDateStr = "", bool IgnoreException = true)
        {
            MoveInService oService = null;
            try
            {
                string sMessage = "";
                MoveIn oServiceObject = null;
                ResultStatus oResulstStatus = null;
                oService = new MoveInService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent("Setting input data for MoveIn ...", System.Diagnostics.EventLogEntryType.Information, 2);
                oServiceObject = new MoveIn() { Container = new ContainerRef(ContainerName) };
                if (ResourceName != "") oServiceObject.Resource = new NamedObjectRef(ResourceName);
                if (DataPoints != null)
                {
                    if (DataCollectionName != "")
                    {
                        oServiceObject.DataCollectionDef = new RevisionedObjectRef() { Name = DataCollectionName, Revision = DataCollectionRev, RevisionOfRecord = (DataCollectionRev == "") };
                        oServiceObject.ParametricData = SetDataPointSummary(oServiceObject.DataCollectionDef, DataPoints);
                    }
                    else
                    {
                        DataPointSummary oDataPointSummaryRef = GetDataPointSummaryRef(oService, oServiceObject, new MoveIn_Request(), new MoveIn_Info(), ref DataCollectionName, ref DataCollectionRev);
                        oServiceObject.ParametricData = SetDataPointSummary(oDataPointSummaryRef, DataPoints);
                    }
                }
                if (Comments != "") oServiceObject.Comments = Comments;
                if (EmployeeName != null) oServiceObject.Employee = new NamedObjectRef(EmployeeName);
                if (IsDate(TxnDateStr)) oServiceObject.TxnDate = Convert.ToDateTime(TxnDateStr);

                // Execute Transaction
                EventLogUtil.LogEvent("Execution a Move In", System.Diagnostics.EventLogEntryType.Information, 2);
                oResulstStatus = oService.ExecuteTransaction(oServiceObject);

                // Process Result
                bool statusMoveIn = ProcessResult(oResulstStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 2);
                return statusMoveIn;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        public bool ExecuteMoveStd(string ContainerName, string ToResourceName = "", string Resource = "", string DataCollectionName = "", string DataCollectionRev = "", DataPointDetails[] DataPoints = null, string Comments = "", string EmployeeName = "", string TxnDateStr = "", bool IgnoreException = true)
        {
            MoveStdService oService = null;
            try
            {
                string sMessage = "";
                MoveStd oServiceObject = null;
                ResultStatus oResultStatus = null;
                oService = new MoveStdService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent("Setting input data for MoveStd ...", System.Diagnostics.EventLogEntryType.Information, 2);
                oServiceObject = new MoveStd() { Container = new ContainerRef(ContainerName) };
                if (Resource != "")
                {
                    oServiceObject.Resource = new NamedObjectRef() { Name = Resource };
                }
                if (ToResourceName != "")
                {
                    oServiceObject.ToResource = new NamedObjectRef() { Name = ToResourceName };
                }
                if (DataPoints != null)
                {
                    if (DataCollectionName != "")
                    {
                        oServiceObject.DataCollectionDef = new RevisionedObjectRef() { Name = DataCollectionName, Revision = DataCollectionRev, RevisionOfRecord = (DataCollectionRev == "") };
                        oServiceObject.ParametricData = SetDataPointSummary(oServiceObject.DataCollectionDef, DataPoints);
                    }
                    else
                    {
                        DataPointSummary oDataPointSummaryRef = GetDataPointSummaryRef(oService, oServiceObject, new MoveStd_Request(), new MoveStd_Info(), ref DataCollectionName, ref DataCollectionRev);
                        oServiceObject.ParametricData = SetDataPointSummary(oDataPointSummaryRef, DataPoints);
                    }
                }

                if (Comments != "") oServiceObject.Comments = Comments;
                if (EmployeeName != "") oServiceObject.Employee = new NamedObjectRef(EmployeeName);
                if (IsDate(TxnDateStr)) oServiceObject.TxnDate = DateTime.Parse(TxnDateStr);
                EventLogUtil.LogEvent("Execution MoveStd ....", System.Diagnostics.EventLogEntryType.Information, 2);
                oResultStatus = oService.ExecuteTransaction(oServiceObject);
                bool statusMoveStd = ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 2);
                return statusMoveStd;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        public bool ExecuteComponentIssue(string ContainerName, List<dynamic> IssueDetailList = null, bool IgnoreException = true)
        {
            ComponentIssueService oService = null;
            try
            {
                string sMessage = "";
                ComponentIssue oServiceObject = null;
                ResultStatus oResultStatus = null;
                oService = new ComponentIssueService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent("Setting input data for ComponentIssue ...", System.Diagnostics.EventLogEntryType.Information, 2);
                oServiceObject = new ComponentIssue() { Container = new ContainerRef(ContainerName)};
                if (IssueDetailList != null)
                {
                    if (IssueDetailList.Count > 0)
                    {
                        var oIssueDetailList = oServiceObject.IssueActualDetails;
                        Array.Resize(ref oIssueDetailList, IssueDetailList.Count);
                        oServiceObject.IssueActualDetails = oIssueDetailList;
                        for (int index = 0; index < IssueDetailList.Count; index++)
                        {
                            Console.WriteLine(IssueDetailList[index].Product.Name);
                            oServiceObject.IssueActualDetails[index] = IssueDetailList[index];
                        }
                    }
                }
                EventLogUtil.LogEvent("Execution ComponentIssue ....", System.Diagnostics.EventLogEntryType.Information, 2);
                oResultStatus = oService.ExecuteTransaction(oServiceObject);
                bool statusComponentIssue= ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 3);
                return statusComponentIssue;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        public bool ContainerExists(string ContainerName, bool IgnoreException = true)
        {
            ViewContainerStatusService oService = null;
            try
            {
                oService = new ViewContainerStatusService(AppSettings.ExCoreUserProfile);

                //Set input Data
                ViewContainerStatus oServiceObject = new ViewContainerStatus();
                oServiceObject.Container = new ContainerRef(ContainerName);
                ViewContainerStatus_Request oServiceRequest = new ViewContainerStatus_Request();
                oServiceRequest.Info = new ViewContainerStatus_Info();
                oServiceRequest.Info.ContainerName = new Info(true);

                //Request the Data
                ViewContainerStatus_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.ExecuteTransaction(oServiceObject, oServiceRequest, out oServiceResult);

                //Return the Result
                string sMessage = "";
                return ProcessResult(oResultStatus, ref sMessage, true);
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        public CurrentContainerStatus GetContainerStatusDetails(string ContainerName, string DataCollectionName = "", string DataCollectionRev = "", bool IgnoreException = true)
        {
            ContainerTxnService oService = null;
            try
            {
                oService = new ContainerTxnService(AppSettings.ExCoreUserProfile);

                //Set input Data
                ContainerTxn oServiceObject = new ContainerTxn();
                oServiceObject.Container = new ContainerRef(ContainerName);
                if (DataCollectionName != "")
                {
                    oServiceObject.DataCollectionDef = new RevisionedObjectRef() { Name = DataCollectionName, Revision = DataCollectionRev, RevisionOfRecord = (DataCollectionRev == "") };
                }
                ContainerTxn_Request oServiceRequest = new ContainerTxn_Request();
                oServiceRequest.Info = new ContainerTxn_Info();
                oServiceRequest.Info.CurrentContainerStatus= new CurrentContainerStatus_Info();
                oServiceRequest.Info.CurrentContainerStatus.RequestValue = true;

                //Requets the Data
                ContainerTxn_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.GetEnvironment(oServiceObject, oServiceRequest, out oServiceResult);
                Console.WriteLine(oResultStatus.IsSuccess);

                //Return Result
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, false))
                {
                    return oServiceResult.Value.CurrentContainerStatus;
                } else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return null;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        #endregion

        #region RESOURCE TXN FUNCTION
        public ResourceStatusDetails GetResourceStatusDetails(string ResourceName, bool IgnoreException = true)
        {
            ResourceTxnService oService = null;
            try
            {
                oService = new ResourceTxnService(AppSettings.ExCoreUserProfile);

                //Set Input Data
                ResourceTxn oServiceObject = new ResourceTxn();
                oServiceObject.Resource = new NamedObjectRef(ResourceName);
                ResourceTxn_Request oServiceRequest = new ResourceTxn_Request();
                oServiceRequest.Info = new ResourceTxn_Info();
                oServiceRequest.Info.ResourceStatusDetails = new ResourceStatusDetails_Info();
                oServiceRequest.Info.ResourceStatusDetails.RequestValue = true;

                //Request the Data
                ResourceTxn_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.GetEnvironment(oServiceObject, oServiceRequest, out oServiceResult);

                //Return the Result
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.ResourceStatusDetails;
                } else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return null;
            }
            finally
            {
                if (!(oService is null)) oService.Close();
            }
        }
        public bool ExecuteResourceSetup(string ResourceName, string Status = "", string Reason = "", string Comments = "", string EmployeeName = "", string TxnDate = "", bool IgnoreException = true)
        {
            ResourceSetupService oService = null;
            try
            {
                string sMessage = "";
                ResourceSetup oServiceObject = null;
                ResultStatus oResultStatus = null;
                oService = new ResourceSetupService(AppSettings.ExCoreUserProfile);

                //Set Input Data
                EventLogUtil.LogEvent("Setting input data for ResourceSetup", System.Diagnostics.EventLogEntryType.Information , 3);
                oServiceObject = new ResourceSetup() { Resource = new NamedObjectRef(ResourceName), ResourceStatusCode = new NamedObjectRef(Status), ResourceStatusReason = new NamedObjectRef(Reason) };
                if (Comments != "") oServiceObject.Comments = Comments;
                if (EmployeeName != "") oServiceObject.Employee = new NamedObjectRef(EmployeeName);
                if (IsDate(TxnDate)) oServiceObject.TxnDate = Convert.ToDateTime(TxnDate);

                //Execute Transaction
                EventLogUtil.LogEvent("Executing ResourceSetup", System.Diagnostics.EventLogEntryType.Information, 3);
                oResultStatus = oService.ExecuteTransaction(oServiceObject);

                //Process Result
                bool statusResourceSetup = ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 3);
                return statusResourceSetup;
            } catch  (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (oService != null) oService.Close();
            }
        }
        public bool ExecuteResourceThruput(string ResourceName, double Qty = 0, string UOM = "", string ProductName = "", string ProductRevision = "", string Comments = "", string EmployeeName = "", string TxnDate = "", bool IgnoreException = true)
        {
            ResourceThruputService oService = null;
            try
            {
                string sMessage = "";
                ResourceThruput oServiceObject = null;
                ResultStatus oResultStatus = null;
                //Set Profile
                oService = new ResourceThruputService(AppSettings.ExCoreUserProfile);

                //Set Input Data into Resource Thruput
                EventLogUtil.LogEvent("Setting input data for Resource Thruput", System.Diagnostics.EventLogEntryType.Information, 3);
                oServiceObject = new ResourceThruput() { Resource = new NamedObjectRef(ResourceName) };
                if (UOM != "") oServiceObject.UOM = new NamedObjectRef(UOM);
                if (ProductName != "" && ProductRevision != "")
                {
                    oServiceObject.Product = new RevisionedObjectRef(ProductName, ProductRevision);
                }
                else if (ProductName != "")
                {
                    oServiceObject.Product = new RevisionedObjectRef(ProductName);
                }
                if (Qty > 0) oServiceObject.Qty = new Primitive<double>() { Value = Qty };
                if (Comments != "") oServiceObject.Comments = Comments;
                if (EmployeeName != "") oServiceObject.Employee = new NamedObjectRef(EmployeeName);
                if (IsDate(TxnDate)) oServiceObject.TxnDate = Convert.ToDateTime(TxnDate);

                //Execute Transaction
                EventLogUtil.LogEvent("Execution ResourceThruput", System.Diagnostics.EventLogEntryType.Information, 3);
                oResultStatus = oService.ExecuteTransaction(oServiceObject);

                //Process Result
                bool statusResourceThruput = ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 3);
                return statusResourceThruput;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (oService != null) oService.Close();
            }
        }
        #endregion

        #region TEAM TRACKING TXN FUNCTION
        public bool AddTeamMembers(string TeamName, List<dynamic> EmployeeList, bool IgnoreException = true)
        {
            mdAddTeamMembersService oService = null;
            try
            {
                string sMessage = "";
                mdAddTeamMembers oServiceObject = null;
                ResultStatus oResultStatus = null;

                //set Profile to Service
                oService = new mdAddTeamMembersService(AppSettings.ExCoreUserProfile);

                //Set Input Data into Resource Thruput
                EventLogUtil.LogEvent("Setting input data for Add Team Members", System.Diagnostics.EventLogEntryType.Information, 3);
                oServiceObject = new mdAddTeamMembers() { Team = new NamedObjectRef(TeamName) };
                if (EmployeeList != null)
                {
                    if (EmployeeList.Count > 0)
                    {
                        var oEmployeeList = oServiceObject.Employees;
                        Array.Resize(ref oEmployeeList, EmployeeList.Count);
                        oServiceObject.Employees = oEmployeeList;
                        for(int index = 0; index < EmployeeList.Count; index++)
                        {
                            Console.WriteLine(EmployeeList[index].Name);
                            oServiceObject.Employees[index] = EmployeeList[index];
                        }
                    }
                }

                //Execute Transaction
                EventLogUtil.LogEvent("Execution Add Team Members to the Team", System.Diagnostics.EventLogEntryType.Information, 3);
                oResultStatus = oService.ExecuteTransaction(oServiceObject);

                //Process Result
                bool statusAddTeamMembers = ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 3);
                return statusAddTeamMembers;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (oService != null) oService.Close();
            }
        }
        public bool RemoveTeamMembers(string TeamName, List<dynamic> EmployeeList, bool IgnoreException = true)
        {
            mdRemoveTeamMembersService oService = null;
            try
            {
                string sMessage = "";
                mdRemoveTeamMembers oServiceObject = null;
                ResultStatus oResultStatus = null;

                //Set Profile to the Service
                oService = new mdRemoveTeamMembersService(AppSettings.ExCoreUserProfile);
                //Set Input Data into Resource Thruput
                EventLogUtil.LogEvent("Setting input data for Remove Team Members", System.Diagnostics.EventLogEntryType.Information, 3);
                oServiceObject = new mdRemoveTeamMembers() { Team = new NamedObjectRef(TeamName) };
                if (EmployeeList != null)
                {
                    if (EmployeeList.Count > 0)
                    {
                        var oEmployeeList = oServiceObject.Employees;
                        Array.Resize(ref oEmployeeList, EmployeeList.Count);
                        oServiceObject.Employees = oEmployeeList;
                        for (int index = 0; index < EmployeeList.Count; index++)
                        {
                            Console.WriteLine(EmployeeList[index].Name);
                            oServiceObject.Employees[index] = EmployeeList[index];
                        }
                    }
                }

                //Execute Transaction
                EventLogUtil.LogEvent("Execution Remove Team Members to the Team", System.Diagnostics.EventLogEntryType.Information, 3);
                oResultStatus = oService.ExecuteTransaction(oServiceObject);

                //Process Result
                bool statusAddTeamMembers = ProcessResult(oResultStatus, ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 3);
                return statusAddTeamMembers;
            }
            catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return false;
            }
            finally
            {
                if (oService != null) oService.Close();
            }
        }
        #endregion
    }
}