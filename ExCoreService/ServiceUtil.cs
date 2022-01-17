using System;
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
        public ProductChanges GetProduct(string ProductName, string ProductRevision = "", bool IgnoreException = true)
        {
            ProductMaintService oService = null;
            try
            {
                oService = new ProductMaintService(AppSettings.ExCoreUserProfile);
                ProductMaint oServiceObject = new ProductMaint();
                if (ProductName != "" && ProductRevision != "")
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(ProductName, ProductRevision);
                }
                else if (ProductName != "")
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(ProductName);
                }
                ProductMaint_Request oServiceRequest = new ProductMaint_Request();
                oServiceRequest.Info = new ProductMaint_Info();
                oServiceRequest.Info.ObjectChanges = new ProductChanges_Info();
                oServiceRequest.Info.ObjectChanges.ProductType = new Info(true);
                oServiceRequest.Info.ObjectChanges.Workflow = new Info(true);
                oServiceRequest.Info.ObjectChanges.Name = new Info(true);
                oServiceRequest.Info.ObjectChanges.Description = new Info(true);
                oServiceRequest.Info.ObjectChanges.ERPBOM = new Info(true);
                oServiceRequest.Info.ObjectChanges.BOM = new Info(true);
                oServiceRequest.Info.ObjectChanges.ERPBOM = new Info(true);
                oServiceRequest.Info.ObjectChanges.StdStartQty = new Info(true);
                oServiceRequest.Info.ObjectChanges.UOM = new Info(true);
                oServiceRequest.Info.ObjectChanges.sswStartOwner = new Info(true);
                oServiceRequest.Info.ObjectChanges.sswStartReason = new Info(true);
                oServiceRequest.Info.ObjectChanges.sswStartLevel = new Info(true);

                ProductMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.Load(oServiceObject, oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                if (oServiceResult.Value.ObjectChanges != null)
                {
                    return oServiceResult.Value.ObjectChanges;
                }
                else
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
        public WorkflowChanges GetWorkflow(string WorkflowName, string WorkflowRevision = "", bool IgnoreException = true)
        {
            WorkflowMaintService oService = null;
            try
            {
                oService = new WorkflowMaintService(AppSettings.ExCoreUserProfile);
                WorkflowMaint oServiceObject = new WorkflowMaint();
                if (WorkflowName != "" && WorkflowRevision != "")
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(WorkflowName, WorkflowRevision);
                }
                else if (WorkflowName != "")
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(WorkflowName);
                }
                WorkflowMaint_Request oServiceRequest = new WorkflowMaint_Request();
                oServiceRequest.Info = new WorkflowMaint_Info();
                oServiceRequest.Info.ObjectChanges = new WorkflowChanges_Info();
                oServiceRequest.Info.ObjectChanges.Name = new Info(true);
                oServiceRequest.Info.ObjectChanges.Description = new Info(true);
                oServiceRequest.Info.ObjectChanges.ERPRoute = new Info(true);
                StepChanges_Info stepChanges_Info = new StepChanges_Info();
                stepChanges_Info.Name = new Info(true);
                stepChanges_Info.RouteStep = new Info(true);
                oServiceRequest.Info.ObjectChanges.Steps = stepChanges_Info;
                oServiceRequest.Info.ObjectChanges.Status = new Info(true);

                WorkflowMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.Load(oServiceObject, oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                if (oServiceResult.Value.ObjectChanges != null)
                {
                    return oServiceResult.Value.ObjectChanges;
                }
                else
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
        public ERPRouteChanges GetERPRouteFromMfgOrder(MfgOrderChanges oMfgOrder, bool IgnoreException = true)
        {
            try
            {
                if (oMfgOrder != null)
                {
                    if (oMfgOrder.Product != null)
                    {
                        ProductChanges oProduct = GetProduct(oMfgOrder.Product.Name);
                        if (oProduct.Workflow != null)
                        {
                            WorkflowChanges oWorkflow = GetWorkflow(oProduct.Workflow.Name);
                            if (oWorkflow.ERPRoute != null)
                            {
                                ERPRouteChanges oERPRoute = GetERPRoute(oWorkflow.ERPRoute.Name);
                                return oERPRoute;
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return null;
                    }
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
        }
        public ERPRouteChanges GetERPRoute(string ERPRouteName, string ERPRouteRevision = "", bool IgnoreException = true)
        {
            ERPRouteMaintService oService = null;
            try
            {
                oService = new ERPRouteMaintService(AppSettings.ExCoreUserProfile);
                ERPRouteMaint oServiceObject = new ERPRouteMaint();
                if (ERPRouteName != "" && ERPRouteRevision != "")
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(ERPRouteName, ERPRouteRevision);
                }
                else if (ERPRouteName != "")
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(ERPRouteName);
                }

                ERPRouteMaint_Request oServiceRequest = new ERPRouteMaint_Request();
                oServiceRequest.Info = new ERPRouteMaint_Info();
                oServiceRequest.Info.ObjectChanges = new ERPRouteChanges_Info();
                oServiceRequest.Info.ObjectChanges.Name = new Info(true);
                oServiceRequest.Info.ObjectChanges.Description = new Info(true);
                oServiceRequest.Info.ObjectChanges.RouteStepItem = new Info(true);
                RouteStepChanges_Info routeStepChanges_Info = new RouteStepChanges_Info();
                routeStepChanges_Info.Name = new Info(true);
                routeStepChanges_Info.ERPOperation = new Info(true);
                routeStepChanges_Info.Sequence = new Info(true);
                oServiceRequest.Info.ObjectChanges.RouteSteps = routeStepChanges_Info;
                oServiceRequest.Info.ObjectChanges.Status = new Info(true);

                ERPRouteMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.Load(oServiceObject ,oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                if (oServiceResult.Value.ObjectChanges != null)
                {
                    return oServiceResult.Value.ObjectChanges;
                }
                else
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
        public MfgOrderChanges GetMfgOrder(string MfgOrderName, bool IgnoreException = true)
        {
            MfgOrderMaintService oService = null;
            try
            {
                oService = new MfgOrderMaintService(AppSettings.ExCoreUserProfile);
                MfgOrderMaint oServiceObject = new MfgOrderMaint();
                oServiceObject.ObjectToChange = new NamedObjectRef(MfgOrderName);

                MfgOrderMaint_Request oServiceRequest = new MfgOrderMaint_Request();
                oServiceRequest.Info = new MfgOrderMaint_Info();
                oServiceRequest.Info.ObjectChanges = new MfgOrderChanges_Info();
                oServiceRequest.Info.ObjectChanges.RequestValue = true;

                MfgOrderMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.Load(oServiceObject, oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.ObjectChanges;
                }
                else
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
        public MfgOrderChanges GetMfgOrderDispatch(bool IgnoreException = true)
        {
            try
            {
                ServiceUtil oServiceUtil = new ServiceUtil();
                NamedObjectRef[] oNameMfgList = oServiceUtil.GetListMfgOrder();
                List<MfgOrderChanges> oMfgList = new List<MfgOrderChanges>();
                List<MfgOrderChanges> oNewMfgList = new List<MfgOrderChanges>(oNameMfgList.Count());
                foreach (var MfgItem in oNameMfgList)
                {
                    oMfgList.Add(oServiceUtil.GetMfgOrder(MfgItem.Name));
                }
                oMfgList.ForEach((value) =>
                {
                    if (value.PlannedStartDate != null && value.isWorkflow != null && value.isWorkflow.Name == AppSettings.Workflow && value.OrderStatus.Name == "Released")
                    {
                        oNewMfgList.Add(value);
                    }
                });
                oNewMfgList = oNewMfgList.OrderBy(x => Convert.ToDateTime(x.PlannedStartDate.ToString())).ToList();
                if (oNewMfgList.Count > 0)
                {
                    return oNewMfgList[0];
                }
                else
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
        }
        public NamedObjectRef[] GetListMfgOrder(bool IgnoreException = true)
        {
            MfgOrderMaintService oService = null;
            try
            {
                oService = new MfgOrderMaintService(AppSettings.ExCoreUserProfile);
                MfgOrderMaint oServiceObject = new MfgOrderMaint();

                MfgOrderMaint_Request oServiceRequest = new MfgOrderMaint_Request();
                oServiceRequest.Info = new MfgOrderMaint_Info();
                oServiceRequest.Info.ObjectListInquiry = new Info(true);
                oServiceRequest.Info.ObjectListInquiry.RequestValue = true;

                MfgOrderMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.GetEnvironment(oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.ObjectListInquiry;
                }
                else
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
        public bool SaveMfgOrder(string Name, string Description = "", string Notes = "", string ProductName = "", string ProductRevision = "", string WorkflowName = "", string WorkflowRevision = "", double Qty = 0, List<dynamic> MaterialList = null, string ERPRoute = "", string PlannedStartDate = "", string PlannedCompletedDate = "", string ReleaseDate = "", string OrderStatus = "", bool AutoCreateQueue = false, bool IgnoreException = true)
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
                if (ERPRoute != "") oServiceObject.ObjectChanges.ERPRoute = new RevisionedObjectRef(ERPRoute);
                if (OrderStatus != "") oServiceObject.ObjectChanges.OrderStatus = new NamedObjectRef(OrderStatus);
                if (AutoCreateQueue != false) oServiceObject.ObjectChanges.sswAutoCreateQueue = AutoCreateQueue;
                if (ProductName != "" && ProductRevision != "" && ObjectExists(new ProductMaintService(AppSettings.ExCoreUserProfile), new ProductMaint(), ProductName, ProductRevision))
                {
                    oServiceObject.ObjectChanges.Product = new RevisionedObjectRef(ProductName, ProductRevision);
                }
                else if (ProductName != "" && ObjectExists(new ProductMaintService(AppSettings.ExCoreUserProfile), new ProductMaint(), ProductName, ""))
                {
                    oServiceObject.ObjectChanges.Product = new RevisionedObjectRef(ProductName);
                }

                if (WorkflowName != "" && WorkflowRevision != "" && ObjectExists(new WorkflowMaintService(AppSettings.ExCoreUserProfile), new WorkflowMaint(), WorkflowName, WorkflowRevision))
                {
                    oServiceObject.ObjectChanges.isWorkflow = new RevisionedObjectRef(WorkflowName, WorkflowRevision);
                }
                else if (WorkflowName != "" && ObjectExists(new WorkflowMaintService(AppSettings.ExCoreUserProfile), new WorkflowMaint(), WorkflowName, ""))
                {
                    oServiceObject.ObjectChanges.isWorkflow = new RevisionedObjectRef(WorkflowName);
                }
                if (Qty > 0) oServiceObject.ObjectChanges.Qty = new Primitive<double>() { Value = Qty };
                if (Description != "") oServiceObject.ObjectChanges.Description = new Primitive<string>() { Value = Description };
                if (Notes != "") oServiceObject.ObjectChanges.Notes = new Primitive<string>() { Value = Notes };
                if (MaterialList != null)
                {
                    if (MaterialList.Count > 0)
                    {
                        oServiceObject.ObjectChanges.ReplaceDetails = true;
                        oServiceObject.ObjectChanges.ReplaceDetailsSubentityListNames = "MaterialList";
                        var oMaterialList = oServiceObject.ObjectChanges.MaterialList;
                        Array.Resize(ref oMaterialList, MaterialList.Count);
                        oServiceObject.ObjectChanges.MaterialList = oMaterialList;
                        for (int index = 0; index < MaterialList.Count; index++)
                        {
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
        public bool SaveProduct(string ProductName, string Revision, string IsRevOfRcd = "", string Description = "", string Notes ="", string ProductType = "", string DocumentSet = "", string WorkflowName = "", string WorkflowRevision = "", string BOMName = "", string BOMRevision = "", bool IgnoreException = true)
        {
            ProductMaintService oService = null;
            try
            {
                ProductMaint oServiceObject = null;
                // CheckObject Exists
                oService = new ProductMaintService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent($"Checking Product {ProductName} : {Revision}", System.Diagnostics.EventLogEntryType.Information, 3);
                bool bBaseExists = ObjectExists(oService, new ProductMaint(), ProductName, "");
                bool bObjectExists = ObjectExists(oService, new ProductMaint(), ProductName, Revision);
                // Prepare Object
                EventLogUtil.LogEvent($"Preparing Product {ProductName} : {Revision}", System.Diagnostics.EventLogEntryType.Information, 3);
                oServiceObject = new ProductMaint();
                if (bObjectExists)
                {
                    oServiceObject.ObjectToChange = new RevisionedObjectRef(ProductName, Revision);
                    oService.BeginTransaction();
                    oService.Load(oServiceObject);
                } else if (bBaseExists)
                {
                    oService.BeginTransaction();
                    oServiceObject.BaseToChange = new RevisionedObjectRef();
                    oServiceObject.BaseToChange.Name = ProductName;
                    oService.NewRev(oServiceObject);
                }
                // PrepareInput Data
                oServiceObject = new ProductMaint();
                oServiceObject.ObjectChanges = new ProductChanges();
                oServiceObject.ObjectChanges.Name = new Primitive<string>() { Value = ProductName};
                oServiceObject.ObjectChanges.Revision = new Primitive<string>() { Value = Revision };
                if (IsRevOfRcd != "") oServiceObject.ObjectChanges.IsRevOfRcd = new Primitive<bool>() { Value = Convert.ToBoolean(IsRevOfRcd) };
                if (Description != "") oServiceObject.ObjectChanges.Description = new Primitive<string>() { Value = Description };
                if (Notes != "") oServiceObject.ObjectChanges.Notes = new Primitive<string>() { Value = Notes };
                if (ProductType != "") oServiceObject.ObjectChanges.ProductType = new NamedObjectRef(ProductType);
                if (DocumentSet != "") oServiceObject.ObjectChanges.DocumentSet = new NamedObjectRef(DocumentSet);
                if (WorkflowName != "" && WorkflowRevision == "")
                {
                    oServiceObject.ObjectChanges.Workflow = new RevisionedObjectRef(WorkflowName);
                } else if (WorkflowName != "" && WorkflowRevision != "")
                {
                    oServiceObject.ObjectChanges.Workflow = new RevisionedObjectRef(WorkflowName, WorkflowRevision);
                }
                if (BOMName != "" && BOMRevision == "")
                {
                    oServiceObject.ObjectChanges.BOM = new RevisionedObjectRef(BOMName);
                } else if (BOMName != "" && BOMRevision != "")
                {
                    oServiceObject.ObjectChanges.BOM = new RevisionedObjectRef(BOMName, BOMRevision);
                }
                // Save the Data
                if (bObjectExists)
                {
                    EventLogUtil.LogEvent($"Updating Product {ProductName} : {Revision}", System.Diagnostics.EventLogEntryType.Information, 3);
                    oService.ExecuteTransaction(oServiceObject);
                } else if (bBaseExists)
                {
                    EventLogUtil.LogEvent($"Creating Product {ProductName} : {Revision}", System.Diagnostics.EventLogEntryType.Information, 3);
                    oService.ExecuteTransaction(oServiceObject);
                } else
                {
                    EventLogUtil.LogEvent($"Creating Product {ProductName} : {Revision}", System.Diagnostics.EventLogEntryType.Information, 3);
                    oService.BeginTransaction();
                    oService.New(oServiceObject);
                    oService.ExecuteTransaction();
                }
                string sMessage = "";
                bool statusProduct = ProcessResult(oService.CommitTransaction(), ref sMessage, false);
                EventLogUtil.LogEvent(sMessage, System.Diagnostics.EventLogEntryType.Information, 2);
                return statusProduct;
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

        #region CONTAINER TXN FUNCTION
        public bool ExecuteRework(string ContainerName, string ReworkReason, string Path = "", string Resource = "", bool IgnoreException = true)
        {
            ReworkService oService = null;
            try
            {
                oService = new ReworkService(AppSettings.ExCoreUserProfile);
                EventLogUtil.LogEvent("Setting input data for Rework Txn", System.Diagnostics.EventLogEntryType.Information, 3);
                Rework oServiceObject = new Rework();
                oServiceObject.Container = new ContainerRef(ContainerName);
                oServiceObject.ReworkReason = new NamedObjectRef(ReworkReason);
                if (Path != "") oServiceObject.Path = new NamedSubentityRef(Path);
                if (Resource != "") oServiceObject.Resource = new NamedObjectRef(Resource);
                EventLogUtil.LogEvent("Execute Rework Txn", System.Diagnostics.EventLogEntryType.Information, 3);
                ResultStatus oResultStatus = null;
                string sMessage = "";
                oResultStatus = oService.ExecuteTransaction(oServiceObject);
                bool statusMoveIn = ProcessResult(oResultStatus, ref sMessage, false);
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
                }
                else if (ProductName != "")
                {
                    oServiceObject.Details.Product = new RevisionedObjectRef(ProductName);
                }
                if (WorkflowName != "" && WorkflowRevision != "")
                {
                    oServiceObject.CurrentStatusDetails.Workflow = new RevisionedObjectRef(WorkflowName, WorkflowRevision);
                }
                else if (WorkflowName != "")
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
                oServiceObject = new ComponentIssue() { Container = new ContainerRef(ContainerName) };
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
                bool statusComponentIssue = ProcessResult(oResultStatus, ref sMessage, false);
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
                oServiceRequest.Info.CurrentContainerStatus = new CurrentContainerStatus_Info();
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
                }
                else
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
        public ResStatusReasonGroupChanges GetResourceStatusReasonGroup(string StatusCodeName, bool IgnoreException = true)
        {
            ResStatusReasonGroupMaintService oService = null;
            try
            {
                oService = new ResStatusReasonGroupMaintService(AppSettings.ExCoreUserProfile);
                ResStatusReasonGroupMaint oServiceObject = new ResStatusReasonGroupMaint();
                oServiceObject.ObjectToChange = new NamedObjectRef(StatusCodeName);

                ResStatusReasonGroupMaint_Request oServiceRequest = new ResStatusReasonGroupMaint_Request();
                oServiceRequest.Info = new ResStatusReasonGroupMaint_Info();
                oServiceRequest.Info.ObjectChanges = new ResStatusReasonGroupChanges_Info();
                oServiceRequest.Info.ObjectChanges.RequestValue = true;

                ResStatusReasonGroupMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.Load(oServiceObject, oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.ObjectChanges;
                }
                else
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
        public ResourceStatusCodeChanges GetResourceStatusCode(string StatusCodeName, bool IgnoreException = true)
        {
            ResourceStatusCodeMaintService oService = null;
            try
            {
                oService = new ResourceStatusCodeMaintService(AppSettings.ExCoreUserProfile);
                ResourceStatusCodeMaint oServiceObject = new ResourceStatusCodeMaint();
                oServiceObject.ObjectToChange = new NamedObjectRef(StatusCodeName);

                ResourceStatusCodeMaint_Request oServiceRequest = new ResourceStatusCodeMaint_Request();
                oServiceRequest.Info = new ResourceStatusCodeMaint_Info();
                oServiceRequest.Info.ObjectChanges = new ResourceStatusCodeChanges_Info();
                oServiceRequest.Info.ObjectChanges.RequestValue = true;

                ResourceStatusCodeMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.Load(oServiceObject, oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.ObjectChanges;
                }
                else
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
        public NamedObjectRef[] GetListResourceStatusCode(bool IgnoreException = true)
        {
            ResourceStatusCodeMaintService oService = null;
            try
            {
                oService = new ResourceStatusCodeMaintService(AppSettings.ExCoreUserProfile);
                ResourceStatusCodeMaint oServiceObject = new ResourceStatusCodeMaint();

                ResourceStatusCodeMaint_Request oServiceRequest = new ResourceStatusCodeMaint_Request();
                oServiceRequest.Info = new ResourceStatusCodeMaint_Info();
                oServiceRequest.Info.ObjectListInquiry = new Info(true);
                oServiceRequest.Info.ObjectListInquiry.RequestValue = true;

                ResourceStatusCodeMaint_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.GetEnvironment(oServiceRequest, out oServiceResult);

                EventLogUtil.LogEvent(oResultStatus.Message, System.Diagnostics.EventLogEntryType.Information, 3);
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.ObjectListInquiry;
                }
                else
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
                }
                else
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
        public GetMaintenanceStatusDetails[] GetGetMaintenanceStatus(string ResourceName, bool IgnoreException = true)
        {
            GetMaintenanceStatusesService oService = null;
            try
            {
                oService = new GetMaintenanceStatusesService(AppSettings.ExCoreUserProfile);

                //Set Input Data
                GetMaintenanceStatuses oServiceObject = new GetMaintenanceStatuses();
                oServiceObject.Resource = new NamedObjectRef(ResourceName);
                GetMaintenanceStatuses_Request oServiceRequest = new GetMaintenanceStatuses_Request();
                oServiceRequest.Info = new GetMaintenanceStatuses_Info();
                oServiceRequest.Info.StatusDetails = new GetMaintenanceStatusDetails_Info();
                oServiceRequest.Info.StatusDetails.RequestValue = true;

                //Request the Data
                GetMaintenanceStatuses_Result oServiceResult = null;
                ResultStatus oResultStatus = oService.GetEnvironment(oServiceObject, oServiceRequest, out oServiceResult);

                //Return the Result
                string sMessage = "";
                if (ProcessResult(oResultStatus, ref sMessage, true))
                {
                    return oServiceResult.Value.StatusDetails;
                }
                else
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
                EventLogUtil.LogEvent("Setting input data for ResourceSetup", System.Diagnostics.EventLogEntryType.Information, 3);
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
                        for (int index = 0; index < EmployeeList.Count; index++)
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
