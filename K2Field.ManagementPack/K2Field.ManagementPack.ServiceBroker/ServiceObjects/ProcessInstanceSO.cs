﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using K2Field.ManagementPack.ServiceBroker.Helpers;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using SourceCode.Workflow.Client;

namespace K2Field.ManagementPack.ServiceBroker.ServiceObjects
{
    public class ProcessInstanceSO : ServiceObjectBase
    {
        private Connection _wfClient = new Connection();
        public ProcessInstanceSO(ServiceBroker api) : base(api) { }
        public override string ServiceFolder => "ProcessInstance";
        
        public override List<ServiceObject> DescribeServiceObjects()
        {
            var so = new ServiceObjectBuilder("ProcessInstanceClient", "Exposes functionality to start the workflow.", true);
            so.CreateProperty(Constants.SoProperties.ProcessInstance.ProcessFolio,
                    "The folio to use for the process.", SoType.Text)
                .CreateProperty(Constants.SoProperties.ProcessInstance.ProcessName,
                    "The full name of the process.", SoType.Text)
                .CreateProperty(Constants.SoProperties.ProcessInstance.StartSync,
                    "Start the process synchronously or not.", SoType.YesNo)
                .CreateProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId,
                    "The process instance ID.", SoType.Number)
                .CreateProperty(Constants.SoProperties.ProcessInstance.ProcessVersion,
                    "The version number to start. Leave empty for default.", SoType.Number)
                .CreateProperty(Constants.SoProperties.ProcessInstance.DataFieldName,
                    "The name of the DataField.", SoType.Text)
                .CreateProperty(Constants.SoProperties.ProcessInstance.DataFieldValue,
                    "The value of the DataField.", SoType.Memo)
                .CreateProperty(Constants.SoProperties.ProcessInstance.XmlFieldName, "The name of the XML field.",
                    SoType.Text)
                .CreateProperty(Constants.SoProperties.ProcessInstance.XmlFieldValue,
                    "The value of the XML field.", SoType.Text);

            var updateFolio = new ServiceObjectMethodBuilder(Constants.Methods.ProcessInstance.UpdateFolio, "Updates the folio of a running process instance", MethodType.Update)
                .AddProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true, true, false)
                .AddProperty(Constants.SoProperties.ProcessInstance.ProcessFolio, true, true, false);
            so.AddMethod(updateFolio);

            var updateDataField  = new ServiceObjectMethodBuilder(Constants.Methods.ProcessInstance.UpdateDataField,
                    "Updates the DataField of a running process instance", MethodType.Update)
                .AddProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true, true, false)
                .AddProperty(Constants.SoProperties.ProcessInstance.DataFieldName, true, true, false)
                .AddProperty(Constants.SoProperties.ProcessInstance.DataFieldValue, true, true, false);
            so.AddMethod(updateDataField);

            var listDataFields = new ServiceObjectMethodBuilder(Constants.Methods.ProcessInstance.ListDataFields,
                    "Lists the data fields with values from the Process Instance", MethodType.List)
                .AddProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true, true, false)
                .AddProperty(Constants.SoProperties.ProcessInstance.DataFieldName, false, false, true)
                .AddProperty(Constants.SoProperties.ProcessInstance.DataFieldValue, false, false, true);
            so.AddMethod(listDataFields);

            var updateXmlField  = new ServiceObjectMethodBuilder(Constants.Methods.ProcessInstance.SetXmlField,
                    "Updates the XmlField of a running process instance", MethodType.Update)
                .AddProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true, true, false)
                .AddProperty(Constants.SoProperties.ProcessInstance.XmlFieldName, true, true, false)
                .AddProperty(Constants.SoProperties.ProcessInstance.XmlFieldValue, true, true, false);
            so.AddMethod(updateXmlField);

            

            return new List<ServiceObject>() { so };
        }
        public override void Execute()
        {
            switch (ServiceBroker.Service.ServiceObjects[0].Methods[0].Name)
            {
                case Constants.Methods.ProcessInstance.UpdateFolio:
                    UpdateFolio();
                    break;
                case Constants.Methods.ProcessInstance.UpdateDataField:
                    UpdateDataField();
                    break;
                case Constants.Methods.ProcessInstance.ListDataFields:
                    ListDataFields();
                    break;
                    //default:
                    //    StartProcessInstance(true);
                    //    break;
            }
        }

        private void UpdateFolio()
        {
            var folio = GetStringProperty(Constants.SoProperties.ProcessInstance.ProcessFolio, true);
            var procId = GetIntProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true);

            using (_wfClient = ServiceBroker.K2Connection.GetWorkflowClientConnection())
            {
                var pi = _wfClient.OpenProcessInstance(procId);
                pi.Folio = folio;
                pi.Update();
            }
        }
        private void UpdateDataField()
        {
            var dataFieldName = GetStringProperty(Constants.SoProperties.ProcessInstance.DataFieldName, true);
            var dataFieldValue = GetStringProperty(Constants.SoProperties.ProcessInstance.DataFieldValue);
            var procId = GetIntProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true);

            using (_wfClient = ServiceBroker.K2Connection.GetWorkflowClientConnection())
            {
                var pi = _wfClient.OpenProcessInstance(procId);
                var dataField = pi.DataFields[dataFieldName];
                switch(dataField.ValueType)
                {
                    case DataType.TypeBinary:
                        dataField.Value = Convert.FromBase64String(dataFieldValue);
                        break;
                    case DataType.TypeBoolean:
                        dataField.Value = Convert.ToBoolean(dataFieldValue);
                        break;
                    case DataType.TypeDate:
                        dataField.Value = Convert.ToDateTime(dataFieldValue);
                        break;
                    case DataType.TypeDecimal:
                        dataField.Value = Convert.ToDecimal(dataFieldValue);
                        break;
                    case DataType.TypeDouble:
                        dataField.Value = Convert.ToDouble(dataFieldValue);
                        break;
                    case DataType.TypeInteger:
                        dataField.Value = Convert.ToInt32(dataFieldValue);
                        break;
                    case DataType.TypeLong:
                        dataField.Value = Convert.ToInt64(dataFieldValue);
                        break;
                    default:
                        dataField.Value = dataFieldValue;
                        break;
                }
                pi.Update();
            }
        }
        private void ListDataFields()
        {
            var procId = GetIntProperty(Constants.SoProperties.ProcessInstance.ProcessInstanceId, true);
            ServiceBroker.Service.ServiceObjects[0].Properties.InitResultTable();
            var dt = ServiceBroker.ServicePackage.ResultTable;

            using (_wfClient = ServiceBroker.K2Connection.GetWorkflowClientConnection())
            {
                var pi = _wfClient.OpenProcessInstance(procId);
                foreach (DataField dataField in pi.DataFields)
                {
                    var dRow = dt.NewRow();
                    dRow[Constants.SoProperties.ProcessInstance.DataFieldName] = dataField.Name;
                    string dataFieldValue;
                    switch (dataField.ValueType)
                    {
                        case DataType.TypeBinary:
                            dataFieldValue= Convert.ToBase64String((byte[])dataField.Value);
                            break;
                        case DataType.TypeDate:
                            dataFieldValue = Convert.ToDateTime(dataField.Value).ToString("yyyy-MM-dd HH:mm:ss");
                            break;
                        default:
                            dataFieldValue = Convert.ToString(dataField.Value);
                            break;
                    }
                    dRow[Constants.SoProperties.ProcessInstance.DataFieldValue] = dataFieldValue;
                    dt.Rows.Add(dRow);
                }
            }
        }
        //private void StartProcessInstance(bool startGeneric)
        //{
        //    string processName = ServiceBroker.Service.ServiceObjects[0].Methods[0].Name;
        //    if (!startGeneric)
        //    {
        //        processName = GetStringProperty(Constants.SoProperties.ProcessInstanceClient.ProcessName, true);
        //    }
        //    int processVersion = GetIntProperty(Constants.SoProperties.ProcessInstanceClient.ProcessVersion);

        //    ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
        //    serviceObject.Properties.InitResultTable();
        //    DataTable results = ServiceBroker.ServicePackage.ResultTable;


        //    using (CLIENT.Connection k2Con = this.ServiceBroker.K2Connection.GetWorkflowClientConnection())
        //    {

        //        CLIENT.ProcessInstance pi;

        //        if (processVersion > 0)
        //        {
        //            pi = k2Con.CreateProcessInstance(processName, processVersion);
        //        }
        //        else
        //        {
        //            pi = k2Con.CreateProcessInstance(processName);
        //        }

        //        string folio = GetStringProperty(Constants.SoProperties.ProcessInstanceClient.ProcessFolio);
        //        if (!string.IsNullOrEmpty(folio))
        //        {
        //            pi.Folio = folio;
        //        }


        //        if (startGeneric)
        //        {
        //            MethodParameters mParams = ServiceBroker.Service.ServiceObjects[0].Methods[0].MethodParameters;
        //            foreach (CLIENT.DataField df in pi.DataFields)
        //            {
        //                df.Value = GetDataFieldValue(mParams[df.Name].Value, df.ValueType);
        //            }
        //        }

        //        k2Con.StartProcessInstance(pi, GetBoolProperty(Constants.SoProperties.ProcessInstanceClient.StartSync));

        //        DataRow dr = results.NewRow();
        //        dr[Constants.SoProperties.ProcessInstanceClient.ProcessInstanceId] = pi.ID;
        //        dr[Constants.SoProperties.ProcessInstanceClient.ProcessFolio] = pi.Folio;
        //        results.Rows.Add(dr);

        //        k2Con.Close();
        //    }
        //}

        ///// <summary>
        ///// Used to map Workflow DataField types to SoType of SMOs.
        ///// </summary>
        ///// <param name="type">Type of the datafield</param>
        ///// <returns></returns>
        //private SoType GetDataFieldType(int type)
        //{
        //    switch (type)
        //    {
        //        case 1: //Boolean
        //            return SoType.YesNo;
        //        case 2: //DateTime
        //            return SoType.DateTime;
        //        case 3: //Decimal
        //            return SoType.Decimal;
        //        case 4: //Double
        //            return SoType.Decimal;
        //        case 5: //Integer
        //            return SoType.Number;
        //        case 6: //Long
        //            return SoType.Number;
        //        case 7: //String
        //            return SoType.Text;
        //        case 8: //Binary
        //            return SoType.Memo;
        //        default:
        //            return SoType.Memo;
        //    }
        //}
        ///// <summary>
        ///// Used to convert parameter value types to datafield ones.
        ///// </summary>
        ///// <param name="val">Value of the parameter</param>
        ///// <param name="dType">DataType of the DataField</param>
        ///// <returns></returns>
        //private object GetDataFieldValue(object val, CLIENT.DataType dType)
        //{
        //    switch (dType)
        //    {
        //        case CLIENT.DataType.TypeBinary:
        //            return Convert.FromBase64String(Convert.ToString(val));
        //        default:
        //            return val;
        //    }
        //}
    }
}
