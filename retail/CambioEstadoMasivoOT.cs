using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Configuration;
using VisualD.GlobalVid;
using VisualD.SBOFunctions;
using VisualD.vkBaseForm;
using VisualD.MultiFunctions;
using VisualD.vkFormInterface;
using VisualD.SBOObjectMg1;
using VisualD.Main;
using VisualD.MainObjBase;
using System.Threading;
using System.Data.SqlClient;
using SAPbouiCOM;
using SAPbobsCOM;
using System.IO;
using System.Data;
using VisualD.ADOSBOScriptExecute;
using VID_Retail.Utils;
using VID_Retail.AprobacionRecepcionTienda;

namespace VID_Retail.CambioEstadoMasivoOT
{
    public class TCambioEstadoMasivoOT : TvkBaseForm, IvkFormInterface
    {
        private List<string> Lista;
        private SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.Form oForm;
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        private String s;

        public new bool InitForm(string uid, string xmlPath, ref Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions SBOFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool Result = base.InitForm(uid, xmlPath, ref application, ref company, ref SBOFunctions, ref _GlobalSettings);

            oRecordSet = (SAPbobsCOM.Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            //Funciones.SBO_f = FSBOf;
            try
            {
                Lista = new List<string>();
                FSBOf.LoadForm(xmlPath, "CambioMasivoEstadoOT.srf", uid);
                EnableCrystal = false;

                oForm = FSBOApp.Forms.Item(uid);
                oForm.AutoManaged = true;
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", false);

                oForm.DataSources.UserDataSources.Add("EstadoNac", BoDataType.dt_SHORT_TEXT, 50);
                ((ComboBox)oForm.Items.Item("EstadoNac").Specific).DataBind.SetBound(true, "", "EstadoNac");

                oForm.DataSources.UserDataSources.Add("EstadoImp", BoDataType.dt_SHORT_TEXT, 50);
                ((ComboBox)oForm.Items.Item("EstadoImp").Specific).DataBind.SetBound(true, "", "EstadoImp");

                oForm.DataSources.UserDataSources.Add("Codbarra", BoDataType.dt_SHORT_TEXT, 50);
                ((EditText)oForm.Items.Item("CodBarra").Specific).DataBind.SetBound(true, "", "CodBarra");

                /*s = GlobalSettings.RunningUnderSQLServer ?
                       "Select Count(*) Cant from [@VIDR_PARAM] " :
                       "Select Count(*) \"Cant\" from \"@VIDR_PARAM\" ";
                oRecordSet.DoQuery(s);

                if (((Int32)oRecordSet.Fields.Item("Cant").Value) == 0)
                {
                    oForm.Mode = BoFormMode.fm_ADD_MODE;
                    oForm.SupportedModes = (Int32)BoAutoFormMode.afm_Add + (Int32)BoAutoFormMode.afm_Ok;
                }
                else
                {
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oForm.DataSources.DBDataSources.Item("@VIDR_PARAM").Query(null);
                    oForm.SupportedModes = (Int32)BoAutoFormMode.afm_Ok;
                }

                oUtil = new TUtils(ref oRecordSet, ref _GlobalSettings, false);
                */
                //oForm.Mode = BoFormMode.fm_ADD_MODE;
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable = oForm.DataSources.DataTables.Add("odt");

                s = @"SELECT ""Code"", ""Name""
                        FROM ""@VK_ESTANAC"" T0
                       ORDER BY ""Code"" ASC";
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(((ComboBox)oForm.Items.Item("EstadoNac").Specific), ref oRecordSet, true);

                s = @"SELECT ""Code"", ""Name""
                        FROM ""@VK_ESTAIMP"" T0
                       ORDER BY ""Code"" ASC";
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(((ComboBox)oForm.Items.Item("EstadoImp").Specific), ref oRecordSet, true);

                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("DocNum", BoFieldsType.ft_Integer, 12);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("callID", BoFieldsType.ft_Integer, 12);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("itemCode", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("itemName", BoFieldsType.ft_AlphaNumeric, 100);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("CodeOrigen", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("Origen", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("Marca", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("Modelo", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("EstadoOri", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("CodeEstadoOri", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("EstadoNvo", BoFieldsType.ft_AlphaNumeric, 50);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("CodeEstadoNvo", BoFieldsType.ft_AlphaNumeric, 50);


                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNum").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNum").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNum").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNum").RightJustified = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNum").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNum").TitleObject.Caption = "Nro Llamada";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").RightJustified = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").TitleObject.Caption = "Llave Llamada";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode").TitleObject.Caption = "Código Articulo";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName").TitleObject.Caption = "Descripción";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeOrigen").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeOrigen").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeOrigen").Visible = false;

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Origen").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Origen").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Origen").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Origen").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Origen").TitleObject.Caption = "Origen";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Marca").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Marca").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Marca").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Marca").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Marca").TitleObject.Caption = "Marca";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Modelo").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Modelo").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Modelo").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Modelo").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Modelo").TitleObject.Caption = "Modelo";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoOri").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoOri").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoOri").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoOri").TitleObject.Caption = "Estado Original";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeEstadoOri").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeEstadoOri").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeEstadoOri").Visible = false;

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoNvo").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoNvo").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoNvo").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("EstadoNvo").TitleObject.Caption = "Estado Nuevo";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeEstadoNvo").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeEstadoNvo").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("CodeEstadoNvo").Visible = false;


                return (Result);

            }
            catch (Exception e)
            {
                OutLog("InitForm: " + e.Message + " ** Trace: " + e.StackTrace);
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
            }
            finally
            {
                if (oForm != null)
                {
                    oForm.Freeze(false);
                    oForm.Visible = true;
                }
            }


            return Result;
        }//fin InitForm


        public new void FormEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, ref Boolean BubbleEvent)
        {
            base.FormEvent(FormUID, ref pVal, ref BubbleEvent);

            try
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if ((pVal.ItemUID == "btnBorrar") && (!pVal.BeforeAction))
                    {
                        BubbleEvent = false;
                        BorrarLinea();
                    }

                    if ((pVal.ItemUID == "btn_1") && (!pVal.BeforeAction))
                        ActualizarServicio();
                }

                if ((pVal.EventType == BoEventTypes.et_VALIDATE) && (pVal.ItemUID == "CodBarra") && (pVal.BeforeAction) && ((System.String)oForm.DataSources.UserDataSources.Item("CodBarra").Value).Trim() != "")
                {
                    BubbleEvent = false;
                    var CodigoBarra = ((System.String)oForm.DataSources.UserDataSources.Item("CodBarra").Value).Trim();
                    if (CodigoBarra != "")
                    {
                        try
                        {
                            oForm.Freeze(true);
                            BuscarCodigo(CodigoBarra);
                            oForm.DataSources.UserDataSources.Item("CodBarra").Value = "";
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_COMBO_SELECT) && (!pVal.BeforeAction) && (pVal.ItemUID == "EstadoNac"))
                {
                    if (((System.String)oForm.DataSources.UserDataSources.Item("EstadoNac").Value).Trim() != "")
                        ActualizarEstado(true, ((System.String)((ComboBox)oForm.Items.Item("EstadoNac").Specific).Selected.Value).Trim(), ((System.String)((ComboBox)oForm.Items.Item("EstadoNac").Specific).Selected.Description).Trim());
                }

                if ((pVal.EventType == BoEventTypes.et_COMBO_SELECT) && (!pVal.BeforeAction) && (pVal.ItemUID == "EstadoImp"))
                {
                    if (((System.String)oForm.DataSources.UserDataSources.Item("EstadoImp").Value).Trim() != "")
                        ActualizarEstado(false, ((System.String)((ComboBox)oForm.Items.Item("EstadoImp").Specific).Selected.Value).Trim(), ((System.String)((ComboBox)oForm.Items.Item("EstadoImp").Specific).Selected.Description).Trim());
                }
            }
            catch (Exception e)
            {
                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("FormEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }

        }//fin FormEvent


        public new void FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, ref Boolean BubbleEvent)
        {
            base.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);

            try
            {

            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("FormDataEvent: " + e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("FormDataEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }//fin FormDataEvent


        public new void MenuEvent(ref MenuEvent pVal, ref Boolean BubbleEvent)
        {
            //Int32 Entry;
            base.MenuEvent(ref pVal, ref BubbleEvent);
            try
            {
                //1281 Buscar; 
                //1282 Crear
                //1284 cancelar; 
                //1285 Restablecer;     
                //1286 Cerrar; 
                //1288 Registro siguiente;
                //1289 Registro anterior; 
                //1290 Primer Registro; 
                //1291 Ultimo Registro; 

            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
                OutLog("MenuEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }//fin MenuEvent


        private void ActualizarEstado(Boolean bEstadoNac, String CodeEstado, String EstadoDesc)
        {
            try
            {
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("grid").Specific).Rows.Count; i++)
                {
                    if ((bEstadoNac) && (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeOrigen", i)).Trim() == "1"))
                    {
                        ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("CodeEstadoNvo", i, CodeEstado);
                        ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("EstadoNvo", i, EstadoDesc);
                    }
                    if ((!bEstadoNac) && (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeOrigen", i)).Trim() == "2"))
                    {
                        ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("CodeEstadoNvo", i, CodeEstado);
                        ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("EstadoNvo", i, EstadoDesc);
                    }
                }
                ((Grid)oForm.Items.Item("grid").Specific).AutoResizeColumns();
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("ActualizarEstado: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }

        private void BuscarCodigo(String Codigo)
        {
            try
            {
                oForm.Freeze(true);
                s = @"SELECT T0.""DocNum""
                          ,T0.""callID""
                          ,T0.""itemCode""
                          ,T0.""itemName""
                          ,T1.""U_VK_Origen_articulo"" ""CodeOrigen""
                          ,IFNULL((SELECT B.""Descr""
                                     FROM ""CUFD"" A
                                     JOIN ""UFD1"" B ON B.""TableID"" = A.""TableID""
                                                  AND B.""FieldID"" = A.""FieldID""
                                    WHERE B.""TableID"" = 'OITM'
                                      AND A.""AliasID"" = 'VK_Origen_articulo'
                                      AND B.""FldValue"" = T1.""U_VK_Origen_articulo""),'') ""Origen""
                          ,T1.""U_VK_Marca"" ""Marca""
                          ,T1.""U_VK_Modelo"" ""Modelo""
                          ,CASE WHEN T1.""U_VK_Origen_articulo"" = '1' THEN IFNULL((SELECT ""Name"" FROM ""@VK_ESTANAC"" WHERE ""Code"" = T0.""U_VK_Estado_Nacional""),'')
                                ELSE IFNULL((SELECT ""Name"" FROM ""@VK_ESTAIMP"" WHERE ""Code"" = T0.""U_VK_Estado_Importado""),'')
                           END ""EstadoOri""
                          ,T0.""U_VK_Estado_Importado"" ""CodeEstadoOri""
                          ,CAST(' ' AS VARCHAR(50)) ""EstadoNvo""
                          ,CAST(' ' AS VARCHAR(50)) ""CodeEstadoNvo""
                      FROM ""OSCL"" T0
                      JOIN ""OITM"" T1 ON T1.""ItemCode"" = T0.""itemCode""
                      WHERE 1=1
                        AND T0.""DocNum"" =  {0}";
                s = String.Format(s, Codigo);
                oRecordSet.DoQuery(s);


                if (oRecordSet.RecordCount == 0)
                    FSBOApp.StatusBar.SetText("No se ha encontrado Llamada de servicio " + Codigo, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                else
                {
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Add(1);
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocNum", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count-1, ((System.Int32)oRecordSet.Fields.Item("DocNum").Value));
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("callID", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.Int32)oRecordSet.Fields.Item("callID").Value));
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("itemCode", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("itemCode").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("itemName", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("itemName").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("CodeOrigen", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("CodeOrigen").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("Origen", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("Origen").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("Modelo", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("Modelo").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("Marca", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("Marca").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("EstadoOri", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("EstadoOri").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("CodeEstadoOri", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("CodeEstadoOri").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("EstadoNvo", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("EstadoNvo").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("CodeEstadoNvo", ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count - 1, ((System.String)oRecordSet.Fields.Item("CodeEstadoNvo").Value).Trim());
                }

                ((Grid)oForm.Items.Item("grid").Specific).AutoResizeColumns();

            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
                OutLog("BuscarCodigo: " + e.Message + " ** Trace: " + e.StackTrace);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        

        private void BorrarLinea()
        {
            try
            {
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("grid").Specific).Rows.Count; i++)
                {
                    if (((Grid)oForm.Items.Item("grid").Specific).Rows.IsSelected(i))
                    {
                        ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Remove(i);
                        ((Grid)oForm.Items.Item("grid").Specific).AutoResizeColumns();
                        return;
                    }
                    else
                        continue;
                }
                FSBOApp.StatusBar.SetText("No ha seleccionado ninguna linea", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            catch (Exception x)
            {
                FSBOApp.MessageBox(x.Message + " ** Trace: " + x.StackTrace, 1, "Ok", "", "");
                OutLog("BorrarLinea: " + x.Message + " ** Trace: " + x.StackTrace);
            }
        }

        private void ActualizarServicio()
        {
            SAPbobsCOM.ServiceCalls oServiceCalls;
            Int32 lRetCode;
            Int32 errCode;
            String errMsg;
            try
            {
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("grid").Specific).Rows.Count; i++)
                {
                    try
                    {
                        if ((((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeEstadoOri", i)).Trim() != ((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeEstadoNvo", i)).Trim())
                           && (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeEstadoNvo", i)).Trim() != ""))
                        {
                            oServiceCalls = ((SAPbobsCOM.ServiceCalls)FCmpny.GetBusinessObject(BoObjectTypes.oServiceCalls));
                            if (oServiceCalls.GetByKey(((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("callID", i))))
                            {
                                if (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeOrigen", i)).Trim() == "1")//Nacional
                                    oServiceCalls.UserFields.Fields.Item("U_VK_Estado_Nacional").Value = ((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeEstadoNvo", i)).Trim();
                                else if (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeOrigen", i)).Trim() == "2")//Importado
                                    oServiceCalls.UserFields.Fields.Item("U_VK_Estado_Importado").Value = ((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("CodeEstadoNvo", i)).Trim();

                                lRetCode = oServiceCalls.Update();
                                if (lRetCode != 0)
                                {
                                    FCmpny.GetLastError(out errCode, out errMsg);
                                    FSBOApp.StatusBar.SetText("No se ha actualizado Llamada de servicio " + ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocNum", i)).ToString() + ", " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                                else
                                    FSBOApp.StatusBar.SetText("Se ha actualizado Llamada de servicio " + ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocNum", i)).ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            }
                        }

                    }
                    catch (Exception x)
                    {
                        FSBOApp.StatusBar.SetText("No se ha actualizado Llamada de servicio " + ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocNum", i)).ToString() + ", " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        OutLog("No se ha actualizado Llamada de servicio " + ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocNum", i)).ToString() + ", " + x.Message + ", ***TRACE " + x.StackTrace); 
                    }
                    finally
                    {
                        oServiceCalls = null;
                    }
                }
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Clear();
                FSBOApp.StatusBar.SetText("Termino de proceso", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception x)
            {
                FSBOApp.MessageBox(x.Message + " ** Trace: " + x.StackTrace, 1, "Ok", "", "");
                OutLog("ActualizarServicio: " + x.Message + " ** Trace: " + x.StackTrace);
            }
        }

    }//fin class
}
