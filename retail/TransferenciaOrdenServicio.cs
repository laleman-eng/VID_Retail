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

namespace VID_Retail.TransferenciaOrdenServicio
{
    public class TTransferenciaOrdenServicio : TvkBaseForm, IvkFormInterface
    {
        private List<string> Lista;
        private SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.Form oForm;
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        private SAPbouiCOM.Grid oGrid;
        private String s;

        public new bool InitForm(string uid, string xmlPath, ref Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions SBOFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool Result = base.InitForm(uid, xmlPath, ref application, ref company, ref SBOFunctions, ref _GlobalSettings);

            oRecordSet = (SAPbobsCOM.Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            //Funciones.SBO_f = FSBOf;
            try
            {
                Lista = new List<string>();
                FSBOf.LoadForm(xmlPath, "TransferenciaOrdenServicio.srf", uid);
                EnableCrystal = false;

                oForm = FSBOApp.Forms.Item(uid);
                oForm.AutoManaged = true;
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", false);

                oForm.DataSources.UserDataSources.Add("cbxOrigen", BoDataType.dt_SHORT_TEXT, 50);
                ((ComboBox)oForm.Items.Item("cbxOrigen").Specific).DataBind.SetBound(true, "", "cbxOrigen");

                oForm.DataSources.UserDataSources.Add("cbxDestino", BoDataType.dt_SHORT_TEXT, 50);
                ((ComboBox)oForm.Items.Item("cbxDestino").Specific).DataBind.SetBound(true, "", "cbxDestino");

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

                s = @"SELECT ""WhsCode"" ""Code"", ""WhsName"" ""Name""
                        FROM ""OWHS"" T0
                        JOIN ""@VIDR_TIENDA"" T1 ON T1.""Code"" = T0.""U_VID_Tda""
                        ORDER BY ""WhsName""";
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(((ComboBox)oForm.Items.Item("cbxOrigen").Specific), ref oRecordSet, true);


                s = @"SELECT ""WhsCode"" ""Code"", ""WhsName"" ""Name""
                        FROM ""OWHS"" T0
                        JOIN ""@VIDR_TIENDA"" T1 ON T1.""Code"" = T0.""U_VID_Tda""
                        ORDER BY ""WhsName""";
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(((ComboBox)oForm.Items.Item("cbxDestino").Specific), ref oRecordSet, true);

                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("Orden", BoFieldsType.ft_Integer, 12);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("callID", BoFieldsType.ft_Integer, 12);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("createDate", BoFieldsType.ft_Date, 40);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("customer", BoFieldsType.ft_AlphaNumeric, 50);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("custmrName", BoFieldsType.ft_AlphaNumeric, 150);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("itemCode", BoFieldsType.ft_AlphaNumeric, 50);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("itemName", BoFieldsType.ft_AlphaNumeric, 150);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("DocEntryT", BoFieldsType.ft_AlphaNumeric, 20);
                ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Columns.Add("DocNumT", BoFieldsType.ft_AlphaNumeric, 20);

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Orden").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Orden")).Editable = true;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Orden")).TitleObject.Caption = "Orden de Servicio";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("callID")).Visible = false;

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("createDate").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("createDate")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("createDate")).TitleObject.Caption = "Fecha Creación";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("customer").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("customer")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("customer")).TitleObject.Caption = "Código Cliente";
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("customer")).LinkedObjectType = "2";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("custmrName").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("custmrName")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("custmrName")).TitleObject.Caption = "Nombre Cliente";
                //((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("custmrName")).LinkedObjectType = "2";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode")).TitleObject.Caption = "Código Articulo";
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode")).LinkedObjectType = "4";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemName")).TitleObject.Caption = "Descripción Articulo";
                //((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("itemCode")).LinkedObjectType = "4";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocEntryT").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocEntryT")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocEntryT")).Visible = false;

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNumT").Type = BoGridColumnType.gct_EditText;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNumT")).Editable = false;
                ((EditTextColumn)((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocNumT")).Visible = false;

                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Add(1);
                ((Grid)oForm.Items.Item("grid").Specific).AutoResizeColumns();

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
                    oForm.Freeze(false);
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
                    if ((pVal.ItemUID == "btn_1") && (!pVal.BeforeAction))
                    {
                        BubbleEvent = false;
                        var BodegaOrigen = ((System.String)((ComboBox)oForm.Items.Item("cbxOrigen").Specific).Selected.Value).Trim();
                        var BodegaDestino = ((System.String)((ComboBox)oForm.Items.Item("cbxDestino").Specific).Selected.Value).Trim();
                        if (ValidarExistanDatos())
                            if ((BodegaOrigen == "") || (BodegaDestino == ""))
                                FSBOApp.StatusBar.SetText("Debe seleccionar Bodega origen y destino", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            else if (BodegaOrigen == BodegaDestino)
                                FSBOApp.StatusBar.SetText("Bodega origen y destino deben ser diferentes", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            else
                                GenerarTransferencia();
                    }

                    if ((pVal.ItemUID == "btnBorrar") && (!pVal.BeforeAction))
                    {
                        BubbleEvent = false;
                        BorrarLinea();
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_VALIDATE) && (pVal.ItemUID == "grid") && (pVal.ColUID == "Orden") && (pVal.BeforeAction)
                    && (((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", pVal.Row)) != 0))
                {
                    var OrdenServicio = ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", pVal.Row));
                    if ((OrdenServicio != 0) && (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("itemCode", pVal.Row)).Trim() == ""))
                    {
                        BubbleEvent = false;
                        try
                        {
                            oForm.Freeze(true);
                            BuscarLlamadasServicio(OrdenServicio.ToString(), pVal.Row);
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }
                    else
                        BubbleEvent = true;
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


        private Boolean ValidarExistanDatos()
        {

            try
            {
                for (Int32 i = 0; i < ((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count; i++)
                {
                    if (((System.Int32)((SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", i)) != 0)
                        return true;
                }

                return false;
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("ValidarExisteDatos: " + e.Message + " ** Trace: " + e.StackTrace);
                return false;
            }
        }

        private void BuscarLlamadasServicio(String Orden, Int32 iRow)
        {
            try
            {
                s = @"SELECT T0.""DocNum"" ""Orden""
                          ,T0.""callID"" ""callID""
                          ,T0.""createDate""
                          ,T0.""customer""
                          ,T0.""custmrName""
                          ,T0.""itemCode""
                          ,T0.""itemName""
                          ,T1.""InvntItem""
                      FROM ""OSCL"" T0
                      JOIN ""OITM"" T1 ON T1.""ItemCode"" = T0.""itemCode""
                     WHERE T0.""DocNum"" = {0}";
                s = String.Format(s, Orden);
                oRecordSet.DoQuery(s);

                if (oRecordSet.RecordCount == 0)
                {
                    FSBOApp.StatusBar.SetText("No se ha encontrado Orden de Servicio con el numero " + Orden, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("Orden", iRow, 0);
                    return;
                }
                else if (((System.String)oRecordSet.Fields.Item("InvntItem").Value).Trim() == "N")
                {
                    FSBOApp.StatusBar.SetText("Se encontro Orden de Servicio con el numero " + Orden + " pero el articulos no es inventariable", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("Orden", iRow, 0);
                    return;
                }
                else
                {
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("Orden", iRow, ((System.Int32)oRecordSet.Fields.Item("Orden").Value));
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("callID", iRow, ((System.Int32)oRecordSet.Fields.Item("callID").Value));
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("createDate", iRow, ((System.DateTime)oRecordSet.Fields.Item("createDate").Value));
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("customer", iRow, ((System.String)oRecordSet.Fields.Item("customer").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("custmrName", iRow, ((System.String)oRecordSet.Fields.Item("custmrName").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("itemCode", iRow, ((System.String)oRecordSet.Fields.Item("itemCode").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("itemName", iRow, ((System.String)oRecordSet.Fields.Item("itemName").Value).Trim());
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocEntryT", iRow, "");
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocNumT", iRow, "");

                }

                s = ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count.ToString();
                var oGrid = ((Grid)oForm.Items.Item("grid").Specific);
                //if (iRow+1 == ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count)
                //    oGrid.DataTable.Rows.Add(1);
                oGrid.DataTable.Rows.Add(1);
                oGrid.Columns.Item("Orden").Click(oGrid.Rows.Count - 1, false, 0);
                //oGrid.SetCellFocus(oGrid.Rows.Count - 1, 1);
                oGrid.CommonSetting.SetCellEditable(iRow + 1, 1, false);
                ((Grid)oForm.Items.Item("grid").Specific).AutoResizeColumns();
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("BuscarOrdenServicio: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }

        private void GenerarTransferencia()
        {
            Int32 Lineas;
            Int32 iLineas;
            Int32 iCantDesde = 0;
            Int32 iCantHasta = 0;
            Int32 lRetCode;
            Int32 errCode;
            String errMsg;
            Boolean Paso = true;
            String BodegaOrigen;
            String BodegaDestino;
            SAPbobsCOM.StockTransfer oStock = null;
            SAPbobsCOM.ServiceCalls oService = null;
            try
            {
                //Cantidad maximo de lineas para los documentos
                s = @"SELECT IFNULL(""U_CantLineas"",0) ""Lineas""
                        FROM ""@VID_FEPROCED""
                        WHERE ""U_TipoDoc"" = '52'
                        AND ""U_Habili"" = 'Y'";
                oRecordSet.DoQuery(s);
                if (oRecordSet.RecordCount == 0)
                    Lineas = 15;
                else
                    Lineas = ((Int32)oRecordSet.Fields.Item("Lineas").Value);

                BodegaOrigen = ((System.String)((ComboBox)oForm.Items.Item("cbxOrigen").Specific).Selected.Value).Trim();
                BodegaDestino = ((System.String)((ComboBox)oForm.Items.Item("cbxDestino").Specific).Selected.Value).Trim();

                iLineas = 0;
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count; i++)
                {
                    if ((((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("itemCode", i)).Trim() == "") || (((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", i)) == 0))
                        continue;
                    if (iLineas == 0)
                    {
                        Paso = false;
                        iCantDesde = i;
                        oStock = ((SAPbobsCOM.StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oStockTransfer));
                        oStock.DocDate = DateTime.Now;
                        //oStock.CardCode = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Tienda", i)).Trim();
                        oStock.Comments = "Cargado por Transferencias Orden servicio";
                        oStock.FromWarehouse = BodegaOrigen;
                        oStock.ToWarehouse = BodegaDestino;
                    }
                    else
                        oStock.Lines.Add();

                    oStock.Lines.ItemCode = ((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("itemCode", i)).Trim();
                    oStock.Lines.Quantity = 1; // ((System.Double)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Quantity", i));
                    oStock.Lines.FromWarehouseCode = BodegaOrigen;
                    oStock.Lines.WarehouseCode = BodegaDestino;
                    //oStock.Lines.UserFields.Fields.Item("U_CR_Nro_Contenedor").Value = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Contenedor", i));

                    //oStock.Lines.UserFields.Fields.Item("U_CR_Nro_Solicitud_Original").Value = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("DocEntry", i)).ToString().Trim();
                    //oStock.Lines.UserFields.Fields.Item("U_CR_Linea_ST_Origen").Value = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("LineNum", i)).ToString().Trim();
                    //oStock.Lines.BaseEntry = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("DocEntry", i));
                    //oStock.Lines.BaseLine = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("LineNum", i));
                    //oStock.Lines.BaseType = InvBaseDocTypeEnum.WarehouseTransfers;

                    iLineas++;
                    iCantHasta = i;
                    if (iLineas == Lineas)
                    {
                        Paso = true;
                        iLineas = 0;
                        if (i == Lineas - 1)
                            FCmpny.StartTransaction();
                        lRetCode = oStock.Add();
                        if (lRetCode != 0)
                        {
                            FCmpny.GetLastError(out errCode, out errMsg);
                            FSBOApp.StatusBar.SetText("Transferencia no ha sido creada, " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            if (FCmpny.InTransaction)
                                FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                            return;
                        }
                        else
                        {
                            var NewKey = FCmpny.GetNewObjectKey();
                            s = @"SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = {0}";
                            s = String.Format(s, NewKey);
                            oRecordSet.DoQuery(s);
                            var NewKeyNum = ((System.Int32)oRecordSet.Fields.Item("DocNum").Value).ToString().Trim();
                            FSBOApp.StatusBar.SetText("Transferencia creada " + NewKeyNum, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            oStock = null;
                            for (Int32 x = iCantDesde; x <= iCantHasta; x++)
                                if ((((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("itemCode", x)).Trim() != "") && (((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", x)) != 0))
                                {
                                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocEntryT", x, NewKey);
                                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocNumT", x, NewKeyNum);
                                }
                        }
                    }

                }//Fin for


                if (Paso == false) //para el ultimos registros restantes que no cuadran con tope de lineas
                {
                    if (!FCmpny.InTransaction)
                        FCmpny.StartTransaction();
                    lRetCode = oStock.Add();
                    if (lRetCode != 0)
                    {
                        FCmpny.GetLastError(out errCode, out errMsg);
                        FSBOApp.StatusBar.SetText("Transferencia no ha sido creada, " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        if (FCmpny.InTransaction)
                            FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                        return;
                    }
                    else
                    {
                        var NewKey = FCmpny.GetNewObjectKey();
                        s = @"SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = {0}";
                        s = String.Format(s, NewKey);
                        oRecordSet.DoQuery(s);
                        var NewKeyNum = ((System.Int32)oRecordSet.Fields.Item("DocNum").Value).ToString().Trim();
                        FSBOApp.StatusBar.SetText("Transferencia creada " + NewKeyNum, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        oStock = null;
                        for (Int32 x = iCantDesde; x <= iCantHasta; x++)
                            if ((((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("itemCode", x)).Trim() != "") && (((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", x)) != 0))
                            {
                                ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocEntryT", x, NewKey);
                                ((Grid)oForm.Items.Item("grid").Specific).DataTable.SetValue("DocNumT", x, NewKeyNum);
                            }
                    }
                }

                //actualizara las llamadas de servicio
                for (Int32 x = 0; x < ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Count; x++)
                {
                    if (((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocEntryT", x)).Trim() == "")
                        continue;
                    oService = ((SAPbobsCOM.ServiceCalls)FCmpny.GetBusinessObject(BoObjectTypes.oServiceCalls));
                    if (oService.GetByKey(((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("callID", x))))
                    {
                        if (oService.Expenses.Count > 1)
                            oService.Expenses.Add();
                        oService.Expenses.DocumentType = BoSvcEpxDocTypes.edt_StockTransfer;
                        oService.Expenses.DocEntry = Convert.ToInt32(((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocEntryT", x)).Trim().Replace(",","."), _nf);
                        oService.Expenses.DocumentNumber = Convert.ToInt32(((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocNumT", x)).Trim().Replace(",", "."), _nf);
                        oService.Expenses.StockTransferDirection = BoStckTrnDir.bos_TransferFromTechnician;
                        lRetCode = oService.Update();
                        if (lRetCode != 0)
                        {
                            FCmpny.GetLastError(out errCode, out errMsg);
                            FSBOApp.StatusBar.SetText("Llamada de servicio no se ha actualizado -> " + ((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", x).ToString().Trim() + ", " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            if (FCmpny.InTransaction)
                                FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                            return;
                        }
                        else
                        {
                            FSBOApp.StatusBar.SetText("Llamada de servicio actualizada -> " + ((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", x).ToString().Trim(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        }
                    }
                    oService = null;
                }//fin for

                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_Commit);

                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Clear();
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.Rows.Add(1);
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Orden").Click(oGrid.Rows.Count - 1, false, 0);
                FSBOApp.StatusBar.SetText("Proceso terminado - Transferencias creadas ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("GenerarTransferencia: " + e.Message + " ** Trace: " + e.StackTrace);
                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
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
                        for (Int32 x = i; x < ((Grid)oForm.Items.Item("grid").Specific).Rows.Count; x++)
                            if ((((System.String)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("itemCode", x)).Trim() == "") || (((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Orden", x)) == 0))
                                ((Grid)oForm.Items.Item("grid").Specific).CommonSetting.SetCellEditable(x + 1, 1, true);
                            else
                                ((Grid)oForm.Items.Item("grid").Specific).CommonSetting.SetCellEditable(x + 1, 1, false);
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

    }//fin class
}

