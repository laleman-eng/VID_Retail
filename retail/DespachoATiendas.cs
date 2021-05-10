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

namespace VID_Retail.DespachoATiendas
{
    public class TDespachoATiendas : TvkBaseForm, IvkFormInterface
    {
        private List<string> Lista;
        private SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.Form oForm;
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        private String s;
        private String TiendaTransito = "";

        public new bool InitForm(string uid, string xmlPath, ref Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions SBOFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool Result = base.InitForm(uid, xmlPath, ref application, ref company, ref SBOFunctions, ref _GlobalSettings);

            oRecordSet = (SAPbobsCOM.Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
            //Funciones.SBO_f = FSBOf;
            try
            {
                Lista = new List<string>();
                FSBOf.LoadForm(xmlPath, "DespachoATiendas.srf", uid);
                EnableCrystal = false;

                oForm = FSBOApp.Forms.Item(uid);
                oForm.AutoManaged = true;

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
                oForm.Mode = BoFormMode.fm_ADD_MODE;
                ((SAPbouiCOM.Grid)oForm.Items.Item("gridT").Specific).DataTable = oForm.DataSources.DataTables.Add("odtT");
                ((SAPbouiCOM.Grid)oForm.Items.Item("gridD").Specific).DataTable = oForm.DataSources.DataTables.Add("odtD");


                CargarTiendas();
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
                    if ((pVal.ItemUID == "btnRefr") && (!pVal.BeforeAction))
                    {
                        s = "Al actualizar listado tiendas se perdera cualquier selección anterior" + Environment.NewLine
                            + "¿ Desea continuar ?";

                        if (FSBOApp.MessageBox(s, 1, "Ok", "Cancelar", "") == 1)
                        {
                            ((Grid)oForm.Items.Item("gridT").Specific).DataTable.Clear();
                            ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Clear();
                            CargarTiendas();
                        }
                    }

                    if ((pVal.ItemUID == "gridD") && (pVal.ColUID == "Selec") && (!pVal.BeforeAction) && (pVal.Row >= 0))
                    {
                        var pos = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).GetDataTableRowIndex(pVal.Row));
                        var Contenedor = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Contenedor", pos));
                        var Valor = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Selec", pos)).Trim();
                        MarcarGrupoContenedor(Contenedor.ToString(), Valor);
                    }

                    if ((pVal.ItemUID == "1") && (pVal.BeforeAction))
                    {
                        BubbleEvent = false;
                        if (ValidarContenedoresSeleccionados())
                            CrearSolicitudes();
                            
                    }

                }

                if (pVal.EventType == BoEventTypes.et_CLICK)
                {
                    if ((pVal.ItemUID == "gridT") && ((pVal.ColUID == "RowsHeader") || (pVal.ColUID == "Tienda")) && (!pVal.BeforeAction) && (pVal.Row >= 0))
                    {
                        ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Clear();
                        var Tienda = ((System.String)((Grid)oForm.Items.Item("gridT").Specific).DataTable.GetValue("Tienda", pVal.Row)).Trim();
                        CargarDetalle(Tienda);
                        ((Grid)oForm.Items.Item("gridT").Specific).Rows.SelectedRows.Add(pVal.Row);
                    }

                    if ((pVal.ItemUID == "gridT") && ((pVal.ColUID == "RowsHeader") || (pVal.ColUID == "Tienda")) && (pVal.BeforeAction) && (pVal.Row >= 0))
                    {
                        if (((Grid)oForm.Items.Item("gridD").Specific).DataTable.Rows.Count > 0)
                        {
                            s = "Al actualizar listado tiendas se perdera cualquier selección anterior" + Environment.NewLine
                                + "¿ Desea continuar ?";

                            if (FSBOApp.MessageBox(s, 1, "Ok", "Cancelar", "") == 1)
                                BubbleEvent = true;
                            else
                                BubbleEvent = false;
                        }
                    }
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


        private void CargarTiendas()
        {
            try
            {
                s = @"SELECT ""U_WhsCodTR"" FROM ""@VIDR_PARAM"" ";
                oRecordSet.DoQuery(s);
                if (oRecordSet.RecordCount == 0)
                {
                    FSBOApp.StatusBar.SetText("Debe ingresar bodega de transito en los parametros del addon", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }
                else
                {
                    if (((System.String)oRecordSet.Fields.Item("U_WhsCodTR").Value).Trim() == "")
                    {
                        FSBOApp.StatusBar.SetText("Debe ingresar bodega de transito en los parametros del addon", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return;
                    }
                    else
                        TiendaTransito = ((System.String)oRecordSet.Fields.Item("U_WhsCodTR").Value).Trim();
                }

                s = @"SELECT S0.""CardCode"" ""Tienda""
                            ,S0.""CardName"" ""Descripcion""
                            ,COUNT(DISTINCT T1.""U_CR_Nro_Contenedor"") ""Contenedores""
                        FROM ""OWTR"" T0
                        JOIN ""WTR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                        JOIN ""WTQ1"" S1 ON S1.""DocEntry"" = T1.""BaseEntry""
                                    AND S1.""LineNum"" = T1.""BaseLine""
                                    AND S1.""ObjType"" = T1.""BaseType""
                        JOIN ""OWTQ"" S0 ON S0.""DocEntry"" = S1.""DocEntry""
                        JOIN ""OWHS"" W0 ON W0.""U_VR_SN"" = S0.""CardCode""
                        WHERE IFNULL(T1.""BaseType"",'-1') <> '-1'
                        AND IFNULL(T1.""U_CR_Nro_Contenedor"",0) <> 0
                        AND T1.""WhsCode"" = '{0}'
                        AND IFNULL(S0.""CardCode"",'') <> ''
                        AND IFNULL(T1.""U_CR_Estado"",'A') = 'A'
                        GROUP BY S0.""CardCode"", S0.""CardName""
                        ORDER BY S0.""CardCode"" ASC";
                s = String.Format(s, TiendaTransito);
                ((Grid)oForm.Items.Item("gridT").Specific).DataTable.ExecuteQuery(s);


                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Tienda").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Tienda").Editable = false;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Tienda").Visible = true;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Tienda").TitleObject.Sortable = true;

                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Descripcion").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Descripcion").Editable = false;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Descripcion").Visible = true;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Descripcion").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Descripcion").TitleObject.Caption = "Descripción";

                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Contenedores").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Contenedores").Editable = false;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Contenedores").Visible = true;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Contenedores").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("gridT").Specific).Columns.Item("Contenedores").RightJustified = true;

                ((Grid)oForm.Items.Item("gridT").Specific).AutoResizeColumns();

            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
                OutLog("CargarTiendas: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }

        private void CargarDetalle(String Tienda)
        {
            try
            {
                oForm.Freeze(true);
                s = @"SELECT T1.""U_CR_Nro_Contenedor"" ""Contenedor""
                            ,T1.""ItemCode""
                            ,T1.""Dscription""
                            ,T1.""Quantity""
                            ,'N' ""Selec""
                            ,T0.""DocNum"" 
  		                    ,T0.""DocEntry""
                            ,T1.""LineNum""
                            ,T1.""ObjType""
                            ,T1.""BaseEntry"" ""SolicitudOriginal""
                            ,T1.""BaseLine"" ""LineaOriginal""
                            ,S0.""CardCode"" ""Tienda""
                            ,W0.""WhsCode"" ""BodegaDestino""
                            ,T1.""WhsCode""  ""Transito""
                        FROM ""OWTR"" T0
                        JOIN ""WTR1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                        JOIN ""WTQ1"" S1 ON S1.""DocEntry"" = T1.""BaseEntry""
                                    AND S1.""LineNum"" = T1.""BaseLine""
                                    AND S1.""ObjType"" = T1.""BaseType""
                        JOIN ""OWTQ"" S0 ON S0.""DocEntry"" = S1.""DocEntry""
                        JOIN ""OWHS"" W0 ON W0.""U_VR_SN"" = S0.""CardCode""
                        WHERE IFNULL(T1.""BaseType"",'-1') <> '-1'
                        AND IFNULL(T1.""U_CR_Nro_Contenedor"",0) <> 0
                        --AND T1.""FromWhsCod"" = ''
                        AND T1.""WhsCode"" = '{1}'
                        AND IFNULL(S0.""CardCode"",'') = '{0}'
                        AND IFNULL(T1.""U_CR_Estado"",'A') = 'A'
                        ORDER BY T1.""U_CR_Nro_Contenedor""";
                s = String.Format(s, Tienda, TiendaTransito);
                ((Grid)oForm.Items.Item("gridD").Specific).DataTable.ExecuteQuery(s);

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Contenedor").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Contenedor").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Contenedor").Visible = true;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Contenedor").TitleObject.Sortable = true;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ItemCode").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ItemCode").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ItemCode").Visible = true;
                //((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ItemCode").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ItemCode").TitleObject.Caption = "Código SKU";

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Dscription").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Dscription").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Dscription").Visible = true;
                //((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Dscription").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Dscription").RightJustified = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Dscription").TitleObject.Caption = "Descripción SKU";

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Quantity").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Quantity").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Quantity").Visible = true;
                //((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Quantity").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Quantity").RightJustified = true;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Quantity").TitleObject.Caption = "Cantidad";

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Selec").Type = BoGridColumnType.gct_CheckBox;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Selec").Editable = true;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Selec").Visible = true;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("DocNum").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("DocNum").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("DocNum").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("DocEntry").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("DocEntry").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("DocEntry").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("LineNum").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("LineNum").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("LineNum").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ObjType").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ObjType").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("ObjType").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("SolicitudOriginal").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("SolicitudOriginal").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("SolicitudOriginal").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("LineaOriginal").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("LineaOriginal").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("LineaOriginal").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Tienda").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Tienda").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Tienda").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("BodegaDestino").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("BodegaDestino").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("BodegaDestino").Visible = false;

                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Transito").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Transito").Editable = false;
                ((Grid)oForm.Items.Item("gridD").Specific).Columns.Item("Transito").Visible = false;


                ((Grid)oForm.Items.Item("gridD").Specific).CollapseLevel = 1;
                ((Grid)oForm.Items.Item("gridD").Specific).AutoResizeColumns();

            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
                OutLog("CargarTiendas: " + e.Message + " ** Trace: " + e.StackTrace);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void MarcarGrupoContenedor(String Contenedor, String Valor)
        {
            try
            {
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Rows.Count; i++)
                {
                    if (((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Contenedor", i)).ToString().Trim() == Contenedor)
                    {
                        ((Grid)oForm.Items.Item("gridD").Specific).DataTable.SetValue("Selec", i, Valor);
                        if (i < ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Rows.Count-1)
                        {
                            if (((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Contenedor", i)).ToString().Trim() != ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Contenedor", i + 1)).ToString().Trim())
                                return;
                        }
                        else
                            continue;
                    }
                }
            }
            catch (Exception x)
            {
                FSBOApp.MessageBox(x.Message + " ** Trace: " + x.StackTrace, 1, "Ok", "", "");
                OutLog("MarcarGrupoContenedor: " + x.Message + " ** Trace: " + x.StackTrace);
            }
        }


        private Boolean ValidarContenedoresSeleccionados()
        {
            try
            {
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Rows.Count; i++)
                {
                    if (((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Selec", i)).Trim() == "Y")
                        return true;
                    else
                        continue;
                }
                FSBOApp.StatusBar.SetText("No ha seleccionado ningún Contenedor", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
            catch (Exception x)
            {
                FSBOApp.MessageBox(x.Message + " ** Trace: " + x.StackTrace, 1, "Ok", "", "");
                OutLog("ValidarContenedoresSeleccionados: " + x.Message + " ** Trace: " + x.StackTrace);
                return false;
            }
        }


        private void CrearSolicitudes ()
        {
            Int32 Lineas;
            Int32 iLineas;
            Int32 lRetCode;
            Int32 errCode;
            String errMsg;
            String Tienda;
            Boolean Paso = true;
            System.Data.DataTable dtResumen = new System.Data.DataTable();
            SAPbobsCOM.StockTransfer oStock = null;
            try
            {
                var dtcolumn = new System.Data.DataColumn();
                dtcolumn.DataType = System.Type.GetType("System.Double");
                dtcolumn.ColumnName = "DocEntry";
                dtResumen.Columns.Add(dtcolumn);

                dtcolumn = new System.Data.DataColumn();
                dtcolumn.DataType = System.Type.GetType("System.Int32");
                dtcolumn.ColumnName = "LineNum";
                dtResumen.Columns.Add(dtcolumn);

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

                Tienda = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Tienda", 0)).Trim();
                iLineas = 0;
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Rows.Count; i++)
                {
                    if (((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Selec", i)).Trim() == "Y")
                    {
                        if (iLineas == 0)
                        {
                            Paso = false;
                            oStock = ((SAPbobsCOM.StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oInventoryTransferRequest));
                            oStock.DocDate = DateTime.Now;
                            oStock.CardCode = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Tienda", i)).Trim();
                            oStock.Comments = "Cargado por Addon Despacho a Tienda";
                            oStock.FromWarehouse = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Transito", i)).Trim();
                            oStock.ToWarehouse = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("BodegaDestino", i)).Trim();
                        }
                        else
                            oStock.Lines.Add();

                        oStock.Lines.ItemCode = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("ItemCode", i)).Trim();
                        oStock.Lines.Quantity = ((System.Double)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Quantity", i));
                        oStock.Lines.FromWarehouseCode = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Transito", i)).Trim();
                        oStock.Lines.WarehouseCode = ((System.String)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("BodegaDestino", i)).Trim();
                        oStock.Lines.UserFields.Fields.Item("U_CR_Nro_Contenedor").Value = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("Contenedor", i));

                        //oStock.Lines.UserFields.Fields.Item("U_CR_Nro_Solicitud_Original").Value = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("DocEntry", i)).ToString().Trim();
                        //oStock.Lines.UserFields.Fields.Item("U_CR_Linea_ST_Origen").Value = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("LineNum", i)).ToString().Trim();
                        oStock.Lines.BaseEntry = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("DocEntry", i));
                        oStock.Lines.BaseLine = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("LineNum", i));
                        oStock.Lines.BaseType = InvBaseDocTypeEnum.WarehouseTransfers;


                        var dRow = dtResumen.NewRow();
                        dRow["DocEntry"] = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("DocEntry", i));
                        dRow["LineNum"] = ((System.Int32)((Grid)oForm.Items.Item("gridD").Specific).DataTable.GetValue("LineNum", i));
                        dtResumen.Rows.Add(dRow);
                        iLineas++;
                        if (iLineas == Lineas)
                        {
                            Paso = true;
                            iLineas = 0;
                            if (i == Lineas-1)
                                FCmpny.StartTransaction();
                            lRetCode = oStock.Add();
                            if (lRetCode != 0)
                            {
                                FCmpny.GetLastError(out errCode, out errMsg);
                                FSBOApp.StatusBar.SetText("Solicitud no ha sido creada, " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                if (FCmpny.InTransaction)
                                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                                return;
                            }
                            else
                            {
                                var NewKey = FCmpny.GetNewObjectKey();
                                FSBOApp.StatusBar.SetText("Solicitud creada " + NewKey, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                oStock = null;
                            }
                        }
                    }

                }//Fin for


                if (Paso == false) //para el ultimos registros restantes que no cuadran con tope de lineas
                {
                    lRetCode = oStock.Add();
                    if (lRetCode != 0)
                    {
                        FCmpny.GetLastError(out errCode, out errMsg);
                        FSBOApp.StatusBar.SetText("Solicitud no ha sido creada, " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        if (FCmpny.InTransaction)
                            FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                        return;
                    }
                    else
                    {
                        var NewKey = FCmpny.GetNewObjectKey();
                        FSBOApp.StatusBar.SetText("Solicitud creada " + NewKey, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                }

                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_Commit);
                //actualizo las transferencias originales cambian el CDU U_CR_Estado a D
                foreach (DataRow orow in dtResumen.Rows)
                {
                    var DocEntry = orow["DocEntry"].ToString().Trim();
                    var LineNum = orow["LineNum"].ToString().Trim();
                    s = @"UPDATE ""WTR1"" SET ""U_CR_Estado"" = 'D' WHERE ""DocEntry"" = {0} AND ""LineNum"" = {1}";
                    s = String.Format(s, DocEntry, LineNum);
                    oRecordSet.DoQuery(s);
                }

                //vuelvo a cargar la grid detalle
                //CargarDetalle(Tienda);
                ((Grid)oForm.Items.Item("gridD").Specific).DataTable.Clear();
                CargarTiendas();
                FSBOApp.StatusBar.SetText("Proceso terminado - solicitudes creadas ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception x)
            {
                FSBOApp.MessageBox(x.Message + " ** Trace: " + x.StackTrace, 1, "Ok", "", "");
                OutLog("CrearSolicitudes: " + x.Message + " ** Trace: " + x.StackTrace);
                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
            }
            finally
            {
                oStock = null;
            }

        }
    }//fin class
}
