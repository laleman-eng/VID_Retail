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

namespace VID_Retail.RecepcionenTiendas
{
    public class TRecepcionenTiendas : TvkBaseForm, IvkFormInterface
    {
        private List<string> Lista;
        private SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.DBDataSource oDBDSHeader;
        private SAPbouiCOM.DBDataSource oDBDSDetalle;
        private SAPbouiCOM.Form oForm;
        private TUtils oUtil; //= new TUtils();
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        private String s;

        public new bool InitForm(string uid, string xmlPath, ref Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions SBOFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool Result = base.InitForm(uid, xmlPath, ref application, ref company, ref SBOFunctions, ref _GlobalSettings);

            oRecordSet = (SAPbobsCOM.Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));            
            try
            {
                oUtil = new TUtils(ref oRecordSet, ref _GlobalSettings, false);
                oUtil.SBO_f = FSBOf;

                Lista = new List<string>();
                FSBOf.LoadForm(xmlPath, "RecepcionenTienda.srf", uid);
                EnableCrystal = false;

                oForm = FSBOApp.Forms.Item(uid);
                oForm.AutoManaged = true;
                ((Matrix)oForm.Items.Item("mtx").Specific).AutoResizeColumns();

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

                                            // Ok Ad  Fnd Vw Rq Sec
                Lista.Add("cbxTienda         , f,  t,  t,  f, n, 1");
                Lista.Add("CodBarra          , f,  t,  f,  f, n, 1");
                Lista.Add("chkBarra          , f,  t,  f,  f, n, 1");
                Lista.Add("DocEntry          , f,  f,  t,  f, n, 1");
                Lista.Add("btnBuscar         , f,  t,  f,  f, n, 1");
                Lista.Add("Contenedor        , f,  t,  t,  f, n, 1");
                FSBOf.SetAutoManaged(ref oForm, Lista);

                s = @"SELECT ""WhsCode"" ""Code"", ""WhsName"" ""Name""
                        FROM ""OWHS"" T0
                        JOIN ""@VIDR_TIENDA"" T1 ON T1.""Code"" = T0.""U_VID_Tda""
                        WHERE IFNULL(T0.""U_VR_SN"",'') <> ''
                        ORDER BY ""WhsName""";
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(((ComboBox)oForm.Items.Item("cbxTienda").Specific), ref oRecordSet, true);

                oForm.DataSources.UserDataSources.Add("CodBarra", BoDataType.dt_LONG_TEXT, 100);
                ((EditText)oForm.Items.Item("CodBarra").Specific).DataBind.SetBound(true, "", "CodBarra");

                oForm.DataSources.UserDataSources.Add("chkBarra", BoDataType.dt_SHORT_TEXT, 1);
                ((CheckBox)oForm.Items.Item("chkBarra").Specific).DataBind.SetBound(true, "", "chkBarra");
                ((CheckBox)oForm.Items.Item("chkBarra").Specific).ValOn = "Y";
                ((CheckBox)oForm.Items.Item("chkBarra").Specific).ValOff = "N";

                oDBDSHeader = ((DBDataSource)oForm.DataSources.DBDataSources.Item("@VID_RECTDA"));
                oDBDSDetalle = ((DBDataSource)oForm.DataSources.DBDataSources.Item("@VID_RECTDAD"));
                oForm.DataBrowser.BrowseBy = "DocEntry";

                ((Grid)oForm.Items.Item("grid").Specific).DataTable = oForm.DataSources.DataTables.Add("odt");

                ((CheckBox)oForm.Items.Item("chkBarra").Specific).Checked = true;
                ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Visible = false;
                ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyRec").Editable = false;
                oForm.Items.Item("3").Visible = true;
                oForm.Items.Item("CodBarra").Visible = true;

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
                    if ((pVal.ItemUID == "1") && (pVal.BeforeAction) && (oForm.Mode == BoFormMode.fm_ADD_MODE))
                    {
                        BubbleEvent = false;
                        CrearRegistro();
                    }

                    if ((pVal.ItemUID == "btnBuscar") && (!pVal.BeforeAction))
                    {
                        if (((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim() == "")
                            FSBOApp.StatusBar.SetText("Debe seleccionar Tienda", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        else if (((System.String)((EditText)oForm.Items.Item("Contenedor").Specific).Value).Trim() == "")
                            FSBOApp.StatusBar.SetText("Debe ingresar Contenedor", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        else
                            BuscarDetalleContenedor(((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim(), ((System.String)((EditText)oForm.Items.Item("Contenedor").Specific).Value).Trim());
                    }

                    if ((pVal.ItemUID == "chkBarra") && (!pVal.BeforeAction))
                    {
                        if (((System.Boolean)((CheckBox)oForm.Items.Item("chkBarra").Specific).Checked))
                        {
                            BuscarDetalleContenedor(((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim(), ((System.String)((EditText)oForm.Items.Item("Contenedor").Specific).Value).Trim());
                            ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Visible = false;
                            ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyRec").Editable = false;
                            oForm.Items.Item("3").Visible = true;
                            oForm.Items.Item("CodBarra").Visible = true;
                            ((EditText)oForm.Items.Item("CodBarra").Specific).Active = true;
                        }
                        else
                        {
                            ((EditText)oForm.Items.Item("xxx").Specific).Active = true;
                            BuscarDetalleContenedor(((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim(), ((System.String)((EditText)oForm.Items.Item("Contenedor").Specific).Value).Trim());
                            ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Visible = true;
                            ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyRec").Editable = true;
                            oForm.Items.Item("3").Visible = false;
                            oForm.Items.Item("CodBarra").Visible = false;
                        }
                    }

                    if ((pVal.ItemUID == "chkBarra") && (pVal.BeforeAction))
                    {
                        s = "Al actualizar forma recepción se perdera cualquier selección anterior" + Environment.NewLine
                               + "¿ Desea continuar ?";

                        if (FSBOApp.MessageBox(s, 1, "Ok", "Cancelar", "") == 1)
                            BubbleEvent = true;
                        else
                            BubbleEvent = false;
                    }

                    //if ((pVal.ItemUID == "grid") && ((pVal.ColUID == "RowsHeader") || (pVal.ColUID == "Contenedor")) && (!pVal.BeforeAction) && (pVal.Row >= 0))
                    if ((pVal.ItemUID == "grid") && (!pVal.BeforeAction) && (pVal.Row >= 0))
                    {
                        //((EditText)oForm.Items.Item("Contenedor").Specific).Value = ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Contenedor", pVal.Row)).ToString().Trim();
                        oDBDSHeader.SetValue("U_Contenedor", 0, ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Contenedor", pVal.Row)).ToString().Trim());
                        BuscarDetalleContenedor(((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim(), ((System.String)((EditText)oForm.Items.Item("Contenedor").Specific).Value).Trim());
                        ((Grid)oForm.Items.Item("grid").Specific).Rows.SelectedRows.Add(pVal.Row);
                    }
                }


                if ((pVal.EventType == BoEventTypes.et_COMBO_SELECT) && (pVal.ItemUID == "cbxTienda") && (!pVal.BeforeAction))
                {
                    if (((ComboBox)oForm.Items.Item("cbxTienda").Specific) != null)
                    {
                        if (((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim() != "")
                            BuscarContenedores();
                    }
                }

                if ((pVal.EventType == BoEventTypes.et_VALIDATE) && (pVal.ItemUID == "Contenedor") && (!pVal.BeforeAction) && (((System.String)oDBDSHeader.GetValue("U_Contenedor", 0)).Trim() != ""))
                {
                    for (Int32 i = 0; i < ((Grid)oForm.Items.Item("grid").Specific).Rows.Count; i++)
                    {
                        if (((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("Contenedor", i)).ToString().Trim() == ((System.String)oDBDSHeader.GetValue("U_Contenedor", 0)).Trim())
                        {
                            ((Grid)oForm.Items.Item("grid").Specific).Rows.SelectedRows.Add(i);
                            break;
                        }
                    }
                    BuscarDetalleContenedor(((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim(), ((System.String)oDBDSHeader.GetValue("U_Contenedor", 0)).Trim());
                    //oForm.Items.Item("CodBarra").Click(BoCellClickType.ct_Regular);
                    BubbleEvent = false;
                }

                if ((pVal.EventType == BoEventTypes.et_VALIDATE) && (pVal.ItemUID == "CodBarra") && (pVal.BeforeAction) && ((System.String)oForm.DataSources.UserDataSources.Item("CodBarra").Value).Trim() != "")
                {
                    BubbleEvent = false;
                    var CodigoBarraSKU = ((System.String)oForm.DataSources.UserDataSources.Item("CodBarra").Value).Trim();
                    if (CodigoBarraSKU != "")
                    {
                        try
                        {
                            oForm.Freeze(true);
                            BuscarSKU(CodigoBarraSKU);
                            //((EditText)oForm.Items.Item("CodBarra").Specific).Value = " ";
                            oForm.DataSources.UserDataSources.Item("CodBarra").Value = "";
                            //((EditText)oForm.Items.Item("xxx").Specific).Active = true;
                            //oForm.Items.Item("xxx").Click(BoCellClickType.ct_Regular);
                            //((EditText)oForm.Items.Item("CodBarra").Specific).Active = true;


                        }
                        finally
                        {
                            oForm.Freeze(false);
                            ///oForm.Items.Item("xxx").Click(BoCellClickType.ct_Regular);
                        }
                    }

                }
                if ((pVal.EventType == BoEventTypes.et_CLICK) && (pVal.ItemUID == "xxx") && (!pVal.BeforeAction))
                {
                    ;// oForm.Items.Item("CodBarra").Click(BoCellClickType.ct_Regular);
                    //((EditText)oForm.Items.Item("CodBarra").Specific).Active = true;
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

                if ((pVal.MenuUID != "") && (pVal.BeforeAction == false))
                {
                    if (pVal.MenuUID == "1282")
                    {
                        ((CheckBox)oForm.Items.Item("chkBarra").Specific).Checked = true;
                    }

                    if ((pVal.MenuUID == "1288") || (pVal.MenuUID == "1289") || (pVal.MenuUID == "1290") || (pVal.MenuUID == "1291"))
                        ((Matrix)oForm.Items.Item("mtx").Specific).AutoResizeColumns();
                }
            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message, 1, "Ok", "", "");
                OutLog("MenuEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }//fin MenuEvent


        private void BuscarDetalleContenedor(String Tienda, String Contenedor)
        {
            String TiendaTransito = "";
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

                s = @"SELECT T0.""DocEntry""
                            ,T1.""LineNum"" ""BaseLine""
                            ,T1.""ItemCode""
                            ,T1.""Dscription""
                            ,T1.""Quantity"" ""Qty""
                            ,0.0 ""QtyRec""
                            ,T1.""Quantity"" ""QtyCon""
                            ,IFNULL(T0.""FolioNum"",0) ""FolioNum""
                            ,CAST(' ' AS VARCHAR(100)) ""Comentario""
                        FROM ""OWTQ"" T0
                        JOIN ""WTQ1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                        JOIN ""OWHS"" W1 ON W1.""WhsCode"" = T1.""WhsCode"" -->para recepcion
                                         AND W1.""U_VR_SN"" = T0.""CardCode""
                        WHERE IFNULL(T1.""BaseType"",'-1') <> '-1'
                        AND T1.""U_CR_Nro_Contenedor"" = '{1}'  --Contenedor
                        AND IFNULL(W1.""U_VID_Tda"",'') <> '' --> para recepcion
                        AND T1.""WhsCode"" = '{2}'
                        --AND IFNULL(T0.""CardCode"",'') = '{2}'
                        AND IFNULL(T1.""U_CR_Estado"",'A') = 'A'
                        AND T1.""FromWhsCod"" = '{0}' --Tienda Transito
                        ORDER BY T1.""DocEntry"" ASC, T1.""VisOrder"" ASC; ";
                s = String.Format(s, TiendaTransito, Contenedor, Tienda);

                oDBDSDetalle.Clear();
                ((Matrix)oForm.Items.Item("mtx").Specific).LoadFromDataSource();
                oRecordSet.DoQuery(s);
                var i = 0;
                while (!oRecordSet.EoF)
                {
                    oDBDSDetalle.InsertRecord(i);
                    oDBDSDetalle.SetValue("U_ItemCode", i, ((System.String)oRecordSet.Fields.Item("ItemCode").Value).Trim());
                    oDBDSDetalle.SetValue("U_Descrip", i, ((System.String)oRecordSet.Fields.Item("Dscription").Value).Trim());
                    oDBDSDetalle.SetValue("U_Qty", i, ((System.Double)oRecordSet.Fields.Item("Qty").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_QtyRec", i, ((System.Double)oRecordSet.Fields.Item("QtyRec").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_QtyAcep", i, ((System.Double)oRecordSet.Fields.Item("QtyRec").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_QtyCon", i, ((System.Double)oRecordSet.Fields.Item("QtyCon").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_FolioNum", i, ((System.Int32)oRecordSet.Fields.Item("FolioNum").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_DocEntry", i, ((System.Int32)oRecordSet.Fields.Item("DocEntry").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_BaseLine", i, ((System.Int32)oRecordSet.Fields.Item("BaseLine").Value).ToString().Replace(",", "."));
                    oDBDSDetalle.SetValue("U_Comentario", i, ((System.String)oRecordSet.Fields.Item("Comentario").Value).Trim());
                    oDBDSDetalle.SetValue("U_Estado", i, "R");
                    oDBDSDetalle.SetValue("U_EstadoAn", i, "R");
                    i++;
                    oRecordSet.MoveNext();
                }

                ((Matrix)oForm.Items.Item("mtx").Specific).LoadFromDataSource();
                ((Matrix)oForm.Items.Item("mtx").Specific).AutoResizeColumns();


            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("Error BuscarDetalleContenedor: " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(x.Message + " , **TRACE " + x.StackTrace);
            }
        }


        private void BuscarSKU(String CodigoBarra)
        {
            SAPbouiCOM.Matrix mtx;
            Boolean bExiste = false;
            String SKU;
            try
            {
                mtx = ((Matrix)oForm.Items.Item("mtx").Specific);
                mtx.FlushToDataSource();
                oForm.Freeze(true);

                s = @"SELECT ""ItemCode"" from ""OBCD"" WHERE ""BcdCode"" = '{0}'";
                s = String.Format(s, CodigoBarra);
                oRecordSet.DoQuery(s);
                if (oRecordSet.RecordCount > 0)
                    SKU = ((System.String)oRecordSet.Fields.Item("ItemCode").Value).Trim();
                else
                {
                    FSBOApp.StatusBar.SetText("Codigo Barra " + CodigoBarra + " no existe en Maestro articulos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }

                for (Int32 i = 0; i < mtx.RowCount; i++)
                {
                    if (((System.String)oDBDSDetalle.GetValue("U_ItemCode", i)).Trim() == SKU)
                    {
                        oDBDSDetalle.SetValue("U_QtyRec", i, (Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyRec", i)).Replace(",", "."), _nf) + 1).ToString().Replace(",", "."));
                        oDBDSDetalle.SetValue("U_QtyAcep", i, (Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyAcep", i)).Replace(",", "."), _nf) + 1).ToString().Replace(",", "."));
                        oDBDSDetalle.SetValue("U_QtyCon", i, (Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyCon", i)).Replace(",", "."), _nf) - 1).ToString().Replace(",", "."));
                        bExiste = true;
                    }
                }

                if (!bExiste)
                {
                    s = @"SELECT COUNT(*) ""Cant"" FROM ""OITM"" WHERE ""ItemCode"" = '{0}' ";
                    s = String.Format(s, SKU);
                    oRecordSet.DoQuery(s);
                    if (((System.Int32)oRecordSet.Fields.Item("Cant").Value) == 0)
                        FSBOApp.StatusBar.SetText("SKU " + SKU + " no existe en Maestro articulos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    else
                    {
                        s = "SKU " + SKU + " no se encuentra incluido en el contenedor" + Environment.NewLine
                            + "¿ Desea ingresarlo al listado de recepción ?";

                        if (FSBOApp.MessageBox(s, 1, "Si", "Cancelar", "") == 1)
                        {

                            s = @"SELECT ""ItemName"" FROM ""OITM"" WHERE ""ItemCode"" = '{0}' ";
                            s = String.Format(s, SKU);
                            oRecordSet.DoQuery(s);
                            oDBDSDetalle.InsertRecord(mtx.RowCount);
                            oDBDSDetalle.SetValue("U_ItemCode", mtx.RowCount, SKU);
                            oDBDSDetalle.SetValue("U_Descrip", mtx.RowCount, ((System.String)oRecordSet.Fields.Item("ItemName").Value).Trim());
                            oDBDSDetalle.SetValue("U_Qty", mtx.RowCount, "0");
                            oDBDSDetalle.SetValue("U_QtyRec", mtx.RowCount, "1");
                            oDBDSDetalle.SetValue("U_QtyAcep", mtx.RowCount, "1");
                            oDBDSDetalle.SetValue("U_QtyCon", mtx.RowCount, "-1");
                            oDBDSDetalle.SetValue("U_FolioNum", mtx.RowCount, "0");
                            oDBDSDetalle.SetValue("U_DocEntry", mtx.RowCount, "0");
                            oDBDSDetalle.SetValue("U_BaseLine", mtx.RowCount, "0");
                            oDBDSDetalle.SetValue("U_Comentario", mtx.RowCount, "");
                            oDBDSDetalle.SetValue("U_Estado", mtx.RowCount, "R");
                            oDBDSDetalle.SetValue("U_EstadoAn", mtx.RowCount, "R");
                        }
                    }
                }

                mtx.LoadFromDataSource();
                mtx.AutoResizeColumns();

            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("Error BuscarSKU: " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(x.Message + " , **TRACE " + x.StackTrace);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void BuscarContenedores()
        {
            String TiendaTransito;
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

                s = @"SELECT T1.""U_CR_Nro_Contenedor"" ""Contenedor""
                            ,IFNULL(T0.""FolioNum"",0) ""FolioNum""
                            ,COUNT(DISTINCT T1.""ItemCode"") ""Cant""
                        FROM ""OWTQ"" T0
                        JOIN ""WTQ1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                        JOIN ""OWHS"" W1  ON W1.""WhsCode"" = T1.""WhsCode"" -->para recepcion
                                         AND W1.""U_VR_SN"" = T0.""CardCode""
                        WHERE IFNULL(T1.""BaseType"",'-1') <> '-1'
                        AND IFNULL(W1.""U_VID_Tda"",'') <> '' --> para recepcion
                        AND T1.""WhsCode"" = '{1}'
                        AND IFNULL(T1.""U_CR_Estado"",'A') = 'A'
                        AND T1.""FromWhsCod"" = '{0}' --Tienda Transito
                        GROUP BY T1.""U_CR_Nro_Contenedor"", IFNULL(T0.""FolioNum"",0)
                        ORDER BY T1.""U_CR_Nro_Contenedor""; ";
                s = String.Format(s, TiendaTransito, ((System.String)((ComboBox)oForm.Items.Item("cbxTienda").Specific).Selected.Value).Trim());

                ((Grid)oForm.Items.Item("grid").Specific).DataTable.ExecuteQuery(s);

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").TitleObject.Caption = "Contenedor";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("FolioNum").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("FolioNum").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("FolioNum").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("FolioNum").TitleObject.Caption = "Nro. Guia";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cant").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cant").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cant").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cant").TitleObject.Caption = "Cant. SKU";
            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText(x.Message + " ** Trace: " + x.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("BuscarContenedores: " + x.Message + " ** Trace: " + x.StackTrace);
            }
        }


        private Boolean CrearRegistro()
        {
            SAPbouiCOM.Matrix mtx;
            try
            {
                ((Matrix)oForm.Items.Item("mtx").Specific).FlushToDataSource();
                mtx = ((Matrix)oForm.Items.Item("mtx").Specific);

                if (((CheckBox)oForm.Items.Item("chkBarra").Specific).Checked == false)
                {
                    for (Int32 i = 0; i < mtx.RowCount; i++)
                    {
                        oDBDSDetalle.SetValue("U_QtyAcep", i, Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyRec", i)).Replace(",", "."), _nf).ToString().Replace(",", "."));
                        oDBDSDetalle.SetValue("U_QtyCon", i, (Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyCon", i)).Replace(",", "."), _nf) - Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyRec", i)).Replace(",", "."), _nf)).ToString().Replace(",", "."));
                    }
                }

                var lRetCode = oUtil.AddDataSourceInt1("VID_RECTDA", oDBDSHeader, "VID_RECTDAD", oDBDSDetalle, "", null, "", null);
                if (lRetCode == 0)
                {
                    FSBOApp.StatusBar.SetText("No se ha creado registro, revisar vd.log", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else
                {
                    var i = 0;
                    while (i < oDBDSDetalle.Size)
                    {
                        var DocEntry = ((System.String)oDBDSDetalle.GetValue("U_DocEntry", i)).Trim();
                        var LineNum = ((System.String)oDBDSDetalle.GetValue("U_BaseLine", i)).Trim();
                        s = @"UPDATE ""WTQ1"" SET ""U_CR_Estado"" = 'D' WHERE ""DocEntry"" = {0} AND ""LineNum"" = {1}";
                        s = String.Format(s, DocEntry, LineNum);
                        oRecordSet.DoQuery(s);
                        i++;
                    }
                    ((Grid)oForm.Items.Item("grid").Specific).DataTable.Clear();
                    if (!((System.Boolean)((CheckBox)oForm.Items.Item("chkBarra").Specific).Checked))
                    {
                        ((EditText)oForm.Items.Item("xxx").Specific).Active = true;
                        ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Visible = false;
                        ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyRec").Editable = false;
                    }
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    return true;
                }
            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("Error BuscarSKU: " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(x.Message + " , **TRACE " + x.StackTrace);
                return false;
            }

        }
    }//fin class
}
