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

namespace VID_Retail.AprobacionRecepcionTienda
{
    public class TAprobacionRecepcionTienda : TvkBaseForm, IvkFormInterface
    {
        private List<string> Lista;
        private SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.DBDataSource oDBDSHeader;
        private SAPbouiCOM.DBDataSource oDBDSDetalle;
        private SAPbouiCOM.Form oForm;
        private TUtils oUtil; //= new TUtils();
        private CultureInfo _nf = new System.Globalization.CultureInfo("en-US");
        private String s;
        private System.Data.DataTable dts;

        public static Int32 DocNum
        { get; set; }

        public new bool InitForm(string uid, string xmlPath, ref Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions SBOFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool Result = base.InitForm(uid, xmlPath, ref application, ref company, ref SBOFunctions, ref _GlobalSettings);

            oRecordSet = (SAPbobsCOM.Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));            
            try
            {
                oUtil = new TUtils(ref oRecordSet, ref _GlobalSettings, false);
                oUtil.SBO_f = FSBOf;

                Lista = new List<string>();
                FSBOf.LoadForm(xmlPath, "RecepcionenTiendaAprobacion.srf", uid);
                EnableCrystal = false;

                oForm = FSBOApp.Forms.Item(uid);
                oForm.AutoManaged = true;
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", false);

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
                Lista.Add("cbxTienda         , f,  f,  t,  f, n, 1");
                Lista.Add("DocEntry          , f,  f,  t,  f, n, 1");
                Lista.Add("btnBuscar         , f,  f,  f,  f, n, 1");
                Lista.Add("Contenedor        , f,  f,  t,  f, n, 1");
                Lista.Add("btnApro           , t,  t,  f,  f, n, 1");
                FSBOf.SetAutoManaged(ref oForm, Lista);


                oDBDSHeader = ((DBDataSource)oForm.DataSources.DBDataSources.Item("@VID_RECTDA"));
                oDBDSDetalle = ((DBDataSource)oForm.DataSources.DBDataSources.Item("@VID_RECTDAD"));
                //oForm.DataBrowser.BrowseBy = "DocEntry";

                //((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Visible = false;
                //((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyRec").Editable = false;
                oForm.Mode = BoFormMode.fm_FIND_MODE;
                ((EditText)oForm.Items.Item("DocEntry").Specific).Value = DocNum.ToString();
                s = "1";
                oForm.Items.Item(s).Click(BoCellClickType.ct_Regular);

                //FSBOApp.ActivateMenuItem("1291");
                ((Matrix)oForm.Items.Item("mtx").Specific).AutoResizeColumns();

                dts = new System.Data.DataTable();
                var dtcolumn = new System.Data.DataColumn();
                dtcolumn.DataType = System.Type.GetType("System.String");
                dtcolumn.ColumnName = "ItemCode";
                dts.Columns.Add(dtcolumn);

                dtcolumn = new System.Data.DataColumn();
                dtcolumn.DataType = System.Type.GetType("System.String");
                dtcolumn.ColumnName = "BaseEntry";
                dts.Columns.Add(dtcolumn);

                dtcolumn = new System.Data.DataColumn();
                dtcolumn.DataType = System.Type.GetType("System.String");
                dtcolumn.ColumnName = "BaseLine";
                dts.Columns.Add(dtcolumn);

                dtcolumn = new System.Data.DataColumn();
                dtcolumn.DataType = System.Type.GetType("System.Double");
                dtcolumn.ColumnName = "Quantity";
                dts.Columns.Add(dtcolumn);

                Bloquear();

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
                    if ((pVal.ItemUID == "1") && (pVal.BeforeAction) && (oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                    {
                        BubbleEvent = false;
                        if (ActualizarRegistro())
                            ;//FSBOApp.ActivateMenuItem("1304");
                    }

                    if ((pVal.ItemUID == "btnApro") && (!pVal.BeforeAction))
                    {
                        BubbleEvent = false;
                        AprobarTodo();
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
                //1304 Actualizar
                if ((pVal.MenuUID != "") && (pVal.BeforeAction == false))
                {
                    if ((pVal.MenuUID == "1288") || (pVal.MenuUID == "1289") || (pVal.MenuUID == "1290") || (pVal.MenuUID == "1291") || (pVal.MenuUID == "1304"))
                    {
                        ((Matrix)oForm.Items.Item("mtx").Specific).AutoResizeColumns();
                        Bloquear();
                    }
                }
            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message, 1, "Ok", "", "");
                OutLog("MenuEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }//fin MenuEvent


        private Boolean ActualizarRegistro()
        {
            SAPbobsCOM.StockTransfer oStock = null;
            Int32 i, l;
            Int32 errCode;
            String errMsg;
            Int32 lRetCode;
            String TiendaTransito;
            String CardCode;
            Int32 Aceptados;
            Boolean bCrearT = false;
            Boolean bCrearTDif = false;
            String BodegaDiferencia;
            SAPbouiCOM.Matrix mtx;

            try
            {
                dts.Rows.Clear();
                mtx = ((Matrix)oForm.Items.Item("mtx").Specific);
                mtx.FlushToDataSource();
                s = @"SELECT ""U_WhsCodTR"", ""U_WhsCodLF"" FROM ""@VIDR_PARAM"" ";
                oRecordSet.DoQuery(s);
                if (oRecordSet.RecordCount == 0)
                {
                    FSBOApp.StatusBar.SetText("Debe ingresar bodega de transito y diferencia de inventario en los parametros del addon", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else
                {
                    if (((System.String)oRecordSet.Fields.Item("U_WhsCodTR").Value).Trim() == "")
                    {
                        FSBOApp.StatusBar.SetText("Debe ingresar bodega de transito en los parametros del addon", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else if (((System.String)oRecordSet.Fields.Item("U_WhsCodLF").Value).Trim() == "")
                    {
                        FSBOApp.StatusBar.SetText("Debe ingresar bodega diferencia inventario en los parametros del addon", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else
                    {
                        TiendaTransito = ((System.String)oRecordSet.Fields.Item("U_WhsCodTR").Value).Trim();
                        BodegaDiferencia = ((System.String)oRecordSet.Fields.Item("U_WhsCodLF").Value).Trim();
                    }
                }

                s = @"SELECT ""U_VR_SN"" FROM ""OWHS"" WHERE ""WhsCode"" = '{0}'";
                s = String.Format(s, ((System.String)oDBDSHeader.GetValue("U_Tienda", 0)).Trim());
                oRecordSet.DoQuery(s);
                CardCode = ((System.String)oRecordSet.Fields.Item("U_VR_SN").Value).Trim();


                //Debe crear las transferencias para las lineas que esten aprobadas y que el Estado halla cambiado
                oStock = ((SAPbobsCOM.StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oStockTransfer));
                oStock.DocDate = DateTime.Now;
                oStock.CardCode = CardCode;
                oStock.Comments = "Cargado por Addon Recepción en Tienda";
                oStock.FromWarehouse = TiendaTransito;
                oStock.ToWarehouse = ((System.String)oDBDSHeader.GetValue("U_Tienda", 0)).Trim();
                i = 0;
                l = 0;
                Aceptados = 0;
                while (i < oDBDSDetalle.Size)
                {
                    if ((((System.String)oDBDSDetalle.GetValue("U_Estado", i)).Trim() == "A") && (((System.String)oDBDSDetalle.GetValue("U_Estado", i)).Trim() != ((System.String)oDBDSDetalle.GetValue("U_EstadoAn", i)).Trim()))
                    {
                        bCrearT = true;
                        if (l > 0)
                            oStock.Lines.Add();
                        oStock.Lines.ItemCode = ((System.String)oDBDSDetalle.GetValue("U_ItemCode", i)).Trim();
                        oStock.Lines.Quantity = Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyAcep", i)).Trim().Replace(",", "."), _nf);
                        oStock.Lines.FromWarehouseCode = TiendaTransito;
                        oStock.Lines.WarehouseCode = ((System.String)oDBDSHeader.GetValue("U_Tienda", 0)).Trim();
                        oStock.Lines.UserFields.Fields.Item("U_CR_Nro_Contenedor").Value = Convert.ToInt32(((System.String)oDBDSHeader.GetValue("U_Contenedor", 0)).Trim().Replace(",", "."), _nf);

                        if ((((System.String)oDBDSDetalle.GetValue("U_DocEntry", i)).Trim() != "") && (((System.String)oDBDSDetalle.GetValue("U_DocEntry", i)).Trim() != "0"))
                        {
                            oStock.Lines.BaseEntry = Convert.ToInt32(((System.String)oDBDSDetalle.GetValue("U_DocEntry", i)).Trim().Replace(",", "."), _nf);
                            oStock.Lines.BaseLine = Convert.ToInt32(((System.String)oDBDSDetalle.GetValue("U_BaseLine", i)).Trim().Replace(",", "."), _nf);
                            oStock.Lines.BaseType = InvBaseDocTypeEnum.InventoryTransferRequest;

                            var QAcep = Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_QtyAcep", i)).Trim().Replace(",", "."), _nf);
                            var QSol = Convert.ToDouble(((System.String)oDBDSDetalle.GetValue("U_Qty", i)).Trim().Replace(",", "."), _nf);
                            if ((QAcep - QSol) < 0)
                            {//Bodega R1017
                                bCrearTDif = true;
                                var dRow = dts.NewRow();
                                dRow["ItemCode"] = ((System.String)oDBDSDetalle.GetValue("U_ItemCode", i)).Trim();
                                dRow["BaseEntry"] = ((System.String)oDBDSDetalle.GetValue("U_DocEntry", i)).Trim().Replace(",", ".");
                                dRow["BaseLine"] = ((System.String)oDBDSDetalle.GetValue("U_BaseLine", i)).Trim().Replace(",", ".");
                                dRow["Quantity"] = (QAcep - QSol) * -1;
                                dts.Rows.Add(dRow);
                            }

                        }
                        l++;
                    }

                    if (((System.String)oDBDSDetalle.GetValue("U_Estado", i)).Trim() == "A")
                        Aceptados++;
                    i++;
                }

                if ((Aceptados) == oDBDSDetalle.Size)
                    oDBDSHeader.SetValue("U_Estado", 0, "C");
                else if (Aceptados == 0)
                    oDBDSHeader.SetValue("U_Estado", 0, "A");
                else
                    oDBDSHeader.SetValue("U_Estado", 0, "P");

                if (bCrearT)
                {
                    FCmpny.StartTransaction();
                    lRetCode = oStock.Add();
                    if (lRetCode != 0)
                    {
                        FCmpny.GetLastError(out errCode, out errMsg);
                        FSBOApp.StatusBar.SetText("No se ha creado Transferencia, " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        OutLog("No se ha creado Transferencia, " + errMsg);
                        if (FCmpny.InTransaction)
                            FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                        return false;
                    }
                    else
                    {//se cra transferencia por articulos recepcionados aprobados
                        var NewKey = FCmpny.GetNewObjectKey();
                        s = @"SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = {0}";
                        s = String.Format(s, NewKey);
                        oRecordSet.DoQuery(s);
                        NewKey = ((System.Int32)oRecordSet.Fields.Item("DocNum").Value).ToString().Trim();

                        if (bCrearTDif)//crea transferncia por recibir menos cantidad de lo solicitado
                        {
                            oStock = null;
                            oStock = ((SAPbobsCOM.StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oStockTransfer));
                            oStock.DocDate = DateTime.Now;
                            oStock.CardCode = CardCode;
                            oStock.Comments = "Cargado por Addon Recepción en Tienda, diferencia Solicitado y Recepcionado";
                            oStock.FromWarehouse = TiendaTransito;
                            oStock.ToWarehouse = BodegaDiferencia;
                            l = 0;
                            foreach (System.Data.DataRow orows in dts.Rows)
                            {
                                if (l > 0)
                                    oStock.Lines.Add();
                                oStock.Lines.ItemCode = ((System.String)orows["ItemCode"]).Trim();
                                oStock.Lines.Quantity = ((System.Double)orows["Quantity"]);
                                oStock.Lines.FromWarehouseCode = TiendaTransito;
                                oStock.Lines.WarehouseCode = BodegaDiferencia;
                                oStock.Lines.UserFields.Fields.Item("U_CR_Nro_Contenedor").Value = Convert.ToInt32(((System.String)oDBDSHeader.GetValue("U_Contenedor", 0)).Trim().Replace(",", "."), _nf);
                                oStock.Lines.BaseEntry = Convert.ToInt32(((System.String)orows["BaseEntry"]).Trim().Replace(",", "."), _nf);
                                oStock.Lines.BaseLine = Convert.ToInt32(((System.String)orows["BaseLine"]).Trim().Replace(",", "."), _nf);
                                oStock.Lines.BaseType = InvBaseDocTypeEnum.InventoryTransferRequest;
                                l++;
                            }

                            lRetCode = oStock.Add();
                            if (lRetCode != 0)
                            {
                                FCmpny.GetLastError(out errCode, out errMsg);
                                FSBOApp.StatusBar.SetText("No se ha creado Transferencia de diferencia, " + errMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                OutLog("No se ha creado Transferencia de diferencia, " + errMsg);
                                if (FCmpny.InTransaction)
                                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                                return false;
                            }
                            else
                            {
                                var NewKeyD = FCmpny.GetNewObjectKey();
                                s = @"SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = {0}";
                                s = String.Format(s, NewKeyD);
                                oRecordSet.DoQuery(s);
                                NewKeyD = ((System.Int32)oRecordSet.Fields.Item("DocNum").Value).ToString().Trim();
                                FSBOApp.StatusBar.SetText("Se ha creado Transferencia " + NewKey, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                FSBOApp.StatusBar.SetText("Se ha creado Transferencia " + NewKeyD + " por diferencia en solicitado y recepcionado", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            }
                        }
                        else
                            FSBOApp.StatusBar.SetText("Se ha creado Transferencia " + NewKey, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                }
                else
                {
                    FSBOApp.StatusBar.SetText("No ha realizado ningun cambio", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return false;
                }


                lRetCode = oUtil.UpdDataSourceInt1("VID_RECTDA", oDBDSHeader, "VID_RECTDAD", oDBDSDetalle, "", null, "", null);
                if (lRetCode == 0)
                {
                    FSBOApp.StatusBar.SetText("No se ha actualizado registro, revisar vd.log", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    if (FCmpny.InTransaction)
                        FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                    return false;
                }
                else
                {
                    oForm.Mode = BoFormMode.fm_OK_MODE;

                    if (FCmpny.InTransaction)
                        FCmpny.EndTransaction(BoWfTransOpt.wf_Commit);
                    oForm.Items.Item("Comentario").Click(BoCellClickType.ct_Regular);
                    Bloquear();
                    s = @"UPDATE ""@VID_RECTDAD"" SET ""U_EstadoAn"" = ""U_Estado"" WHERE ""DocEntry"" = {0}";
                    s = String.Format(s, ((System.String)oDBDSHeader.GetValue("DocEntry", 0)).Trim());
                    oRecordSet.DoQuery(s);
                    return true;
                }
            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("Error BuscarSKU: " + x.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(x.Message + " , **TRACE " + x.StackTrace);
                if (FCmpny.InTransaction)
                    FCmpny.EndTransaction(BoWfTransOpt.wf_RollBack);
                return false;
            }
            finally
            {
                oStock = null;
            }
        }

        private void Bloquear()
        {
            try
            {
                if (((System.String)oDBDSHeader.GetValue("U_Estado", 0)).Trim() == "C")
                {
                    ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyAcep").Editable = false;
                    ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Estado").Editable = false;
                    ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Editable = false;
                }
                else
                {
                    ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("QtyAcep").Editable = true;
                    ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Estado").Editable = true;
                    ((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Comentario").Editable = true;

                    for (Int32 numfila = 1; numfila <= ((Matrix)oForm.Items.Item("mtx").Specific).RowCount; numfila++)
                    {
                        s = ((ComboBox)((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Estado").Cells.Item(numfila).Specific).Value.ToString().Trim();
                        if (s == "A")
                        {
                            ((Matrix)oForm.Items.Item("mtx").Specific).CommonSetting.SetCellEditable(numfila, 7, false);
                            ((Matrix)oForm.Items.Item("mtx").Specific).CommonSetting.SetCellEditable(numfila, 8, false);
                        }
                        else
                        {
                            ((Matrix)oForm.Items.Item("mtx").Specific).CommonSetting.SetCellEditable(numfila, 7, true);
                            ((Matrix)oForm.Items.Item("mtx").Specific).CommonSetting.SetCellEditable(numfila, 8, true);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error Bloquear -> " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Error Bloquear -> " + e.Message + ", TRACE " + e.StackTrace);
            }
        }


        private void AprobarTodo()
        {
            oForm.Freeze(true);
            try
            {
                ((Matrix)oForm.Items.Item("mtx").Specific).FlushToDataSource();
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                for (Int32 numfila = 0; numfila < ((Matrix)oForm.Items.Item("mtx").Specific).RowCount; numfila++)
                {
                    s = ((ComboBox)((Matrix)oForm.Items.Item("mtx").Specific).Columns.Item("Estado").Cells.Item(numfila+1).Specific).Value.ToString().Trim();
                    if (s != "A")
                        oDBDSDetalle.SetValue("U_Estado", numfila, "A");
                }
                ((Matrix)oForm.Items.Item("mtx").Specific).LoadFromDataSource();
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("AprobarTodo: " + e.Message + " ** Trace: " + e.StackTrace);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }//fin class
}

