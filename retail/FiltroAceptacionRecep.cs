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

namespace VID_Retail.FiltroAceptacionRecep
{
    public class TFiltroAceptacionRecep : TvkBaseForm, IvkFormInterface
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
                FSBOf.LoadForm(xmlPath, "FiltroAceptacionRecep.srf", uid);
                EnableCrystal = false;

                oForm = FSBOApp.Forms.Item(uid);
                oForm.AutoManaged = true;
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", false);

                oForm.DataSources.UserDataSources.Add("cbxTienda", BoDataType.dt_SHORT_TEXT, 50);
                ((ComboBox)oForm.Items.Item("cbxTienda").Specific).DataBind.SetBound(true, "", "cbxTienda");

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
                       WHERE IFNULL(T0.""U_VR_SN"",'') <> ''
                        ORDER BY ""WhsName""";
                oRecordSet.DoQuery(s);
                FSBOf.FillCombo(((ComboBox)oForm.Items.Item("cbxTienda").Specific), ref oRecordSet, true);

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
                        var Lin = ValidarContenedoresSeleccionados();
                        if (Lin >= 0)
                            AbrirAprobacion(Lin);
                    }

                    if ((pVal.ItemUID == "btnBuscar") && (!pVal.BeforeAction))
                        CargarContenedores(((System.String)oForm.DataSources.UserDataSources.Item("cbxTienda").Value).Trim());
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


        public void CargarContenedores(String Tienda)
        {
            try
            {
                oForm.Freeze(true);
                s = @"SELECT T0.""U_Contenedor"" ""Contenedor""
                              ,IFNULL((SELECT B.""Descr""
                                         FROM ""CUFD"" A
                                         JOIN ""UFD1"" B ON B.""TableID"" = A.""TableID""
                                                      AND B.""FieldID"" = A.""FieldID""
                                        WHERE A.""TableID"" = '@VID_RECTDA'
                                          AND A.""AliasID"" = 'Estado'
                                          AND B.""FldValue"" = T0.""U_Estado""),'') ""Estado""
                              ,SUM(CASE WHEN T1.""U_Estado"" = 'R' THEN 1 ELSE 0 END) ""Cantidad""
                              ,T0.""DocEntry""
                          FROM ""@VID_RECTDA"" T0
                          JOIN ""@VID_RECTDAD"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
                         WHERE T0.""U_Tienda"" = '{0}'
                           AND T0.""U_Estado"" <> 'C'
                         GROUP BY T0.""U_Contenedor""
                              ,T0.""U_Estado""
                              ,T0.""DocEntry"" ";
                s = String.Format(s, Tienda);
                ((Grid)oForm.Items.Item("grid").Specific).DataTable.ExecuteQuery(s);

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Contenedor").TitleObject.Sortable = true;

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Estado").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Estado").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Estado").Visible = true;

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cantidad").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cantidad").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cantidad").Visible = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cantidad").RightJustified = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cantidad").TitleObject.Sortable = true;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("Cantidad").TitleObject.Caption = "Cantidad SKU por Aprobar";

                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocEntry").Type = BoGridColumnType.gct_EditText;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocEntry").Editable = false;
                ((Grid)oForm.Items.Item("grid").Specific).Columns.Item("DocEntry").Visible = false;
                ((Grid)oForm.Items.Item("grid").Specific).AutoResizeColumns();

            }
            catch (Exception e)
            {
                FSBOApp.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");
                OutLog("CargarContenedores: " + e.Message + " ** Trace: " + e.StackTrace);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void AbrirAprobacion(Int32 Linea)
        {
            String oUid;
            IvkFormInterface oFormVk = null;
            try
            {
                var DocNum = ((System.Int32)((Grid)oForm.Items.Item("grid").Specific).DataTable.GetValue("DocEntry", Linea));

                s = @"SELECT COUNT(*) ""Cant""
                          FROM ""@VID_RECTDAD""
                         WHERE ""DocEntry"" = {0}
                           AND IFNULL(""U_Estado"",'R') = 'R';";
                s = String.Format(s, DocNum);
                oRecordSet.DoQuery(s);
                if (((System.Int32)oRecordSet.Fields.Item("Cant").Value) == 0)
                {
                    FSBOApp.StatusBar.SetText("Contenedor tiene Aprobados todos sus SKU, se actualizara lista de contenedores", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    CargarContenedores(((System.String)oForm.DataSources.UserDataSources.Item("cbxTienda").Value).Trim());
                }
                else
                {
                    oFormVk = (IvkFormInterface)(new TAprobacionRecepcionTienda());
                    TAprobacionRecepcionTienda.DocNum = DocNum;
                    oUid = FSBOf.generateFormId(FGlobalSettings.SBOSpaceName, FGlobalSettings);
                    oFormVk.InitForm(oUid, "forms\\", ref FSBOApp, ref FCmpny, ref FSBOf, ref FGlobalSettings);
                    FoForms.Add(oFormVk);
                }
            }
            catch (Exception x)
            {
                FSBOApp.StatusBar.SetText("AbrirAprobacion: " + x.Message + " ** Trace: " + x.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("AbrirAprobacion: " + x.Message + " ** Trace: " + x.StackTrace);
            }
        }


        private Int32 ValidarContenedoresSeleccionados()
        {
            try
            {
                for (Int32 i = 0; i < ((Grid)oForm.Items.Item("grid").Specific).Rows.Count; i++)
                {
                    if (((Grid)oForm.Items.Item("grid").Specific).Rows.IsSelected(i))
                        return i;
                    else
                        continue;
                }
                FSBOApp.StatusBar.SetText("No ha seleccionado ningún Contenedor", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return -1;
            }
            catch (Exception x)
            {
                FSBOApp.MessageBox(x.Message + " ** Trace: " + x.StackTrace, 1, "Ok", "", "");
                OutLog("ValidarContenedoresSeleccionados: " + x.Message + " ** Trace: " + x.StackTrace);
                return -1;
            }
        }

    }//fin class
}
