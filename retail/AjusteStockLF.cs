using System;
using System.Text;
using System.Collections.Generic;
using SAPbouiCOM;
using SAPbobsCOM;
using VisualD.GlobalVid;
using VisualD.SBOFunctions;
using VisualD.vkBaseForm;
using VisualD.SBOGeneralService;
using VisualD.MasterDataMatrixForm;
using VisualD.vkFormInterface;
using VID_Retail.Periodos;


namespace VID_Retail.AjusteStockLF
{
    class TAjusteStockLF : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.Application R_application;
        SAPbobsCOM.Company R_company;
        CSBOFunctions R_sboFunctions;
        TGlobalVid R_GlobalSettings;
        string BodegaTienda = "";
        string BodegaTrastienda = "";
        string BodegaAjustes = "";
        string BodegaTiendaName = "";
        string BodegaTrastiendaName = "";
        string BodegaAjustesName = "";
        string oConsulta = "";

        public TAjusteStockLF()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        private SAPbouiCOM.DataTable oTable = null;
        private bool Primeravez;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;
            SAPbouiCOM.Matrix mtx0;

            R_application = application;
            R_company = company;
            R_sboFunctions = sboFunctions;
            R_GlobalSettings = _GlobalSettings;

            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                try
                {
                    FSBOf.LoadForm(xmlPath, "AjusteStockLF.srf", uid);
                    EnableCrystal = false;

                    oForm = FSBOApp.Forms.Item(uid);
                    oForm.AutoManaged = true;
                    oForm.SupportedModes = 1;             // afm_All
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oForm.PaneLevel = 1;

                    oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                    oForm.Items.Item("mtx0").AffectsFormMode = true;
                    mtx0 = (Matrix)(oForm.Items.Item("mtx0").Specific);

                    oForm.DataSources.UserDataSources.Add("DSTienda", BoDataType.dt_SHORT_TEXT, 50);
                    ((ComboBox)oForm.Items.Item("Tienda").Specific).DataBind.SetBound(true, "", "DSTienda");

                    oSql = GlobalSettings.RunningUnderSQLServer ?
                           "select Code, Name from [@VIDR_TIENDA] order by Name" :
                           "select \"Code\", \"Name\" from \"@VIDR_TIENDA\" order by \"Name\" ";
                    oRS.DoQuery(oSql);
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("Tienda").Specific)), ref oRS, false);

                    return (oResult);
                }
                catch (Exception e)
                {
                    FSBOApp.StatusBar.SetText(e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    OutLog(e.Message + " - " + e.StackTrace);
                    return (false);
                }
            }
            finally
            {
                if (oForm != null)
                    oForm.Visible = true;
            }
        }

        public new void FormEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, ref Boolean BubbleEvent)
        {
            base.FormEvent(FormUID, ref pVal, ref BubbleEvent);
            SAPbouiCOM.Form oForm = FSBOApp.Forms.Item(FormUID);
            string s;

            try
            {
                oForm.Freeze(true);
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "1") && (!pVal.BeforeAction) && ((pVal.FormMode == (Int32)BoFormMode.fm_ADD_MODE) || (pVal.FormMode == (Int32)BoFormMode.fm_UPDATE_MODE)))
                        {
                            //                            ValidarDatos();
                            //                            insertarDatos();
                        }
                        if ((pVal.ItemUID == "btnProc") && (!pVal.BeforeAction))
                        {
                            if (oConsulta == "")
                                return;

                            doTransfer();

                            ((Matrix)(oForm.Items.Item("mtx0").Specific)).Clear();
                            oConsulta = "";
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        {
                            if ((pVal.ItemUID == "Tienda") && (!pVal.BeforeAction))
                            {
                                s = oForm.DataSources.UserDataSources.Item("DSTienda").ValueEx.Trim();
                                ((Matrix)(oForm.Items.Item("mtx0").Specific)).Clear();
                                getBodegas(s);
                                oConsulta = FillMtx();
                            }
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(e.Message + " - " + e.StackTrace);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public new void MenuEvent(ref MenuEvent pVal, ref Boolean BubbleEvent)
        {
            //Int32 Entry;
            base.MenuEvent(ref pVal, ref BubbleEvent);
        }

        private void getBodegas(string oTda)
        {
            string oSql;

            BodegaAjustes = "";
            BodegaTienda = "";
            BodegaTrastienda = "";
            BodegaAjustesName = "";
            BodegaTiendaName = "";
            BodegaTrastiendaName = "";
            oSql = "Select \"WhsCode\", \"WhsName\", U_GSP_LOGISTICO     " +
                   "  from OWHS                             " +
                   " where U_GSP_LOGISTICO in ('B', 'L')    " +
                   "   and \"U_VID_Tda\" = '{0}'            ";
            oRS.DoQuery(string.Format(oSql, oTda));
            while (!oRS.EoF)
            {
                if (((String)oRS.Fields.Item("U_GSP_LOGISTICO").Value).Trim() == "B")
                {
                    BodegaTienda = ((String)oRS.Fields.Item("WhsCode").Value).Trim();
                    BodegaTiendaName = ((String)oRS.Fields.Item("WhsName").Value).Trim();
                }
                if (((String)oRS.Fields.Item("U_GSP_LOGISTICO").Value).Trim() == "L")
                {
                    BodegaTrastienda = ((String)oRS.Fields.Item("WhsCode").Value).Trim();
                    BodegaTrastiendaName = ((String)oRS.Fields.Item("WhsName").Value).Trim();
                }
                oRS.MoveNext();
            }

            oSql = "Select a.\"U_WhsCodLF\", b.\"WhsName\"    " +
                   "  from \"@VIDR_PARAM\"  a inner join OWHS b on a.\"U_WhsCodLF\" = b.\"WhsCode\"";
            oRS.DoQuery(oSql);
            if (!oRS.EoF)
            {
                BodegaAjustes = ((String)oRS.Fields.Item("U_WhsCodLF").Value).Trim();
                BodegaAjustesName = ((String)oRS.Fields.Item("WhsName").Value).Trim();
            }

        }

        private string FillMtx()
        {
            String oSql;
            SAPbouiCOM.Matrix mtx0;

            try
            {
                mtx0 = (Matrix)(oForm.Items.Item("mtx0").Specific);

                oForm.Freeze(true);
                oSql = "Select a.\"Tienda\", a.\"ItemCode\", i.\"ItemName\", a.\"Qty\", a.\"QtyOpen\",                                       " +
                       "        IfNull(bt.\"OnHand\", 0) \"EnTienda\", IfNull(bl.\"OnHand\",0) \"EnTrasTienda\", IfNull(bd.\"OnHand\",0) \"EnDiferencias\", " +
                       "        CASE WHEN a.\"Qty\" - IfNull(bt.\"OnHand\", 0) - IfNull(bl.\"OnHand\", 0) > 0                                " +
                       "             THEN IfNull(bl.\"OnHand\", 0)                                                                           " +
                       "             ELSE a.\"Qty\" - IfNull(bt.\"OnHand\", 0)                                                               " +
                       "             END \"Transf_Trastienda\",                                                                              " +
                       "        CASE WHEN a.\"Qty\" - IfNull(bt.\"OnHand\", 0) - IfNull(bl.\"OnHand\", 0) - IfNull(bd.\"OnHand\",0) > 0      " +
                       "             THEN IfNull(bd.\"OnHand\", 0)                                                                           " +
                       "             ELSE CASE WHEN a.\"Qty\" - IfNull(bt.\"OnHand\", 0) - IfNull(bl.\"OnHand\", 0) > 0                      " +
                       "                       THEN a.\"Qty\" - IfNull(bt.\"OnHand\", 0) - IfNull(bl.\"OnHand\", 0)                          " +
                       "                       ELSE 0 END                                                                                    " +
                       "             END \"Transf_Diferencias\"                                                                              " +
                       "   from                                                                                                              " +
                       "        ( Select b.\"Tienda\", b.\"ItemCode\", SUM(b.\"Qty\") \"Qty\", SUM(b.\"QtyOpen\") \"QtyOpen\"                " +
                       "            from                                                                                                     " +
                       "                 ( Select h.\"Code\", h.U_GSP_CABOTI \"Tienda\", h.U_GSP_CACAIX \"Caja\", h.U_GSP_CANUME \"Numero\", " +
                       "                          h.U_GSP_CADATA \"Fecha\", h.U_GSP_CAHORA \"Hora\", h.U_GSP_ERROR \"Error\",                " +
                       "                          h.U_GSP_CADOCU \"TipoDoc\",                                                                " +
                       "                          d.U_GSP_LIARTI \"ItemCode\", d.U_GSP_LIDES1 \"ItemName\", d.U_GSP_LIQUAN \"Qty\",          " +
                       "                          d.U_GSP_LIQUAN_OPEN \"QtyOpen\"                                                            " +
                       "                     from \"@GSP_TPVCAP\" h inner join \"@GSP_TPVLIN\" d on h.\"Code\" = d.\"U_GSP_DOCCODE\"         " +
                       "                    where h.U_GSP_ERROR like '-10 La cantidad recae en un inventario negativo%'                      " +
                       "                      and h.U_GSP_CABOTI = '{0}'                                                                     " +
                       "                    order by h.\"Code\"                                                                              " +
                       "                 ) b                                                                                                 " +
                       "           group by b.\"Tienda\", b.\"ItemCode\"                                                                     " +
                       "        ) a                                                                                                          " +
                       "            left outer join OITW bt on bt.\"WhsCode\" = a.\"Tienda\" and bt.\"ItemCode\" = a.\"ItemCode\"            " +
                       "            left outer join OITW bl on bl.\"WhsCode\" = '{1}'  and bl.\"ItemCode\" = a.\"ItemCode\"                  " +
                       "            left outer join OITW bd on bd.\"WhsCode\" = '{2}' and bd.\"ItemCode\" = a.\"ItemCode\"                   " +
                       "                 inner join OITM i  on  i.\"ItemCode\" = a.\"ItemCode\"                                              " +
                       "  Where a.\"Qty\" - bt.\"OnHand\" > 0                                                                                " +
                       "  Order by 1 , 3                                                                                                     ";
                oSql = string.Format(oSql, BodegaTienda, BodegaTrastienda, BodegaAjustes);

                if (oTable == null)
                {
                    Primeravez = true;
                    oForm.DataSources.DataTables.Add("dtData");
                }

                oTable = oForm.DataSources.DataTables.Item("dtData");
                oTable.Clear();
                oTable.ExecuteQuery(oSql);

                if (oTable.Rows.Count <= 0)
                {
                    FSBOApp.MessageBox("No existen registros a procesar !!!", 1, "Ok");
                    return "";
                }

                if (oTable.Rows.Count > 0)
                {
                    if (!Primeravez)
                    {
                        //mtx0.Columns.Item("Select").DataBind.UnBind();
                        mtx0.Columns.Item("Tienda").DataBind.UnBind();
                        mtx0.Columns.Item("ItemCode").DataBind.UnBind();
                        mtx0.Columns.Item("ItemName").DataBind.UnBind();
                        mtx0.Columns.Item("Quantity").DataBind.UnBind();
                        mtx0.Columns.Item("OnHandT").DataBind.UnBind();
                        mtx0.Columns.Item("OnHandB").DataBind.UnBind();
                        mtx0.Columns.Item("OnHandD").DataBind.UnBind();
                        mtx0.Columns.Item("TransfTT").DataBind.UnBind();
                        mtx0.Columns.Item("TransfBA").DataBind.UnBind();
                    }

                    //mtx0.Columns.Item("Select").DataBind.Bind("dtData", "Select");
                    mtx0.Columns.Item("Tienda").DataBind.Bind("dtData", "Tienda");
                    mtx0.Columns.Item("ItemCode").DataBind.Bind("dtData", "ItemCode");
                    mtx0.Columns.Item("ItemName").DataBind.Bind("dtData", "ItemName");
                    mtx0.Columns.Item("Quantity").DataBind.Bind("dtData", "QtyOpen");
                    mtx0.Columns.Item("OnHandT").DataBind.Bind("dtData", "EnTienda");
                    mtx0.Columns.Item("OnHandB").DataBind.Bind("dtData", "EnTrasTienda");
                    mtx0.Columns.Item("OnHandD").DataBind.Bind("dtData", "EnDiferencias");
                    mtx0.Columns.Item("TransfTT").DataBind.Bind("dtData", "Transf_Trastienda");
                    mtx0.Columns.Item("TransfBA").DataBind.Bind("dtData", "Transf_Diferencias");


                    oForm.Mode = BoFormMode.fm_OK_MODE;
                }

                mtx0.LoadFromDataSource();

                return oSql;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void doTransfer()
        {
            bool bTrastienda = false;
            bool bBodegaDiff = false;
            SAPbobsCOM.StockTransfer Trastienda = (StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oStockTransfer);
            SAPbobsCOM.StockTransfer BodegaDiff = (StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oStockTransfer);
            int linTrastienda = -1;
            int linBodegaDiff = -1;
            int nErr;
            string sErr;

            if (oConsulta == "")
            {
                FSBOApp.MessageBox("No se ha seleccionado una tienda.", 1, "Ok");
                return;
            }

            oRS.DoQuery(oConsulta);
            while (!oRS.EoF)
            {
                if ((double)oRS.Fields.Item("Transf_Trastienda").Value > 0)
                {
                    if (!bTrastienda)
                    {
                        Trastienda.FromWarehouse = BodegaTrastienda;
                        Trastienda.ToWarehouse = BodegaTienda;
                        Trastienda.DocDate = DateTime.Now;
                        Trastienda.UserFields.Fields.Item("U_VK_Almacen_Origen").Value = BodegaTrastiendaName;
                        Trastienda.UserFields.Fields.Item("U_VK_AlmacenDestino").Value = BodegaTiendaName;
                    }
                    bTrastienda = true;

                    linTrastienda++;
                    if (linTrastienda > 0)
                        Trastienda.Lines.Add();

                    Trastienda.Lines.SetCurrentLine(linTrastienda);
                    Trastienda.Lines.ItemCode = ((string)oRS.Fields.Item("ItemCode").Value).Trim();
                    Trastienda.Lines.Quantity = ((double)oRS.Fields.Item("Transf_Trastienda").Value);
                    Trastienda.Lines.WarehouseCode = BodegaTienda;
                }
                if ((double)oRS.Fields.Item("Transf_Diferencias").Value > 0)
                {
                    if (!bBodegaDiff)
                    {
                        BodegaDiff.FromWarehouse = BodegaAjustes;
                        BodegaDiff.ToWarehouse = BodegaTienda;
                        BodegaDiff.DocDate = DateTime.Now;
                        BodegaDiff.UserFields.Fields.Item("U_VK_Almacen_Origen").Value = BodegaAjustesName;
                        BodegaDiff.UserFields.Fields.Item("U_VK_AlmacenDestino").Value = BodegaTiendaName;
                    }
                    bBodegaDiff = true;

                    linBodegaDiff++;
                    if (linBodegaDiff > 0)
                        BodegaDiff.Lines.Add();

                    BodegaDiff.Lines.SetCurrentLine(linBodegaDiff);
                    BodegaDiff.Lines.ItemCode = ((string)oRS.Fields.Item("ItemCode").Value).Trim();
                    BodegaDiff.Lines.Quantity = ((double)oRS.Fields.Item("Transf_Diferencias").Value);
                    BodegaDiff.Lines.WarehouseCode = BodegaTienda;
                }
                oRS.MoveNext();
            }

            if (bTrastienda)
                if (Trastienda.Add() != 0)
                {
                    FCmpny.GetLastError(out nErr, out sErr);
                    throw new Exception("Error en transferencia desde trastienda: " + nErr.ToString() + " - " + sErr);
                }

            if (bBodegaDiff)
                if (BodegaDiff.Add() != 0)
                {
                    FCmpny.GetLastError(out nErr, out sErr);
                    throw new Exception("Error en transferencia desde Bodega ajustes: " + nErr.ToString() + " - " + sErr);
                }

            if ((bTrastienda) || (bBodegaDiff))
                FSBOApp.StatusBar.SetText("Transferencias realizadas.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
        }
    }
}
