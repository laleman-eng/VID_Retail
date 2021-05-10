using System;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Threading;
using System.IO;
using SAPbouiCOM;
using SAPbobsCOM;
using VisualD.GlobalVid;
using VisualD.SBOFunctions;
using VisualD.vkBaseForm;
using VisualD.SBOGeneralService;
using VisualD.MasterDataMatrixForm;
using VisualD.vkFormInterface;
using VID_Retail.Periodos;


namespace VID_Retail.Precios
{
    class TPrecios : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.Application R_application;
        SAPbobsCOM.Company R_company;
        CSBOFunctions R_sboFunctions;
        TGlobalVid R_GlobalSettings;
        List<object> R_oForms;
        List<Int32> oListasPrecio = new List<int>();
        Boolean MargenBasePrecioVenta = false;
        Int32 BaseListPrice = -10;

        public TPrecios(List<object> oForms)
        {
            R_oForms = oForms;
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        private SAPbouiCOM.DataTable oTable = null;
        private SAPbouiCOM.DataTable oTable2 = null;
        private Int32 mtx3Height;
        private Int32 Rectang1Height;
        private Int32 Rectang2Height;
        private Int32 oRowSelect = -1;
        private String oFechaini, oFechafin;
        private string CodeCategoria = "";
        private string CodeGrupo = "";
        private string CodeFamilia = "";
        private string ItemCodesFilter = "";
        private bool useItemCodeFilter = false;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;

            R_application = application;
            R_company = company;
            R_sboFunctions = sboFunctions;
            R_GlobalSettings = _GlobalSettings;

            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                try
                {
                    FSBOf.LoadForm(xmlPath, "Precios.srf", uid);
                    EnableCrystal = false;

                    oForm = FSBOApp.Forms.Item(uid);
                    oForm.AutoManaged = true;
                    oForm.SupportedModes = 1;             // ok 
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oForm.PaneLevel = 1;
                    oForm.Items.Item("tab1").Click();

                    mtx3Height = oForm.Items.Item("mtx3").Height;
                    Rectang1Height = oForm.Items.Item("Rectang1").Height;
                    Rectang2Height = oForm.Items.Item("Rectang2").Height;

                    oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                    oForm.DataSources.UserDataSources.Add("DSDpto", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSCategor", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSGrupo", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSFamilia", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSMarca", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSItemCode", BoDataType.dt_SHORT_TEXT, 50);
                    ((ComboBox)(oForm.Items.Item("cb_dpto").Specific)).DataBind.SetBound(true, "", "DSDpto");
                    ((EditText)(oForm.Items.Item("Grupo").Specific)).DataBind.SetBound(true, "", "DSGrupo");
                    ((EditText)(oForm.Items.Item("Categoria").Specific)).DataBind.SetBound(true, "", "DSCategor");
                    ((EditText)(oForm.Items.Item("Familia").Specific)).DataBind.SetBound(true, "", "DSFamilia");
                    ((EditText)(oForm.Items.Item("Marca").Specific)).DataBind.SetBound(true, "", "DSMarca");
                    ((EditText)(oForm.Items.Item("ItemCode").Specific)).DataBind.SetBound(true, "", "DSItemCode");

                    //Margen
                    oSql = GlobalSettings.RunningUnderSQLServer ?
                            "Select U_MgnPVta, IsNull(U_PriceLst, -10) U_PriceLst from [@VIDR_PARAM] " :
                            "Select \"U_MgnPVta\", IfNull(\"U_PriceLst\", -10) \"U_PriceLst\" from \"@VIDR_PARAM\" ";
                    oRS.DoQuery(oSql);
                    if (!oRS.EoF)
                    {
                        if ((String)oRS.Fields.Item("U_MgnPVta").Value == "Y")
                            MargenBasePrecioVenta = true;
                        BaseListPrice = (Int32)oRS.Fields.Item("U_PriceLst").Value;
                    }

                    setMtx1();
                    setMtx2();
                    setMtx3();
                    SetupPeriodos();
                    FillMtx3();

                    ////Filtros
                    //Departamento
                    oSql = GlobalSettings.RunningUnderSQLServer ?
                            "Select Code, Name from [@VK_DPTO] order by Name" :
                            "Select \"Code\", \"Name\" from \"@VIDR_DPTO\" order by \"Name\" ";
                    oRS.DoQuery(oSql);
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("cb_dpto").Specific)), ref oRS, true);

                    AddChooseFromList();
                    //Categoria
                    ((EditText)(oForm.Items.Item("Categoria").Specific)).ChooseFromListUID = "CFLCategoria";
                    ((EditText)(oForm.Items.Item("Categoria").Specific)).ChooseFromListAlias = "Name";

                    //Grupo
                    ((EditText)(oForm.Items.Item("Grupo").Specific)).ChooseFromListUID = "CFLGrupo";
                    ((EditText)(oForm.Items.Item("Grupo").Specific)).ChooseFromListAlias = "Name";

                    //Familia
                    ((EditText)(oForm.Items.Item("Familia").Specific)).ChooseFromListUID = "CFLFamilia";
                    ((EditText)(oForm.Items.Item("Familia").Specific)).ChooseFromListAlias = "Name";

                    oForm.Mode = BoFormMode.fm_OK_MODE;

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

        private void AddChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            SAPbouiCOM.Conditions oConds;
            SAPbouiCOM.Condition oCond;

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)(FSBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams));

            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "VIDR_CATEGORIA";
            oCFLCreationParams.UniqueID = "CFLCategoria";
            oCFL = oCFLs.Add(oCFLCreationParams);
            oConds = oCFL.GetConditions();
            oCond = oConds.Add();
            oCond.Alias = "U_Depto";
            oCond.Operation = BoConditionOperation.co_EQUAL;
            oCond.CondVal = "";
            oCFL.SetConditions(oConds);

            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "VIDR_GRUPO";
            oCFLCreationParams.UniqueID = "CFLGrupo";
            oCFL = oCFLs.Add(oCFLCreationParams);
            oConds = oCFL.GetConditions();
            oCond = oConds.Add();
            oCond.Alias = "U_Categoria";
            oCond.Operation = BoConditionOperation.co_EQUAL;
            oCond.CondVal = "";
            oCFL.SetConditions(oConds);

            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "VIDR_FAMILIA";
            oCFLCreationParams.UniqueID = "CFLFamilia";
            oCFL = oCFLs.Add(oCFLCreationParams);
            oConds = oCFL.GetConditions();
            oCond = oConds.Add();
            oCond.Alias = "U_Grupo";
            oCond.Operation = BoConditionOperation.co_EQUAL;
            oCond.CondVal = "";
            oCFL.SetConditions(oConds);
        }

        private void AddChooseFromListDinamico(String sCFL, String oParam)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.Conditions oCons;
            SAPbouiCOM.Condition oCon;
            Int32 i;

            oCFLs = oForm.ChooseFromLists;

            oCFL = oCFLs.Item(sCFL);
            i = 0;
            oCons = oCFL.GetConditions();
            oCon = oCons.Item(i);
            oCon.CondVal = oParam;
            oCFL.SetConditions(oCons);
        }

        public new void FormEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, ref Boolean BubbleEvent)
        {
            base.FormEvent(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.DataTable oDataTable;
            SAPbouiCOM.Matrix mtx1;
            String sValue;
            String sAux;
            String sVal;
            Double dVal;
            Double dTab;
            Double dTax;
            Double UltCompra;
            Double CostProm;
            Int32 offS;
            Boolean isGrossPrice;

            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "tab1") && (!pVal.BeforeAction))
                            oForm.PaneLevel = 1;
                        else if ((pVal.ItemUID == "tab2") && (!pVal.BeforeAction))
                            oForm.PaneLevel = 2;
                        else if ((pVal.ItemUID == "mtx3") && (pVal.BeforeAction))
                        {
                            if ((pVal.Row == 0) || (pVal.ColUID != "V_-1"))
                                BubbleEvent = false;
                            else if ((oRowSelect != -1) && (oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                                BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea cambiar periodo seleccionado? - perdera las modificaciones realizadas", 2, "Ok", "Cancel"));
                        }
                        else if ((pVal.ItemUID == "btnFltPrd") && (pVal.BeforeAction) && (oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            if (oRowSelect != -1)
                                BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea cambiar periodo seleccionado? - perdera las modificaciones realizadas", 2, "Ok", "Cancel"));
                        }
                        else if ((pVal.ItemUID == "btnFltPrd") && (!pVal.BeforeAction))
                        {
                            useItemCodeFilter = false;
                            if (getDetallePeriods())
                            {
                                getDetalleProducts();
                                oForm.Mode = BoFormMode.fm_OK_MODE;
                            }
                        }
                        else if ((pVal.ItemUID == "btnPer") && (!pVal.BeforeAction))
                            Periodos();
                        else if ((pVal.ItemUID == "btnLoad") && (!pVal.BeforeAction))
                        {
                            useItemCodeFilter = true;
                            if (LoadProducts())
                                if (getDetallePeriods())
                                {
                                    getDetalleProducts();
                                    oForm.Mode = BoFormMode.fm_OK_MODE;
                                }
                        }
                        else if ((pVal.ItemUID == "1") && (pVal.BeforeAction) && (pVal.FormMode == (Int32)BoFormMode.fm_UPDATE_MODE))
                            Updatemtx1();
                        break;
                    case BoEventTypes.et_FORM_RESIZE:
                        if (!pVal.BeforeAction)
                        {
                            oForm.Items.Item("mtx3").Height = mtx3Height;
                            oForm.Items.Item("Rectang1").Height = Rectang1Height;
                            oForm.Items.Item("Rectang2").Height = Rectang2Height;
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        if ((pVal.ItemUID == "cb_dpto") && (!pVal.BeforeAction))
                        {
                            oForm.Freeze(true);
                            oForm.DataSources.UserDataSources.Item("DSCategor").ValueEx = "";
                            oForm.DataSources.UserDataSources.Item("DSGrupo").ValueEx = "";
                            oForm.DataSources.UserDataSources.Item("DSFamilia").ValueEx = "";
                            CodeCategoria = "";
                            CodeGrupo = "";
                            CodeFamilia = "";
                            oForm.Freeze(false);
                        }
                        break;
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.BeforeAction)
                        {
                            oForm.Freeze(true);
                            if (pVal.ItemUID == "Categoria")
                            {
                                AddChooseFromListDinamico("CFLCategoria", oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim());
                            }
                            else if (pVal.ItemUID == "Grupo")
                            {
                                AddChooseFromListDinamico("CFLGrupo", CodeCategoria.Trim());
                            }
                            else if (pVal.ItemUID == "Familia")
                            {
                                AddChooseFromListDinamico("CFLFamilia", CodeGrupo.Trim());
                            }
                            oForm.Freeze(false);
                        }
                        else if (!pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "Categoria")
                            {
                                oDataTable = ((SAPbouiCOM.ChooseFromListEvent)(pVal)).SelectedObjects;
                                try
                                {
                                    oForm.Freeze(true);
                                    sValue = (String)(oDataTable.GetValue("Name", 0));
                                    CodeCategoria = (String)(oDataTable.GetValue("Code", 0));
                                    oForm.DataSources.UserDataSources.Item("DSCategor").ValueEx = sValue;
                                    oForm.DataSources.UserDataSources.Item("DSGrupo").ValueEx = "";
                                    oForm.DataSources.UserDataSources.Item("DSFamilia").ValueEx = "";
                                    CodeGrupo = "";
                                    CodeFamilia = "";
                                }
                                catch { }
                                finally
                                {
                                    oForm.Freeze(false);
                                }
                            }
                            else if (pVal.ItemUID == "Grupo")
                            {
                                oDataTable = ((SAPbouiCOM.ChooseFromListEvent)(pVal)).SelectedObjects;
                                try
                                {
                                    oForm.Freeze(true);
                                    sValue = (String)(oDataTable.GetValue("Name", 0));
                                    CodeGrupo = (String)(oDataTable.GetValue("Code", 0));
                                    oForm.DataSources.UserDataSources.Item("DSGrupo").ValueEx = sValue;
                                    oForm.DataSources.UserDataSources.Item("DSFamilia").ValueEx = "";
                                    CodeFamilia = "";
                                }
                                catch { }
                                finally
                                {
                                    oForm.Freeze(false);
                                }
                            }
                            else if (pVal.ItemUID == "Familia")
                            {
                                oDataTable = ((SAPbouiCOM.ChooseFromListEvent)(pVal)).SelectedObjects;
                                try
                                {
                                    oForm.Freeze(true);
                                    sValue = (String)(oDataTable.GetValue("Name", 0));
                                    CodeFamilia = (String)(oDataTable.GetValue("Code", 0));
                                    oForm.DataSources.UserDataSources.Item("DSFamilia").ValueEx = sValue;
                                }
                                catch { }
                                finally
                                {
                                    oForm.Freeze(false);
                                }
                            }
                        }
                        break;
                    case BoEventTypes.et_VALIDATE:
                        if ((pVal.ItemUID == "mtx1") && (pVal.ColUID.Substring(0, 1) == "P") && (pVal.BeforeAction))
                        {
                            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
                            sAux = pVal.ColUID.Substring(1, pVal.ColUID.Length - 1);
                            offS = Int32.Parse(((EditText)(mtx1.Columns.Item("RowId").Cells.Item(pVal.Row).Specific)).Value) - 1;

                            dVal = ((Double)oTable.GetValue("SPrice" + sAux, offS));
                            sVal = ((String)oTable.GetValue("Tipo" + sAux, offS));
                            if (sVal == "OP") //Ya tiene precio especial y no puede modificarse
                            {
                                oTable.Rows.Offset = offS;
                                mtx1.SetLineData(pVal.Row);
                                FSBOApp.StatusBar.SetText("Precio asignado o superpuesto para otro periodo. No se puede modificar.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                        }
                        if ((pVal.ItemUID == "mtx1") && (pVal.ColUID.Substring(0, 1) == "P") && (!pVal.BeforeAction))
                        {
                            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
                            sAux = pVal.ColUID.Substring(1, pVal.ColUID.Length - 1);

                            offS = Int32.Parse(((EditText)(mtx1.Columns.Item("RowId").Cells.Item(pVal.Row).Specific)).Value) - 1;

                            sVal = ((EditText)(mtx1.Columns.Item("P" + sAux).Cells.Item(pVal.Row).Specific)).Value;
                            dVal = Double.Parse(sVal, System.Globalization.CultureInfo.InvariantCulture);

                            dTab = ((Double)oTable.GetValue("CPrice" + sAux, offS));
                            UltCompra = ((Double)oTable.GetValue("LastPurPrc", offS));
                            CostProm = ((Double)oTable.GetValue("LstEvlPric", offS));
                            dTax = ((Double)oTable.GetValue("Rate", offS));
                            isGrossPrice = (((String)oTable.GetValue("GrossPrice" + sAux, offS)) == "Y");

                            if (dTab != dVal)
                            {
                                oTable.SetValue("CPrice" + sAux, offS, dVal);
                                oTable.SetValue("Tipo" + sAux, offS, "M");
                                oTable.SetValue("Modif", offS, "Y");
                                oTable.SetValue("MargenCompra" + sAux, offS, Margen(dVal, UltCompra, dTax, isGrossPrice));
                                oTable.SetValue("MargenCosto" + sAux, offS, Margen(dVal, CostProm, dTax, isGrossPrice));
                                oTable.Rows.Offset = offS;
                                mtx1.SetLineData(pVal.Row);
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
        }

        public new void MenuEvent(ref MenuEvent pVal, ref Boolean BubbleEvent)
        {
            //Int32 Entry;
            base.MenuEvent(ref pVal, ref BubbleEvent);
        }

        private Double Margen(Double pVenta, Double pCosto, Double TaxRate, Boolean GrossPrice)
        {
            if (MargenBasePrecioVenta)
            {
                if (GrossPrice)
                    return pVenta != 0.0 ? 100 * (pVenta / (1 + TaxRate) - pCosto) / (pVenta / (1 + TaxRate)) : 0.0;
                else
                    return pVenta != 0.0 ? 100 * (pVenta - pCosto) / pVenta : 0.0;
            }
            else
            {
                if (GrossPrice)
                    return pCosto != 0.0 ? 100 * (pVenta / (1 + TaxRate) - pCosto) / pCosto : 0.0;
                else
                    return pCosto != 0.0 ? 100 * (pVenta - pCosto) / pCosto : 0.0;
            }
        }

        private void setMtx1()
        {
            SAPbouiCOM.Matrix mtx1;
            String oSql;
            Int32 i;

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);

            mtx1.CommonSetting.FixedColumnsCount = 3;

            mtx1.Columns.Item("ItemCode").Editable = false;
            mtx1.Columns.Item("ItemName").Editable = false;
            mtx1.Columns.Item("UltCompra").Editable = false;
            mtx1.Columns.Item("CostProm").Editable = false;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select 0 Ord, ListNum, ListName from OPLN Where Lisnum = {0}  " + 
                   " Union All " +
                   "Select 1 Ord, ListNum, ListName from OPLN Where Lisnum <> {0} order by 1, 3 " :
                   "Select 0 \"Ord\", \"ListNum\", \"ListName\" from OPLN Where \"ListNum\" = {0} " +
                   " Union All " +
                   "Select 1 \"Ord\", \"ListNum\", \"ListName\" from OPLN Where \"ListNum\" != {0} order by 1, 3 ";
            oSql = String.Format(oSql, BaseListPrice.ToString());
            oRS.DoQuery(oSql);

            while (!oRS.EoF)
            {
                i = (Int32)(oRS.Fields.Item("ListNum").Value);

                mtx1.Columns.Add("P" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                mtx1.Columns.Item("P" + i.ToString()).Width = 70;
                mtx1.Columns.Item("P" + i.ToString()).RightJustified = true;
                mtx1.Columns.Item("P" + i.ToString()).TitleObject.Caption = (String)(oRS.Fields.Item("ListName").Value);

                mtx1.Columns.Add("MP" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                mtx1.Columns.Item("MP" + i.ToString()).Width = 70;
                mtx1.Columns.Item("MP" + i.ToString()).TitleObject.Caption = "Margen Compra";
                mtx1.Columns.Item("MP" + i.ToString()).RightJustified = true;
                mtx1.Columns.Item("MP" + i.ToString()).Editable = false;
                //mtx1.Columns.Item("MP" + i.ToString()).

                mtx1.Columns.Add("MC" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                mtx1.Columns.Item("MC" + i.ToString()).Width = 70;
                mtx1.Columns.Item("MC" + i.ToString()).TitleObject.Caption = "Margen Costo";
                mtx1.Columns.Item("MC" + i.ToString()).RightJustified = true;
                mtx1.Columns.Item("MC" + i.ToString()).Editable = false;

                mtx1.Columns.Add("T" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                mtx1.Columns.Item("T" + i.ToString()).Width = 30;
                mtx1.Columns.Item("T" + i.ToString()).TitleObject.Caption = "Tipo";
                mtx1.Columns.Item("T" + i.ToString()).Editable = false;

                oRS.MoveNext();
            }
        }

        private void setMtx2()
        {
            SAPbouiCOM.Matrix mtx2;
            String oSql;
            Int32 i;

            mtx2 = (Matrix)(oForm.Items.Item("mtx2").Specific);

            mtx2.Columns.Item("ItemCode").Editable = false;
            mtx2.Columns.Item("ItemName").Editable = false;
            mtx2.Columns.Item("Dpto").Editable = false;
            mtx2.Columns.Item("Cat").Editable = false;
            mtx2.Columns.Item("Grupo").Editable = false;
            mtx2.Columns.Item("Familia").Editable = false;
            mtx2.Columns.Item("Fecini").Editable = false;
            mtx2.Columns.Item("Fecfin").Editable = false;
            mtx2.Columns.Item("Periodo").Editable = false;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select 0 Ord, ListNum, ListName from OPLN Where Lisnum = {0}  " +
                   " Union All " +
                   "Select 1 Ord, ListNum, ListName from OPLN Where Lisnum <> {0} order by 1, 3 " :
                   "Select 0 \"Ord\", \"ListNum\", \"ListName\" from OPLN Where \"ListNum\" = {0} " +
                   " Union All " +
                   "Select 1 \"Ord\", \"ListNum\", \"ListName\" from OPLN Where \"ListNum\" != {0} order by 1, 3 ";
            oSql = String.Format(oSql, BaseListPrice.ToString());
            oRS.DoQuery(oSql);

            while (!oRS.EoF)
            {
                i = (Int32)(oRS.Fields.Item("ListNum").Value);
                mtx2.Columns.Add("P" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                mtx2.Columns.Item("P" + i.ToString()).Width = 70;
                mtx2.Columns.Item("P" + i.ToString()).RightJustified = true;
                mtx2.Columns.Item("P" + i.ToString()).TitleObject.Caption = (String)(oRS.Fields.Item("ListName").Value);
                mtx2.Columns.Item("P" + i.ToString()).Editable = false;

                //oForm.DataSources.UserDataSources.Add("DSLN" + i.ToString(), BoDataType.dt_SHORT_TEXT, 100);
                //oForm.DataSources.UserDataSources.Item("DSLN" + i.ToString()).ValueEx = (String)(oRS.Fields.Item("ListName").Value);
                //mtx2.Columns.Add("LN" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //mtx2.Columns.Item("LN" + i.ToString()).DataBind.SetBound(true, "", "DSLN" + i.ToString());
                //mtx2.Columns.Item("LN" + i.ToString()).Width = 70;
                //mtx2.Columns.Item("LN" + i.ToString()).Editable = false;

                oRS.MoveNext();
            }
        }

        private void setMtx3()
        {
            SAPbouiCOM.Matrix mtx3;

            mtx3 = (Matrix)(oForm.Items.Item("mtx3").Specific);
            oForm.DataSources.UserDataSources.Add("DS3Periodo", BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("DS3Fecini", BoDataType.dt_DATE, 11);
            oForm.DataSources.UserDataSources.Add("DS3Fecfin", BoDataType.dt_DATE, 11);
            oForm.DataSources.UserDataSources.Add("DS3QtyDias", BoDataType.dt_LONG_NUMBER, 10);
            oForm.DataSources.UserDataSources.Add("DS1QtyWEnd", BoDataType.dt_LONG_NUMBER, 10);
            mtx3.Columns.Item("Code").DataBind.SetBound(true, "", "DS3Periodo");
            mtx3.Columns.Item("Fecini").DataBind.SetBound(true, "", "DS3Fecini");
            mtx3.Columns.Item("Fecfin").DataBind.SetBound(true, "", "DS3Fecfin");
            mtx3.Columns.Item("QtyDias").DataBind.SetBound(true, "", "DS3QtyDias");
            mtx3.Columns.Item("QtyWEnd").DataBind.SetBound(true, "", "DS1QtyWEnd");
        }

        private void SetupPeriodos()
        {
            String oSql;

            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.CompanyService oCompService;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select Count(*) Cant from [@VIDR_PERIODO]" :
                   "Select Count(*) \"Cant\" from \"@VIDR_PERIODO\"";
            oRS.DoQuery(oSql);

            if ((Int32)oRS.Fields.Item("Cant").Value == 0)
            {
                oCompService = FCmpny.GetCompanyService();
                oGeneralService = oCompService.GetGeneralService("VIDR_PERIODO");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                oGeneralData.SetProperty("Code", "Normal");
                oGeneralData.SetProperty("Name", "Normal");
                DateTime d = new DateTime(2000, 1, 1);
                oGeneralData.SetProperty("U_Fecini", d);
                d = new DateTime(2099, 12, 31);
                oGeneralData.SetProperty("U_Fecfin", d);
                oGeneralService.Add(oGeneralData);
            }
        }

        private void FillMtx3()
        {
            String oSql;

            oForm.Freeze(true);
            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select Code, U_Fecini, U_Fecfin,            " +
                   "       DATEDIFF(d, U_Fecini, U_Fecfin)/7+1 + CASE WHEN DATEPART(dw,U_Fecini) IN (1,7) THEN -1 ELSE 0 END + CASE WHEN DATEPART(dw, U_Fecfin) IN (1,7) THEN -1 ELSE 0 END WeekEnd, " +
                   "       DATEDIFF(d, U_Fecini, U_Fecfin) Dias " +
                   "  from [@VIDR_PERIODO] where U_Fecfin > GETDATE()-1 order by 2 " :
                   "Select \"Code\", \"U_Fecini\", \"U_Fecfin\",  " +
                   "       TO_Integer(DAYS_BETWEEN( \"U_Fecini\", \"U_Fecfin\")/7)+1 + CASE WHEN WEEKDAY(\"U_Fecini\") IN (5,6) THEN -1 ELSE 0 END + CASE WHEN WEEKDAY(\"U_Fecfin\") IN (5,6) THEN -1 ELSE 0 END \"WeekEnd\", " +
                   "       DAYS_BETWEEN( \"U_Fecini\", \"U_Fecfin\") Dias " +
                   "  from \"@VIDR_PERIODO\" where \"U_Fecfin\" > ADD_DAYS(CURRENT_DATE, -1) order by 2 ";
            oRS.DoQuery(oSql);

            while (!oRS.EoF)
            {
                oForm.DataSources.UserDataSources.Item("DS3Periodo").ValueEx = (String)oRS.Fields.Item("Code").Value;
                oForm.DataSources.UserDataSources.Item("DS3Fecini").ValueEx = ((DateTime)(oRS.Fields.Item("U_Fecini").Value)).ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("DS3Fecfin").ValueEx = ((DateTime)(oRS.Fields.Item("U_Fecfin").Value)).ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("DS3QtyDias").ValueEx = ((Int32)(oRS.Fields.Item("Dias").Value)).ToString();
                oForm.DataSources.UserDataSources.Item("DS1QtyWEnd").ValueEx = ((Int32)(oRS.Fields.Item("WeekEnd").Value)).ToString();
                ((Matrix)(oForm.Items.Item("mtx3").Specific)).AddRow(1, -1);
                oRS.MoveNext();
            }
            oForm.Freeze(false);
        }

        private void Periodos()
        {
            try
            {
                IvkFormInterface oForm;

                oForm = (IvkFormInterface)(new TPeriodos());

                if (oForm != null)
                {

                    if (oForm.InitForm(R_sboFunctions.generateFormId(GlobalSettings.SBOSpaceName, GlobalSettings), "forms\\", ref  R_application, ref  R_company, ref R_sboFunctions, ref R_GlobalSettings))
                    { R_oForms.Add(oForm); }
                    else
                    {
                        R_application.Forms.Item(oForm.getFormId()).Close();
                        oForm = null;
                    }
                }

            }
            catch (Exception e)
            {
                R_application.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok", "", "");  // Captura errores no manejados
                OutLog("MenuEventExt: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }

        private bool getDetallePeriods()
        {
            String oSql, oSql0, oWhere ;
            String aux0, aux1, aux2, aux3;
            String oCond1, oCond2, oCond3, oCond4, oCond5, oCond6;
            Int32 i;
            SAPbouiCOM.Matrix mtx1;
            SAPbouiCOM.Matrix mtx3; //Periodos
            Boolean Primeravez = false;
            Boolean primerRegistro = true;
            Boolean precioBase = false;
            Boolean GrossPrice = false;
            String  NoEsPreciosBase = "(1 = 1)";

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            mtx3 = (Matrix)(oForm.Items.Item("mtx3").Specific);

            oCond1 = oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim() : "%";
            oCond2 = CodeCategoria.Trim() != "" ? CodeCategoria.Trim() : "%";
            oCond3 = CodeGrupo.Trim() != "" ? CodeGrupo.Trim() : "%";
            oCond4 = CodeFamilia.Trim() != "" ? CodeFamilia.Trim() : "%";
            oCond5 = oForm.DataSources.UserDataSources.Item("DSMarca").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSMarca").ValueEx.Trim() : "%";
            oCond6 = oForm.DataSources.UserDataSources.Item("DSItemCode").ValueEx.Trim() != "" ? "%"+oForm.DataSources.UserDataSources.Item("DSItemCode").ValueEx.Trim()+"%" : "%";

            if ((oCond1 == "%") && (oCond2 == "%") && (oCond3 == "%") && (oCond4 == "%") && (oCond5 == "%") && (oCond6 == "%") && (!useItemCodeFilter))
            {
                FSBOApp.MessageBox("Debe seleccionar al menos un filtro !!!", 1, "Ok");
                return false;
            }

            oRowSelect = -1;
            oFechafin = "";
            oFechaini = "";
            oListasPrecio.Clear();
            for (i = 1; i <= mtx3.RowCount; i++)
                if (mtx3.IsRowSelected(i))
                {
                    oRowSelect = i;
                    mtx3.GetLineData(i);
                    oFechaini = oForm.DataSources.UserDataSources.Item("DS3Fecini").ValueEx;
                    oFechafin = oForm.DataSources.UserDataSources.Item("DS3Fecfin").ValueEx;
                    if ((oFechaini == "20000101") && (oFechafin == "20991231"))
                    {
                        NoEsPreciosBase = "(1 = 2)";
                        precioBase = true;
                    }
                    break;
                }
            if ((mtx3.RowCount < 1) || (oRowSelect == -1))
            {
                FSBOApp.MessageBox("Debe seleccionar un periodo !!!", 1, "Ok");
                return false;
            }
            mtx1.Clear();

            oSql0 = GlobalSettings.RunningUnderSQLServer ?
                   "Select " :
                   "Select 0 \"Ord\", \"ListNum\", \"ListName\", \"IsGrossPrc\"  from OPLN Where \"ListNum\" = {0} " +
                   " Union All " +
                   "Select 1 \"Ord\", \"ListNum\", \"ListName\", \"IsGrossPrc\" from OPLN Where \"ListNum\" != {0} order by 1, 3 ";
            oSql0 = String.Format(oSql0, BaseListPrice.ToString());
            oRS.DoQuery(oSql0);

            aux0 = "";
            aux3 = "";
            while (!oRS.EoF)
            {
                i = (Int32)(oRS.Fields.Item("ListNum").Value);
                GrossPrice = (((String)(oRS.Fields.Item("IsGrossPrc").Value)) == "Y");
                aux1 = "(Select p.\"ItemCode\", p.\"Price\" as \"Price{0}\", s.\"Price\" as \"SPrice{0}\", p.\"Currency\" as \"Currency{0}\", " +
                       "        Case When IfNull(s.\"Price\", 0) = 0 Then p.\"Price\" Else s.\"Price\" End \"CPrice{0}\",                     " +
                       "        Case When IfNull(s.\"Price\", 0) = 0 Then 'B'                                                                 " +
                       "             Else Case When s.\"FromDate\" = TO_DATE('" + oFechaini + "','YYYYMMDD') and s.\"ToDate\" = TO_DATE('" + oFechafin + "','YYYYMMDD') Then 'P' " +
                       "             Else 'OP' End                                                                                            " + 
                       "             End \"Tipo{0}\"                                                                                          " + 
                       "   From ITM1 p left outer join OPLN l on p.\"PriceList\" = l.\"ListNum\"                                              " +
                       "               left outer join SPP1 s on p.\"PriceList\" = s.\"ListNum\" and p.\"ItemCode\" = s.\"ItemCode\" and ( s.\"FromDate\" between TO_DATE('" + oFechaini + "','YYYYMMDD') and TO_DATE('" + oFechafin + "','YYYYMMDD') or " +
                       "                                                                                                                   s.\"ToDate\"   between TO_DATE('" + oFechaini + "','YYYYMMDD') and TO_DATE('" + oFechafin + "','YYYYMMDD') or " +
                       "                                                                                                                  (s.\"FromDate\" < TO_DATE('" + oFechaini + "','YYYYMMDD') and s.\"ToDate\" > TO_DATE('" + oFechafin + "','YYYYMMDD') )) " +
                       "                                                                                                             and  " + NoEsPreciosBase + 
                       "  Where l.\"ListNum\" = {0} ) T{0} on i.\"ItemCode\" = T{0}.\"ItemCode\" ";
                if (MargenBasePrecioVenta)
                {
                    if (GrossPrice)
                        aux2 = " , \"Price{0}\", \"SPrice{0}\", \"CPrice{0}\", \"Tipo{0}\", \"Currency{0}\", 'Y' \"GrossPrice{0}\",                                                                                        " +
                               "  Case When \"CPrice{0}\" <> 0 Then 100 * ( ( (\"CPrice{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"AvgPrice\"  ) / (\"CPrice{0}\" / (1+ifNull(t.\"Rate\", 0)/100)) ) Else 0 End MargenCosto{0},   " +
                               "  Case When \"CPrice{0}\" <> 0 Then 100 * ( ( (\"CPrice{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"LastPurPrc\") / (\"CPrice{0}\" / (1+ifNull(t.\"Rate\", 0)/100)) ) Else 0 End MargenCompra{0} ";
                    else
                        aux2 = " , \"Price{0}\", \"SPrice{0}\", \"CPrice{0}\", \"Tipo{0}\", \"Currency{0}\", 'N' \"GrossPrice{0}\",                        " +
                               "  Case When \"CPrice{0}\" <> 0 Then 100 * ((\"CPrice{0}\" - i.\"AvgPrice\") / \"CPrice{0}\") Else 0 End MargenCosto{0},   " +
                               "  Case When \"CPrice{0}\" <> 0 Then 100 * ((\"CPrice{0}\" - i.\"LastPurPrc\") / \"CPrice{0}\") Else 0 End MargenCompra{0} ";
                }
                else
                {
                    if (GrossPrice)
                        aux2 = " , \"Price{0}\", \"SPrice{0}\", \"CPrice{0}\", \"Tipo{0}\", \"Currency{0}\", 'Y' \"GrossPrice{0}\",                                                               " +
                               "  Case When i.\"AvgPrice\"   <> 0 Then 100 * ( ( (\"CPrice{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"AvgPrice\"  ) / i.\"AvgPrice\"  ) Else 0 End MargenCosto{0},       " +
                               "  Case When i.\"LastPurPrc\" <> 0 Then 100 * ( ( (\"CPrice{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"LastPurPrc\") / i.\"LastPurPrc\") Else 0 End MargenCompra{0} ";
                    else
                        aux2 = " , \"Price{0}\", \"SPrice{0}\", \"CPrice{0}\", \"Tipo{0}\", \"Currency{0}\", 'N' \"GrossPrice{0}\",                              " +
                               "  Case When i.\"AvgPrice\" <> 0 Then 100 * ((\"CPrice{0}\" - i.\"AvgPrice\") / i.\"AvgPrice\") Else 0 End MargenCosto{0},       " +
                               "  Case When i.\"LastPurPrc\" <> 0 Then 100 * ((\"CPrice{0}\" - i.\"LastPurPrc\") / i.\"LastPurPrc\") Else 0 End MargenCompra{0} ";
                }

                aux0 = aux0 + " left outer join  " + String.Format(aux1, i.ToString());
                aux3 = aux3 + String.Format(aux2, i.ToString());

                oListasPrecio.Add(i);
                oRS.MoveNext();
            }

            if (useItemCodeFilter)
                oWhere = " where i.\"ItemCode\" in " + ItemCodesFilter;
            else
                oWhere = " where IfNull(i.\"U_VK_Categoria\",'') like '{0}'            " +
                         "   and IfNull(i.\"U_VK_Familia\",'') like '{1}'              " +
                         "   and IfNull(i.\"U_VK_Grupo\",'') like '{2}'                " +
                         "   and IfNull(i.\"U_VK_Marca\",'') like '{3}'                " +
                         "   and IfNull(i.\"U_VK_Departamento\",'') like '{4}'         " +
                         "   and i.\"ItemCode\" like '{5}'                             ";

            oSql = "Select i.\"ItemCode\",   i.\"ItemName\", 'N' \"Update\",     " +
                   "       i.\"LastPurPrc\", i.\"LastPurCur\", i.\"LastPurDat\", " +
                   "       i.\"AvgPrice\",   i.\"LstEvlPric\", i.\"LstEvlDate\", " +
                   "       i.\"AvgPrice\",   t.\"Rate\" / 100 \"Rate\",          " +
                   "       'N' \"Modif\",                                        " +
                   "       ROW_NUMBER() OVER (ORDER BY i.\"ItemCode\") \"RowId\" " +
                   aux3 +
                   "  from OITM i left outer join OSTC t on i.\"TaxCodeAR\" = t.\"Code\" " +
                   aux0 +
                   oWhere +
                   " Order by 1 ";


            if (oTable == null)
            {
                Primeravez = true;
                oForm.DataSources.DataTables.Add("dtData");
            }

            oSql = String.Format(oSql, oCond2, oCond4, oCond3, oCond5, oCond1, oCond6);
            oTable = oForm.DataSources.DataTables.Item("dtData");
            oTable.Clear();
            oTable.ExecuteQuery(oSql);

            if (oTable.Rows.Count <= 0)
            {
                FSBOApp.MessageBox("No existen registros coincidentes !!!", 1, "Ok");
                return false;
            }

            if (oTable.Rows.Count > 0)
            {
                if (!Primeravez)
                {
                    mtx1.Columns.Item("Modif").DataBind.UnBind();
                    mtx1.Columns.Item("RowId").DataBind.UnBind();
                    mtx1.Columns.Item("Update").DataBind.UnBind();
                    mtx1.Columns.Item("ItemCode").DataBind.UnBind();
                    mtx1.Columns.Item("ItemName").DataBind.UnBind();
                    mtx1.Columns.Item("UltCompra").DataBind.UnBind();
                    mtx1.Columns.Item("CostProm").DataBind.UnBind();

                    oRS.DoQuery(oSql0);
                    while (!oRS.EoF)
                    {
                        i = (Int32)(oRS.Fields.Item("ListNum").Value);
                        mtx1.Columns.Item("P" + i.ToString()).DataBind.UnBind();
                        mtx1.Columns.Item("MP" + i.ToString()).DataBind.UnBind();
                        mtx1.Columns.Item("MC" + i.ToString()).DataBind.UnBind();
                        mtx1.Columns.Item("T" + i.ToString()).DataBind.UnBind();
                        oRS.MoveNext();
                    }
                }

                mtx1.Columns.Item("Modif").DataBind.Bind("dtData", "Modif");
                mtx1.Columns.Item("RowId").DataBind.Bind("dtData", "RowId");
                mtx1.Columns.Item("Update").DataBind.Bind("dtData", "Update");
                mtx1.Columns.Item("ItemCode").DataBind.Bind("dtData", "ItemCode");
                mtx1.Columns.Item("ItemName").DataBind.Bind("dtData", "ItemName");
                mtx1.Columns.Item("UltCompra").DataBind.Bind("dtData", "LastPurPrc");
                mtx1.Columns.Item("CostProm").DataBind.Bind("dtData", "AvgPrice");

                oRS.DoQuery(oSql0);
                while (!oRS.EoF)
                {
                    i = (Int32)(oRS.Fields.Item("ListNum").Value);
                    mtx1.Columns.Item("P" + i.ToString()).DataBind.Bind("dtData", "CPrice" + i.ToString());
                    mtx1.Columns.Item("MP" + i.ToString()).DataBind.Bind("dtData", "MargenCompra" + i.ToString());
                    mtx1.Columns.Item("MC" + i.ToString()).DataBind.Bind("dtData", "MargenCosto" + i.ToString());
                    mtx1.Columns.Item("T" + i.ToString()).DataBind.Bind("dtData", "Tipo" + i.ToString());
                    if (precioBase)
                    {
                        if ((BaseListPrice != -10) && (!primerRegistro))
                            mtx1.Columns.Item("P" + i.ToString()).Editable = false;
                        primerRegistro = false;
                    }
                    else
                        mtx1.Columns.Item("P" + i.ToString()).Editable = true;

                    oRS.MoveNext();
                }
            }

            mtx1.LoadFromDataSource();

            return true;
        }

        private void Updatemtx1()
        {
            Int32 oRow, i, j;
            Boolean precioBase = false;
            Boolean oUpdate = true;
            SAPbobsCOM.Items oItems = null;
            SAPbobsCOM.Items_Prices oItemPrice;
            SAPbobsCOM.SpecialPrices oSpp = null;
            Int32 nErr;
            String sErr;

            if (2 == FSBOApp.MessageBox("¿Desea actualizar la información de precios?", 2, "Ok", "No"))
                return;

            if ((oForm.DataSources.UserDataSources.Item("DS3Fecini").ValueEx == "20000101") || (oForm.DataSources.UserDataSources.Item("DS3Fecfin").ValueEx == "20993112"))
            {
                oItems = (SAPbobsCOM.Items)FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                precioBase = true;
            }

            for (oRow = 0; oRow < oTable.Rows.Count; oRow++)
            {
                if ((String)oTable.GetValue("Modif", oRow) != "Y")
                    continue;

                if (precioBase)
                {
                    oItems.GetByKey((String)oTable.GetValue("ItemCode", oRow));
                    oItemPrice = oItems.PriceList;

                    for (i = 0; i < oListasPrecio.Count; i++)
                    {
                        if ((Double)oTable.GetValue("CPrice" + oListasPrecio[i].ToString(), oRow) == (Double)oTable.GetValue("Price" + oListasPrecio[i].ToString(), oRow))
                            continue;
                        else
                            for (j = 0; j < oItemPrice.Count; j++)
                            {
                                oItemPrice.SetCurrentLine(j);
                                if (oItemPrice.PriceList == oListasPrecio[i])
                                    oItemPrice.Price = (Double)oTable.GetValue("CPrice" + oListasPrecio[i].ToString(), oRow);
                            }
                    }
                    oItems.Update();
                }
                else
                {
                    for (i = 0; i < oListasPrecio.Count; i++)
                    {
                        if ((String)oTable.GetValue("Tipo" + oListasPrecio[i].ToString(), oRow) == "B")
                            continue;
                        if ((Double)oTable.GetValue("CPrice" + oListasPrecio[i].ToString(), oRow) == (Double)oTable.GetValue("SPrice" + oListasPrecio[i].ToString(), oRow))
                            continue;

                        oSpp = (SAPbobsCOM.SpecialPrices)FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices);

                        oUpdate = true;
                        if (!oSpp.GetByKey((String)oTable.GetValue("ItemCode", oRow), "*" + oListasPrecio[i].ToString()))
                        {
                            oSpp.PriceListNum = oListasPrecio[i];
                            oSpp.ItemCode = (String)oTable.GetValue("ItemCode", oRow);
                            oUpdate = false;
                        }

                        for (j = 0; j < oSpp.SpecialPricesDataAreas.Count; j++)
                        {
                            oSpp.SpecialPricesDataAreas.SetCurrentLine(j);
                            if ((oSpp.SpecialPricesDataAreas.DateFrom == DateTime.ParseExact(oFechaini, "yyyyMMdd", null)) && (oSpp.SpecialPricesDataAreas.Dateto == DateTime.ParseExact(oFechafin, "yyyyMMdd", null)))
                            {
                                oSpp.SpecialPricesDataAreas.Delete();
                                break;
                            }
                        }

                        if ((Double)oTable.GetValue("CPrice" + oListasPrecio[i].ToString(), oRow) > 0)
                        {
                            if ((oSpp.SpecialPricesDataAreas.Count > 1) || ((oSpp.SpecialPricesDataAreas.PriceListNo != 0) && (oSpp.SpecialPricesDataAreas.SpecialPrice != 0.0)))
                                oSpp.SpecialPricesDataAreas.Add();
                            oSpp.SpecialPricesDataAreas.DateFrom = DateTime.ParseExact(oFechaini, "yyyyMMdd", null);
                            oSpp.SpecialPricesDataAreas.Dateto = DateTime.ParseExact(oFechafin, "yyyyMMdd", null);
                            oSpp.SpecialPricesDataAreas.PriceListNo = oListasPrecio[i];
                            oSpp.SpecialPricesDataAreas.PriceCurrency = (String)oTable.GetValue("Currency" + oListasPrecio[i].ToString(), oRow);
                            oSpp.SpecialPricesDataAreas.SpecialPrice = (Double)oTable.GetValue("CPrice" + oListasPrecio[i].ToString(), oRow);
                        }

                        if (!oUpdate)
                            nErr = oSpp.Add();
                        else
                            nErr = oSpp.Update();
                        if (nErr != 0)
                        {
                            FCmpny.GetLastError(out nErr, out sErr);
                            oSpp = null;
                            throw new Exception("Error " + nErr.ToString() + " : " + sErr);
                        }
                        oSpp = null;
                    }
                }
            }
            oForm.Mode = BoFormMode.fm_OK_MODE;
        }

        private void getDetalleProducts() // Solo view
        {
            String oSql, oSql0, oWhere;
            String aux0, aux1, aux2, aux3;
            String oCond1, oCond2, oCond3, oCond4, oCond5, oCond6;
            Int32 i;
            SAPbouiCOM.Matrix mtx2; //Productos
            Boolean Primeravez = false;
            Boolean GrossPrice = false;

            mtx2 = (Matrix)(oForm.Items.Item("mtx2").Specific);

            oCond1 = oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim() : "%";
            oCond2 = CodeCategoria.Trim() != "" ? CodeCategoria.Trim() : "%";
            oCond3 = CodeGrupo.Trim() != "" ? CodeGrupo.Trim() : "%";
            oCond4 = CodeFamilia.Trim() != "" ? CodeFamilia.Trim() : "%";
            oCond5 = oForm.DataSources.UserDataSources.Item("DSMarca").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSMarca").ValueEx.Trim() : "%";
            oCond6 = oForm.DataSources.UserDataSources.Item("DSItemCode").ValueEx.Trim() != "" ? "%" + oForm.DataSources.UserDataSources.Item("DSItemCode").ValueEx.Trim() + "%" : "%";

            if ((oCond1 == "%") && (oCond2 == "%") && (oCond3 == "%") && (oCond4 == "%") && (oCond5 == "%") && (oCond6 == "%") && (!useItemCodeFilter))
            {
                FSBOApp.MessageBox("Debe seleccionar al menos un filtro !!!", 1, "Ok");
                return;
            }

            mtx2.Clear();

            oSql0 = GlobalSettings.RunningUnderSQLServer ?
                   "Select  " :
                   "Select 0 \"Ord\", \"ListNum\", \"ListName\", \"IsGrossPrc\" from OPLN Where \"ListNum\" = {0} " +
                   " Union All " +
                   "Select 1 \"Ord\", \"ListNum\", \"ListName\", \"IsGrossPrc\" from OPLN Where \"ListNum\" != {0} order by 1, 3 ";
            oSql0 = String.Format(oSql0, BaseListPrice.ToString());
            oRS.DoQuery(oSql0);

            aux0 = "";
            aux3 = "";
            while (!oRS.EoF)
            {
                i = (Int32)(oRS.Fields.Item("ListNum").Value);
                GrossPrice = (((String)(oRS.Fields.Item("IsGrossPrc").Value)) == "Y");
                aux1 = "(Select p.\"ItemCode\", p.\"Price\" as \"Price{0}\", p.\"Currency\" as \"Currency{0}\",          " +
                       "        TO_DATE('20000101', 'YYYYMMDD') \"FromDate\" ,                                           " +
                       "        TO_DATE('20991231', 'YYYYMMDD') \"ToDate\"                                               " +
                       "   From ITM1 p left outer join OPLN l on p.\"PriceList\" = l.\"ListNum\"                         " +
                       "  Where l.\"ListNum\" = {0}                                                                      " +
                       "    UNION ALL                                                                                    " +
                       " Select s.\"ItemCode\", s.\"Price\" as \"Price{0}\", s.\"Currency\" as \"Currency{0}\",          " +
                       "        s.\"FromDate\" as \"FromDate{0}\", s.\"ToDate\" as \"ToDate{0}\"                         " +
                       "   From SPP1 s left outer join OPLN l on s.\"ListNum\" = l.\"ListNum\" and  s.\"ToDate\" >= CURRENT_DATE " +
                       "  Where l.\"ListNum\" = {0} ) T{0} on i.\"ItemCode\" = T{0}.\"ItemCode\"    and                  " +
                       "                                      a.\"U_Fecini\" = T{0}.\"FromDate\"    and                  " +
                       "                                      a.\"U_Fecfin\" = T{0}.\"ToDate\"                           ";

                if (MargenBasePrecioVenta)
                {
                    if (GrossPrice)
                        aux2 = " , \"Price{0}\", 'Y' \"GrossPrice{0}\",                                                                                                                                                 " +
                               "  Case When \"Price{0}\" <> 0 Then 100 * ( ( (\"Price{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"AvgPrice\"  ) / (\"Price{0}\" / (1+ifNull(t.\"Rate\", 0)/100)) ) Else 0 End MargenCosto{0},   " +
                               "  Case When \"Price{0}\" <> 0 Then 100 * ( ( (\"Price{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"LastPurPrc\") / (\"Price{0}\" / (1+ifNull(t.\"Rate\", 0)/100)) ) Else 0 End MargenCompra{0} ";
                    else
                        aux2 = " , \"Price{0}\", 'N' \"GrossPrice{0}\",                                                                                 " +
                               "  Case When \"Price{0}\" <> 0 Then 100 * ((\"Price{0}\" - i.\"AvgPrice\") / \"Price{0}\") Else 0 End MargenCosto{0},   " +
                               "  Case When \"Price{0}\" <> 0 Then 100 * ((\"Price{0}\" - i.\"LastPurPrc\") / \"Price{0}\") Else 0 End MargenCompra{0} ";
                }
                else
                {
                    if (GrossPrice)
                        aux2 = " , \"Price{0}\", 'Y' \"GrossPrice{0}\",                                                                                                                          " +
                               "  Case When i.\"AvgPrice\"   <> 0 Then 100 * ( ( (\"Price{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"AvgPrice\"  ) / i.\"AvgPrice\"  ) Else 0 End MargenCosto{0},       " +
                               "  Case When i.\"LastPurPrc\" <> 0 Then 100 * ( ( (\"Price{0}\" / (1+ifNull(t.\"Rate\", 0)/100) ) - i.\"LastPurPrc\") / i.\"LastPurPrc\") Else 0 End MargenCompra{0} ";
                    else
                        aux2 = " , \"Price{0}\", 'N' \"GrossPrice{0}\",                                                                                         " +
                               "  Case When i.\"AvgPrice\" <> 0 Then 100 * ((\"Price{0}\" - i.\"AvgPrice\") / i.\"AvgPrice\") Else 0 End MargenCosto{0},       " +
                               "  Case When i.\"LastPurPrc\" <> 0 Then 100 * ((\"Price{0}\" - i.\"LastPurPrc\") / i.\"LastPurPrc\") Else 0 End MargenCompra{0} ";
                }

                aux0 = aux0 + " left outer join  " + String.Format(aux1, i.ToString());
                aux3 = aux3 + String.Format(aux2, i.ToString());

                oRS.MoveNext();
            }

            if (useItemCodeFilter)
                oWhere = " where i.\"ItemCode\" in " + ItemCodesFilter;
            else
                oWhere = " where IfNull(i.\"U_VK_Categoria\",'') like '{0}'            " +
                         "   and IfNull(i.\"U_VK_Familia\",'') like '{1}'              " +
                         "   and IfNull(i.\"U_VK_Grupo\",'') like '{2}'                " +
                         "   and IfNull(i.\"U_VK_Marca\",'') like '{3}'                " +
                         "   and IfNull(i.\"U_VK_Departamento\",'') like '{4}'         " +
                         "   and i.\"ItemCode\" like '{5}'                             ";

            oSql = "Select i.\"ItemCode\",   i.\"ItemName\", i.\"U_VK_Marca\",   " +
                   "       ca.\"Name\" \"U_VK_Categoria\", fa.\"Name\" \"U_VK_Familia\", gr.\"Name\" \"U_VK_Grupo\", de.\"Name\" \"U_VK_Departamento\",  " +
                   "       a.\"U_Fecini\", a.\"U_Fecfin\", a.\"Code\", t.\"Rate\" \"Rate\",  " +
                   "       i.\"LastPurPrc\", i.\"LastPurCur\", i.\"LastPurDat\",             " +
                   "       i.\"AvgPrice\",   i.\"LstEvlPric\", i.\"LstEvlDate\"              " +
                   aux3 +
                   "  from OITM i left outer join OSTC t on i.\"TaxCodeAR\" = t.\"Code\"                         " +
                   "              left outer join \"@VIDR_DPTO\" de on i.\"U_VK_Departamento\" = de.\"Code\"     " +
                   "              left outer join \"@VIDR_CATEGORIA\" ca on i.\"U_VK_Categoria\" = ca.\"Code\"   " +
                   "              left outer join \"@VIDR_FAMILIA\" fa on i.\"U_VK_Familia\" = fa.\"Code\"       " +
                   "              left outer join \"@VIDR_GRUPO\" gr on i.\"U_VK_Grupo\" = gr.\"Code\"           " +
                   "              Cross join  (Select \"Code\", \"U_Fecini\", \"U_Fecfin\" from \"@VIDR_PERIODO\" Where \"U_Fecfin\" >= CURRENT_DATE order by \"U_Fecini\") a " +
                   aux0 +
                   oWhere +
                   " Order by 1,3                                                ";


            if (oTable2 == null)
            {
                Primeravez = true;
                oForm.DataSources.DataTables.Add("dtData2");
            }

            oSql = String.Format(oSql, oCond2, oCond4, oCond3, oCond5, oCond1, oCond6);
            oTable2 = oForm.DataSources.DataTables.Item("dtData2");
            oTable2.Clear();
            oTable2.ExecuteQuery(oSql);

            if (oTable2.Rows.Count <= 0)
            {
                FSBOApp.MessageBox("No existen registros coincidentes !!!", 1, "Ok");
                return;
            }

            if (oTable2.Rows.Count > 0)
            {
                if (!Primeravez)
                {
                    mtx2.Columns.Item("ItemCode").DataBind.UnBind();
                    mtx2.Columns.Item("ItemName").DataBind.UnBind();
                    mtx2.Columns.Item("Dpto").DataBind.UnBind();
                    mtx2.Columns.Item("Cat").DataBind.UnBind();
                    mtx2.Columns.Item("Grupo").DataBind.UnBind();
                    mtx2.Columns.Item("Familia").DataBind.UnBind();
                    mtx2.Columns.Item("Fecini").DataBind.UnBind();
                    mtx2.Columns.Item("Fecfin").DataBind.UnBind();
                    mtx2.Columns.Item("Periodo").DataBind.UnBind();

                    oRS.DoQuery(oSql0);
                    while (!oRS.EoF)
                    {
                        i = (Int32)(oRS.Fields.Item("ListNum").Value);
                        mtx2.Columns.Item("P" + i.ToString()).DataBind.UnBind();
                        oRS.MoveNext();
                    }
                }

                mtx2.Columns.Item("ItemCode").DataBind.Bind("dtData2", "ItemCode");
                mtx2.Columns.Item("ItemName").DataBind.Bind("dtData2", "ItemName");
                mtx2.Columns.Item("Dpto").DataBind.Bind("dtData2", "U_VK_Departamento");
                mtx2.Columns.Item("Cat").DataBind.Bind("dtData2", "U_VK_Categoria");
                mtx2.Columns.Item("Grupo").DataBind.Bind("dtData2", "U_VK_Grupo");
                mtx2.Columns.Item("Familia").DataBind.Bind("dtData2", "U_VK_Familia");
                mtx2.Columns.Item("Fecini").DataBind.Bind("dtData2", "U_Fecini");
                mtx2.Columns.Item("Fecfin").DataBind.Bind("dtData2", "U_Fecfin");
                mtx2.Columns.Item("Periodo").DataBind.Bind("dtData2", "Code");

                oRS.DoQuery(oSql0);
                while (!oRS.EoF)
                {
                    i = (Int32)(oRS.Fields.Item("ListNum").Value);
                    mtx2.Columns.Item("P" + i.ToString()).DataBind.Bind("dtData2", "Price" + i.ToString());
                    oRS.MoveNext();
                }
            }

            mtx2.LoadFromDataSource();

            return;
        }

        private bool LoadProducts()
        {
            String sFileName;

            ItemCodesFilter = "";

            TGetOpenFileClass OpenXlsfile = new TGetOpenFileClass();
            String dir = Directory.GetCurrentDirectory();
            //OpenXlsfile.Filter = "Archivos Texto (*.*)|*.*";
            OpenXlsfile.Filter = "Archivo con Articulos (*.xls)|*.xls";

            Thread threadOpenFile = new Thread(new ThreadStart(OpenXlsfile.Open));
            threadOpenFile.SetApartmentState(ApartmentState.STA);
            try
            {
                threadOpenFile.Start();
                while (!threadOpenFile.IsAlive) ;

                Thread.Sleep(1);
                //**threadOpenFile.Sleep(1); 
                threadOpenFile.Join();
                Directory.SetCurrentDirectory(dir);
                sFileName = OpenXlsfile.FileName;

            }
            catch (Exception r)
            {
                FSBOApp.StatusBar.SetText(r.Message , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Importar : " + r.Message + " ** Trace: " + r.StackTrace);
                return false;
            }

            try
            {
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                 "Data Source=" + sFileName + ";" +
                                 "Extended Properties=Excel 8.0;" +
                                 "Mode=Share Deny Write";

                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT T0.* FROM [" + "promocion$" + "] T0 ", strConn);
                DataSet ADOQueryExcel = new DataSet();

                adapter.Fill(ADOQueryExcel, "PRODUCTOS");
                int NroLineas = ADOQueryExcel.Tables["PRODUCTOS"].Rows.Count;

                if (NroLineas < 2)
                    throw new Exception("Archivo de productos sin datos.");

                bool primeraVez = true;
                var DataExcel = ADOQueryExcel.Tables["PRODUCTOS"].AsEnumerable();
                foreach (DataRow oRow in DataExcel)
                {
                    if (oRow["ItemCode"] == System.DBNull.Value)
                        continue;
                    if (oRow["ItemCode"].ToString().Trim() != "")
                    {
                        if (primeraVez)
                        {
                            primeraVez = false;
                            ItemCodesFilter = ItemCodesFilter + "'" + oRow["ItemCode"].ToString().Trim() + "'";
                        }
                        else
                            ItemCodesFilter = ItemCodesFilter + ", '" + oRow["ItemCode"].ToString().Trim() + "'";
                    }
                }
                ItemCodesFilter = " (" + ItemCodesFilter + ") ";

                return true;
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("Importar : " + e.Message + " ** Trace: " + e.StackTrace);
                return false;
            }
        }
    }
}
