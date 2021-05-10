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


namespace VID_Retail.Surtido
{
    class TSurtido : TvkBaseForm, IvkFormInterface
    {
        public TSurtido()
        {
        }

        private Int32 mtx1Height;
        private Int32 Rectang1Height;
        private Int32 Rectang2Height;
        private Int32 Rectang3Height;
        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.DataTable oTable = null;
        private SAPbouiCOM.DataTable oTabl2 = null;
        private SAPbouiCOM.DataTable oTblGrid = null;
        private SAPbouiCOM.Form oForm = null;
        private Int32 nTiendas;
        private Int32 oRowSelect;
        private String oCodeSelect;
        private List<Boolean> ItemInicial = new List<bool>();
        private List<Boolean> ItemFinal = new List<bool>();
        private string CodeCategoria = "";
        private string CodeGrupo = "";
        private string CodeFamilia = "";

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;
            SAPbouiCOM.Matrix mtx1;
            SAPbouiCOM.Matrix mtx2;
            SAPbouiCOM.Matrix mtx3;
            bool oResult;

            oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                try
                {
                    FSBOf.LoadForm(xmlPath, "Surtido.srf", uid);
                    EnableCrystal = false;

                    oForm = FSBOApp.Forms.Item(uid);
                    oForm.AutoManaged = true;
                    oForm.SupportedModes = 9;             // ok + view
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oForm.PaneLevel = 2;
                    oForm.Items.Item("tab2").Click();

                    // Ocultar tab1 temporalmente
                    oForm.Items.Item("tab1").Visible = true;

                    mtx1Height = oForm.Items.Item("mtx1").Height;
                    Rectang1Height = oForm.Items.Item("Rectang1").Height;
                    Rectang2Height = oForm.Items.Item("Rectang2").Height;
                    Rectang3Height = oForm.Items.Item("Rectang3").Height;

                    oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                    oForm.DataSources.UserDataSources.Add("DSDpto", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSCategor", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSGrupo", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSFamilia", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSMarca", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSEstado", BoDataType.dt_SHORT_TEXT, 10);
                    oForm.DataSources.UserDataSources.Add("DSTipoArt", BoDataType.dt_SHORT_TEXT, 10);
                    oForm.DataSources.UserDataSources.Add("DSFecIni", BoDataType.dt_DATE, 0);
                    oForm.DataSources.UserDataSources.Add("DSFecFin", BoDataType.dt_DATE, 0);
                    oForm.DataSources.UserDataSources.Add("DSTienda", BoDataType.dt_SHORT_TEXT, 50);
                    ((ComboBox)(oForm.Items.Item("cb_dpto").Specific)).DataBind.SetBound(true, "", "DSDpto");
                    ((EditText)(oForm.Items.Item("Marca").Specific)).DataBind.SetBound(true, "", "DSMarca");
                    ((EditText)(oForm.Items.Item("Grupo").Specific)).DataBind.SetBound(true, "", "DSGrupo");
                    ((EditText)(oForm.Items.Item("Categoria").Specific)).DataBind.SetBound(true, "", "DSCategor");
                    ((EditText)(oForm.Items.Item("Familia").Specific)).DataBind.SetBound(true, "", "DSFamilia");
                    ((EditText)(oForm.Items.Item("Estado").Specific)).DataBind.SetBound(true, "", "DSEstado");
                    ((EditText)(oForm.Items.Item("Estado2").Specific)).DataBind.SetBound(true, "", "DSEstado");
                    ((EditText)(oForm.Items.Item("TipoArt").Specific)).DataBind.SetBound(true, "", "DSTipoArt");
                    ((EditText)(oForm.Items.Item("FecIni").Specific)).DataBind.SetBound(true, "", "DSFecIni");
                    ((EditText)(oForm.Items.Item("FecFin").Specific)).DataBind.SetBound(true, "", "DSFecFin");
                    ((EditText)(oForm.Items.Item("FecIni3").Specific)).DataBind.SetBound(true, "", "DSFecIni");
                    ((EditText)(oForm.Items.Item("FecFin3").Specific)).DataBind.SetBound(true, "", "DSFecFin");
                    ((ComboBox)(oForm.Items.Item("Tienda").Specific)).DataBind.SetBound(true, "", "DSTienda");

                     oSql = GlobalSettings.RunningUnderSQLServer ?
                           "Select Code, Name from [@VIDR_TIENDA] order by Code" :
                           "Select \"Code\", \"Name\" from \"@VIDR_TIENDA\" order by \"Code\"";

                    oForm.Items.Item("mtx3").AffectsFormMode = true;
                    mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
                    mtx2 = (Matrix)(oForm.Items.Item("mtx2").Specific);
                    mtx3 = (Matrix)(oForm.Items.Item("mtx3").Specific);

                    //mtx3
                    mtx3.Columns.Item("ItemCode").Editable = false;
                    mtx3.Columns.Item("ItemName").Editable = false;
                    mtx3.Columns.Item("Dpto").Editable = false;
                    mtx3.Columns.Item("Cat").Editable = false;
                    mtx3.Columns.Item("Grupo").Editable = false;
                    mtx3.Columns.Item("Familia").Editable = false;
                    mtx3.Columns.Item("Marca").Editable = false;

                    //mtx1
                    oForm.DataSources.UserDataSources.Add("DSCode", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSName", BoDataType.dt_SHORT_TEXT, 100);
                    mtx1.Columns.Item("Code").DataBind.SetBound(true, "", "DSCode");
                    mtx1.Columns.Item("Name").DataBind.SetBound(true, "", "DSName");
                    mtx1.Columns.Item("Code").Editable = false;
                    mtx1.Columns.Item("Name").Editable = false;

                    oRS.DoQuery(oSql);
                    SetMtx(mtx1, "");
                    FillMtx1();

                    //mtx2
                    mtx2.Columns.Item("ItemCode").Editable = false;
                    mtx2.Columns.Item("ItemName").Editable = false;
                    mtx2.Columns.Item("Dpto").Editable = false;
                    mtx2.Columns.Item("Cat").Editable = false;
                    mtx2.Columns.Item("Grupo").Editable = false;
                    mtx2.Columns.Item("Familia").Editable = false;
                    mtx2.Columns.Item("Marca").Editable = false;
                    mtx2.Columns.Item("Estado").Editable = false;
                    mtx2.Columns.Item("Precio").Editable = false;
                    mtx2.Columns.Item("Unidades").Editable = false;
                    mtx2.Columns.Item("Total").Editable = false;
                    mtx2.Columns.Item("Cluster").Editable = true;

                    // oRS.DoQuery(oSql);
                    // SetMtx(mtx2, "2");

                    oSql = GlobalSettings.RunningUnderSQLServer ?
                            "Select Code, Name from [@VIDR_CLUSTER] order by Name" :
                            "Select \"Code\", \"Name\" from \"@VIDR_CLUSTER\" order by \"Name\" ";
                    oRS.DoQuery(oSql);
                    FSBOf.FillComboMtx((Column)mtx2.Columns.Item("Cluster"), ref oRS, true);

                    // grid


                    //Filtros
                    //Tiendas
                    oSql = GlobalSettings.RunningUnderSQLServer ?
                            "Select Code, Name from [@VIDR_TIENDA] order by Name" :
                            "Select \"Code\", \"Name\" from \"@VIDR_TIENDA\" order by \"Name\" ";
                    oRS.DoQuery(oSql);
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("Tienda").Specific)), ref oRS, true);

                    //Departamento
                    oSql = GlobalSettings.RunningUnderSQLServer ?
                            "Select Code, Name from [@VIDR_DPTO] order by Name" :
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
            SAPbouiCOM.Matrix mtx2;
            String sValue;
            String nVal;
            Int32 offS;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_CLICK:
                    if ((pVal.ItemUID == "tab1") && (!pVal.BeforeAction))
                        oForm.PaneLevel = 1;
                    else if ((pVal.ItemUID == "tab2") && (!pVal.BeforeAction))
                        oForm.PaneLevel = 2;
                    else if ((pVal.ItemUID == "tab3") && (!pVal.BeforeAction))
                        oForm.PaneLevel = 3;
                    else if ((pVal.ItemUID == "btnFltPrd") && (!pVal.BeforeAction))
                        FillDetalleProducto();
                    else if ((pVal.ItemUID == "btnFltGrd") && (!pVal.BeforeAction))
                        FillDetalleGrilla();
                    else if ((pVal.ItemUID == "mtx1") && (!pVal.BeforeAction))
                        FillDetalleCluster();
                    else if ((pVal.ItemUID == "1") && (pVal.BeforeAction) && (pVal.FormMode == (Int32)BoFormMode.fm_UPDATE_MODE))
                        Updatemtx3();
                    break;
                case BoEventTypes.et_FORM_RESIZE:
                    if (!pVal.BeforeAction)
                    {
                        oForm.Items.Item("mtx1").Height = mtx1Height;
                        oForm.Items.Item("Rectang1").Height = Rectang1Height;
                        oForm.Items.Item("Rectang2").Height = Rectang2Height;
                        oForm.Items.Item("Rectang3").Height = Rectang3Height;
                    }
                    break;
                case BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "Categoria")
                            {
                                AddChooseFromListDinamico("CFLCategoria", oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim());
                            }
                            else if (pVal.ItemUID == "Grupo")
                            {
                                AddChooseFromListDinamico("CFLGrupo", CodeCategoria);
                            }
                            else if (pVal.ItemUID == "Familia")
                            {
                                AddChooseFromListDinamico("CFLFamilia", CodeGrupo);
                            }
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
                    else if ((pVal.ItemUID == "mtx2") && (pVal.ColUID == "Cluster") && (!pVal.BeforeAction))
                    {
                        mtx2 = (Matrix)(oForm.Items.Item("mtx2").Specific);

                        offS = Int32.Parse(((EditText)(mtx2.Columns.Item("RowId").Cells.Item(pVal.Row).Specific)).Value) - 1;
                        nVal = ((ComboBox)(mtx2.Columns.Item("Cluster").Cells.Item(pVal.Row).Specific)).Value;
                        
                        oTabl2.SetValue("U_Cluster", offS, nVal);
                        oTabl2.Rows.Offset = offS;
                        mtx2.SetLineData(pVal.Row);
                    }
                    break;
            }
        }

        public new void MenuEvent(ref MenuEvent pVal, ref Boolean BubbleEvent)
        {
            //Int32 Entry;
            base.MenuEvent(ref pVal, ref BubbleEvent);
        }

        private void SetMtx(SAPbouiCOM.Matrix omtx, String prefix)
        {
            Int32 i;

            i = 0;
            nTiendas = 0;
            while (!oRS.EoF)
            {
                i++;
                nTiendas++;
                oForm.DataSources.UserDataSources.Add("DS" + prefix + "Tda" + i.ToString(), BoDataType.dt_SHORT_TEXT, 1);
                omtx.Columns.Add("Tda" + i.ToString(), BoFormItemTypes.it_CHECK_BOX);
                omtx.Columns.Item("Tda" + i.ToString()).DataBind.SetBound(true, "", ("DS" + prefix + "Tda" + i.ToString()));
                omtx.Columns.Item("Tda" + i.ToString()).TitleObject.Caption = (String)oRS.Fields.Item("Code").Value;
                omtx.Columns.Item("Tda" + i.ToString()).Editable = false;

                oRS.MoveNext();
            }
        }

        private void FillMtx1()
        {
            String oSql;
            String sCluster;
            Int32 oLine;
            Int32 i;
            SAPbouiCOM.Matrix mtx1;

            try
            {
                oForm.Freeze(true);
                mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);

                oSql = GlobalSettings.RunningUnderSQLServer ?
                       "select c.Code, c.Name, IsNull(c.U_Tienda, '') Tienda     " +
                       "  from [@VIDR_TIENDA] t right outer join                 " +
                       "       (Select h.Code, h.Name, d.U_Tienda, d.U_Activa    " +
                       "	      from [@VIDR_CLUSTER] h left outer join [@VIDR_CLUSTERD] d on h.Code = d.Code) c " +
                       "	   on t.Code = c.U_Tienda                            " +
                       "Order by 1, 3                                            " :
                       "select c.\"Code\", c.\"Name\", IfNull(c.\"U_Tienda\", '') \"Tienda\"     " +
                       "  from \"@VIDR_TIENDA\" t right outer join                               " +
                       "       (Select h.\"Code\", h.\"Name\", d.\"U_Tienda\", d.\"U_Activa\"    " +
                       "	      from \"@VIDR_CLUSTER\" h left outer join \"@VIDR_CLUSTERD\" d on h.\"Code\" = d.\"Code\") c " +
                       "	   on t.\"Code\" = c.\"U_Tienda\"                                    " +
                       "Order by 1, 3                                                            ";
                oRS.DoQuery(oSql);

                mtx1.Clear();
                oLine = 0;
                sCluster = "";
                while (!oRS.EoF)
                {
                    if (sCluster != (String)oRS.Fields.Item("Code").Value)
                    {
                        sCluster = (String)oRS.Fields.Item("Code").Value;
                        oForm.DataSources.UserDataSources.Item("DSCode").ValueEx = (String)oRS.Fields.Item("Code").Value;
                        oForm.DataSources.UserDataSources.Item("DSName").ValueEx = (String)oRS.Fields.Item("Name").Value;
                        for (i = 1; i <= nTiendas; i++)
                            oForm.DataSources.UserDataSources.Item("DSTda" + i.ToString()).ValueEx = "N";
                    }

                    while ((!oRS.EoF) && (sCluster == (String)oRS.Fields.Item("Code").Value))
                    {
                        if (((String)oRS.Fields.Item("Tienda").Value).Trim() != "")
                        {
                            oLine = 0;
                            for (i = 1; i <= mtx1.Columns.Count - 1; i++)
                            {
                                if (mtx1.Columns.Item(i).DataBind.Alias.Substring(0, 5) == "DSTda")
                                    oLine++;
                                if (mtx1.Columns.Item(i).TitleObject.Caption == ((String)oRS.Fields.Item("Tienda").Value).Trim())
                                    break;
                            }
                            oForm.DataSources.UserDataSources.Item("DSTda" + oLine.ToString()).ValueEx = "Y";
                        }
                        oRS.MoveNext();
                    }
                    mtx1.AddRow(1, -1);
                }

                ClearDS();
                mtx1.AddRow(1, -1);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void ClearDS()
        {
            SAPbouiCOM.Matrix mtx1;

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);

            oForm.DataSources.UserDataSources.Item("DSCode").ValueEx = "";
            oForm.DataSources.UserDataSources.Item("DSName").ValueEx = "";
            for (int i = 1; i <= nTiendas; i++)
                oForm.DataSources.UserDataSources.Item("DSTda" + i.ToString()).ValueEx = "";
        }

        private void FillDetalleCluster()
        {
            String oSql;
            Boolean Primeravez = false;
            SAPbouiCOM.Matrix mtx1;
            SAPbouiCOM.Matrix mtx3;

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            mtx3 = (Matrix)(oForm.Items.Item("mtx3").Specific);
                                        
            oRowSelect = -1;
            oCodeSelect = "";
            for (int i = 1; i <= mtx1.RowCount; i++)
                if (mtx1.IsRowSelected(i))
                {
                    oRowSelect = i;
                    mtx1.GetLineData(i);
                    oCodeSelect = oForm.DataSources.UserDataSources.Item("DSCode").ValueEx;
                    break;
                }

            if ((mtx1.RowCount < 1) || (oRowSelect == -1))
            {
                return;
            }

            mtx3.Clear();

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select " :
                   "Select i.\"ItemCode\", i.\"ItemName\", i.\"U_VK_Marca\",                                     " +
                   "       ca.\"Name\" \"U_VK_Categoria\", fa.\"Name\" \"U_VK_Familia\", gr.\"Name\" \"U_VK_Grupo\", de.\"Name\" \"U_VK_Departamento\",  " +
                   "	   i.\"U_Cluster\", i.\"U_Tipo_Articulo\", i.\"U_Estado\"                                " +
                   "  from OITM i left outer join \"@VIDR_DPTO\" de on i.\"U_VK_Departamento\" = de.\"Code\"     " +
                   "              left outer join \"@VIDR_CATEGORIA\" ca on i.\"U_VK_Categoria\" = ca.\"Code\"   " +
                   "              left outer join \"@VIDR_FAMILIA\" fa on i.\"U_VK_Familia\" = fa.\"Code\"       " +
                   "              left outer join \"@VIDR_GRUPO\" gr on i.\"U_VK_Grupo\" = gr.\"Code\"           " +
                   " where IfNull(i.\"U_Cluster\",'') like '{0}'                                                 ";
            oSql = String.Format(oSql, oCodeSelect);

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
                FSBOApp.MessageBox("No existen registros coincidentes !!!", 1, "Ok");
                return;
            }

            if (oTable.Rows.Count > 0)
            {
                if (!Primeravez)
                {
                    mtx3.Columns.Item("ItemCode").DataBind.UnBind();
                    mtx3.Columns.Item("ItemName").DataBind.UnBind();
                    mtx3.Columns.Item("Dpto").DataBind.UnBind();
                    mtx3.Columns.Item("Cat").DataBind.UnBind();
                    mtx3.Columns.Item("Grupo").DataBind.UnBind();
                    mtx3.Columns.Item("Familia").DataBind.UnBind();
                    mtx3.Columns.Item("Marca").DataBind.UnBind();
                }

                mtx3.Columns.Item("ItemCode").DataBind.Bind("dtData", "ItemCode");
                mtx3.Columns.Item("ItemName").DataBind.Bind("dtData", "ItemName");
                mtx3.Columns.Item("Dpto").DataBind.Bind("dtData", "U_VK_Departamento");
                mtx3.Columns.Item("Cat").DataBind.Bind("dtData", "U_VK_Categoria");
                mtx3.Columns.Item("Grupo").DataBind.Bind("dtData", "U_VK_Grupo");
                mtx3.Columns.Item("Familia").DataBind.Bind("dtData", "U_VK_Familia");
                mtx3.Columns.Item("Marca").DataBind.Bind("dtData", "U_VK_Marca");

                mtx3.LoadFromDataSource();
                oForm.Mode = BoFormMode.fm_OK_MODE;
            }
        }

        private void FillDetalleProducto()
        {
            String oSql;
            String oDpto, oCategor, oGrupo, oFamilia, oMarca, oEstado, oTipoArt;
            String oFecini, oFecfin;
            SAPbouiCOM.Matrix mtx2;
            Boolean Primeravez = false;

            mtx2 = (Matrix)(oForm.Items.Item("mtx2").Specific);

            oDpto = oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSDpto").ValueEx.Trim() : "%";
            oCategor = CodeCategoria.Trim() != "" ? CodeCategoria.Trim() : "%";
            oGrupo = CodeGrupo.Trim() != "" ? CodeGrupo.Trim() : "%";
            oFamilia = CodeFamilia.Trim() != "" ? CodeFamilia.Trim() : "%";
            oMarca = oForm.DataSources.UserDataSources.Item("DSMarca").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSMarca").ValueEx.Trim() : "%";
            oEstado = oForm.DataSources.UserDataSources.Item("DSEstado").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSEstado").ValueEx.Trim() : "%";
            oTipoArt = oForm.DataSources.UserDataSources.Item("DSTipoArt").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSTipoArt").ValueEx.Trim() : "%";
            oFecini = oForm.DataSources.UserDataSources.Item("DSFecIni").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSFecIni").ValueEx.Trim() : "20000101";
            oFecfin = oForm.DataSources.UserDataSources.Item("DSFecFin").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSFecFin").ValueEx.Trim() : "20991231";

            if (2 == FSBOApp.MessageBox("Desea recuperar productos?", 1, "Ok", "Cancel"))
                return;

            if ((oDpto == "%") && (oCategor == "%") && (oGrupo == "%") && (oFamilia == "%") && (oMarca == "%") && (oEstado == "%") && (oTipoArt == "%"))
            {
                FSBOApp.MessageBox("Debe seleccionar al menos un filtro !!!", 1, "Ok");
                return;
            }

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select " :
                   "Select i.\"ItemCode\", i.\"ItemName\", i.\"U_VK_Marca\",                                       " +
                   "       ca.\"Name\" \"U_VK_Categoria\", fa.\"Name\" \"U_VK_Familia\", gr.\"Name\" \"U_VK_Grupo\", de.\"Name\" \"U_VK_Departamento\",  " +
                   "	   i.\"U_Cluster\", i.\"U_Tipo_Articulo\", i.\"U_Estado\", pr.\"Price\",                 " +
                   "       i.\"U_Cluster\" \"OldCluster\",                                                       " +
                   "       f.\"INMPrice\" - nc.\"INMPrice\" + f.\"VatSum\" - nc.\"VatSum\" \"Venta\",            " +
                   "       f.\"VatSum\" - nc.\"VatSum\" \"Impuestos\",                                           " +
                   "       f.\"Quantity\" - nc.\"Quantity\" \"Unidades\",                                        " +
                   "       ROW_NUMBER() OVER (ORDER BY i.\"ItemCode\") \"RowId\"                                 " +
                   "  from OITM i left outer join ITM1 pr on i.\"ItemCode\" = pr.\"ItemCode\" and pr.\"PriceList\" = 1                          " +
                   "              left outer join ( Select d.\"ItemCode\", SUM(d.\"INMPrice\") \"INMPrice\", SUM(d.\"VatSum\") \"VatSum\", SUM(d.\"Quantity\") \"Quantity\" " +
                   "                                  From OINV h inner join INV1 d on h.\"DocEntry\" = d.\"DocEntry\"                          " +
                   "                                 Where h.CANCELED = 'N'                                                                     " +
                   "                                   and h.\"TaxDate\" between TO_DATE('{7}', 'YYYYMMDD') and TO_DATE('{8}', 'YYYYMMDD')      " +
                   "                                 Group by d.\"ItemCode\" ) f                                                                " +
                   "                on i.\"ItemCode\" = f.\"ItemCode\"                                                                          " +
                   "              left outer join ( Select d.\"ItemCode\", SUM(d.\"INMPrice\") \"INMPrice\", SUM(d.\"VatSum\") \"VatSum\", SUM(d.\"Quantity\") \"Quantity\" " +
                   "                                  From ORIN h inner join RIN1 d on h.\"DocEntry\" = d.\"DocEntry\"                          " +
                   "                                 Where h.CANCELED = 'N'                                                                     " +
                   "                                   and h.\"TaxDate\" between TO_DATE('{7}', 'YYYYMMDD') and TO_DATE('{8}', 'YYYYMMDD')      " +
                   "                                 Group by d.\"ItemCode\" ) nc                                                               " +
                   "                on i.\"ItemCode\" = nc.\"ItemCode\"                                                                         " +
                   "              left outer join \"@VIDR_DPTO\" de on i.\"U_VK_Departamento\" = de.\"Code\"     " +
                   "              left outer join \"@VIDR_CATEGORIA\" ca on i.\"U_VK_Categoria\" = ca.\"Code\"     " +
                   "              left outer join \"@VIDR_FAMILIA\" fa on i.\"U_VK_Familia\" = fa.\"Code\"     " +
                   "              left outer join \"@VIDR_GRUPO\" gr on i.\"U_VK_Grupo\" = gr.\"Code\"     " +
                   " where IfNull(i.\"U_VK_Categoria\",'') like '{0}'                                            " +
                   "   and IfNull(i.\"U_VK_Familia\",'') like '{1}'                                              " +
                   "   and IfNull(i.\"U_VK_Grupo\",'') like '{2}'                                                " +
                   "   and IfNull(i.\"U_VK_Marca\",'') like '{3}'                                                " +
                   "   and IfNull(i.\"U_VK_Departamento\",'') like '{4}'                                         " +
                   "   and IfNull(i.\"U_Estado\",'') like '{5}'                                                  " +
                   "   and IfNull(i.\"U_Tipo_Articulo\",'') like '{6}'                                           " +
                   "   and i.\"frozenFor\" = 'N'                                                                 ";

            oSql = String.Format(oSql, oCategor, oFamilia, oGrupo, oMarca, oDpto, oEstado, oTipoArt, oFecini, oFecfin);

            if (oTabl2 == null)
            {
                Primeravez = true;
                oForm.DataSources.DataTables.Add("dtData2");
            }

            mtx2.Clear();
            oTabl2 = oForm.DataSources.DataTables.Item("dtData2");
            oTabl2.Clear();
            oTabl2.ExecuteQuery(oSql);

            if (oTabl2.Rows.Count <= 0)
            {
                FSBOApp.MessageBox("No existen registros coincidentes !!!", 1, "Ok");
                return;
            }

            if (oTabl2.Rows.Count > 0)
            {
                if (!Primeravez)
                {
                    mtx2.Columns.Item("ItemCode").DataBind.UnBind();
                    mtx2.Columns.Item("ItemName").DataBind.UnBind();
                    mtx2.Columns.Item("Dpto").DataBind.UnBind();
                    mtx2.Columns.Item("Cat").DataBind.UnBind();
                    mtx2.Columns.Item("Grupo").DataBind.UnBind();
                    mtx2.Columns.Item("Familia").DataBind.UnBind();
                    mtx2.Columns.Item("Marca").DataBind.UnBind();
                    mtx2.Columns.Item("Estado").DataBind.UnBind();
                    mtx2.Columns.Item("Precio").DataBind.UnBind();
                    mtx2.Columns.Item("Unidades").DataBind.UnBind();
                    mtx2.Columns.Item("Total").DataBind.UnBind();
                    mtx2.Columns.Item("Cluster").DataBind.UnBind();
                    mtx2.Columns.Item("RowId").DataBind.UnBind();
                }

                mtx2.Columns.Item("ItemCode").DataBind.Bind("dtData2", "ItemCode");
                mtx2.Columns.Item("ItemName").DataBind.Bind("dtData2", "ItemName");
                mtx2.Columns.Item("Dpto").DataBind.Bind("dtData2", "U_VK_Departamento");
                mtx2.Columns.Item("Cat").DataBind.Bind("dtData2", "U_VK_Categoria");
                mtx2.Columns.Item("Grupo").DataBind.Bind("dtData2", "U_VK_Grupo");
                mtx2.Columns.Item("Familia").DataBind.Bind("dtData2", "U_VK_Familia");
                mtx2.Columns.Item("Marca").DataBind.Bind("dtData2", "U_VK_Marca");
                mtx2.Columns.Item("Estado").DataBind.Bind("dtData2", "U_Estado");
                mtx2.Columns.Item("Precio").DataBind.Bind("dtData2", "Price");
                mtx2.Columns.Item("Unidades").DataBind.Bind("dtData2", "Unidades");
                mtx2.Columns.Item("Total").DataBind.Bind("dtData2", "Venta");
                mtx2.Columns.Item("Cluster").DataBind.Bind("dtData2", "U_Cluster");
                mtx2.Columns.Item("RowId").DataBind.Bind("dtData2", "RowId");

                mtx2.LoadFromDataSource();
                oForm.Mode = BoFormMode.fm_OK_MODE;
            }
        }

        private void FillDetalleGrilla()
        {
            SAPbouiCOM.Grid grid0;
            String oSql;
            String oTienda;
            String oFecini, oFecfin;

            if (1 != FSBOApp.MessageBox("¿Desea recuperar la información de surtidos?", 1, "Ok", "No"))
                return;

            oTienda = oForm.DataSources.UserDataSources.Item("DSTienda").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSTienda").ValueEx.Trim() : "%";
            oFecini = oForm.DataSources.UserDataSources.Item("DSFecIni").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSFecIni").ValueEx.Trim() : "20000101";
            oFecfin = oForm.DataSources.UserDataSources.Item("DSFecFin").ValueEx.Trim() != "" ? oForm.DataSources.UserDataSources.Item("DSFecFin").ValueEx.Trim() : "20991231";

            grid0 = (Grid)(oForm.Items.Item("grid0").Specific);

            if (oTblGrid == null)
            {
                //Primeravez = true;
                oForm.DataSources.DataTables.Add("dtDataGrid");
            }

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select " :
                   "Select de.\"Name\" \"U_VK_Departamento\", ca.\"Name\" \"U_VK_Categoria\", gr.\"Name\" \"U_VK_Grupo\", fa.\"Name\" \"U_VK_Familia\",  " +
                   "       i.\"ItemCode\", i.\"ItemName\", i.\"U_VK_Marca\",                                     " +
                   "	   i.\"U_Cluster\", i.\"U_Tipo_Articulo\", i.\"U_Estado\", pr.\"Price\",                 " +
                   "       f.\"INMPrice\" - nc.\"INMPrice\" + f.\"VatSum\" - nc.\"VatSum\" \"Venta\",            " +
                   "       f.\"VatSum\" - nc.\"VatSum\" \"Impuestos\",                                           " +
                   "       f.\"Quantity\" - nc.\"Quantity\" \"Unidades\"                                         " +
                   "      ,(Select SUM(\"OnHand\") \"OnHand\"                                                    " +
                   "          from OITW T0 inner join OWHS T1 on T0.\"WhsCode\" = T1.\"WhsCode\"                 " +
                   "         Where T0.\"ItemCode\" = i.\"ItemCode\"                                              " +
                   "           and T1.\"U_VID_Tda\" = '{0}'        ) \"Stock\"                                   " +
//                   "      , ROW_NUMBER() OVER (ORDER BY i.\"ItemCode\") \"RowId\"                                 " +
                   "  from OITM i left outer join ITM1 pr on i.\"ItemCode\" = pr.\"ItemCode\" and pr.\"PriceList\" = 1                          " +
                   "              left outer join ( Select d.\"ItemCode\", SUM(d.\"INMPrice\") \"INMPrice\", SUM(d.\"VatSum\") \"VatSum\", SUM(d.\"Quantity\") \"Quantity\" " +
                   "                                  From OINV h inner join INV1 d on h.\"DocEntry\" = d.\"DocEntry\"                          " +
                   "                                 Where h.CANCELED = 'N'                                                                     " +
                   "                                   and h.\"TaxDate\" between TO_DATE('{1}', 'YYYYMMDD') and TO_DATE('{2}', 'YYYYMMDD')      " +
                   "                                 Group by d.\"ItemCode\" ) f                                                                " +
                   "                on i.\"ItemCode\" = f.\"ItemCode\"                                                                          " +
                   "              left outer join ( Select d.\"ItemCode\", SUM(d.\"INMPrice\") \"INMPrice\", SUM(d.\"VatSum\") \"VatSum\", SUM(d.\"Quantity\") \"Quantity\" " +
                   "                                  From ORIN h inner join RIN1 d on h.\"DocEntry\" = d.\"DocEntry\"                          " +
                   "                                 Where h.CANCELED = 'N'                                                                     " +
                   "                                   and h.\"TaxDate\" between TO_DATE('{1}', 'YYYYMMDD') and TO_DATE('{2}', 'YYYYMMDD')      " +
                   "                                 Group by d.\"ItemCode\" ) nc                                                               " +
                   "                on i.\"ItemCode\" = nc.\"ItemCode\"                                                                         " +
                   "              left outer join \"@VIDR_DPTO\" de on i.\"U_VK_Departamento\" = de.\"Code\"                                    " +
                   "              left outer join \"@VIDR_CATEGORIA\" ca on i.\"U_VK_Categoria\" = ca.\"Code\"                                  " +
                   "              left outer join \"@VIDR_FAMILIA\" fa on i.\"U_VK_Familia\" = fa.\"Code\"                                      " +
                   "              left outer join \"@VIDR_GRUPO\" gr on i.\"U_VK_Grupo\" = gr.\"Code\"                                          " +
                   " Where IfNull(i.\"U_Cluster\",'') in (Select \"Code\" from \"@VIDR_CLUSTERD\" where \"U_Tienda\" = '{0}' )                  " +
                   "   and i.\"frozenFor\" = 'N'                                                                                                " +
                   " order by i.\"U_VK_Departamento\", i.\"U_VK_Categoria\", i.\"U_VK_Grupo\", i.\"U_VK_Familia\", i.\"ItemCode\"               ";
            oSql = String.Format(oSql, oTienda, oFecini, oFecfin);

            //grid0.();
            oTblGrid = oForm.DataSources.DataTables.Item("dtDataGrid");
            oTblGrid.Clear();
            oTblGrid.ExecuteQuery(oSql);

            grid0.DataTable = oTblGrid;
            grid0.Columns.Item("U_VK_Departamento").TitleObject.Caption = "Departamento";
            grid0.Columns.Item("U_VK_Categoria").TitleObject.Caption = "Categoria";
            grid0.Columns.Item("U_VK_Grupo").TitleObject.Caption = "Grupo";
            grid0.Columns.Item("U_VK_Familia").TitleObject.Caption = "Familia";
            grid0.Columns.Item("ItemCode").TitleObject.Caption = "Producto";
            grid0.Columns.Item("ItemName").TitleObject.Caption = "Descripción";
            grid0.Columns.Item("U_VK_Marca").TitleObject.Caption = "Marca";
            grid0.Columns.Item("U_Cluster").TitleObject.Caption = "Cluster";
            grid0.Columns.Item("U_Tipo_Articulo").TitleObject.Caption = "Tipo artículo";
            grid0.Columns.Item("U_Estado").TitleObject.Caption = "Estado";
            grid0.Columns.Item("Price").TitleObject.Caption = "Precio";
            grid0.Columns.Item("Venta").TitleObject.Caption = "Venta";
            grid0.Columns.Item("Stock").TitleObject.Caption = "Stock";
            grid0.Columns.Item("Unidades").TitleObject.Caption = "Unidades";
            grid0.AutoResizeColumns();
            grid0.CollapseLevel = 4;

            grid0.Rows.CollapseAll();
//            grid0.CommonSetting.FixedColumnsCount = 1;

        }

        private Boolean Updatemtx3()
        {
            SAPbouiCOM.Matrix mtx2;
            SAPbobsCOM.Items oItem;
            String oVal, nVal;
            String sErr;
            Int32 nErr;

            if (2 == FSBOApp.MessageBox("¿Desea actualizar la información de surtidos?", 2, "Ok", "No"))
                return false;

            mtx2 = (Matrix)(oForm.Items.Item("mtx2").Specific);

            for (int i = 0; i < oTabl2.Rows.Count; i++)
            {
                oVal = ((String)oTabl2.GetValue("OldCluster", i));
                nVal = ((String)oTabl2.GetValue("U_Cluster", i));
                if (oVal != nVal)
                {
                    oItem = (SAPbobsCOM.Items)FCmpny.GetBusinessObject(BoObjectTypes.oItems);
                    if (oItem.GetByKey(((String)oTabl2.GetValue("ItemCode", i)).Trim()))
                    {
                        oItem.UserFields.Fields.Item("U_Cluster").Value = nVal;
                        nErr = oItem.Update();
                        if (nErr != 0)
                        {
                            FCmpny.GetLastError(out nErr, out sErr);
                            if (FSBOApp.MessageBox("Fila: " + (i+1).ToString() + " - " + sErr, 1, "Continuar", "Detener actualización", "") == 2)
                                return false;
                        }
                    }
                }
            }

            oForm.Mode = BoFormMode.fm_OK_MODE;
            return true;
        }
    }
}