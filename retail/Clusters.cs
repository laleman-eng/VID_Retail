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


namespace VID_Retail.Clusters
{
    class TClusters : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.Application R_application;
        SAPbobsCOM.Company R_company;
        CSBOFunctions R_sboFunctions;
        TGlobalVid R_GlobalSettings;

        public TClusters()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        private Int32 nTiendas;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;
            Int32 i;
            SAPbouiCOM.Matrix mtx1;

            R_application = application;
            R_company = company;
            R_sboFunctions = sboFunctions;
            R_GlobalSettings = _GlobalSettings;

            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                try
                {
                    FSBOf.LoadForm(xmlPath, "Clusters.srf", uid);
                    EnableCrystal = false;

                    oForm = FSBOApp.Forms.Item(uid);
                    oForm.AutoManaged = true;
                    oForm.SupportedModes = -1;             // afm_All
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oForm.PaneLevel = 1;

                    oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                    oSql = GlobalSettings.RunningUnderSQLServer ?
                           "Select Code, Name from [@VIDR_TIENDA] order by Code" :
                           "Select \"Code\", \"Name\" from \"@VIDR_TIENDA\" order by \"Code\"";
                    oRS.DoQuery(oSql);

                    oForm.Items.Item("mtx1").AffectsFormMode = true;
                    mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);         
           
                    oForm.DataSources.UserDataSources.Add("DSCode", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSName", BoDataType.dt_SHORT_TEXT, 100);
                    mtx1.Columns.Item("Code").DataBind.SetBound(true, "", "DSCode");
                    mtx1.Columns.Item("Name").DataBind.SetBound(true, "", "DSName");

                    i = 0;
                    nTiendas = 0;
                    while (!oRS.EoF)
                    {
                        i++;
                        nTiendas++;
                        oForm.DataSources.UserDataSources.Add("DSTda" + i.ToString(), BoDataType.dt_SHORT_TEXT, 1);
                        mtx1.Columns.Add("Tda" + i.ToString(), BoFormItemTypes.it_CHECK_BOX);
                        mtx1.Columns.Item("Tda" + i.ToString()).DataBind.SetBound(true, "", ("DSTda" + i.ToString()));
                        mtx1.Columns.Item("Tda" + i.ToString()).TitleObject.Caption = (String)oRS.Fields.Item("Code").Value;
                        oRS.MoveNext();
                    }

                    FillMtx();

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
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "1") && (pVal.BeforeAction) && ((pVal.FormMode == (Int32)BoFormMode.fm_ADD_MODE) || (pVal.FormMode == (Int32)BoFormMode.fm_UPDATE_MODE)))
                        {
                            BubbleEvent = false;
                            ValidarDatos();
                            BubbleEvent = true;
                        }
                        if ((pVal.ItemUID == "1") && (!pVal.BeforeAction) && ((pVal.FormMode == (Int32)BoFormMode.fm_ADD_MODE) || (pVal.FormMode == (Int32)BoFormMode.fm_UPDATE_MODE)))
                        {
                            insertarDatos();
                        }
                        break;
                    case BoEventTypes.et_FORM_RESIZE:
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

        private void FillMtx()
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

        private void insertarDatos()
        {
            SAPbouiCOM.Matrix mtx1;
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.GeneralDataCollection oGeneralLines;
            SAPbobsCOM.GeneralData oGeneralLine;
            SAPbobsCOM.CompanyService oCompService;
            String oDSUser;
            String oSql;
            String oCode;

            oCompService = FCmpny.GetCompanyService();
            oGeneralService = oCompService.GetGeneralService("VIDR_CLUSTER");

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select Code from [@VIDR_CLUSTER] where Code = '{0}'" :
                   "Select \"Code\" from \"@VIDR_CLUSTER\" where \"Code\" = '{0}'";
            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            for (int i = 1; i <= mtx1.RowCount; i++)
            {
                mtx1.GetLineData(i);
                oCode = oForm.DataSources.UserDataSources.Item("DSCode").ValueEx.Trim();
                if (oCode == "")
                    continue;
                oRS.DoQuery(String.Format(oSql, oCode));

                if (oRS.EoF)
                {
                    oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                    oGeneralData.SetProperty("Code", oCode);
                    oGeneralData.SetProperty("Name", oForm.DataSources.UserDataSources.Item("DSName").ValueEx.Trim());
                    oGeneralLines = oGeneralData.Child("VIDR_CLUSTERD");
                    for (int j = 1; j <= mtx1.Columns.Count - 1; j++)
                    {
                        oDSUser = mtx1.Columns.Item(j).DataBind.Alias;
                        if ((mtx1.Columns.Item(j).DataBind.Alias.Substring(0, 5) == "DSTda") && (oForm.DataSources.UserDataSources.Item(oDSUser).ValueEx.Trim() == "Y"))
                        {
                            oGeneralLine = oGeneralLines.Add();
                            oGeneralLine.SetProperty("U_Tienda", mtx1.Columns.Item(j).TitleObject.Caption);
                        }
                    }
                    oGeneralService.Add(oGeneralData);
                }
                else
                {
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("Code", oCode);
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oGeneralData.SetProperty("Name", oForm.DataSources.UserDataSources.Item("DSName").ValueEx.Trim());
                    oGeneralLines = oGeneralData.Child("VIDR_CLUSTERD");
                    for (int j = oGeneralLines.Count - 1; j >= 0; j--)
                        oGeneralLines.Remove(j);

                    for (int j = 1; j <= mtx1.Columns.Count - 1; j++)
                    {
                        oDSUser = mtx1.Columns.Item(j).DataBind.Alias;
                        if ((mtx1.Columns.Item(j).DataBind.Alias.Substring(0, 5) == "DSTda") && (oForm.DataSources.UserDataSources.Item(oDSUser).ValueEx.Trim() == "Y"))
                        {
                            oGeneralLine = oGeneralLines.Add();
                            oGeneralLine.SetProperty("U_Tienda", mtx1.Columns.Item(j).TitleObject.Caption);
                        }
                    }
                    oGeneralService.Update(oGeneralData);
                }
            }
            oGeneralService = null;
        }

        private void ValidarDatos()
        {
            SAPbouiCOM.Matrix mtx1;
            String oCode;
            String oDSUser;
            String sAux;

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);

            List<String> oMtx = new List<String>();

            for (int i = 1; i <= mtx1.RowCount; i++)
            {
                mtx1.GetLineData(i);
                oCode = oForm.DataSources.UserDataSources.Item("DSCode").ValueEx.Trim();
                if (oCode == "")
                    continue;

                sAux = "";
                for (int j = 3; j <= mtx1.Columns.Count - 1; j++)
                {
                    oDSUser = mtx1.Columns.Item(j).DataBind.Alias;
                    if ((mtx1.Columns.Item(j).DataBind.Alias.Substring(0, 5) == "DSTda") && (oForm.DataSources.UserDataSources.Item(oDSUser).ValueEx.Trim() == "Y"))
                        sAux = sAux + "1";
                    else
                        sAux = sAux + "0";
                }
                oMtx.Add(sAux);
            }

            for (int i = 0; i < oMtx.Count; i++)
                for (int j = i+1; j < oMtx.Count; j++)
                {
//                    OutLog("i: " + i.ToString() + "  j: " + j.ToString());
//                    OutLog("i -> " + oMtx[i] + "   j -> " + oMtx[j]);
                    if (oMtx[i] == oMtx[j])
                        throw new Exception("Cluster duplicados linea " + (i+1).ToString() + " y " + (j+1).ToString());
                }
        }
    }
}
