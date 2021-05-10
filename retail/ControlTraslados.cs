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


namespace VID_Retail.ControlTraslados
{
    class TControlTraslados : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.Application R_application;
        SAPbobsCOM.Company R_company;
        CSBOFunctions R_sboFunctions;
        TGlobalVid R_GlobalSettings;

        public TControlTraslados()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;
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
                    FSBOf.LoadForm(xmlPath, "ControlTrasladosBodegas.srf", uid);
                    EnableCrystal = false;
                    VID_DelRow = true;
                    VID_DelRowOK = true;

                    oForm = FSBOApp.Forms.Item(uid);
                    oForm.AutoManaged = true;
                    oForm.SupportedModes = -1;             // afm_All
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oForm.PaneLevel = 1;

                    oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                    
                    oForm.Items.Item("mtx1").AffectsFormMode = true;
                    mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);

                    oForm.DataSources.UserDataSources.Add("DSOrigen", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSDestino", BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("DSRegla", BoDataType.dt_SHORT_TEXT, 100);
                    ((ComboBox)oForm.Items.Item("Origen").Specific).DataBind.SetBound(true, "", "DSOrigen");
                    mtx1.Columns.Item("Destino").DataBind.SetBound(true, "", "DSDestino");
                    mtx1.Columns.Item("Regla").DataBind.SetBound(true, "", "DSRegla");

                    oSql = GlobalSettings.RunningUnderSQLServer ?
                           "Select WhsCode Code, WhsName Name from OWHS order by WhsCode" :
                           "Select \"WhsCode\" \"Code\", \"WhsName\" \"Name\" from \"OWHS\" order by \"WhsCode\"";
                    oRS.DoQuery(oSql);
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("Origen").Specific)), ref oRS, false);
                    oRS.MoveFirst();
                    FSBOf.FillComboMtx(mtx1.Columns.Item("Destino"), ref oRS, true);

                    mtx1.Columns.Item("Regla").ValidValues.Add("", "");
                    mtx1.Columns.Item("Regla").ValidValues.Add("BQ", "Bloqueado Siempre");
                    mtx1.Columns.Item("Regla").ValidValues.Add("HB", "Habilitado Siempre");
                    mtx1.Columns.Item("Regla").ValidValues.Add("BG", "Bloqueado Cod. Gerona");
                    mtx1.Columns.Item("Regla").ValidValues.Add("HG", "Habilitado Cod. Gerona");
                    mtx1.Columns.Item("Regla").ValidValues.Add("BR", "Bloqueado Cod. C Royal");
                    mtx1.Columns.Item("Regla").ValidValues.Add("HR", "Habilitado Cod. C Royal");

                    ((ComboBox)(oForm.Items.Item("Origen").Specific)).ValidValues.Item(0);
                    mtx1.Columns.Item("Destino").ValidValues.Item(0);

                    //FillMtx();
                    mtx1.AddRow(1, -1);

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
                oForm.Freeze(true);
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "1") && (!pVal.BeforeAction) && ((pVal.FormMode == (Int32)BoFormMode.fm_ADD_MODE) || (pVal.FormMode == (Int32)BoFormMode.fm_UPDATE_MODE)))
                        {
                            ValidarDatos();
                            insertarDatos();
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
                        if ((pVal.ItemUID == "mtx1") && (pVal.ColUID == "Destino") && (!pVal.BeforeAction))
                            ValidateDestino(pVal.Row);
                        else if ((pVal.ItemUID == "mtx1") && (pVal.ColUID == "Regla") && (!pVal.BeforeAction))
                            ValidateRegla(pVal.Row);
                        else if ((pVal.ItemUID == "Origen") && (pVal.BeforeAction) && (oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            BubbleEvent = false;
                            if (1 == FSBOApp.MessageBox("Perderá modificaciones realizadas. ¿Desea continuar?", 2, "Ok", "Cancel"))
                                BubbleEvent = true;
                        }
                        else if ((pVal.ItemUID == "Origen") && (!pVal.BeforeAction))
                            FillMtx(oForm.DataSources.UserDataSources.Item("DSOrigen").ValueEx.Trim());
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

        private void FillMtx(String oWhsOrigen)
        {
            String oSql;
            SAPbouiCOM.Matrix mtx1;

            try
            {
                oForm.Freeze(true);
                mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);

                oSql = GlobalSettings.RunningUnderSQLServer ?
                       "Select U_WhsCodeO, U_WhsCodeD, U_Regla    " +
                       "  from [@VIDR_TRASLADO]                   " +
                       " where U_WhsCode = '{0}'                  " :
                       "Select \"U_WhsCodeO\", \"U_WhsCodeD\", \"U_Regla\"     " +
                       "  from \"@VIDR_TRASLADO\"                              " +
                       " where \"U_WhsCodeO\" = '{0}'                          " ;
                oRS.DoQuery(String.Format(oSql, oWhsOrigen));

                mtx1.Clear();
                while (!oRS.EoF)
                {
                    oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx = (String)oRS.Fields.Item("U_WhsCodeD").Value;
                    oForm.DataSources.UserDataSources.Item("DSRegla").ValueEx = (String)oRS.Fields.Item("U_Regla").Value;
                    mtx1.AddRow(1, -1);
                    oRS.MoveNext();
                }

                oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("DSRegla").ValueEx = "";
                mtx1.AddRow(1, -1);

                oForm.Mode = BoFormMode.fm_OK_MODE;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void ValidateDestino(Int32 oRow)
        {
            SAPbouiCOM.Matrix mtx1;
            String Destino = "";

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            mtx1.GetLineData(oRow);
            Destino = oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx.Trim();

            if (Destino == "")
                return;
            if (oForm.DataSources.UserDataSources.Item("DSOrigen").ValueEx.Trim() == "")
            {
                FSBOApp.StatusBar.SetText("Origen no definido.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return;
            }
            if (oForm.DataSources.UserDataSources.Item("DSOrigen").ValueEx.Trim() == Destino)
            {
                FSBOApp.StatusBar.SetText("Origen no puede ser igual a Destino.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return;
            }

            for (Int32 i = 1; i <= mtx1.RowCount; i++)
            {
                if (i == oRow)
                    continue;
                mtx1.GetLineData(i);
                if (oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx == Destino)
                {
                    FSBOApp.StatusBar.SetText("Destino ya existe en linea : " + i.ToString() + ".", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }
            }

            if (oRow == mtx1.RowCount)
            {
                oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("DSRegla").ValueEx = "";
                mtx1.AddRow(1, -1);
            }
        }

        private void ValidateRegla(Int32 oRow)
        {
            SAPbouiCOM.Matrix mtx1;
            String Regla = "";

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            mtx1.GetLineData(oRow);
            Regla = oForm.DataSources.UserDataSources.Item("DSRegla").ValueEx.Trim();

            if (Regla == "")
                FSBOApp.StatusBar.SetText("Debe seleccionar una regla de operación.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        }

        private void ValidarDatos()
        {
            SAPbouiCOM.Matrix mtx1;
            String Destino = "";
            String Origen = "";
            String Regla = "";

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            Origen = oForm.DataSources.UserDataSources.Item("DSOrigen").ValueEx.Trim();

            if (Origen == "")
                throw new Exception("Origen no valido");

            for (Int32 i = 1; i <= mtx1.RowCount; i++)
            {
                mtx1.GetLineData(i);
                Destino = oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx;
                Regla = oForm.DataSources.UserDataSources.Item("DSRegla").ValueEx;
                if (Destino == Origen)
                    throw new Exception("Origen y Destino no pueden ser iguales, linea " + i.ToString());
                if ((Destino != "") && (Regla == ""))
                    throw new Exception("Regla no definida en linea " + i.ToString());
                for (Int32 j = 1; j <= mtx1.RowCount; j++)
                {
                    if (i == j)
                        continue;
                    mtx1.GetLineData(j);
                    if (Destino == oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx)
                        throw new Exception("Destino duplicado en linea " + i.ToString());
                }
            }
        }

        private void insertarDatos()
        {
            SAPbouiCOM.Matrix mtx1;
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.CompanyService oCompService;
            String oCode;
            String oSql;
            String Destino = "";
            String Origen = "";

            oCompService = FCmpny.GetCompanyService();
            oGeneralService = oCompService.GetGeneralService("VIDR_TRASLADO");
            Origen = oForm.DataSources.UserDataSources.Item("DSOrigen").ValueEx.Trim();

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select Code from [@VIDR_TRASLADO] where Code = '{0}'" :
                   "Select \"Code\" from \"@VIDR_TRASLADO\" where \"Code\" like '{0}-%'";
            oRS.DoQuery(String.Format(oSql, Origen));
            while (!oRS.EoF)
            {
                oCode = (String)oRS.Fields.Item("Code").Value;
                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", oCode);
                oGeneralService.Delete(oGeneralParams);
                oRS.MoveNext();
            }

            mtx1 = (Matrix)(oForm.Items.Item("mtx1").Specific);
            for (int i = 1; i <= mtx1.RowCount; i++)
            {
                mtx1.GetLineData(i);
                Destino = oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx.Trim();
                if (Destino == "")
                    continue;

                oCode = Origen + "-" + Destino;
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                oGeneralData.SetProperty("Code", oCode);
                oGeneralData.SetProperty("Name", oCode);
                oGeneralData.SetProperty("U_WhsCodeO", oForm.DataSources.UserDataSources.Item("DSOrigen").ValueEx.Trim());
                oGeneralData.SetProperty("U_WhsCodeD", oForm.DataSources.UserDataSources.Item("DSDestino").ValueEx.Trim());
                oGeneralData.SetProperty("U_Regla", oForm.DataSources.UserDataSources.Item("DSRegla").ValueEx.Trim());

                oGeneralService.Add(oGeneralData);
            }
            oGeneralService = null;
        }
    }
}
