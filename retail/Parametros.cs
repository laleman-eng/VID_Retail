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
using VID_Retail.Utils;


namespace VID_Retail.Parametros
{
    class TParametros : TvkBaseForm, IvkFormInterface
    {
        public TParametros()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        private TUtils oUtil = null; 

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;

            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                try
                {
                    FSBOf.LoadForm(xmlPath, "Parametros.srf", uid);
                    EnableCrystal = false;

                    oForm = FSBOApp.Forms.Item(uid);
                    oForm.AutoManaged = true;
                    oForm.PaneLevel = 1;
                    oForm.Items.Item("tab2").Click();

                    oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                    oForm.DataSources.UserDataSources.Add("DSUnPrcLst", BoDataType.dt_SHORT_TEXT, 1);
                    ((CheckBox)(oForm.Items.Item("UnPriceLst").Specific)).DataBind.SetBound(true, "", "DSUnPrcLst");


                    ((ComboBox)(oForm.Items.Item("PriceLst").Specific)).ValidValues.Add("-10", "");
                    oSql = GlobalSettings.RunningUnderSQLServer ?
                            "Select '-10' Code, '' Name " +
                            " UNION ALL " +
                            "Select ListNum Code, ListName Name from OPLN order by Name " :
                            "Select '-10' \"Code\", '' \"Name\" from Dummy " +
                            " UNION ALL " +
                            "Select TO_ALPHANUM(\"ListNum\") \"Code\", \"ListName\" \"Name\" from OPLN order by \"Name\" ";
                    oRS.DoQuery(oSql);
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("PriceLst").Specific)), ref oRS, false);


                    oSql = GlobalSettings.RunningUnderSQLServer ?
                           "Select Count(*) Cant, IsNull(MAX(U_PriceLst), -10) PriceLst from [@VIDR_PARAM] " :
                           "Select Count(*) \"Cant\", IfNull(MAX(\"U_PriceLst\") , -10) \"PriceLst\" from \"@VIDR_PARAM\" ";
                    oRS.DoQuery(oSql);

                    if (((Int32)oRS.Fields.Item("PriceLst").Value) == -10)
                        oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx = "N";
                    else
                        oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx = "Y";

                    if (((Int32)oRS.Fields.Item("Cant").Value) == 0)
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

                    oSql = GlobalSettings.RunningUnderSQLServer ?
                           "Select WhsCode Code, WhsName Name from OWHS order by WhsCode" :
                           "Select \"WhsCode\" \"Code\", \"WhsName\" \"Name\" from \"OWHS\" order by \"WhsCode\"";
                    oRS.DoQuery(oSql);
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("WhsCodLF").Specific)), ref oRS, false);
                    oRS.MoveFirst();
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("WhsCodCD").Specific)), ref oRS, false);
                    oRS.MoveFirst();
                    FSBOf.FillCombo(((ComboBox)(oForm.Items.Item("WhsCodTR").Specific)), ref oRS, false);

                    oUtil = new TUtils(ref oRS, ref _GlobalSettings, false);

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
            SAPbouiCOM.DBDataSource oDS;
            String sVal;

            try
            {
                oForm.Freeze(true);
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "tab1") && (!pVal.BeforeAction))
                            oForm.PaneLevel = 1;
                        else if ((pVal.ItemUID == "tab2") && (!pVal.BeforeAction))
                            oForm.PaneLevel = 2;
                        else if ((pVal.ItemUID == "UnPriceLst") && (pVal.BeforeAction))
                        {
                            sVal = oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx;
                            if (sVal == "N")
                                BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea cambiar operación a: Lista de Precios Base Única? - Verifique que no existan precios manuales en las listas dependientes", 2, "Ok", "Cancel"));
                            else
                                BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea cambiar operación a: multiples Listas de Precios Base? - Esto le permite manejar precios manuales por lista", 2, "Ok", "Cancel"));
                        }
                        else if ((pVal.ItemUID == "UnPriceLst") && (!pVal.BeforeAction))
                        {
                            sVal = oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx;
                            if (sVal == "Y")
                            {
                                oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx = "N";
                                oForm.DataSources.DBDataSources.Item("@VIDR_PARAM").SetValue("U_PriceLst", 0, "-10");
                            }
                            else
                                oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx = "Y";
                            oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                        /*
                        else if ((pVal.ItemUID == "btnConnect") && (!pVal.BeforeAction))
                        {
                            if (oUtil != null)
                            {
                                if (oUtil.SetOtherSBOCompany())
                                {
                                    FSBOApp.StatusBar.SetText("Conexión exitosa....", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    oUtil.disconnectOtherSBOCompany();
                                }
                                else
                                    FSBOApp.StatusBar.SetText("No se pudo establecer conexión...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                        }
                        */
                        else if ((pVal.ItemUID == "1") && (pVal.BeforeAction) && ((oForm.Mode == BoFormMode.fm_ADD_MODE) || (oForm.Mode == BoFormMode.fm_UPDATE_MODE)))
                        {
                            oDS = oForm.DataSources.DBDataSources.Item("@VIDR_PARAM");
                            oDS.SetValue("Code", 0, "1");
                            BubbleEvent = ValidateData();
                        }
                        else if ((pVal.ItemUID == "1") && (!pVal.BeforeAction) && (oForm.Mode == BoFormMode.fm_ADD_MODE) && (pVal.ActionSuccess))
                        {
                        }
                        break;
                    case BoEventTypes.et_COMBO_SELECT:
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

        public new void FormDataEvent(ref BusinessObjectInfo oBusinessObjectInfo, ref bool BubbleEvent)
        {
            base.FormDataEvent(ref oBusinessObjectInfo, ref BubbleEvent);

            switch (oBusinessObjectInfo.EventType)
            {
                case BoEventTypes.et_FORM_DATA_LOAD:
                    break;
                case BoEventTypes.et_FORM_DATA_ADD:
                    break;
            }
        }

        private bool ValidateData()
        {
            if ((oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx == "Y") && (oForm.DataSources.DBDataSources.Item("@VIDR_PARAM").GetValue("U_PriceLst", 0) == "-10"))
            {
                FSBOApp.StatusBar.SetText("Debe seleccionar una lista de precios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
            if ((oForm.DataSources.UserDataSources.Item("DSUnPrcLst").ValueEx == "N") && (oForm.DataSources.DBDataSources.Item("@VIDR_PARAM").GetValue("U_PriceLst", 0) != "-10"))
            {
                FSBOApp.StatusBar.SetText("La lista de precios deba estar vacia...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
            return true;
        }

    }
}
