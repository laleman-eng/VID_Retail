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

namespace VID_Retail.Password
{
    class TPassword : TvkBaseForm, IvkFormInterface
    {
        public TPassword()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        TUtils oUtil;


        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);

            oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

            try
            {
                FSBOf.LoadForm(xmlPath, "Password.srf", uid);
                EnableCrystal = false;
                ModalForm = true;

                oForm = FSBOApp.Forms.Item(uid);

                oForm.DataSources.UserDataSources.Add("DSUser", BoDataType.dt_SHORT_TEXT, 20);
                oForm.DataSources.UserDataSources.Add("DSPw"  , BoDataType.dt_SHORT_TEXT, 20);
                ((EditText)(oForm.Items.Item("Usuario").Specific)).DataBind.SetBound(true, "", "DSUser");
                ((EditText)(oForm.Items.Item("Password").Specific)).DataBind.SetBound(true, "", "DSPw");

                oForm.DataSources.UserDataSources.Item("DSUser").ValueEx = FCmpny.UserName;

                oForm.Mode = BoFormMode.fm_OK_MODE;
                oForm.SupportedModes = (Int32)BoAutoFormMode.afm_Ok;

                oUtil = new TUtils(ref oRS, ref _GlobalSettings, true);

                return (oResult);
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(e.Message + " - " + e.StackTrace);
                return (false);
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
                        if ((pVal.ItemUID == "1") && (pVal.BeforeAction))
                        {
                            // Test pass
                            GlobalSettings.Pw = ((EditText)(oForm.Items.Item("Password").Specific)).String.Trim();
                            BubbleEvent = oUtil.SetOtherSBOCompany();
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
    }
}
