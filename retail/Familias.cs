using System;
using System.Collections.Generic;
using SAPbouiCOM;
using SAPbobsCOM;
using VisualD.GlobalVid;
using VisualD.SBOFunctions;
using VisualD.MasterDataMatrixForm;
using VisualD.SBOGeneralService;
using VisualD.MultiFunctions;
using VisualD.vkFormInterface;


namespace VID_Retail.Familias
{
    public class TFamilias : TMasterDataMatrixForm, IvkFormInterface
    {
        public TFamilias()
        {
        }

        private SAPbouiCOM.Form oForm = null;
        private SAPbobsCOM.Recordset oRS;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool oResult;
            String oSql;

            FormFileName = "Familia.srf";
            TableName = "@VIDR_FAMILIA";
            MatrixName = "mtx0";
            UdoName = "VIDR_FAMILIA";
            MsgUpdate = "¿Desea actualizar la definición de familias?";
            ColumnsNames = new String[] { "Code", "Name", "Grupo" };

            ListaMx.Add("Code      , r , tx");
            ListaMx.Add("Name      , r , tx");
            ListaMx.Add("Grupo     , r , tx");

            VID_DelRow = true;
            VID_DelRowOK = true;

            try
            {
                oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);

                oForm = FSBOApp.Forms.Item(uid);
                oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                oSql = GlobalSettings.RunningUnderSQLServer ?
                       "Select Code, Name from [@VIDR_GRUPO] order by Name" :
                       "Select \"Code\", \"Name\" from \"@VIDR_GRUPO\" order by \"Name\" ";
                oRS.DoQuery(oSql);
                FSBOf.FillComboMtx(((Matrix)(oForm.Items.Item("mtx0").Specific)).Columns.Item("Grupo"), ref oRS, false);
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(e.Message + " - " + e.StackTrace);
                oResult = false;
            }
            finally
            {
                if (oForm != null)
                    oForm.Visible = true;
            }

            return (oResult);
        }

        public new void FormEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, ref Boolean BubbleEvent)
        {
            base.FormEvent(FormUID, ref pVal, ref BubbleEvent);

            try
            {


            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //OutLog("FormEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }

    }
}
