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


namespace VID_Retail.Tiendas
{
    public class TTiendas : TMasterDataMatrixForm, IvkFormInterface
    {
        public TTiendas()
        {
        }

        private SAPbouiCOM.Form oForm = null;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            bool oResult;

            FormFileName = "Tiendas.srf";
            TableName = "@VIDR_TIENDA";
            MatrixName = "mtx0";
            UdoName = "VIDR_TIENDA";
            MsgUpdate = "¿Desea actualizar la definición de tiendas?";
            ColumnsNames = new String[] { "Code", "Name" };

            ListaMx.Add("Code     , r , tx");
            ListaMx.Add("Name     , r , tx");

            VID_DelRow = true;
            VID_DelRowOK = true;

            try
            {
                oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
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
