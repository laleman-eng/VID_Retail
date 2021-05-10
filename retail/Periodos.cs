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


namespace VID_Retail.Periodos
{
    public class TPeriodos : TMasterDataMatrixForm, IvkFormInterface
    {
        public TPeriodos()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            SAPbouiCOM.Matrix omtx;
            SAPbouiCOM.DBDataSource oDBDS;
            bool oResult;

            FormFileName = "Periodos.srf";
            TableName = "@VIDR_PERIODO";
            MatrixName = "mtx0";
            UdoName = "VIDR_PERIODO";
            MsgUpdate = "¿Desea actualizar la definición de periodos?";
            ColumnsNames = new String[] { "Code", "Name", "Fecini", "Fecfin" };

            ListaMx.Add("Code   , r , tx");
            ListaMx.Add("Name   , r , tx");
            ListaMx.Add("Fecini , r , tx");
            ListaMx.Add("Fecfin , r , tx");

            VID_DelRow = true;
            VID_DelRowOK = true;

            try
            {
                oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);

                oForm = FSBOApp.Forms.Item(uid);
                oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                omtx = (Matrix)(oForm.Items.Item("mtx0").Specific);
                oDBDS = oForm.DataSources.DBDataSources.Item("@VIDR_PERIODO");

                //oForm.Freeze(true);
                fillMatrix(omtx, oDBDS);
                //oForm.Freeze(false);
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

        private void fillMatrix(Matrix oMtx, DBDataSource oDS)
        {
            String oSql;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select Code, Name, U_Fecini, U_Fecfin from [@VIDR_PERIODO] order by U_Fecini" :
                   "Select \"Code\", \"Name\", \"U_Fecini\", \"U_Fecfin\" from \"@VIDR_PERIODO\" order by \"U_Fecini\"";
            oRS.DoQuery(oSql);

            oMtx.Clear();
            oDS.Clear();
            oDS.InsertRecord(0);
            while (!oRS.EoF)
            {
                oDS.SetValue("Code", 0, (String)oRS.Fields.Item("Code").Value);
                oDS.SetValue("Name", 0, (String)oRS.Fields.Item("Name").Value);
                oDS.SetValue("U_Fecini", 0, ((DateTime)oRS.Fields.Item("U_Fecini").Value).ToString("yyyyMMdd"));
                oDS.SetValue("U_Fecfin", 0, ((DateTime)oRS.Fields.Item("U_Fecfin").Value).ToString("yyyyMMdd"));
                oMtx.AddRow(1);
                oRS.MoveNext();
            }
            oDS.SetValue("Code", 0, "");
            oDS.SetValue("Name", 0, "");
            oDS.SetValue("U_Fecini", 0, "");
            oDS.SetValue("U_Fecfin", 0, "");
            oMtx.AddRow(1);
        }

        private bool ValidateFecha(Int32 oCurrentRow, String oCol)
        {
            // Validacion desactivada
            return true; 

            Matrix omtx = (Matrix)(oForm.Items.Item("mtx0").Specific);
            DBDataSource oDBDS = oForm.DataSources.DBDataSources.Item("@VIDR_PERIODO");
            String sFecha;
            DateTime fecha;
            DateTime fini;
            DateTime ffin;

            oDBDS.Clear();
            oDBDS.InsertRecord(0);

            omtx.GetLineData(oCurrentRow);
            sFecha = ((EditText)omtx.Columns.Item(oCol).Cells.Item(oCurrentRow).Specific).Value;
            if (sFecha == "")
                return true;
            fecha = DateTime.ParseExact(sFecha, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

            for (int i = 1; i < omtx.RowCount; i++)
            {
                if (i == oCurrentRow)
                    continue;

                omtx.GetLineData(i);
                if (oDBDS.GetValue("Code", 0).Trim() == "Normal")
                    continue;

                fini = DateTime.ParseExact(oDBDS.GetValue("U_Fecini", 0), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                ffin = DateTime.ParseExact(oDBDS.GetValue("U_Fecfin", 0), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                if ((fecha >= fini) && (fecha <= ffin) && (oCol == "Fecini"))
                {
                    FSBOApp.StatusBar.SetText("Fecha inicial del periodo no puede ir dentro de otro periodo.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if ((fecha >= fini) && (fecha <= ffin) && (oCol == "Fecfin"))
                {
                    FSBOApp.StatusBar.SetText("Fecha final del periodo no puede ir dentro de otro periodo.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
            }
            return true;
        }

        public new void FormEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, ref Boolean BubbleEvent)
        {

            try
            {
                base.FormEvent(FormUID, ref pVal, ref BubbleEvent);
                SAPbouiCOM.Form oForm = FSBOApp.Forms.Item(FormUID);

                switch (pVal.EventType)
                {
                    case BoEventTypes.et_VALIDATE:
                        if ((pVal.ItemUID == "mtx0") && (pVal.BeforeAction))
                        {
                            if (pVal.ColUID == "Fecini") 
                                BubbleEvent = ValidateFecha(pVal.Row, pVal.ColUID);
                            if (pVal.ColUID == "Fecfin")
                                BubbleEvent = ValidateFecha(pVal.Row, pVal.ColUID);
                        }
                        break;
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "1") && (!pVal.BeforeAction))
                        {
                        }
                        break;
                    case BoEventTypes.et_FORM_RESIZE:
                        if (pVal.BeforeAction)
                            oForm.Freeze(true);
                        else if (!pVal.BeforeAction)
                            oForm.Freeze(false);
                        break;
                }
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText(e.Message + " ** Trace: " + e.StackTrace, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog("FormEvent: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        }
    }
}
