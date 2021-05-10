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
using VID_Retail.Password;


namespace VID_Retail.NCVentaRelacionada
{
    class TNCVentaRelacionada : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.Application R_application;
        SAPbobsCOM.Company R_company;
        CSBOFunctions R_sboFunctions;
        TGlobalVid R_GlobalSettings;
        List<object> R_oForms;
        SAPbouiCOM.StaticText oTxt;
        SAPbouiCOM.Button oBtn;
        TUtils oUtil;

        public TNCVentaRelacionada(List<object> oForms)
        {
            R_oForms = oForms;
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        private String CCGerona = "";

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;
            SAPbouiCOM.Item oItm;
            SAPbouiCOM.Item oItmRef;
            IvkFormInterface newForm;

            R_application = application;
            R_company = company;
            R_sboFunctions = sboFunctions;
            R_GlobalSettings = _GlobalSettings;

            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                EnableCrystal = false;
                oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                oForm = FSBOApp.Forms.Item(uid);

                oItmRef = sboFunctions.getFormItem(oForm, "70");

                // Estado
                oItm = oForm.Items.Add("Txt_Estado", BoFormItemTypes.it_STATIC);
                oItm.Visible = true;
                oItm.Left = oItmRef.Left;
                oItm.Top = oItmRef.Top + oItmRef.Height + 2;
                oItm.Height = oItmRef.Height;
                oItm.Width = oItmRef.Width * 2;
                oItm.TextStyle = (Int32)BoTextStyle.ts_BOLD;
                oTxt = (StaticText)(oItm.Specific);
                oTxt.Caption = "";

                // Boton aplicar
                oItm = oForm.Items.Add("btnInsCR", BoFormItemTypes.it_BUTTON);
                oItm.Visible = false;
                oItm.Left = oItmRef.Left;
                oItm.Top = oItmRef.Top + (oItmRef.Height + 2) * 2;
                oItm.Height = oItmRef.Height;
                oItm.Width = oItmRef.Width;
                oItm.TextStyle = (Int32)BoTextStyle.ts_BOLD;
                oBtn = (Button)(oItm.Specific);
                oBtn.Caption = "Insertar NC en Gerona";

                oUtil = new TUtils(ref oRS, ref R_GlobalSettings, true);

                oSql = GlobalSettings.RunningUnderSQLServer ?
                      "Update " :
                      "Select \"U_CCGerona\" from \"@VIDR_PARAM\" ";
                oRS.DoQuery(String.Format(oSql));
                CCGerona = ((String)(oRS.Fields.Item("U_CCGerona").Value)).Trim();

                // Set Password Gerona
                if (_GlobalSettings.oCompanyVentaRelacionada != null)
                    if (_GlobalSettings.oCompanyVentaRelacionada.Connected)
                        return (oResult);

                newForm = (IvkFormInterface)(new TPassword());
                if (newForm != null)
                {
                    if (newForm.InitForm(R_sboFunctions.generateFormId(GlobalSettings.SBOSpaceName, GlobalSettings), "forms\\", ref  R_application, ref  R_company, ref R_sboFunctions, ref R_GlobalSettings))
                    { R_oForms.Add(newForm); }
                    else
                    {
                        R_application.Forms.Item(newForm.getFormId()).Close();
                        newForm = null;
                    }
                }

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
                if (oTxt.Caption != "")
                    if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                    {
                        oTxt.Caption = "";
                        oBtn.Item.Visible = false;
                    }

                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        if ((pVal.ItemUID == "btnInsCR") && (pVal.BeforeAction))
                            BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea enviar NC a venta relacionada?", 2, "Ok", "Cancel"));
                        else if ((pVal.ItemUID == "btnInsCR") && (!pVal.BeforeAction))
                        {
                            if ("" == oForm.DataSources.DBDataSources.Item("ORIN").GetValue("FolioNum", 0).Trim())
                            {
                                FSBOApp.StatusBar.SetText("Documento sin folio.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                break;
                            }
                            if ("" == CreateNCGerona(oForm.DataSources.DBDataSources.Item("ORIN"), oForm.DataSources.DBDataSources.Item("RIN1")))
                            {
                                R_application.MessageBox("Nota de Crédito ingresada", 1, "Ok", "", "");
                                oTxt.Caption = "NC en Gerona: " + oForm.DataSources.DBDataSources.Item("ORIN").GetValue("U_VR_DocRel", 0).Trim();
                                oBtn.Item.Visible = false;
                            }
                        }
                        break;
                    case BoEventTypes.et_FORM_RESIZE:
                        if (!pVal.BeforeAction)
                            ;
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

        public new void FormDataEvent(ref BusinessObjectInfo oBusinessObjectInfo, ref bool BubbleEvent)
        {
            base.FormDataEvent(ref oBusinessObjectInfo, ref BubbleEvent);

            try
            {
                switch (oBusinessObjectInfo.EventType)
                {
                    case BoEventTypes.et_FORM_DATA_LOAD:
                        if (!oBusinessObjectInfo.BeforeAction)
                        {
                            if (oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocType", 0).Trim() != "I")
                            {
                                oTxt.Caption = "";
                                oBtn.Item.Visible = false;
                            }
                            else if (oForm.DataSources.DBDataSources.Item("ORIN").GetValue("CardCode", 0).Trim() != GlobalSettings.SNGerona)
                            {
                                oTxt.Caption = "";
                                oBtn.Item.Visible = false;
                            }
                            else if (oForm.DataSources.DBDataSources.Item("ORIN").GetValue("U_VR_DocRel", 0).Trim() == "")
                            {
                                oTxt.Caption = "Documento no enviado";
                                oBtn.Item.Visible = true;
                            }
                            else
                            {
                                oTxt.Caption = "DocNum en Gerona: " + oForm.DataSources.DBDataSources.Item("ORIN").GetValue("U_VR_DocRel", 0).Trim();
                                oBtn.Item.Visible = false;
                            }
                        }
                        break;
                    case BoEventTypes.et_FORM_DATA_ADD:
                        if (!oBusinessObjectInfo.BeforeAction)
                        {
                            if (oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocType", 0).Trim() == "I")
                                if (oForm.DataSources.DBDataSources.Item("ORIN").GetValue("CANCELED", 0).Trim() == "N")
                                    if (oForm.DataSources.DBDataSources.Item("ORIN").GetValue("CardCode", 0).Trim() == GlobalSettings.SNGerona)
                                        resetDocRel(oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0));
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

        private void resetDocRel(String oDocEntry)
        {
            String oSql;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                  "Update ORIN set U_VR_DocRel = NULL where DocEntry = {0} " :
                  "Update ORIN set \"U_VR_DocRel\" = NULL where \"DocEntry\" = {0} ";

            oRS.DoQuery(String.Format(oSql, oDocEntry.Trim()));
        }

        private String CreateNCGerona(SAPbouiCOM.DBDataSource DSHead, SAPbouiCOM.DBDataSource DSDet)
        {
            String sErr = "";
            Int32 oDocEntry;
            Int32 VentaRelacionadaDocEntry;
            Int32 oLine;
            Int32 VR_OVOrig = 0;
            String oDocNewKey;

            if (!oUtil.SetOtherSBOCompany())
                throw new Exception("Error en conexión a base de datos de venta relacionada");

            try
            {
                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)GlobalSettings.oCompanyVentaRelacionada.GetBusinessObject(BoObjectTypes.oPurchaseCreditNotes);

                oDoc.CardCode = GlobalSettings.SN_CR;
                oDoc.HandWritten = BoYesNoEnum.tNO;
                oDoc.DocDate = DateTime.Today;
                oDoc.DocDueDate = DateTime.Today;
                oDoc.TaxDate = DateTime.Today;
                oDoc.Address = DSHead.GetValue("Address", 0);
                oDoc.Address2 = DSHead.GetValue("Address", 0);
                oDoc.Comments = "NC Casa Royal relacionada - " + DSHead.GetValue("DocNum", 0) + "  " + DSHead.GetValue("Comments", 0);
                oDoc.FolioNumber = Int32.Parse(DSHead.GetValue("FolioNum", 0));
                oDoc.FolioPrefixString = DSHead.GetValue("FolioPref", 0);
                if (DSHead.GetValue("U_VR_OVOrig", 0).Trim() != "")
                    VR_OVOrig = Int32.Parse(DSHead.GetValue("U_VR_OVOrig", 0));
                oDoc.UserFields.Fields.Item("U_VR_OVOrig").Value = VR_OVOrig;
                oDoc.UserFields.Fields.Item("U_VR_DocRel").Value = Int32.Parse(DSHead.GetValue("DocNum", 0));
                oDocEntry = Int32.Parse(DSHead.GetValue("DocEntry", 0));

                oLine = -1;
                for (int i = 0; i <= DSDet.Size - 1; i++)
                {
                    if (DSDet.GetValue("ItemCode", i).Trim() == "")
                        continue;

                    oLine++;
                    if (oLine > 0)
                        oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(oLine);

                    oDoc.Lines.ItemCode = DSDet.GetValue("ItemCode", i);
                    oDoc.Lines.Quantity = Double.Parse(DSDet.GetValue("Quantity", i), System.Globalization.CultureInfo.InvariantCulture);
                    oDoc.Lines.Currency = DSDet.GetValue("Currency", i);
                    oDoc.Lines.PriceAfterVAT = Double.Parse(DSDet.GetValue("PriceAfVAT", i), System.Globalization.CultureInfo.InvariantCulture);
                    oDoc.Lines.CostingCode = CCGerona;
                    //oDoc.Lines.Price = Double.Parse(DSDet.GetValue("Price", i), System.Globalization.CultureInfo.InvariantCulture);
                }

                if (GlobalSettings.Debug)
                    oDoc.SaveToFile("oPurchaseCreditNote.xml");
                int nErr = oDoc.Add();
                if (nErr != 0)
                {
                    GlobalSettings.oCompanyVentaRelacionada.GetLastError(out nErr, out sErr);
                    throw new Exception(sErr);
                }
                else
                {   // Update este documento local
                    oDocNewKey = GlobalSettings.oCompanyVentaRelacionada.GetNewObjectKey();
                    VentaRelacionadaDocEntry = Int32.Parse(oDocNewKey.Trim());
                    oDoc.GetByKey(VentaRelacionadaDocEntry);

                    SAPbobsCOM.Documents esteDoc = (SAPbobsCOM.Documents)FCmpny.GetBusinessObject(BoObjectTypes.oCreditNotes);
                    if (!esteDoc.GetByKey(oDocEntry))
                    {
                        FCmpny.GetLastError(out nErr, out sErr);
                        throw new Exception(sErr);
                    }
                    else
                    {
                        esteDoc.UserFields.Fields.Item("U_VR_OVOrig").Value = VR_OVOrig;
                        esteDoc.UserFields.Fields.Item("U_VR_DocRel").Value = oDoc.DocNum;
                        nErr = esteDoc.Update();
                        if (nErr != 0)
                        {
                            FCmpny.GetLastError(out nErr, out sErr);
                            throw new Exception(sErr);
                        }
                    }
                }

                return sErr;
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error: " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(e.Message + " - " + e.StackTrace);
                return e.Message;
            }
            finally
            {
                oUtil.disconnectOtherSBOCompany();
            }
        }
    }
}
