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

namespace VID_Retail.TransferenciaStockDev
{
    public class regPurchaseOrder
    {
        public String WhsCode { get; set; }
        public String WhsName { get; set; }
        public Int32 LineNum { get; set; }
        public String CardCode { get; set; }
        public String ShipToDef { get; set; }
        public String ItemCode { get; set; }
        public Double Quantity { get; set; }
        public Double Price { get; set; }
        public String TrasladoKey { get; set; }
    }

    class TTransferenciaStockDev : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.StaticText oTxt;
        SAPbouiCOM.Button oBtn;

        public TTransferenciaStockDev()
        {
        }

        private SAPbobsCOM.Recordset oRS;
        private SAPbouiCOM.Form oForm = null;
        private String fromWhs;
        private String fromWhsName;
        private String toWhs;
        private String toWhsName;

        public new bool InitForm(string uid, string xmlPath, ref SAPbouiCOM.Application application, ref SAPbobsCOM.Company company, ref CSBOFunctions sboFunctions, ref TGlobalVid _GlobalSettings)
        {
            String oSql;
            SAPbouiCOM.Item oItm;
            SAPbouiCOM.Item oItmRef;
            SAPbouiCOM.Item oItmRefTx;

            bool oResult = base.InitForm(uid, xmlPath, ref application, ref company, ref sboFunctions, ref _GlobalSettings);
            try
            {
                EnableCrystal = false;
                oRS = (Recordset)(FCmpny.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                oForm = FSBOApp.Forms.Item(uid);

                oItmRef = sboFunctions.getFormItem(oForm, "9");
                oItmRefTx = sboFunctions.getFormItem(oForm, "10");

                // Estado
                oItm = oForm.Items.Add("Txt_Estado", BoFormItemTypes.it_STATIC);
                oItm.Visible = true;
                oItm.Left = oItmRefTx.Left;
                oItm.Top = oItmRef.Top + oItmRef.Height + 2;
                oItm.Height = oItmRefTx.Height;
                oItm.Width = oItmRefTx.Width;
                oItm.TextStyle = (Int32)BoTextStyle.ts_BOLD;
                oTxt = (StaticText)(oItm.Specific);
                oTxt.Caption = "";

                // Boton aplicar
                oItm = oForm.Items.Add("btnInsCD", BoFormItemTypes.it_BUTTON);
                oItm.Visible = false;
                oItm.Left = oItmRefTx.Left;
                oItm.Top = oItmRef.Top + oItmRef.Height + oItmRefTx.Height + 2;
                oItm.Height = oItmRefTx.Height;
                oItm.Width = oItmRefTx.Width;
                oItm.TextStyle = (Int32)BoTextStyle.ts_BOLD;
                oBtn = (Button)(oItm.Specific);
                oBtn.Caption = "Solicitar devolución";

                oSql = GlobalSettings.RunningUnderSQLServer ?
                      "Select " :
                      "Select p.\"U_WhsCodTR\", w.\"WhsName\" from \"@VIDR_PARAM\" p left outer join OWHS w on p.\"U_WhsCodTR\" = w.\"WhsCode\" ";
                oRS.DoQuery(String.Format(oSql));
                fromWhs = ((String)(oRS.Fields.Item("U_WhsCodTR").Value)).Trim();
                fromWhsName = ((String)(oRS.Fields.Item("WhsName").Value)).Trim();

                oSql = GlobalSettings.RunningUnderSQLServer ?
                      "Select " :
                      "Select p.\"U_WhsCodCD\", w.\"WhsName\" from \"@VIDR_PARAM\" p left outer join OWHS w on p.\"U_WhsCodCD\" = w.\"WhsCode\" ";
                oRS.DoQuery(String.Format(oSql));
                toWhs = ((String)(oRS.Fields.Item("U_WhsCodCD").Value)).Trim();
                toWhsName = ((String)(oRS.Fields.Item("WhsName").Value)).Trim();

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
                        if ((pVal.ItemUID == "btnInsCD") && (pVal.BeforeAction))
                            BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea generar solicitud de traslado al CD?", 2, "Ok", "Cancel"));
                        else if ((pVal.ItemUID == "btnInsCD") && (!pVal.BeforeAction))
                            if ("" == CreateSolicitudesTraslado(oForm.DataSources.DBDataSources.Item("OWTR"), oForm.DataSources.DBDataSources.Item("WTR1")))
                            {
                                FSBOApp.MessageBox("Solicitudes de traslado al CD generada.", 1, "Ok", "", "");
                                oTxt.Caption = "Solicitud al CD generada.";
                                oBtn.Item.Visible = false;
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

            Int32 oDocEntry = -1;
            String oSql;
            Int32 nVal = -1;
            Int32 nValAnt = -2;
            String Traslados = "";
            Boolean ConTraslado = false;
            Boolean SinTraslado = false;
            try
            {
                switch (oBusinessObjectInfo.EventType)
                {
                    case BoEventTypes.et_FORM_DATA_LOAD:
                        if (!oBusinessObjectInfo.BeforeAction)
                        {
                            oDocEntry = Int32.Parse(oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0));
                            oSql = "Select d.\"U_VR_TrasRl\"     " +
                                   "  from WTR1  d               " +
                                   " where d.\"DocEntry\" = {0}  ";
                            oSql = string.Format(oSql, oDocEntry.ToString());
                            oRS.DoQuery(oSql);
                            while (!oRS.EoF)
                            {
                                nVal = (Int32)(oRS.Fields.Item("U_VR_TrasRl").Value);
                                if (nVal <= 0)
                                    SinTraslado = true;
                                else
                                    ConTraslado = true;
                                if ((SinTraslado) && (ConTraslado))
                                    throw new Exception("Documento con error. Presenta lineas con y sin traslados");

                                if ((nVal > 0) && (nVal != nValAnt))
                                {
                                    if (Traslados == "")
                                        Traslados = nVal.ToString();
                                    else
                                        Traslados = Traslados + " - " + nVal.ToString();
                                    nValAnt = nVal;
                                }
                                oRS.MoveNext();
                            }

                            oSql = GlobalSettings.RunningUnderSQLServer ?
                                  "Select " :
                                  "Select w.\"WhsCode\", w.\"WhsName\" from OWHS w where w.\"U_VR_SN\" = '{0}' ";
                            oSql = String.Format(oSql, oForm.DataSources.DBDataSources.Item("OWTR").GetValue("CardCode", 0));
                            oRS.DoQuery(String.Format(oSql));
                            if (oRS.EoF)
                            {
                                oTxt.Caption = "";
                                oBtn.Item.Visible = false;
                            }
                            else if (oForm.DataSources.DBDataSources.Item("OWTR").GetValue("ToWhsCode", 0).Trim() != fromWhs)
                            {
                                oTxt.Caption = "";
                                oBtn.Item.Visible = false;
                            }
                            else if (Traslados.Trim() == "")
                            {
                                oTxt.Caption = "Solicitud no generada";
                                oBtn.Item.Visible = true;
                            }
                            else
                            {
                                oTxt.Caption = "Solicitud generada: " + Traslados;
                                oBtn.Item.Visible = false;
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

        private void setTrasladoRel(String oDocEntry, String oLineNum, String oDocNum)
        {
            String oSql;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                  "Update" :
                  "Update WTR1 set \"U_VR_TrasRl\" = {2} where \"DocEntry\" = {0} and \"LineNum\" = {1} ";

            oRS.DoQuery(String.Format(oSql, oDocEntry, oLineNum, oDocNum));
        }

        private String CreateSolicitudesTraslado(SAPbouiCOM.DBDataSource DSHead, SAPbouiCOM.DBDataSource DSDet)
        {
            Int32 oDocEntry;
            String sErr = "";
            String oSql;
            Int32 oLine = -1;
            String oDocNewKey;

            try
            {
                oSql = GlobalSettings.RunningUnderSQLServer ?
                      "Select " :
                      "Select w.\"WhsCode\", w.\"WhsName\" from OWHS w where w.\"U_VR_SN\" = '{0}' ";
                oSql = String.Format(oSql, DSHead.GetValue("CardCode", 0));
                oRS.DoQuery(String.Format(oSql));
                if (oRS.EoF)
                    throw new Exception("No se ha definido SN para bodega de tienda.");

                SAPbobsCOM.StockTransfer oDoc = (SAPbobsCOM.StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oInventoryTransferRequest);

                oDoc.CardCode = DSHead.GetValue("CardCode", 0);
                oDoc.DocDate = DateTime.Today;
                oDoc.TaxDate = DateTime.Today;
                oDoc.ShipToCode = DSHead.GetValue("ShipToCode", 0);
                oDoc.FromWarehouse = fromWhs;
                oDoc.UserFields.Fields.Item("U_VK_Almacen_Origen").Value = fromWhsName;
                oDoc.ToWarehouse = toWhs;
                oDoc.UserFields.Fields.Item("U_VK_AlmacenDestino").Value = toWhsName;
                oDoc.Comments = "Solicitud de traslado al CD (devolución): " + DSHead.GetValue("DocNum", 0) + " - " + DSHead.GetValue("Comments", 0);
                oDoc.UserFields.Fields.Item("U_VR_DocRel").Value = Int32.Parse(DSHead.GetValue("DocNum", 0));
                oDoc.UserFields.Fields.Item("U_VK_Tipo_Solicitud").Value = DSHead.GetValue("U_VK_Tipo_Solicitud", 0);
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
                    oDoc.Lines.Price = Double.Parse(DSDet.GetValue("Price", i), System.Globalization.CultureInfo.InvariantCulture);
                    //oDoc.Lines.Price = Double.Parse(DSDet.GetValue("Price", i), System.Globalization.CultureInfo.InvariantCulture);
                }

                if (oDoc != null)
                {
                    if (GlobalSettings.Debug)
                        oDoc.SaveToFile("oSolicitudTransferenciaStockDevolucion.xml");

                    int nErr = oDoc.Add();
                    if (nErr != 0)
                    {
                        FCmpny.GetLastError(out nErr, out sErr);
                        throw new Exception(sErr);
                    }
                    oDocNewKey = FCmpny.GetNewObjectKey();
                    oDoc = null;

                    for (int i = 0; i < DSDet.Size; i++)
                        setTrasladoRel(oDocEntry.ToString(), DSDet.GetValue("LineNum", i), oDocNewKey);
                }

                return sErr;
            }
            catch (Exception e)
            {
                FSBOApp.StatusBar.SetText("Error: " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                OutLog(e.Message + " - " + e.StackTrace);
                return e.Message;
            }
        }
    }
}
