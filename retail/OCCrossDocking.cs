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

namespace VID_Retail.OCCrossDocking
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

    class TOCCrossDocking : TvkBaseForm, IvkFormInterface
    {
        SAPbouiCOM.StaticText oTxt;
        SAPbouiCOM.Button oBtn;

        public TOCCrossDocking()
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
                oItm = oForm.Items.Add("btnInsCD", BoFormItemTypes.it_BUTTON);
                oItm.Visible = false;
                oItm.Left = oItmRef.Left;
                oItm.Top = oItmRef.Top + (oItmRef.Height + 2) * 2;
                oItm.Height = oItmRef.Height;
                oItm.Width = oItmRef.Width;
                oItm.TextStyle = (Int32)BoTextStyle.ts_BOLD;
                oBtn = (Button)(oItm.Specific);
                oBtn.Caption = "Ingrear Sol. traslado";

                oSql = GlobalSettings.RunningUnderSQLServer ?
                      "Select " :
                      "Select p.\"U_WhsCodCD\", w.\"WhsName\" from \"@VIDR_PARAM\" p left outer join OWHS w on p.\"U_WhsCodCD\" = w.\"WhsCode\" ";
                oRS.DoQuery(String.Format(oSql));
                fromWhs = ((String)(oRS.Fields.Item("U_WhsCodCD").Value)).Trim();
                fromWhsName = ((String)(oRS.Fields.Item("WhsName").Value)).Trim();
                oSql = GlobalSettings.RunningUnderSQLServer ?
                      "Select " :
                      "Select p.\"U_WhsCodTR\", w.\"WhsName\" from \"@VIDR_PARAM\" p left outer join OWHS w on p.\"U_WhsCodTR\" = w.\"WhsCode\" ";
                oRS.DoQuery(String.Format(oSql));
                toWhs = ((String)(oRS.Fields.Item("U_WhsCodTR").Value)).Trim();
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
                            BubbleEvent = (1 == FSBOApp.MessageBox("¿Desea generar solicitudes de traslado?", 2, "Ok", "Cancel"));
                        else if ((pVal.ItemUID == "btnInsCD") && (!pVal.BeforeAction))
                            if ("" == CreateSolicitudesTraslado(oForm.DataSources.DBDataSources.Item("OPOR"), oForm.DataSources.DBDataSources.Item("POR1")))
                            {
                                FSBOApp.MessageBox("Solicitudes de traslado generadas.", 1, "Ok", "", "");
                                oTxt.Caption = "Solicitudes de traslado generadas.";
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
                            oDocEntry = Int32.Parse(oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0));
                            oSql = "Select d.\"U_VR_TrasRl\"     " +
                                   "  from POR1  d               " +
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


                            if (oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocType", 0).Trim() != "I")
                            {
                                oTxt.Caption = "";
                                oBtn.Item.Visible = false;
                            }
                            else if (oForm.DataSources.DBDataSources.Item("OPOR").GetValue("U_VK_TipoOC", 0).Trim() != "2") // 2 = CrossDocking
                            {
                                oTxt.Caption = "";
                                oBtn.Item.Visible = false;
                            }
                            else if (Traslados.Trim() == "")
                            {
                                oTxt.Caption = "Solicitudes no generadas";
                                oBtn.Item.Visible = true;
                            }
                            else
                            {
                                oTxt.Caption = "Solicitudes generadas: " + Traslados;
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
                  "Update POR1 set \"U_VR_TrasRl\" = {2} where \"DocEntry\" = {0} and \"LineNum\" = {1} ";

            oRS.DoQuery(String.Format(oSql, oDocEntry, oLineNum, oDocNum));
        }

        private String CreateSolicitudesTraslado(SAPbouiCOM.DBDataSource DSHead, SAPbouiCOM.DBDataSource DSDet)
        {
            Int32 oDocEntry;
            string oSql = "";
            string sErr = "";
            Int32 oLine = -1;
            String oDocNewKey;
            String oWhs = "";
            Int32 j = 0;

            try
            {
                SAPbobsCOM.StockTransfer oDoc = null;
                regPurchaseOrder oPurOrderReg =  null;
                List<regPurchaseOrder> oPurOrder = new List<regPurchaseOrder>();


                oDocEntry = Int32.Parse(DSHead.GetValue("DocEntry", 0));
                oSql = "Select d.\"U_VK_Almacen_Destino\" \"WhsCode\", d.\"ItemCode\", d.\"Dscription\", d.\"Quantity\",     " +
                       "       w.\"WhsName\", IfNull(w.\"U_VR_SN\", '') \"CardCode\",                   " +
                       "       IfNull(c.\"ShipToDef\", '') \"ShipToDef\",                               " +
                       "       IfNull(c.\"Address\", '') \"Address\",                                   " +
                       "       IfNull(c.\"City\", '') \"City\",                                         " +
                       "       d.\"Price\", d.\"LineNum\", d.\"VisOrder\"                               " +
                       "  from OPOR h inner join POR1 d on h.\"DocEntry\" = d.\"DocEntry\"              " +
                       "         left outer join OWHS w on d.\"U_VK_Almacen_Destino\"  = w.\"WhsCode\"  " +
                       "         left outer join OCRD c on w.\"U_VR_SN\"  = c.\"CardCode\"              " +
                       " where h.\"DocEntry\" = {0}                                                     " +
                       " order by 1, 2                                                                  ";
                oSql = string.Format(oSql, oDocEntry.ToString());
                oRS.DoQuery(oSql);
                while (!oRS.EoF)
                {
                    if (((string)(oRS.Fields.Item("CardCode").Value)).Trim() == "")
                        throw new Exception("Cliente no definido para bodega: " + (string)(oRS.Fields.Item("WhsCode").Value));
                    if (((string)(oRS.Fields.Item("WhsCode").Value)).Trim() == "")
                        throw new Exception("Almacen de destino no definido para linea: " + ((Int32)(oRS.Fields.Item("VisOrder").Value)).ToString());
                    if (((string)(oRS.Fields.Item("ShipToDef").Value)).Trim() == "")
                        throw new Exception("Cliente sin direccion de despacho para almacen destino: " + (string)(oRS.Fields.Item("WhsCode").Value) + ", Linea: " + ((Int32)(oRS.Fields.Item("VisOrder").Value)).ToString());

                    oPurOrderReg = new regPurchaseOrder();
                    oPurOrderReg.WhsCode = ((string)(oRS.Fields.Item("WhsCode").Value)).Trim();
                    oPurOrderReg.WhsName = ((string)(oRS.Fields.Item("WhsName").Value)).Trim();
                    oPurOrderReg.LineNum = ((Int32)(oRS.Fields.Item("LineNum").Value));
                    oPurOrderReg.CardCode = ((string)(oRS.Fields.Item("CardCode").Value)).Trim();
                    oPurOrderReg.ShipToDef = ((string)(oRS.Fields.Item("ShipToDef").Value)).Trim();
                    oPurOrderReg.ItemCode = ((string)(oRS.Fields.Item("ItemCode").Value)).Trim();
                    oPurOrderReg.Quantity = ((Double)(oRS.Fields.Item("Quantity").Value));
                    oPurOrderReg.Price = ((Double)(oRS.Fields.Item("Price").Value));
                    oPurOrder.Add(oPurOrderReg);
                    oRS.MoveNext();
                }

                for (int i = 0; i < oPurOrder.Count; i++)
                {
                    if (oWhs != oPurOrder[i].WhsCode)
                    {
                        if (oDoc != null)
                        {
                            if (GlobalSettings.Debug)
                                oDoc.SaveToFile("oTrasladoParaCrossdocking.xml");
                            int nErr = oDoc.Add();
                            if (nErr != 0)
                            {
                                FCmpny.GetLastError(out nErr, out sErr);
                                throw new Exception(sErr);
                            }
                            oDocNewKey = FCmpny.GetNewObjectKey();
                            for (int k = j; k < i; k++)
                                oPurOrder[k].TrasladoKey = oDocNewKey;
                            
                            oDoc = null;
                            j = i;
                        }

                        oLine = -1;
                        oWhs = oPurOrder[i].WhsCode;

                        oDoc = (SAPbobsCOM.StockTransfer)FCmpny.GetBusinessObject(BoObjectTypes.oInventoryTransferRequest);
                        oDoc.CardCode = oPurOrder[i].CardCode;
                        oDoc.DocDate = DateTime.Today;
                        oDoc.TaxDate = DateTime.Today;
                        oDoc.ShipToCode = oPurOrder[i].ShipToDef;
                        oDoc.FromWarehouse = fromWhs;
                        oDoc.UserFields.Fields.Item("U_VK_Almacen_Origen").Value = fromWhsName;
                        oDoc.ToWarehouse = toWhs;
                        oDoc.UserFields.Fields.Item("U_VK_AlmacenDestino").Value = toWhsName;
                        oDoc.UserFields.Fields.Item("U_VK_Tipo_Solicitud").Value = "2";  //Cross docking
                        oDoc.UserFields.Fields.Item("U_CR_OC_CrossDocking").Value = (2200000000 + oDocEntry).ToString(); 
                        oDoc.Comments = "OC cross docking relacionada: " + DSHead.GetValue("DocNum", 0) + " - " + DSHead.GetValue("Comments", 0);
                    }


                    oLine++;
                    if (oLine > 0)
                        oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(oLine);

                    oDoc.Lines.ItemCode = oPurOrder[i].ItemCode;
                    oDoc.Lines.Quantity = oPurOrder[i].Quantity;
                    oDoc.Lines.Price = oPurOrder[i].Price;
                    oDoc.Lines.UserFields.Fields.Item("U_VR_TrasRl").Value = Int32.Parse(DSHead.GetValue("DocNum", 0));
                }
                if (oDoc != null)
                {
                    if (GlobalSettings.Debug)
                        oDoc.SaveToFile("oTrasladoParaCrossdocking.xml");
                    int nErr = oDoc.Add();
                    if (nErr != 0)
                    {
                        FCmpny.GetLastError(out nErr, out sErr);
                        throw new Exception(sErr);
                    }
                    oDocNewKey = FCmpny.GetNewObjectKey();
                    for (int k = j; k < oPurOrder.Count ; k++)
                        oPurOrder[k].TrasladoKey = oDocNewKey;

                    oDoc = null;
                }

                for (int i = 0; i < oPurOrder.Count; i++)                
                    setTrasladoRel(oDocEntry.ToString(), oPurOrder[i].LineNum.ToString(), oPurOrder[i].TrasladoKey);

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
