using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using SAPbouiCOM;
using SAPbobsCOM; 
using VisualD.MainObjBase;
using VisualD.MenuConfFr; 
//using VisualD.GlobalVid;
using VisualD.SBOFunctions;
using VisualD.vkBaseForm;
using VisualD.vkFormInterface;
using VisualD.MultiFunctions;
using System.Xml;
using VID_Retail.Precios;
using VID_Retail.Surtido;
using VID_Retail.Tiendas;
using VID_Retail.Periodos;
using VID_Retail.Clusters;
using VID_Retail.Departamentos;
using VID_Retail.Categorias;
using VID_Retail.Grupos;
using VID_Retail.Familias;
using VID_Retail.Parametros;
using VID_Retail.ControlTraslados;
using VID_Retail.FacturaVentaRelacionada;
using VID_Retail.AjusteStockLF;
using VID_Retail.NCVentaRelacionada;
using VID_Retail.OCCrossDocking;
using VID_Retail.TransferenciaStockDev;
using VID_Retail.DespachoATiendas;
using VID_Retail.RecepcionenTiendas;
using VID_Retail.FiltroAceptacionRecep;
using VID_Retail.CambioEstadoMasivoOT;
using VID_Retail.TransferenciaOrdenServicio;

namespace VID_Retail.RetailObj
{
    public class TRetailObj : TMainObjBase //class(TMainObjBase)
    {
        public override void AddMenus()
        {
            base.AddMenus();
            System.Xml.XmlDocument oXMLDoc;
            //String sImagePath;
            try
            {
                //inherited addMenus;
                oXMLDoc = new System.Xml.XmlDocument();
                //try
                    //sImagePath := TMultiFunctions.ExtractFilePath(TMultiFunctions.ParamStr(0)) + '\Menus\Menu.xml';
                    //oXMLDoc.Load(sImagePath);
                    //StrAux := oXMLDoc.InnerXml;
                    //SBOApplication.LoadBatchActions(var StrAux);
                //except
                //on e: exception do
                    //SBOFunctions.oLog.OutLog('AddMenus err: ' + e.Message + ' ** Trace: ' + e.  StackTrace);
                //end;
            }
            finally
            {
                oXMLDoc = null;
            }
        } //fin AddMenus

        public override void MenuEventExt(List<object> oForms, ref MenuEvent pVal, ref Boolean BubbleEvent)
        {
            IvkFormInterface oForm;
            base.MenuEventExt(oForms, ref pVal, ref BubbleEvent);
            try
            {
                //Inherited MenuEventExt(oForms,var pVal,var BubbleEvent);
                oForm = null;
                if (! pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "VD_RETAIL_01":
                                oForm = (IvkFormInterface)(new TPrecios(oForms));
                                break;
                        case "VD_RETAIL_02":
                                oForm = (IvkFormInterface)(new TSurtido());
                                break;
                        case "VD_RETAIL_03":
                                oForm = (IvkFormInterface)(new TTiendas());
                                break;
                        case "VD_RETAIL_04":
                                oForm = (IvkFormInterface)(new TPeriodos());
                                break;
                        case "VD_RETAIL_05":
                                oForm = (IvkFormInterface)(new TClusters());
                                break;
                        case "VD_RETAIL_06":
                                oForm = (IvkFormInterface)(new TDepartamentos());
                                break;
                        case "VD_RETAIL_07":
                                oForm = (IvkFormInterface)(new TCategorias());
                                break;
                        case "VD_RETAIL_08":
                                oForm = (IvkFormInterface)(new TFamilias());
                                break;
                        case "VD_RETAIL_09":
                                oForm = (IvkFormInterface)(new TGrupos());
                                break;
                        case "VD_RETAIL_10":
                                oForm = (IvkFormInterface)(new TParametros());
                                break;
                        case "VD_RETAIL_11":
                                oForm = (IvkFormInterface)(new TControlTraslados());
                                break;
                        case "VD_RETAIL_12":
                                oForm = (IvkFormInterface)(new TAjusteStockLF());
                                break;
                        case "VD_RETAIL_13":
                                oForm = (IvkFormInterface)(new TDespachoATiendas());
                                break;
                        case "VD_RETAIL_14":
                                oForm = (IvkFormInterface)(new TRecepcionenTiendas());
                                break;
                        case "VD_RETAIL_15":
                                oForm = (IvkFormInterface)(new TFiltroAceptacionRecep());
                                break;
                        case "VD_RETAIL_16":
                                oForm = (IvkFormInterface)(new TCambioEstadoMasivoOT());
                                break;
                        case "VD_RETAIL_17":
                                oForm = (IvkFormInterface)(new TTransferenciaOrdenServicio());
                                break;
                        default:
                                break;
                    }  
            
                    if (oForm != null) 
                    {
                        SAPbouiCOM.Application App = SBOApplication;
                        SAPbobsCOM.Company Cmpny = SBOCompany;
                        VisualD.SBOFunctions.CSBOFunctions SboF = SBOFunctions;
                        VisualD.GlobalVid.TGlobalVid Glob = GlobalSettings;
                        
                        if (oForm.InitForm(SBOFunctions.generateFormId(GlobalSettings.SBOSpaceName, GlobalSettings), "forms\\",ref  App,ref  Cmpny,ref SboF, ref Glob)) 
                        {   oForms.Add(oForm); }
                        else 
                        {
                            SBOApplication.Forms.Item(oForm.getFormId()).Close();
                            oForm = null;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                SBOApplication.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok","","");  // Captura errores no manejados
                oLog.OutLog("MenuEventExt: " + e.Message + " ** Trace: " + e.StackTrace);
            }
        } //MenuEventExt

        public override IvkFormInterface ItemEventExt(IvkFormInterface oIvkForm, List<object> oForms, String LstFrmUID, String FormUID, ref ItemEvent pVal, ref Boolean BubbleEvent)
        {
            SAPbouiCOM.Form oForm;
            SAPbouiCOM.Form oFormParent;
            IvkFormInterface result = null;
            result = base.ItemEventExt(oIvkForm, oForms, LstFrmUID, FormUID, ref pVal, ref BubbleEvent);
            try
            {
                //inherited ItemEventExt(oIvkForm,oForms,LstFrmUID, FormUID, var pVal, var BubbleEvent);   

                result = base.ItemEventExt(oIvkForm, oForms, LstFrmUID, FormUID, ref pVal, ref BubbleEvent);

                if (result != null)
                {
                    return result;
                }
                else
                {
                    if (oIvkForm != null)
                    {
                        return oIvkForm;
                    }
                }

                // CFL Extendido (Enmascara el CFL estandar)
                if ((pVal.BeforeAction) && (pVal.EventType == BoEventTypes.et_FORM_LOAD) && (!string.IsNullOrEmpty(LstFrmUID)))
                {
                    try
                    {
                        oForm = SBOApplication.Forms.Item(LstFrmUID);
                    }
                    catch
                    {
                        oForm = null;
                    }
                }


                if ((!pVal.BeforeAction) && (pVal.FormTypeEx == "0"))
                {
                    if ((oIvkForm == null) && (GlobalSettings.UsrFldsFormActive) && (GlobalSettings.UsrFldsFormUid != "") && (pVal.EventType == BoEventTypes.et_FORM_LOAD))
                    {
                        oForm = SBOApplication.Forms.Item(pVal.FormUID);
                        oFormParent = SBOApplication.Forms.Item(GlobalSettings.UsrFldsFormUid);
                        try
                        {
                            //SBO_App.StatusBar.SetText(oFormParent.Title,BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Warning);
                            SBOFunctions.FillListUserFieldForm(GlobalSettings.ListFormsUserField, oFormParent, oForm);
                        }
                        finally
                        {
                            GlobalSettings.UsrFldsFormUid = "";
                            GlobalSettings.UsrFldsFormActive = false;
                        }
                    }
                    else
                    {
                        if ((pVal.EventType == BoEventTypes.et_FORM_ACTIVATE) || (pVal.EventType == BoEventTypes.et_COMBO_SELECT) || (pVal.EventType == BoEventTypes.et_FORM_RESIZE))
                        {
                            oForm = SBOApplication.Forms.Item(pVal.FormUID);
                            SBOFunctions.DisableListUserFieldsForm(GlobalSettings.ListFormsUserField, oForm);
                        }
                    }

                }


                if ((!pVal.BeforeAction) && (pVal.EventType == BoEventTypes.et_FORM_LOAD) && (oIvkForm == null))
                {
                    switch (pVal.FormTypeEx)
                    {
                        case "133": // Factura deudores
                            result = (IvkFormInterface)(new TFacturaVentaRelacionada(oForms));
                            break;
                        case "142": // Orden de Compra 
                            result = (IvkFormInterface)(new TOCCrossDocking());
                            break;
                        case "179": // NC deudores
                            result = (IvkFormInterface)(new TNCVentaRelacionada(oForms));
                            break;
                        case "940": //  Transferencia de stock 
                            result = (IvkFormInterface)(new TTransferenciaStockDev());
                            break;
                    } //fi  switch
                }


                if (result != null)
                {
                    SAPbouiCOM.Application App = SBOApplication;
                    SAPbobsCOM.Company Cmpny = SBOCompany;
                    VisualD.SBOFunctions.CSBOFunctions SboF = SBOFunctions;
                    VisualD.GlobalVid.TGlobalVid Glob = GlobalSettings;
                    if (result.InitForm(pVal.FormUID, @"forms\\", ref App, ref Cmpny, ref SboF, ref Glob))
                    {
                        oForms.Add(result);
                    }
                    else
                    {
                        SBOApplication.Forms.Item(result.getFormId()).Close();
                        result = null;
                    }
                }

                return result;
            }// fin try
            catch (Exception e)
            {
                oLog.OutLog("ItemEventExt: " + e.Message + " ** Trace: " + e.StackTrace);
                SBOApplication.MessageBox(e.Message + " ** Trace: " + e.StackTrace, 1, "Ok","","");  // Captura errores no manejados
                return null;
            }
    
        } //fin ItemEventExt
    }
}
