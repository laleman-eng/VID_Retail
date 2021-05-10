using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using VisualD.vkBaseForm;
using VisualD.vkFormInterface;
using VisualD.GlobalVid;
using VisualD.SBOGeneralService;
using VisualD.SBOFunctions;


namespace VID_Retail.Utils
{
    class TUtils
    {
        private VisualD.GlobalVid.TGlobalVid GlobalSettings;
        private String DBCmpny;
        private String UserDB, PwDB;
        public VisualD.SBOFunctions.CSBOFunctions SBO_f;

        public TUtils(ref SAPbobsCOM.Recordset oRS, ref VisualD.GlobalVid.TGlobalVid oGlob, Boolean ValidaParams)
        {
            String oSql;

            GlobalSettings = oGlob;

            oSql = GlobalSettings.RunningUnderSQLServer ?
                   "Select U_BD_Gerona, U_Usuario, U_Password, U_UsuarioBD, U_PasswordBD, U_SN_CR, U_Debug " +
                   "  from [@VIDR_PARAM]" :
                   "Select \"U_BD_Gerona\", \"U_Usuario\", \"U_Password\", \"U_UsuarioBD\", \"U_PasswordBD\", \"U_SN_CR\", \"U_SNGerona\", \"U_Debug\"  " +
                   "  from \"@VIDR_PARAM\" ";
            oRS.DoQuery(oSql);
            if ((oRS.EoF) && (ValidaParams))
                throw new Exception("Parametros no definidos");

            DBCmpny = ((String)oRS.Fields.Item("U_BD_Gerona").Value).Trim();
            //User = ((String)oRS.Fields.Item("U_Usuario").Value).Trim();
            //Pw = ((String)oRS.Fields.Item("U_Password").Value).Trim();
            UserDB = ((String)oRS.Fields.Item("U_UsuarioBD").Value).Trim();
            PwDB = ((String)oRS.Fields.Item("U_PasswordBD").Value).Trim();
            oGlob.SNGerona = ((String)oRS.Fields.Item("U_SNGerona").Value).Trim();
            oGlob.SN_CR = ((String)oRS.Fields.Item("U_SN_CR").Value).Trim();
            oGlob.Debug = ((String)oRS.Fields.Item("U_Debug").Value).Trim() == "Y" ? true : false;
        }

        public bool SetOtherSBOCompany()
        {
            Int32 nErr;
            String sErr;

            try
            {
                if (GlobalSettings.SNGerona.Trim() == "")
                    throw new Exception("SN Gerona no definido en parametros");
                if (GlobalSettings.SN_CR.Trim() == "")
                    throw new Exception("SN Casa Royal no definido en parametros");

                if (GlobalSettings.oCompanyVentaRelacionada == null)
                    GlobalSettings.oCompanyVentaRelacionada = new SAPbobsCOM.Company();

                if (GlobalSettings.oCompanyVentaRelacionada != null)
                    if (GlobalSettings.oCompanyVentaRelacionada.Connected)
                        return true;

                GlobalSettings.SBO_f.SBOApp.StatusBar.SetText("Estableciendo conexión con empresa venta relacionada...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                GlobalSettings.oCompanyVentaRelacionada.Server = GlobalSettings.SBO_f.Cmpny.Server;
                GlobalSettings.oCompanyVentaRelacionada.DbServerType = GlobalSettings.SBO_f.Cmpny.DbServerType;
                GlobalSettings.oCompanyVentaRelacionada.DbUserName = UserDB;
                GlobalSettings.oCompanyVentaRelacionada.DbPassword = PwDB;
                GlobalSettings.oCompanyVentaRelacionada.CompanyDB = DBCmpny;
                GlobalSettings.oCompanyVentaRelacionada.UserName = GlobalSettings.SBO_f.Cmpny.UserName;
                GlobalSettings.oCompanyVentaRelacionada.Password = GlobalSettings.Pw;

                nErr = GlobalSettings.oCompanyVentaRelacionada.Connect();

                if (nErr != 0)
                {
                    GlobalSettings.oCompanyVentaRelacionada.GetLastError(out nErr, out sErr);
                    throw new Exception(sErr);
                }

                return true;
            }
            catch (Exception e)
            {
                GlobalSettings.SBO_f.SBOApp.StatusBar.SetText(e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                GlobalSettings.SBO_f.oLog.OutLog(e.Message + " - " + e.StackTrace);
                return false;
            }
        }

        public void disconnectOtherSBOCompany()
        {
            // No desconectar otra BD
            return;
            //if (GlobalSettings.oCompanyVentaRelacionada.Connected)
            //    GlobalSettings.oCompanyVentaRelacionada.Disconnect();
        }


        public Int32 AddDataSourceInt1(String Header, SAPbouiCOM.DBDataSource oDBDSHeader, String Line1, SAPbouiCOM.DBDataSource oDBDSLine1, String Line2, SAPbouiCOM.DBDataSource oDBDSLine2, String Line3, SAPbouiCOM.DBDataSource oDBDSLine3)
        {
            SAPbobsCOM.GeneralService oGeneralServiceAdd = null;
            SAPbobsCOM.GeneralData oGeneralDataAdd = null;
            SAPbobsCOM.GeneralDataCollection oGeneralCollection = null;
            SAPbobsCOM.GeneralDataParams oGeneralDataParameter = null;
            TSBOGeneralService oGen = null;
            SAPbobsCOM.Company Cmpny;
            SAPbobsCOM.CompanyService CmpnyService;

            try
            {
                Cmpny = SBO_f.Cmpny;
                CmpnyService = Cmpny.GetCompanyService();
                oGen = new TSBOGeneralService();

                oGen.SBO_f = SBO_f;
                oGeneralServiceAdd = (SAPbobsCOM.GeneralService)(CmpnyService.GetGeneralService(Header));
                oGeneralDataAdd = (SAPbobsCOM.GeneralData)(oGeneralServiceAdd.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData));
                oGen.SetNewDataSourceHeader_InUDO(oDBDSHeader, oGeneralDataAdd);

                if (oDBDSLine1 != null)
                {
                    if (oDBDSLine1.Size > 0)
                    {
                        oGeneralCollection = oGeneralDataAdd.Child(Line1);
                        oGen.SetNewDataSourceLines_InUDO(oDBDSLine1, oGeneralDataAdd, oGeneralCollection);
                    }
                }

                if (oDBDSLine2 != null)
                {
                    if (oDBDSLine2.Size > 0)
                    {
                        oGeneralCollection = oGeneralDataAdd.Child(Line2);
                        oGen.SetNewDataSourceLines_InUDO(oDBDSLine2, oGeneralDataAdd, oGeneralCollection);
                    }
                }

                if (oDBDSLine3 != null)
                {
                    if (oDBDSLine3.Size > 0)
                    {
                        oGeneralCollection = oGeneralDataAdd.Child(Line3);
                        oGen.SetNewDataSourceLines_InUDO(oDBDSLine3, oGeneralDataAdd, oGeneralCollection);
                    }
                }

                //Cmpny.StartTransaction();
                oGeneralDataParameter = oGeneralServiceAdd.Add(oGeneralDataAdd);
                //Cmpny.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                return (System.Int32)(oGeneralDataParameter.GetProperty("DocEntry"));
            }
            catch (Exception e)
            {
                SBO_f.oLog.OutLog("AddDataSource1: Error-> " + e.Message + " ** Trace: " + e.StackTrace);
                return 0;
            }
            finally
            {
                SBO_f._ReleaseCOMObject(oGeneralServiceAdd);
                SBO_f._ReleaseCOMObject(oGeneralDataAdd);
                SBO_f._ReleaseCOMObject(oGeneralCollection);
                SBO_f._ReleaseCOMObject(oGeneralDataParameter);
                SBO_f._ReleaseCOMObject(oGen);
            }
        }

        public Int32 UpdDataSourceInt1(String Header, SAPbouiCOM.DBDataSource oDBDSHeader, String Line1, SAPbouiCOM.DBDataSource oDBDSLine1, String Line2, SAPbouiCOM.DBDataSource oDBDSLine2, String Line3, SAPbouiCOM.DBDataSource oDBDSLine3)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataCollection oGeneralDataCollection = null;
            SAPbobsCOM.GeneralDataParams oGeneralDataParams = null;
            TSBOGeneralService oGen = null;
            String StrDummy;
            SAPbobsCOM.Company Cmpny;
            SAPbobsCOM.CompanyService CmpnyService;
            Int32 DocEntry;

            oGen = new TSBOGeneralService();
            try
            {
                oGen.SBO_f = SBO_f;
                Cmpny = SBO_f.Cmpny;
                CmpnyService = Cmpny.GetCompanyService();
                oGeneralService = (SAPbobsCOM.GeneralService)(SBO_f.Cmpny.GetCompanyService().GetGeneralService(Header));
                oGeneralDataParams = (SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams));
                StrDummy = "DocEntry";
                DocEntry = Convert.ToInt32(((System.String)oDBDSHeader.GetValue("DocEntry", 0)));
                oGeneralDataParams.SetProperty(StrDummy, DocEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralDataParams);

                oGen.SetNewDataSourceHeader_InUDO(oDBDSHeader, oGeneralData);

                if (oDBDSLine1 != null)
                {
                    if (oDBDSLine1.Size > 0)
                    {
                        oGeneralDataCollection = oGeneralData.Child(Line1);
                        oGen.SetNewDataSourceLines_InUDO(oDBDSLine1, oGeneralData, oGeneralDataCollection);
                    }
                }

                if (oDBDSLine2 != null)
                {
                    if (oDBDSLine2.Size > 0)
                    {
                        oGeneralDataCollection = oGeneralData.Child(Line2);
                        oGen.SetNewDataSourceLines_InUDO(oDBDSLine2, oGeneralData, oGeneralDataCollection);
                    }
                }

                if (oDBDSLine3 != null)
                {
                    if (oDBDSLine3.Size > 0)
                    {
                        oGeneralDataCollection = oGeneralData.Child(Line3);
                        oGen.SetNewDataSourceLines_InUDO(oDBDSLine3, oGeneralData, oGeneralDataCollection);
                    }
                }

                //SBO_f.Cmpny.StartTransaction();
                oGeneralService.Update(oGeneralData);

                //SBO_f.Cmpny.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                return Convert.ToInt32(((System.String)oDBDSHeader.GetValue("DocEntry", 0)));
            }
            catch
            {
                //if (SBO_f.Cmpny.InTransaction)
                //    SBO_f.Cmpny.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                return 0;
            }
            finally
            {
                SBO_f._ReleaseCOMObject(oGeneralService);
                SBO_f._ReleaseCOMObject(oGeneralData);
                SBO_f._ReleaseCOMObject(oGeneralDataCollection);
                SBO_f._ReleaseCOMObject(oGeneralDataParams);
                SBO_f._ReleaseCOMObject(oGen);
            }
        }
    }
}
