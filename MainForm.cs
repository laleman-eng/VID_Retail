using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using VisualD.Main;
using VisualD.MultiFunctions;
using VisualD.MainObjBase;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;
using System.IO;
using VID_Retail.RetailObj;

namespace VID_Retail
{

    public partial class MainForm : System.Windows.Forms.Form
    {

        TMainClassExt MainClass = new TMainClassExt();
        public MainForm()
        {
            InitializeComponent();
            MainClass.MainObj.Add(new TRetailObj());
            MainClass.Init();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ShowInTaskbar = false;
            Hide();
        }

    }

    public class TMainClassExt : TMainClass //class (TMainClass)
    {
        public TMainClassExt()
            : base()
        {
        }

        private void CloseSplash()
        {
            //  if SplashScreen.Visible then SplashScreen.Close(); 
        }

        public override void SetFiltros()
        {
            //SAPbouiCOM.EventFilters oFilters;
            //SAPbouiCOM.EventFilter oFilter;
             
        }


        public override void initApp()
        {
            String XlsFile;
            String s;
            base.initApp();
            Boolean sUser = false;
            SAPbobsCOM.Recordset ors;

            try
            {
                GlobalSettings.SBOMeta = SBOMetaData;

                oLog.DebugLvl = 20;
                GlobalSettings.SBO_f = SBOFunctions;
                MainObj[0].GlobalSettings = GlobalSettings;
                MainObj[0].SBOApplication = SBOApplication;
                MainObj[0].SBOCompany = SBOCompany;
                MainObj[0].oLog = oLog;
                MainObj[0].SBOFunctions = SBOFunctions;

                ors = (SAPbobsCOM.Recordset)(SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                if (GlobalSettings.RunningUnderSQLServer)
                    s = "Select SUPERUSER from OUSR where INTERNAL_K = " + SBOCompany.UserSignature.ToString();
                else
                    s = "Select SUPERUSER from OUSR where INTERNAL_K = " + SBOCompany.UserSignature.ToString();

                ors.DoQuery(s);
                if (! ors.EoF)
                {
                   if ((String)(ors.Fields.Item("SUPERUSER").Value) == "Y") 
                      sUser = true;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
                ors = null;
                GC.Collect();

                if (sUser)
                {
                   XlsFile = System.IO.Path.GetDirectoryName(TMultiFunctions.ParamStr(0)) + "\\Docs\\EDVDRET.xls";
                   if (!SBOFunctions.ValidEstructSHA1(XlsFile))
                   {
                       oLog.OutLog("InitApp: Estructura de datos - VID Retail");
                       SBOApplication.StatusBar.SetText("Inicializando AddOn VID Retail.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                       if (!SBOMetaData.SyncTablasUdos("1.1", XlsFile))
                       {
                           SBOFunctions.DeleteSHA1FromTable("EDVDRET.xls");
                           oLog.OutLog("InitApp: sincronización de Estructura de datos fallo: EDVDRET.xls");
                           CloseSplash();
                           SBOApplication.MessageBox("Estructura de datos con problemas, consulte a soporte...", 1, "Ok", "", "");
                           Halt(0);
                       }
                   }
                   XlsFile = System.IO.Path.GetDirectoryName(TMultiFunctions.ParamStr(0)) + "\\Docs\\EDVDCROYAL.xls";
                   if (!SBOFunctions.ValidEstructSHA1(XlsFile))
                   {
                       oLog.OutLog("InitApp: Estructura de datos - VID Gerona");
                       SBOApplication.StatusBar.SetText("Inicializando AddOn VID Retail.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                       if (!SBOMetaData.SyncTablasUdos("1.1", XlsFile))
                       {
                           SBOFunctions.DeleteSHA1FromTable("EDVDCROYAL.xls");
                           oLog.OutLog("InitApp: sincronización de Estructura de datos fallo: EDVDCROYAL.xls");
                           CloseSplash();
                           SBOApplication.MessageBox("Estructura de datos con problemas, consulte a soporte...", 1, "Ok", "", "");
                           Halt(0);
                       }
                   }
                   XlsFile = System.IO.Path.GetDirectoryName(TMultiFunctions.ParamStr(0)) + "\\Docs\\UDFRECTDA.xls";
                   if (!SBOFunctions.ValidEstructSHA1(XlsFile))
                   {
                       oLog.OutLog("InitApp: Estructura de datos - VID Gerona");
                       SBOApplication.StatusBar.SetText("Inicializando AddOn VID Retail.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                       if (!SBOMetaData.SyncTablasUdos("1.1", XlsFile))
                       {
                           SBOFunctions.DeleteSHA1FromTable("EDVDCROYAL.xls");
                           oLog.OutLog("InitApp: sincronización de Estructura de datos fallo: UDFRECTDA.xls");
                           CloseSplash();
                           SBOApplication.MessageBox("Estructura de datos con problemas, consulte a soporte...", 1, "Ok", "", "");
                           Halt(0);
                       }
                   }
                }


                //SetFiltros();


                MainObj[0].AddMenus();

                InitOK = false;
                oLog.OutLog("App SBO in C# - Init!");
                SBOApplication.StatusBar.SetText("Aplicación Inicializada.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                SetFiltros();

                InitOK = true;
                oLog.OutLog("C# - Shine your crazy diamond!");
                SBOApplication.StatusBar.SetText("Aplicación Inicializada.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                CloseSplash();
            }
        }
    }
}
