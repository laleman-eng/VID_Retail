using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VisualD.SBOFunctions;
using VisualD.MultiFunctions;
using VisualD.SBOObjectMg1;

namespace VisualD.GlobalVid
{
    public struct TUserFieldFormRecord
    {
        public string ParentFormUid;
        public string ChildFormUid;
        public string Table;
        public string Field;
    }

    public class TGlobalAddOnOptions
    {
        public string AddonId;
        public string Opciones;
        public string SQLUsers;
    }

    public class TGlobalVid : ICloneable
    {
        private Int32 fCorrelativo = 0;
        private Boolean fUsarSkinLinkVID = true;
        public Boolean fUsarSkinVID = true;
        public Boolean fUsarSkinVID_Form0 = false;
        public Boolean fCerrarCRForm = false;
        public Boolean fValidarCierre = false;

        public SAPbobsCOM.Company oCompanyVentaRelacionada = null;
        public String SNGerona;
        public String SN_CR;
        public Boolean Debug;
        public String Pw = "";

        private CSBOFunctions FSBO_f;
        public string ThousandSeparator;
        public string CrystalReportFileName;
        public string CRFormPreviewUID = "";
        public string DecimalSeparator;
        public Int32 NumberDecimalDigits = 6;
        public Int32 NumberDecimalDigitsQty = 6;
        public string SBOSpaceName = "VIDRetail";
        public string LastFormUID = "";
        public string PrevFormUID = "";
        public Int32 MinFormWidth = 800;
        public Int32 MinFormHeight = 600;
        public Int32 HeightBtn = 50;
        public string ExePath = "";
        public string GLOB_DocEntry = "";
        public List<string> GLOB_ListDocEntry = new List<string>();
        public Boolean UsrFldsFormActive = false;
        public string UsrFldsFormUid = "";
        public List<TUserFieldFormRecord> ListFormsUserField = new List<TUserFieldFormRecord>();
        public List<Object> FoForms;
        public Boolean GLOB_CenterForm = true;
        public Dictionary<string, string> Menu_List_CrystalReports = new Dictionary<string, string>();
        public Dictionary<string, System.Type> Menu_List_vkBaseForm = new Dictionary<string, System.Type>();
        public Dictionary<string, Dictionary<string, Int32>> FormFolderPanes = new Dictionary<string, Dictionary<string, Int32>>();
        public string Ser_ItemCode = "";
        public string Ser_Serie = "";
        public string GLOB_EncryptSQL = "Addon BPP";
        public TSBOObjectMg SBOMeta { get; set; }

#if HANA
//        public string GLOB_TipoBD = "Hana";
#else
//        public string GLOB_TipoBD = "SQLServer";
#endif

        public Boolean ValidarCierre
        {
            get
            {
                return fValidarCierre;
            }

            set
            {
                fValidarCierre = value;
            }
        }

        public Boolean UsarSkinLinkVID
        {
            get
            {
                return fUsarSkinLinkVID;
            }

            set
            {
                fUsarSkinLinkVID = value;
            }

        }


        public Boolean UsarSkinVID
        {
            get
            {
                return fUsarSkinVID;
            }
            set
            {
                fUsarSkinVID = value;
            }

        }

        public Boolean UsarSkinVID_Form0
        {
            get
            {
                return fUsarSkinVID_Form0;
            }

            set
            {
                fUsarSkinVID_Form0 = value;
            }
        }

        public Boolean CerrarCRForm
        {
            get
            {
                return fCerrarCRForm;
            }

            set
            {
                fCerrarCRForm = value;
            }
        }

        public Int32 iCorrelativo
        {
            get
            {
                return fCorrelativo;
            }
            set
            {
                fCorrelativo = value;
            }

        }

        public CSBOFunctions SBO_f
        {
            get
            {
                return FSBO_f;
            }
            set
            {
                FSBO_f = value;
            }
        }

        public object Clone()
        {
            return this.MemberwiseClone();
        }


        public Boolean RunningUnderSQLServer
        {
            get
            {
                return (FSBO_f.Cmpny.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB);
                //return (GLOB_TipoBD == "SQLServer");
            }
            
        }
    }
}
