using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TB1300
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        
        private static SAPbouiCOM.Application SBO_Application;
        private static SAPbobsCOM.Company diCompany;

        [STAThread]
        static void Main()
        {
            ConnectToUI();
            //CreateForm();
            //SaveFormToXML();
            LoadFormFromXML();
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            Application.Run();
        }

        private static void ConnectToUI()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi;
            string sConnectionString;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();

            SBO_Application.MessageBox("Connected to UI API", 1, "Continue", "Cancel");
            //ConnectwithSSO();
            ConnectwithSharedMemory();
        }
        private static void ConnectwithSSO()
        {
            diCompany = new SAPbobsCOM.Company();
            string cookie = diCompany.GetContextCookie();
            string connInfo = SBO_Application.Company.GetConnectionContext(cookie);

            int ret = diCompany.SetSboLoginContext(connInfo);
            if (ret != 0)
                SBO_Application.MessageBox("DI Connection failed!", 0, "Ok", "", "");
            else
                SBO_Application.MessageBox("Connected with SSO!", 0, "Ok", "", "");
        }
        private static void ConnectwithSharedMemory()
        {
            diCompany = (SAPbobsCOM.Company)Program.SBO_Application.Company.GetDICompany();
            SBO_Application.MessageBox("DI Connected To: " + Program.diCompany.CompanyName, 0, "Ok", "", "");
        }

        public static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    SBO_Application.MessageBox("My is addon disconnected." + Program.diCompany.CompanyName, 0, "Ok", "", "");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Program.diCompany);
                    Application.Exit();       
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        public static void CreateForm()
        {
            try
            {
                SAPbouiCOM.Form oForm;
                SAPbouiCOM.FormCreationParams creationPackage;

                creationPackage = (SAPbouiCOM.FormCreationParams)
                SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                creationPackage.UniqueID = "TB1_DVDAvailability";
                creationPackage.FormType = "TB1_DVDAvailability";
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;


                oForm = SBO_Application.Forms.AddEx(creationPackage);
                oForm.Title = "DVD Availability Check";

                oForm.Left = 400;
                oForm.Top = 100;
                oForm.ClientWidth = 270;
                oForm.ClientHeight = 154;

                //create label - DVD Name
                SAPbouiCOM.Item oItem;
                SAPbouiCOM.StaticText oStatic;
                oItem = oForm.Items.Add("lb_Name", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 5;
                oItem.Top = 20;
                oItem.Width = 80;
                oItem.Height = 14;
                oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStatic.Caption = "DVD Name";

                //create label - DVD Aisle
                oItem = oForm.Items.Add("lb_Aisle", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 5;
                oItem.Top = 39;
                oItem.Width = 80;
                oItem.Height = 14;
                oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStatic.Caption = "DVD Aisle";

                //create label - DVD Section
                oItem = oForm.Items.Add("lb_Section", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 5;
                oItem.Top = 58;
                oItem.Width = 80;
                oItem.Height = 14;
                oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStatic.Caption = "DVD Section";

                //create label - DVD Rented
                oItem = oForm.Items.Add("lb_Rented", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 5;
                oItem.Top = 77;
                oItem.Width = 80;
                oItem.Height = 14;
                oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStatic.Caption = "DVD Rented";

                //create label - Rented To
                oItem = oForm.Items.Add("lb_RentTo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 5;
                oItem.Top = 96;
                oItem.Width = 80;
                oItem.Height = 14;
                oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
                oStatic.Caption = "Rented To";

                //create edit text - DVD Name
                oItem = oForm.Items.Add("tx_Name", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 90;
                oItem.Top = 20;
                oItem.Width = 175;
                oItem.Height = 14;
                oItem.LinkTo = "lb_Name";

                //create edit text - DVD Aisle
                oItem = oForm.Items.Add("tx_Aisle", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 90;
                oItem.Top = 39;
                oItem.Width = 175;
                oItem.Height = 14;
                oItem.LinkTo = "lb_Aisle";

                //create edit text - DVD Section
                oItem = oForm.Items.Add("tx_Section", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 90;
                oItem.Top = 58;
                oItem.Width = 175;
                oItem.Height = 14;
                oItem.LinkTo = "lb_Section";

                //create edit text - DVD Rented
                oItem = oForm.Items.Add("tx_Rented", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 90;
                oItem.Top = 77;
                oItem.Width = 175;
                oItem.Height = 14;
                oForm.Visible = true;
                oItem.LinkTo = "lb_Rented";

                //create edit text - Rented To
                oItem = oForm.Items.Add("tx_RentTo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 90;
                oItem.Top = 96;
                oItem.Width = 175;
                oItem.Height = 14;
                oItem.LinkTo = "lb_RentTo";

                //create OK button
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Top = 130;
                oItem.Width = 65;
                oItem.Height = 19;

                //create Cancel button
                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Top = 130;
                oItem.Width = 65;
                oItem.Height = 19;

                //create DVD rent button
                oItem = oForm.Items.Add("Rent", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 200;
                oItem.Top = 130;
                oItem.Width = 65;
                oItem.Height = 19;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Rent DVD";

                oForm.Visible = true;
            }         
            catch(Exception ex)
            {
                SBO_Application.MessageBox("Exception: " + ex.Message);
            }             
            }

        public static void SaveFormToXML()
        {
            try
            {
                SAPbouiCOM.Form oForm;
                oForm = SBO_Application.Forms.GetForm("TB1_DVDAvailability", 0);

                System.Xml.XmlDocument oXMLDoc = new System.Xml.XmlDocument();
                string sXmlString = oForm.GetAsXML();
                oXMLDoc.LoadXml(sXmlString);
                oXMLDoc.Save("../../../DVDAvailability.xml");
                SBO_Application.MessageBox("Form saved.");
            }
            catch(Exception ex)
            {
                SBO_Application.MessageBox("Exception: " + ex.Message);
            }            
        }

        public static void LoadFormFromXML()
        {
            try
            {
                SAPbouiCOM.Form oForm;
                System.Xml.XmlDocument oXMLDoc = new System.Xml.XmlDocument();
                SAPbouiCOM.FormCreationParams creationPackage;

                oXMLDoc.Load("../../../DVDAvailability.xml");
                creationPackage = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.XmlData = oXMLDoc.InnerXml;

                oForm = SBO_Application.Forms.AddEx(creationPackage);
                oForm.Visible = true;
            }
            catch(Exception ex)
            {
                SBO_Application.MessageBox("Exception: " + ex.Message);
            }
        }

    }
}
