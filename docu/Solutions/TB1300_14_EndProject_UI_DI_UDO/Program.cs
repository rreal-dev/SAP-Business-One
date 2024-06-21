using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace TB1300
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static SAPbobsCOM.Company diCompany;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }

                diCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                UDO.CreateUDT("TB1_CAR", "Car Master Data", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                UDO.CreateUDT("TB1_CAR_D", "Car Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);

                UDO.CreateUDF("TB1_CAR_D", "MODEL", "Car Model", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
                UDO.CreateUDF("TB1_CAR_D", "FUEL", "Fuel Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
                UDO.CreateUDF("TB1_CAR_D", "BODY", "Body Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
                UDO.CreateUDF("TB1_CAR_D", "POWER", "Horse Power", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);

                UDO.CreateUDO();

                UDO.InsertToUDO("01", "BMW", "320i", "Petrol", "Sedan", "110");
                UDO.InsertToUDO("02", "Ford", "Focus", "Diesel", "Hatchback", "120");
                UDO.InsertToUDO("03", "Kia", "Rio", "Petrol", "Tourer", "130");
                UDO.InsertToUDO("04", "Mercedes", "SLS", "Diesel", "Coupe", "140");
                UDO.InsertToUDO("05", "Skoda", "Octavia", "Petrol", "Sedan", "150");
                UDO.InsertToUDO("06", "Alfa Romeo", "Gulia", "Hybrid", "SUV", "160");
                UDO.InsertToUDO("07", "VolksWagen", "Golf", "Petrol", "Coupe", "170");
                UDO.InsertToUDO("08", "Peugeot", "Partner", "Diesel", "Van", "180");
                UDO.InsertToUDO("09", "Lexus", "IS300", "Hybrid", "Sedan", "190");
                UDO.InsertToUDO("10", "Toyota", "Yaris", "Petrol", "Hatchback", "200");

                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
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
    }
}
