using System;
using SAPbouiCOM.Framework;

namespace TB1300
{
    class UDO
    {
        public static void CreateUDT(string MyTableName, string MyTableDescription, SAPbobsCOM.BoUTBTableType MyTableType)
        {
            try
            {
                SAPbobsCOM.UserTablesMD oUDT;

                oUDT = (SAPbobsCOM.UserTablesMD)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUDT.GetByKey(MyTableName) == false)
                {
                    oUDT.TableName = MyTableName;
                    oUDT.TableDescription = MyTableDescription;
                    oUDT.TableType = MyTableType;
                    int ret = oUDT.Add();

                    if (ret == 0)
                    {
                        Application.SBO_Application.MessageBox("Add Table: " + oUDT.TableName + " successfull");
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDT);
                        GC.Collect();

                    }
                    else
                        Application.SBO_Application.MessageBox("Add Table error: " + Program.diCompany.GetLastErrorDescription());
                }
                else
                    Application.SBO_Application.MessageBox("Table: " + MyTableName + " already exists");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        public static void CreateUDF(string MyTableName, string MyFieldName, string MyFieldDescrition, SAPbobsCOM.BoFieldTypes MyFieldType, int MyFieldSize)
        {
            try
            {
                SAPbobsCOM.UserFieldsMD oUDF;
                oUDF = (SAPbobsCOM.UserFieldsMD)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                oUDF.TableName = MyTableName;
                oUDF.Name = MyFieldName;
                oUDF.Description = MyFieldDescrition;
                oUDF.Type = MyFieldType;
                oUDF.EditSize = MyFieldSize;
                int ret = oUDF.Add();


                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                GC.Collect();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Exception: " + ex.Message);
            }
        }

        public static void CreateUDO()
        {
            try
            {
                SAPbobsCOM.UserObjectsMD oUserObjectMD;

                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                oUserObjectMD.Code = "TB1_CAR";
                oUserObjectMD.Name = oUserObjectMD.Code;
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.TableName = "TB1_CAR";

                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

                oUserObjectMD.FindColumns.ColumnAlias = "Code";
                oUserObjectMD.FindColumns.Add();
                oUserObjectMD.FindColumns.ColumnAlias = "Name";
                oUserObjectMD.FindColumns.Add();

                oUserObjectMD.ChildTables.TableName = "TB1_CAR_D";


                int ret = oUserObjectMD.Add();
                if (ret != 0)
                    Application.SBO_Application.MessageBox("error: " + Program.diCompany.GetLastErrorDescription());

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                GC.Collect();

            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        public static void InsertToUDO(string MyCode, string MyName, string MyModel, string MyFuel, string MyBody, string MyPower)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            try
            {
                oCompanyService = Program.diCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("TB1_CAR");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", MyCode);
                oGeneralData.SetProperty("Name", MyName);
                oChildren = oGeneralData.Child("TB1_CAR_D");
                oChild = oChildren.Add();
                oChild.SetProperty("U_MODEL", MyModel);
                oChild.SetProperty("U_FUEL", MyFuel);
                oChild.SetProperty("U_BODY", MyBody);
                oChild.SetProperty("U_POWER", MyPower);
                oGeneralParams = oGeneralService.Add(oGeneralData);

            }
            catch (Exception ex)
            {
                //Application.SBO_Application.MessageBox(ex.Message);
            }

        }
    }
}
