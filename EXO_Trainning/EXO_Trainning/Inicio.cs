using EXO_Generales;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using System.Xml;
using EXO_UIAPI;

namespace Cliente
{
    public class Matriz : EXO_UIAPI.EXO_DLLBase
    {
        public static EXO_UIAPI.EXO_UIAPI oGlobal;
        public static Type TypeMatriz;


        private static SAPbouiCOM.Application SBO_Application;
        private static SAPbobsCOM.Company diCompany;


        public struct ColumnasUDO
        {
            public string Nomcolum;
            public string Descripcion;
            public SAPbobsCOM.BoYesNoEnum Busqueda;
            public SAPbobsCOM.BoYesNoEnum Visible;
            public SAPbobsCOM.BoYesNoEnum Habilitada;
        }


        public Matriz(EXO_UIAPI.EXO_UIAPI general, Boolean actualizar, Boolean usaLicencia, int idDLL)
            : base(general, actualizar, usaLicencia, idDLL)
        {
            oGlobal = this.objGlobal;
            TypeMatriz = this.GetType();
            Object ThisMatriz = this;

            if (actualizar)
            {

                #region Creo tablas y UDO [COMENTADO]
                /*
                if (objGlobal.refDi.comunes.esAdministrador())
                {
                    string fBD = "", cMen = "";

                    #region Campos de usuario pre
                    cMen = "";
                    fBD = Matriz.oGlobal.funciones.leerEmbebido(Matriz.TypeMatriz, "db_ModuloClienteVIDS.xml");
                    if (!objGlobal.refDi.comunes.LoadBDFromXML(fBD, cMen))
                    {
                        Matriz.oGlobal.SBOApp.MessageBox(cMen, 1, "Ok", "", "");
                        Matriz.oGlobal.SBOApp.MessageBox("Error en creacion de campos db_ModuloClienteVIDS.xml", 1, "Ok", "", "");
                    }
                    else
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("Actualizacion de campos db_ModuloClienteVIDS realizada", 1, "Ok", "", "");
                    }
                    #endregion

                }
                else
                {
                    Matriz.oGlobal.SBOApp.MessageBox("Necesita permisos de administrador para actualizar la base de datos.\nCampos no creados", 1, "Ok", "", "");
                }
                */
                #endregion m

                #region Creo tablas
                //db_DVD__UDTyUDF.Create_UDT();
                //db_DVD__UDTyUDF.InsertToUDT();
                #endregion
            }

            //**
            SAPbobsCOM.Recordset oRec = null;


            GC.WaitForPendingFinalizers();


        }

        public override EventFilters filtros()
        {
            SAPbouiCOM.EventFilters oFilters = new SAPbouiCOM.EventFilters();
            #region Mando filtros
            try
            {
                string fXML = Matriz.oGlobal.funciones.leerEmbebido(Matriz.TypeMatriz, "xFiltrosFile.xml");
                oFilters.LoadFromXML(fXML);
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox("Error en carga de xFiltrosFile.xml", 1, "Ok", "", "");
                oFilters = null;
            }
            #endregion                  
            return oFilters;

        }

        public override XmlDocument menus()
        {

            try
            {
                string mXML = "";
                XmlDocument oXML = new XmlDocument();

                mXML = Matriz.oGlobal.funciones.leerEmbebido(Matriz.TypeMatriz, "xMenu_EXO_FORMS.xml");

                oXML.LoadXml(mXML);
                return oXML;
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox("Error en carga de xMenu_EXO_FORMS.xml", 1, "Ok", "", "");
                return null;
            }




        }





        public override bool SBOApp_ItemEvent(ItemEvent infoEvento)
        {
            bool lRetorno = true;

            if (infoEvento.FormTypeEx == "EXO_FORM_TableSearch")
            {
                EXO_TableSearch fFichTableSearch = new EXO_TableSearch();
                lRetorno = fFichTableSearch.ItemEvent(infoEvento);
                fFichTableSearch = null;
            }

            return lRetorno;
        }


        public override bool SBOApp_FormDataEvent(BusinessObjectInfo infoDataEvent)
        {
            bool lRetorno = true;

            return lRetorno;
        }



        public override bool SBOApp_MenuEvent(MenuEvent infoMenuEvent)
        {
            bool lRetorno = true;

            switch (infoMenuEvent.MenuUID)
            {

                case "mFich1":

                    if (!infoMenuEvent.BeforeAction)
                    {
                        EXO_TableSearch fFichForm1 = new EXO_TableSearch(true);
                        fFichForm1 = null;

                    }
                    break;

            }

            return lRetorno;
        }







    }
}

