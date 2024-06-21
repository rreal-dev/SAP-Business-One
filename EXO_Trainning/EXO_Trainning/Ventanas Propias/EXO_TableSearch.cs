using Cliente;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using System.Globalization;

namespace Cliente
{
    public class EXO_TableSearch
    {
        public static string cgCompanyCode;


        public EXO_TableSearch()
        {
        }

        public EXO_TableSearch(bool lCreacion)
        {
            SAPbouiCOM.Form oForm = null;


            #region CargoScreen
            SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.oGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));

            Type Tipo = this.GetType();
            string strXML = Matriz.oGlobal.funciones.leerEmbebido(Matriz.TypeMatriz, "Formularios.EXO_FORM_TableSearch.srf");




            oParametrosCreacion.XmlData = strXML;
            oParametrosCreacion.UniqueID = "";

            try
            {
                oForm = Matriz.oGlobal.SBOApp.Forms.AddEx(oParametrosCreacion);
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }

            oForm.Visible = true;
            oForm.PaneLevel = 1;
            ((Folder)oForm.Items.Item("FLD_OCRD").Specific).Select();

            cargarCombo(oForm, (Item)oForm.Items.Item("cb_OCRD"));
            cargarCombo(oForm, (Item)oForm.Items.Item("cb_OITM"));
            cargarCombo(oForm, (Item)oForm.Items.Item("cb_ORDR"));

            CreateChooseFromList(oForm);
            #endregion







        }



        public bool ItemEvent(ItemEvent infoEvento)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

            switch (infoEvento.EventType)
            {

                case BoEventTypes.et_ITEM_PRESSED:




                    if ((infoEvento.ItemUID == "1")
                         && infoEvento.BeforeAction)
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("boton [ok] pulsado", 1, "Ok", "", "");
                    }

                    if ((infoEvento.ItemUID == "2")
                         && infoEvento.BeforeAction)
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("boton [cancel] pulsado", 1, "Ok", "", "");
                    }

                    if ((infoEvento.ItemUID == "btn_Buscar")
                         && infoEvento.BeforeAction)
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("boton [btn_Buscar] pulsado", 1, "Ok", "", "");
                        mostrarTabla(oForm);
                    }

                    if ((infoEvento.ItemUID == "btn_ALL")
                         && infoEvento.BeforeAction)
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("boton [btn_ALL] pulsado", 1, "Ok", "", "");
                        mostrarTodo(oForm);
                    }
                    break;


                case BoEventTypes.et_CHOOSE_FROM_LIST:

                    if (infoEvento.BeforeAction & infoEvento.ItemUID == "et_N_OCRD")
                    {
                        #region Modifico el Choosefromlist para que salgan solo los del tipo correcto

                        

                        if (infoEvento.ItemUID == "et_N_OCRD")
                        {
                            string cIDChoose = "CFL_OCRD";
                        
                        



                        SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
                        SAPbouiCOM.ChooseFromList oCFL = oCFLs.Item(cIDChoose);



                        oCFL.SetConditions(null);
                        //SAPbouiCOM.Conditions oCons = oCFL.GetConditions();



                        //string cTipoProblema = "";
                        //try
                        //{
                        //    cTipoProblema = ((SAPbouiCOM.ComboBox)oForm.Items.Item("2").Specific).Selected.Value;
                        //}
                        //catch (Exception ex)
                        //{
                        //}



                        ////Se llaman los campos igual
                        //SAPbouiCOM.Condition oCon = oCons.Add();
                        //oCon.Alias = "CardType";
                        //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        //oCon.CondVal = cTipoProblema;
                        ////oCon = oCons.Add();
                        //oCFL.SetConditions(oCons);

                        }
                        #endregion

                    }

                    if (infoEvento.BeforeAction & infoEvento.ItemUID == "et_N_OITM")
                    {
                        #region Modifico el Choosefromlist para que salgan solo los del tipo correcto
                        if (infoEvento.ItemUID == "et_N_OITM")
                        {
                            string cIDChoose = "CFL_OITM";



                            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
                            SAPbouiCOM.ChooseFromList oCFL = oCFLs.Item(cIDChoose);



                            oCFL.SetConditions(null);
                            //SAPbouiCOM.Conditions oCons = oCFL.GetConditions();



                            //string cTipoProblema = "";
                            //try
                            //{
                            //    cTipoProblema = ((SAPbouiCOM.ComboBox)oForm.Items.Item("4").Specific).Selected.Value;
                            //}
                            //catch (Exception ex)
                            //{
                            //}



                            ////Se llaman los campos igual
                            //SAPbouiCOM.Condition oCon = oCons.Add();
                            //oCon.Alias = "U_EXO_TipPro";
                            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            //oCon.CondVal = cTipoProblema;
                            ////oCon = oCons.Add();
                            //oCFL.SetConditions(oCons);
                        }
                        #endregion

                    }

                    if (infoEvento.BeforeAction & infoEvento.ItemUID == "et_N_ORDR")
                    {
                        #region Modifico el Choosefromlist para que salgan solo los del tipo correcto
                        if (infoEvento.ItemUID == "et_N_ORDR")
                        {
                            string cIDChoose = "CFL_ORDR";



                            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
                            SAPbouiCOM.ChooseFromList oCFL = oCFLs.Item(cIDChoose);



                            oCFL.SetConditions(null);
                            //SAPbouiCOM.Conditions oCons = oCFL.GetConditions();



                            //string cTipoProblema = "";
                            //try
                            //{
                            //    cTipoProblema = ((SAPbouiCOM.ComboBox)oForm.Items.Item("17").Specific).Selected.Value;
                            //}
                            //catch (Exception ex)
                            //{
                            //}



                            ////Se llaman los campos igual
                            //SAPbouiCOM.Condition oCon = oCons.Add();
                            //oCon.Alias = "U_EXO_TipPro";
                            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            //oCon.CondVal = cTipoProblema;
                            ////oCon = oCons.Add();
                            //oCFL.SetConditions(oCons);
                        }
                        #endregion

                    }


                    if (!infoEvento.BeforeAction && infoEvento.ActionSuccess)
                    {
                        #region Los Choosefromlist
                        IChooseFromListEvent oCFLEvento = (IChooseFromListEvent)infoEvento;
                        string sCFL_ID = oCFLEvento.ChooseFromListUID;
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        
                        if (oForm.Mode == BoFormMode.fm_OK_MODE) oForm.Mode = BoFormMode.fm_UPDATE_MODE;



                        if (oCFLEvento.SelectedObjects != null)
                        {
                            switch (infoEvento.ItemUID)
                            {
                                case "et_N_OCRD":
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)oForm.Items.Item("et_N_OCRD").Specific).String = Convert.ToString(oCFLEvento.SelectedObjects.GetValue("CardCode", 0));
                                    }
                                    catch (Exception ex)
                                    { }
                                    break;
                                case "et_N_OITM":
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)oForm.Items.Item("et_N_OITM").Specific).String = Convert.ToString(oCFLEvento.SelectedObjects.GetValue("ItemCode", 0));
                                    }
                                    catch (Exception ex)
                                    { }
                                    break;
                                case "et_N_ORDR":
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)oForm.Items.Item("et_N_ORDR").Specific).String = Convert.ToString(oCFLEvento.SelectedObjects.GetValue("DocEntry", 0));
                                    }
                                    catch (Exception ex)
                                    { }
                                    break;


                            }
                        }
                        #endregion
                    }
                    break;
            }

            return true;

        }












        public static void CreateChooseFromList(Form oForm)
        {

            try
            {
                #region CREATE ChooseFromList CFL_OCRD
                SAPbouiCOM.ChooseFromList CFL_OCRD;
                SAPbouiCOM.ChooseFromListCreationParams parametros_CFL_OCRD = (SAPbouiCOM.ChooseFromListCreationParams)Matriz.oGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                parametros_CFL_OCRD.MultiSelection = false;
                parametros_CFL_OCRD.ObjectType = "2";
                parametros_CFL_OCRD.UniqueID = "CFL_OCRD";
                CFL_OCRD = oForm.ChooseFromLists.Add(parametros_CFL_OCRD);
                SAPbouiCOM.EditText etEditText_et_N_OCRD = ((SAPbouiCOM.EditText)oForm.Items.Item("et_N_OCRD").Specific);
                //etEditText_et_N_OCRD.DataBind.SetBound(true, "OCRD", "CardCode");
                etEditText_et_N_OCRD.ChooseFromListUID = "CFL_OCRD";
                etEditText_et_N_OCRD.ChooseFromListAlias = "CardCode";
                #endregion


                #region CREATE ChooseFromList CFL_OITM
                SAPbouiCOM.ChooseFromList CFL_OITM;
                SAPbouiCOM.ChooseFromListCreationParams parametros_CFL_OITM = (SAPbouiCOM.ChooseFromListCreationParams)Matriz.oGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                parametros_CFL_OITM.MultiSelection = false;
                parametros_CFL_OITM.ObjectType = "4";
                parametros_CFL_OITM.UniqueID = "CFL_OITM";
                CFL_OITM = oForm.ChooseFromLists.Add(parametros_CFL_OITM);
                SAPbouiCOM.EditText etEditText_et_N_OITM = ((SAPbouiCOM.EditText)oForm.Items.Item("et_N_OITM").Specific);
              
                etEditText_et_N_OITM.ChooseFromListUID = "CFL_OITM";
                etEditText_et_N_OITM.ChooseFromListAlias = "ItemCode";
                #endregion


                #region CREATE ChooseFromList CFL_ORDR
                SAPbouiCOM.ChooseFromList CFL_ORDR;
                SAPbouiCOM.ChooseFromListCreationParams parametros_CFL_ORDR = (SAPbouiCOM.ChooseFromListCreationParams)Matriz.oGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                parametros_CFL_ORDR.MultiSelection = false;
                parametros_CFL_ORDR.ObjectType = "17";
                parametros_CFL_ORDR.UniqueID = "CFL_ORDR";
                CFL_ORDR = oForm.ChooseFromLists.Add(parametros_CFL_ORDR);
                SAPbouiCOM.EditText etEditText_et_Nombre_ORDR = ((SAPbouiCOM.EditText)oForm.Items.Item("et_N_ORDR").Specific);
               
                etEditText_et_Nombre_ORDR.ChooseFromListUID = "CFL_ORDR";
                etEditText_et_Nombre_ORDR.ChooseFromListAlias = "DocEntry";
                #endregion




            }
            catch (Exception e)
            {
                Matriz.oGlobal.SBOApp.MessageBox(e.Message, 1, "Ok", "", "");
            }
        }


        
        



        public static void cargarCombo(Form oForm, Item oItem)
        {
            ComboBox combo1, combo2, combo3;
            switch (oItem.UniqueID)
            {
                case "cb_OITM":

                    combo1 = (ComboBox)oItem.Specific;
                    combo1.ValidValues.Add("ItemCode", "Código");

                    combo2 = (ComboBox)oItem.Specific;
                    combo2.ValidValues.Add("ItemName", "Nombre");

                    combo3 = (ComboBox)oItem.Specific;
                    combo3.ValidValues.Add("ItmsGrpCod", "Grupo");

                    break;

                case "cb_OCRD":

                    combo1 = (ComboBox)oItem.Specific;
                    combo1.ValidValues.Add("CardCode", "Código");

                    combo2 = (ComboBox)oItem.Specific;
                    combo2.ValidValues.Add("CardName", "Nombre");

                    combo3 = (ComboBox)oItem.Specific;
                    combo3.ValidValues.Add("CardType", "Tipo");

                    break;

                case "cb_ORDR":

                    combo1 = (ComboBox)oItem.Specific;
                    combo1.ValidValues.Add("DocEntry", "Código Pedido");

                    combo2 = (ComboBox)oItem.Specific;
                    combo2.ValidValues.Add("DocType", "Tipo Pedido");

                    combo3 = (ComboBox)oItem.Specific;
                    combo3.ValidValues.Add("CardCode", "CódigoCliente");
                    break;

                default:
                    break;
            }
        }


        public static void Add_Data_To_Matrix_OCRD(String sQuery, Form oForm)
        {
            try
            {
                Matrix matrix;
                matrix = oForm.Items.Item("Mtrx_OCRD").Specific;
                matrix.Clear();
                //for (int i = 0; i < matrix.Columns.Count; i++)
                //{
                //    matrix.Columns.Remove(i);
                //}

                Matriz.oGlobal.SBOApp.MessageBox(sQuery);
                ((DataTable)oForm.DataSources.DataTables.Item("DT_OCRD")).ExecuteQuery(sQuery);

                //try
                //{
                //matrix.Columns.Add("CardCode", BoFormItemTypes.it_EDIT);
                //matrix.Columns.Add("CardName", BoFormItemTypes.it_EDIT);
                //matrix.Columns.Add("CardType", BoFormItemTypes.it_EDIT);
                //matrix.Columns.Add("GroupCode", BoFormItemTypes.it_EDIT);
                //matrix.Columns.Add("CntctPrsn", BoFormItemTypes.it_EDIT);
                //matrix.Columns.Add("Balance", BoFormItemTypes.it_EDIT);
                //matrix.Columns.Add("OrdersBal", BoFormItemTypes.it_EDIT);
                //}
                //catch (Exception) { }

                //matrix.Columns.Item("CardCode").TitleObject.Caption = "CardCode";
                //matrix.Columns.Item("CardName").TitleObject.Caption = "CardName";
                //matrix.Columns.Item("CardType").TitleObject.Caption = "CardType";
                //matrix.Columns.Item("GroupCode").TitleObject.Caption = "GroupCode";
                //matrix.Columns.Item("CntctPrsn").TitleObject.Caption = "CntctPersn";
                //matrix.Columns.Item("Balance").TitleObject.Caption = "Balance";
                //matrix.Columns.Item("OrdersBal").TitleObject.Caption = "OrderBal";

                matrix.Columns.Item("CardCode").DataBind.Bind("DT_OCRD", "CardCode");
                matrix.Columns.Item("CardName").DataBind.Bind("DT_OCRD", "CardName");
                matrix.Columns.Item("CardType").DataBind.Bind("DT_OCRD", "CardType");
                matrix.Columns.Item("GroupCode").DataBind.Bind("DT_OCRD", "GroupCode");
                matrix.Columns.Item("CntctPrsn").DataBind.Bind("DT_OCRD", "CntctPrsn");
                matrix.Columns.Item("Balance").DataBind.Bind("DT_OCRD", "Balance");
                matrix.Columns.Item("OrdersBal").DataBind.Bind("DT_OCRD", "OrdersBal");

                matrix.LoadFromDataSource();
                matrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox("Error en el Metodo Add_Data_To_Matrix_OCRD");
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message);
            }
            
        }

        public static void Add_Data_To_Matrix_OITM(String sQuery, Form oForm)
        {
            try
            {
                Matrix matrix;
                matrix = oForm.Items.Item("Mtrx_OITM").Specific;
                matrix.Clear();
                //for (int i = 0; i < matrix.Columns.Count; i++)
                //{
                //    matrix.Columns.Remove(i);
                //}

                Matriz.oGlobal.SBOApp.MessageBox(sQuery);
                ((DataTable)oForm.DataSources.DataTables.Item("DT_OITM")).ExecuteQuery(sQuery);

                //try
                //{
                //    matrix.Columns.Add("ItemCode", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("ItemName", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("ItmsGrpCod", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("OnHand", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("IsCommited", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("OnOrder", BoFormItemTypes.it_EDIT);
                //}
                //catch (Exception) { }


                //matrix.Columns.Item("ItemCode").TitleObject.Caption = "ItemCode";
                //matrix.Columns.Item("ItemName").TitleObject.Caption = "ItemName";
                //matrix.Columns.Item("ItmsGrpCod").TitleObject.Caption = "ItmsGrpCod";
                //matrix.Columns.Item("OnHand").TitleObject.Caption = "OnHand";
                //matrix.Columns.Item("IsCommited").TitleObject.Caption = "IsCommited";
                //matrix.Columns.Item("OnOrder").TitleObject.Caption = "OnOrder";

                matrix.Columns.Item("ItemCode").DataBind.Bind("DT_OITM", "ItemCode");
                matrix.Columns.Item("ItemName").DataBind.Bind("DT_OITM", "ItemName");
                matrix.Columns.Item("ItmsGrpCod").DataBind.Bind("DT_OITM", "ItmsGrpCod");
                matrix.Columns.Item("OnHand").DataBind.Bind("DT_OITM", "OnHand");
                matrix.Columns.Item("IsCommited").DataBind.Bind("DT_OITM", "IsCommited");
                matrix.Columns.Item("OnOrder").DataBind.Bind("DT_OITM", "OnOrder");

                matrix.LoadFromDataSource();
                matrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox("Error en el Metodo Add_Data_To_Matrix_OITM");
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message);
            }
            
        }

        public static void Add_Data_To_Matrix_ORDR(String sQuery, Form oForm)
        {
            try
            {
                Matrix matrix;
                matrix = oForm.Items.Item("Mtrx_ORDR").Specific;
                matrix.Clear();
                //for (int i = 0; i < matrix.Columns.Count; i++)
                //{
                //    matrix.Columns.Remove(i);
                //}

                Matriz.oGlobal.SBOApp.MessageBox(sQuery);
                ((DataTable)oForm.DataSources.DataTables.Item("DT_ORDR")).ExecuteQuery(sQuery);

                //try
                //{
                //    matrix.Columns.Add("DocEntry", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("CardCode", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("CardName", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("DocType", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("DocDate", BoFormItemTypes.it_EDIT);
                //    matrix.Columns.Add("DocStatus", BoFormItemTypes.it_EDIT);
                //}
                //catch (Exception){}


                //matrix.Columns.Item("DocEntry").TitleObject.Caption = "DocEntry";
                //matrix.Columns.Item("CardCode").TitleObject.Caption = "CardCode";
                //matrix.Columns.Item("CardName").TitleObject.Caption = "CardName";
                //matrix.Columns.Item("DocType").TitleObject.Caption = "DocType";
                //matrix.Columns.Item("DocDate").TitleObject.Caption = "DocDate";
                //matrix.Columns.Item("DocStatus").TitleObject.Caption = "DocStatus";

                matrix.Columns.Item("DocEntry").DataBind.Bind("DT_ORDR", "DocEntry");
                matrix.Columns.Item("CardCode").DataBind.Bind("DT_ORDR", "CardCode");
                matrix.Columns.Item("CardName").DataBind.Bind("DT_ORDR", "CardName");
                matrix.Columns.Item("DocType").DataBind.Bind("DT_ORDR", "DocType");
                matrix.Columns.Item("DocDate").DataBind.Bind("DT_ORDR", "DocDate");
                matrix.Columns.Item("DocStatus").DataBind.Bind("DT_ORDR", "DocStatus");

                matrix.LoadFromDataSource();
                matrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox("Error en el Metodo Add_Data_To_Matrix_ORDR");
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message);
            }
            
        }



        public static Boolean mostrarTodo(Form oForm)
        {
            try
            {
                String sQuery = "";

                
                
                

                if (((Folder)oForm.Items.Item("FLD_OCRD").Specific).Selected)
                {
                    ((EditText)oForm.Items.Item("et_N_OCRD").Specific).Value = "*";
                    sQuery = "SELECT * FROM OCRD";
                    ClearComboBox(((ComboBox)oForm.Items.Item("cb_OCRD").Specific));
                    cargarCombo(oForm, (Item)oForm.Items.Item("cb_OCRD"));
                    Add_Data_To_Matrix_OCRD(sQuery, oForm);
                }

                if (((Folder)oForm.Items.Item("FLD_OITM").Specific).Selected)
                {
                    ((EditText)oForm.Items.Item("et_N_OITM").Specific).Value = "*";
                    sQuery = "SELECT * FROM OITM";
                    ClearComboBox(((ComboBox)oForm.Items.Item("cb_OITM").Specific));
                    cargarCombo(oForm, (Item)oForm.Items.Item("cb_OITM"));
                    Add_Data_To_Matrix_OITM(sQuery, oForm);
                }

                if (((Folder)oForm.Items.Item("FLD_ORDR").Specific).Selected)
                {
                    ((EditText)oForm.Items.Item("et_N_ORDR").Specific).Value = "*";
                    sQuery = "SELECT * FROM ORDR";
                    ClearComboBox(((ComboBox)oForm.Items.Item("cb_ORDR").Specific));
                    cargarCombo(oForm, (Item)oForm.Items.Item("cb_ORDR"));
                    Add_Data_To_Matrix_ORDR(sQuery, oForm);
                    
                }              
                
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message);
                return false;
            }
            return true;
        }


        public static void ClearComboBox(ComboBox combo)
        {
            int Count = combo.ValidValues.Count;

            for (int i = 0; i < Count; i++)

            {

                combo.ValidValues.Remove(combo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);

            }

        }



        public static void mostrarTabla(Form oForm)
        {
            String sQuery = "";
            Matrix matrix;

            try
            {
                if (((Folder)oForm.Items.Item("FLD_OCRD").Specific).Selected)
                {
                    if (((EditText)oForm.Items.Item("et_N_OCRD").Specific).Value != "" && ((ComboBox)oForm.Items.Item("cb_OCRD").Specific).Value != "")
                    {
                        sQuery = "SELECT * FROM OCRD";
                        sQuery += " WHERE " + ((ComboBox)oForm.Items.Item("cb_OCRD").Specific).Value;
                        sQuery += " LIKE '" + ((EditText)oForm.Items.Item("et_N_OCRD").Specific).Value + "'";

                        Add_Data_To_Matrix_OCRD(sQuery, oForm);
                    }

                    else
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("INTRODUCE UN VALOR VÁLIDO", 1, "Ok", "", "");
                    }

                }

                if (((Folder)oForm.Items.Item("FLD_OITM").Specific).Selected)
                {
                    if (((EditText)oForm.Items.Item("et_N_OITM").Specific).Value != "" && ((ComboBox)oForm.Items.Item("cb_OITM").Specific).Value != "")
                    {
                        sQuery = "SELECT * FROM OITM";
                        sQuery += " WHERE " + ((ComboBox)oForm.Items.Item("cb_OITM").Specific).Value;
                        sQuery += " LIKE '" + ((EditText)oForm.Items.Item("et_N_OITM").Specific).Value + "'";

                        Add_Data_To_Matrix_OITM(sQuery, oForm);
                    }

                    else
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("INTRODUCE UN VALOR VÁLIDO", 1, "Ok", "", "");
                    }

                }


                if (((Folder)oForm.Items.Item("FLD_ORDR").Specific).Selected)
                {
                    if (((EditText)oForm.Items.Item("et_N_ORDR").Specific).Value != "" && ((ComboBox)oForm.Items.Item("cb_ORDR").Specific).Value !="")
                    {
                        sQuery = "SELECT * FROM ORDR";
                        sQuery += " WHERE " + ((ComboBox)oForm.Items.Item("cb_ORDR").Specific).Value;
                        sQuery += " LIKE '" + ((EditText)oForm.Items.Item("et_N_ORDR").Specific).Value + "'";

                        Add_Data_To_Matrix_ORDR(sQuery, oForm);
                    }

                    else
                    {
                        Matriz.oGlobal.SBOApp.MessageBox("INTRODUCE UN VALOR VÁLIDO", 1, "Ok", "", "");
                    }
                }

            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }




        }
    }
}
