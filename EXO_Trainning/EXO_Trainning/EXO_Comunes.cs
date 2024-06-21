using System;
using System.Text;
using System.IO;
using System.Threading;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Reflection;
using System.Net;

namespace Cliente
{
    //
    public class FuncionesFTP
    {

        public struct RutasConexion
        {
            public string Cliente;
            public string RutaDescarga;
            public string RutaUpload;
            public string RutaFallidos;
            public string RutaProcesados;
            public string DirFTP;
            public string Usuftp;
            public string PWDFTP;
            public string CarpetaIN;
            public string CarpetaOUT;
        }

        public static RutasConexion LlenoRutas(string cCliente)
        {
            RutasConexion Rutas = new RutasConexion();

            string sql = "SELECT ISNULL(T0.U_EXO_DirFTP, '') AS 'DirFTP', ISNULL(T0.U_EXO_PWDFTP, '') AS 'PWD', ISNULL(T0.U_EXO_UsuFTP, '') AS 'UsuFTP', ";
            sql += " ISNULL(T0.U_EXO_CarpetaIN, '') AS 'CarpetaIN', ISNULL(T0.U_EXO_CarpetaOUT, '') AS 'CarpetaOUT', ";
            sql += " ISNULL(T0.U_EXO_RutaDescarga, '') AS 'RutaDescarga', isnull(T0.U_EXO_RutaProce, '') AS 'RutaProcesados', ";
            sql += " isnull(T0.U_EXO_RutaUpload, '') AS 'RutaUpload', isnull(T0.U_EXO_RutaFallidos, '') AS 'RutaFallidos' ";
            sql += "  FROM [@EXO_CONEXCLI] T0 WHERE T0.Code ='" + cCliente + "'";

            try
            {
                SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);

                if (oRec.EoF)
                {
                    Rutas.Cliente = "";
                    Rutas.CarpetaIN = "";
                    Rutas.CarpetaOUT = "";
                    Rutas.DirFTP = "";
                    Rutas.PWDFTP = "";
                    Rutas.Usuftp = "";
                    Rutas.RutaDescarga = "";
                    Rutas.RutaFallidos = "";
                    Rutas.RutaProcesados = "";
                    Rutas.RutaUpload = "";
                }
                else
                {
                    Rutas.Cliente = cCliente;
                    Rutas.CarpetaIN = oRec.Fields.Item("CarpetaIN").Value;
                    Rutas.CarpetaOUT = oRec.Fields.Item("CarpetaOUT").Value;
                    Rutas.DirFTP = oRec.Fields.Item("DirFTP").Value;
                    Rutas.PWDFTP = oRec.Fields.Item("PWD").Value;
                    Rutas.Usuftp = oRec.Fields.Item("UsuFTP").Value;
                    Rutas.RutaDescarga = oRec.Fields.Item("RutaDescarga").Value;
                    Rutas.RutaFallidos = oRec.Fields.Item("RutaFallidos").Value;
                    Rutas.RutaProcesados = oRec.Fields.Item("RutaProcesados").Value;
                    Rutas.RutaUpload = oRec.Fields.Item("RutaUpload").Value;
                }
            }
            catch (Exception ex)
            {
            }

            return Rutas;

        }

        public static void DescargarFicherosFTP(RutasConexion Rutas)
        {

            string cRuta = string.Format("ftp://{0}/{1}", Rutas.DirFTP, Rutas.CarpetaIN);
            Matriz.oGlobal.SBOApp.SetStatusBarMessage(cRuta, BoMessageTime.bmt_Short, false);
            FtpWebRequest dirFtp = ((FtpWebRequest)FtpWebRequest.Create(cRuta));

            // Los datos del usuario (credenciales)
            NetworkCredential cr = new NetworkCredential(Rutas.Usuftp, Rutas.PWDFTP);
            dirFtp.Credentials = cr;
            dirFtp.KeepAlive = false; // ¿¿

            // El comando a ejecutar
            //dirFtp.Method = "LIST";


            // También usando la enumeración de WebRequestMethods.Ftp
            dirFtp.Method = WebRequestMethods.Ftp.ListDirectory;


            // Obtener el resultado del comando
            StreamReader reader = new StreamReader(dirFtp.GetResponse().GetResponseStream());

            // Leer el stream
            string cLinea = "";
            while ((cLinea = reader.ReadLine()) != null)
            {
                string cNomCorto = System.IO.Path.GetFileName(cLinea);
                if (cNomCorto.Length < 3 || (cNomCorto.Substring(0, 3) != "art" && cNomCorto.Substring(0, 3) != "pem" && cNomCorto.Substring(0, 3) != "pas")) continue;

                string cRutaDescarga = string.Format("ftp://{0}/{1}/{2}", Rutas.DirFTP, Rutas.CarpetaIN, cNomCorto);
                DescargarYBorrar(Rutas, cRutaDescarga);
            }

            // Cerrar el stream abierto.
            reader.Close();

            reader = null;

        }

        public static bool DescargarYBorrar(RutasConexion Rutas, string ficFTP)
        {
            bool lRetorno = false;

            FtpWebRequest dirFtp = ((FtpWebRequest)FtpWebRequest.Create(ficFTP));

            // Los datos del usuario (credenciales)
            NetworkCredential cr = new NetworkCredential(Rutas.Usuftp, Rutas.PWDFTP);
            dirFtp.Credentials = cr;
            dirFtp.KeepAlive = false; // ¿¿

            // El comando a ejecutar usando la enumeración de WebRequestMethods.Ftp
            dirFtp.Method = WebRequestMethods.Ftp.DownloadFile;

            // Obtener el resultado del comando
            StreamReader reader =
                new StreamReader(dirFtp.GetResponse().GetResponseStream());

            // Leer el stream
            string res = reader.ReadToEnd();


            // Guardarlo localmente con la extensión .txt
            string ficLocal = Path.Combine(Rutas.RutaDescarga, Path.GetFileName(ficFTP));
            StreamWriter sw = new StreamWriter(ficLocal, false, Encoding.UTF8);
            sw.Write(res);
            sw.Close();
            reader.Close();
            Matriz.oGlobal.SBOApp.SetStatusBarMessage("Descargado en Fich local " + ficLocal, BoMessageTime.bmt_Short, false);

            if (File.Exists(ficLocal))
            {
                DeleteFileOnServer(Rutas, ficFTP);
            }

            reader = null;
            sw = null;

            return lRetorno;


        }

        public static bool DeleteFileOnServer(RutasConexion Rutas, string cFileBorrar)
        {
            // The serverUri parameter should use the ftp:// scheme.
            // It contains the name of the server file that is to be deleted.
            // Example: 
            // 

            //if (serverUri.Scheme != Uri.UriSchemeFtp)
            //{
            //    return false;
            //}
            // Get the object used to communicate with the server.

            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(cFileBorrar);
            request.Credentials = new NetworkCredential(Rutas.Usuftp, Rutas.PWDFTP);
            request.Method = WebRequestMethods.Ftp.DeleteFile;
            request.KeepAlive = false;   //¿¿

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            response.Close();
            request = null;
            response = null;

            return true;
        }

        public static bool UploadFTP(FuncionesFTP.RutasConexion Rutas, string strFileNameLocal)
        {
            bool lRetorno = false;
            FtpWebRequest ftpRequest;

            // Crea el objeto de conexión del servidor FTP            
            ftpRequest = (FtpWebRequest)WebRequest.Create(string.Format("ftp://{0}/{1}", Rutas.DirFTP, Path.Combine(Rutas.CarpetaOUT, Path.GetFileName(strFileNameLocal))));
            // Asigna las credenciales
            ftpRequest.Credentials = new NetworkCredential(Rutas.Usuftp, Rutas.PWDFTP);
            // Asigna las propiedades
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
            ftpRequest.UsePassive = true;
            ftpRequest.UseBinary = true;
            ftpRequest.KeepAlive = false;

            // Lee el archivo y lo envía            
            FileStream stream = File.OpenRead(strFileNameLocal);
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);
            stream.Close();
            Matriz.oGlobal.SBOApp.SetStatusBarMessage("Leido " + strFileNameLocal, BoMessageTime.bmt_Short, false);
            Stream reqStream = ftpRequest.GetRequestStream();
            reqStream.Write(buffer, 0, buffer.Length);
            reqStream.Flush();
            reqStream.Close();
            
            try
            {
                ftpRequest = (FtpWebRequest)WebRequest.Create(string.Format("ftp://{0}/{1}", Rutas.DirFTP, Path.Combine(Rutas.CarpetaOUT, Path.GetFileName(strFileNameLocal))));
                ftpRequest.Credentials = new NetworkCredential(Rutas.Usuftp, Rutas.PWDFTP);
                ftpRequest.Method = WebRequestMethods.Ftp.GetDateTimestamp;
                ftpRequest.UsePassive = true;
                ftpRequest.UseBinary = true;
                ftpRequest.KeepAlive = false;

                //Si no existe, casca,...no puedo compara la hora pues el servidor puede tener otra hora
                FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();
                DateTime dFechaModif = response.LastModified;                
                response.Close();

                ////Borro el fichero generado y lo paso a procesados
                Matriz.oGlobal.SBOApp.SetStatusBarMessage("Enviado " + strFileNameLocal, BoMessageTime.bmt_Short, false);
                //Paso el fichero a procesados
                if (Rutas.RutaProcesados != "")
                {
                   System.IO.File.Copy(strFileNameLocal, Path.Combine(Rutas.RutaProcesados, "Respuesta-" + Path.GetFileName(strFileNameLocal)), true);
                   System.IO.File.Delete(strFileNameLocal);
                }
                lRetorno = true;
                
            }
            catch (Exception ex)
            {
            }

            return lRetorno;
        }
    }

    public class Conversiones
    {
      
      public static double ValueSAPToDoubleSistema(string Texto)
            {
                string Cadena = Texto;

                double Valor = 0.0;
                System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();
                string SepDecSistema = nfi.NumberGroupSeparator;
                //string SepMilSistema = nfi.NumberDecimalSeparator;

                //En pantalla el separador decimal es .
                if (SepDecSistema != ".")
                {
                    Cadena = Cadena.Replace('.', ',');
                }
                double.TryParse(Cadena, out Valor);
                return Valor;
            }
      
  

   

   
      public static string DateStringSAP(DateTime dFecha)
      {
          string cRetorno = dFecha.Year.ToString("0000") + dFecha.Month.ToString("00") + dFecha.Day.ToString("00");

          return cRetorno;
      }
    }

  

  
    public class Utilidades
    {
       
       #region Lleno Combo de Series
        public static bool LlenoComboSeries(ref SAPbouiCOM.ComboBox oCombo,  BoObjectTypes oTipoObjeto, bool lSinBloqueados, bool lConIndicador, string cFecha)
        {
            string cNumeroObjeto = "", cIndicador = "", sql = "";
            SAPbobsCOM.Recordset oRec;

            switch (oTipoObjeto)
            {
                case BoObjectTypes.oDeliveryNotes:
                      cNumeroObjeto = "15";
                    break;
                case BoObjectTypes.oInvoices:
                      cNumeroObjeto = "13";
                    break;

            }

            if (lConIndicador)
            {
                sql = "SELECT T0.Indicator FROM OFPR T0 WHERE Convert(varchar(10), T0.F_RefDate, 112) <= '" + cFecha + "'  AND ";
                sql += " Convert(varchar(10), T0.T_RefDate, 112) >= '" + cFecha + "'";
                oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
                cIndicador = oRec.Fields.Item(0).Value.ToString().Trim();
            }
           
            sql = "SELECT T0.Series, T0.SeriesName FROM NNM1 T0  ";
            sql += " WHERE T0.ObjectCode = '" + cNumeroObjeto + "'";

            if (lConIndicador) sql += " AND T0.Indicator = '" + cIndicador + "'";
            if (lSinBloqueados) sql += " AND T0.Locked = 'N'";



            //Borro lo que hubiera
            int nCount = oCombo.ValidValues.Count;
            for (int i = 0; i < nCount; i++)
            {
                oCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }
            
            oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString().Trim(), oRec.Fields.Item(1).Value.ToString().Trim());
                oRec.MoveNext();
            }
            

            return true;
        }

        public static bool LlenoComboSeries(ref SAPbouiCOM.Column oColumnCombo, BoObjectTypes oTipoObjeto, bool lSinBloqueados, bool lConIndicador, string cFecha)
        {
            SAPbobsCOM.Recordset oRec;
            oRec = SubLlenoComboSeries(oTipoObjeto, lSinBloqueados, lConIndicador, cFecha);

            //Borro lo que hubiera
            int nCount = oColumnCombo.ValidValues.Count;
            for (int i = 0; i < nCount; i++)
            {
                oColumnCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }

            while (!oRec.EoF)
            {
                oColumnCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString().Trim(), oRec.Fields.Item(1).Value.ToString().Trim());
                oRec.MoveNext();
            }


            return true;
        }

        private static SAPbobsCOM.Recordset SubLlenoComboSeries(BoObjectTypes oTipoObjeto, bool lSinBloqueados, bool lConIndicador, string cFecha)
        {
            string cNumeroObjeto = "", cIndicador = "", sql = "";
            SAPbobsCOM.Recordset oRec = null;

            switch (oTipoObjeto)
            {
                //Albaran de ventas
                case BoObjectTypes.oDeliveryNotes:
                    cNumeroObjeto = "15";
                    break;

                //Facturas de venta
                case BoObjectTypes.oInvoices:
                    cNumeroObjeto = "13";
                    break;

                //Devoluciones de ventas
                case BoObjectTypes.oReturns:
                    cNumeroObjeto = "16";
                    break;
                    
                //Abonos
                case BoObjectTypes.oCreditNotes:
                    cNumeroObjeto = "14";
                    break;

            }

            if (lConIndicador)
            {
                sql = "SELECT T0.Indicator FROM OFPR T0 WHERE Convert(varchar(10), T0.F_RefDate, 112) <= '" + cFecha + "'  AND ";
                sql += " Convert(varchar(10), T0.T_RefDate, 112) >= '" + cFecha + "'";
                oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
                cIndicador = oRec.Fields.Item(0).Value.ToString().Trim();
            }

            sql = "SELECT T0.Series, T0.SeriesName FROM NNM1 T0  ";
            sql += " WHERE T0.ObjectCode = '" + cNumeroObjeto + "'";

            if (lConIndicador) sql += " AND T0.Indicator = '" + cIndicador + "'";
            if (lSinBloqueados) sql += " AND T0.Locked = 'N'";

            oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            return oRec;

            ////Borro lo que hubiera
            //int nCount = oCombo.ValidValues.Count;
            //for (int i = 0; i < nCount; i++)
            //{
            //    oCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            //}

            //oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            //while (!oRec.EoF)
            //{
            //    oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString().Trim(), oRec.Fields.Item(1).Value.ToString().Trim());
            //    oRec.MoveNext();
            //}


            
        }        
        #endregion


        public static void CrearBusquedaFormateada(string FormId, string ItemId, string ColId, BoFormattedSearchActionEnum Accion,
                                                   int Consulta, BoYesNoEnum Refrescar, string FieldId, BoYesNoEnum ForzarRefrescar, BoYesNoEnum PorCampo)
        {
            bool lExiste = false;
            int nRet;

            SAPbobsCOM.FormattedSearches oFormattedSearches = (SAPbobsCOM.FormattedSearches)Matriz.oGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);

            string sql = "SELECT IndexID FROM CSHS WHERE FormId = '" + FormId + "' AND ItemID = '" + ItemId + "' AND ColID = '" + ColId + "'";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            string cAux = oRec.Fields.Item(0).Value.ToString().Trim();
            if (cAux == "") cAux = "0";
                        
            if (Convert.ToInt16(cAux) != 0)
            {
                
                oFormattedSearches.GetByKey(Convert.ToInt32(Convert.ToInt16(cAux)));
                lExiste = true;
            }

            oFormattedSearches.FormID = FormId;
            oFormattedSearches.ItemID = ItemId;
            oFormattedSearches.ColumnID = ColId;
            oFormattedSearches.Action = Accion;
            oFormattedSearches.QueryID = Consulta;
            oFormattedSearches.Refresh = Refrescar;
            oFormattedSearches.FieldID = FieldId;
            oFormattedSearches.ForceRefresh = ForzarRefrescar;
            oFormattedSearches.ByField = PorCampo;

            nRet = lExiste ? oFormattedSearches.Update() : oFormattedSearches.Add();
            if (nRet != 0)
            {
                Matriz.oGlobal.SBOApp.MessageBox(Matriz.oGlobal.compañia.GetLastErrorDescription(), 1, "", "", "");
            }
            else
            {
                Matriz.oGlobal.SBOApp.SetStatusBarMessage((lExiste ? "Actualizada" : "Creada") + " busqueda formateada para Form. " + FormId + " Item " + ItemId + " Col " + ColId + " - consulta " + Consulta.ToString(), BoMessageTime.bmt_Short, false);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearches);
            oFormattedSearches = null;
        }

        public static void BorroLineaMatrix(ref SAPbouiCOM.Matrix oMatrix, ref SAPbouiCOM.Form oFormulario)
        {

            oFormulario.Freeze(true);
            try
            {
                oMatrix.FlushToDataSource();
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (oMatrix.IsRowSelected(i))
                    {
                        oMatrix.DeleteRow(i);
                        if (oFormulario.Mode == BoFormMode.fm_OK_MODE) oFormulario.Mode = BoFormMode.fm_UPDATE_MODE;
                        break;
                    }
                }

                oMatrix.FlushToDataSource();
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }

            oFormulario.Freeze(false);
        }

        public static bool GraboDeDataSource(SAPbouiCOM.DBDataSource oDBDataSource, string cTabla, int Digitos = 0)
        {
            string cCode;
            int nOK;
            DateTime dAuxiliar;
            SAPbobsCOM.UserTable oUserTable = Matriz.oGlobal.compañia.UserTables.Item(cTabla);

            try
            {
                for (int i = 0; i <= oDBDataSource.Size - 1; i++)
                {
                    cCode = oDBDataSource.GetValue("Code", i).Trim();

                    bool lCrearNuevo = true;
                    if (cCode != "")
                    {
                        lCrearNuevo = !oUserTable.GetByKey(cCode);
                    }

                    if (!lCrearNuevo)
                    {
                        oUserTable.GetByKey(cCode);
                        #region si existe...

                        foreach (SAPbouiCOM.Field oField in oDBDataSource.Fields)
                        {
                            if (oField.Name != "Code" && oField.Name != "Name")
                            {
                                switch (oField.Type)
                                {
                                    case (BoFieldsType.ft_Date):
                                        {
                                            if (DateTime.TryParseExact(oDBDataSource.GetValue(oField.Name, i), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dAuxiliar))
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = dAuxiliar;
                                            }
                                            else
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = "";
                                            }
                                        }
                                        break;
                                    case BoFieldsType.ft_Float:
                                    case BoFieldsType.ft_Percent:
                                    case BoFieldsType.ft_Measure:
                                    case BoFieldsType.ft_Price:
                                    case BoFieldsType.ft_Quantity:
                                    case BoFieldsType.ft_Rate:
                                    case BoFieldsType.ft_Sum:
                                        oUserTable.UserFields.Fields.Item(oField.Name).Value = Conversiones.ValueSAPToDoubleSistema(oDBDataSource.GetValue(oField.Name, i));
                                        break;

                                    default:
                                        {
                                            oUserTable.UserFields.Fields.Item(oField.Name).Value = oDBDataSource.GetValue(oField.Name, i);
                                        }
                                        break;
                                }
                            }
                        }

                        nOK = oUserTable.Update();
                        if (nOK != 0)
                        {
                            Matriz.oGlobal.SBOApp.SetStatusBarMessage(Matriz.oGlobal.compañia.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                            return false;
                        }
                        #endregion
                    }
                    else
                    {
                        #region si no existe
                        oUserTable = Matriz.oGlobal.compañia.UserTables.Item(cTabla);

                        foreach (SAPbouiCOM.Field oField in oDBDataSource.Fields)
                        {
                            if (oField.Name != "Code" && oField.Name != "Name")
                            {
                                switch (oField.Type)
                                {
                                    case (BoFieldsType.ft_Date):
                                        {
                                            if (DateTime.TryParseExact(oDBDataSource.GetValue(oField.Name, i), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dAuxiliar))
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = dAuxiliar;
                                            }
                                            else
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = "";
                                            }
                                        }
                                        break;

                                    case BoFieldsType.ft_Float:
                                    case BoFieldsType.ft_Percent:
                                    case BoFieldsType.ft_Measure:
                                    case BoFieldsType.ft_Price:
                                    case BoFieldsType.ft_Quantity:
                                    case BoFieldsType.ft_Rate:
                                    case BoFieldsType.ft_Sum:
                                        oUserTable.UserFields.Fields.Item(oField.Name).Value = Conversiones.ValueSAPToDoubleSistema(oDBDataSource.GetValue(oField.Name, i));
                                        break;

                                    default:
                                        oUserTable.UserFields.Fields.Item(oField.Name).Value = oDBDataSource.GetValue(oField.Name, i);
                                        break;
                                }
                            }
                            else if (oField.Name == "Code")
                            {
                                #region Ultimo code
                                string sqlUlt = "SELECT MAX( CAST(T0.Code AS NUMERIC(10))) FROM [@" + cTabla + "] T0";
                                double nMaxCode = Matriz.oGlobal.refDi.SQL.sqlNumericaB1(sqlUlt);
                                string cNuevoCode = "";
                                if (Digitos == -1)
                                {
                                    cNuevoCode = Convert.ToString((Convert.ToInt32(nMaxCode) + 1));
                                }
                                else if (Digitos == 0)
                                {
                                    string cNumCeros = Convert.ToString(nMaxCode);
                                    string cFormatAux = "".PadRight(cNumCeros.Length, '0');
                                    cNuevoCode = (Convert.ToInt32(nMaxCode) + 1).ToString(cFormatAux);
                                }
                                else
                                {
                                    string cFormatAux = "".PadRight(Digitos, '0');
                                    cNuevoCode = (Convert.ToInt32(nMaxCode) + 1).ToString(cFormatAux);
                                }
                                #endregion
                                oUserTable.Code = cNuevoCode;
                                oUserTable.Name = oUserTable.Code;
                            }
                        }

                        nOK = oUserTable.Add();
                        if (nOK != 0)
                        {
                            Matriz.oGlobal.SBOApp.SetStatusBarMessage(Matriz.oGlobal.compañia.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                            return false;
                        }
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.SetStatusBarMessage(Matriz.oGlobal.compañia.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                
                GC.WaitForPendingFinalizers();
            }

            return true;

        }

        public static void BorroDataTable(ref SAPbouiCOM.DataTable oTablaInf)
        {            
            if (!oTablaInf.IsEmpty)
            {
              int nNumReg = oTablaInf.Rows.Count;
              for (int i = 0; i < nNumReg; i++)
              {
                  oTablaInf.Rows.Remove(0);
              }
            }            
        }


        public static void LLenoComboGenerico(ref SAPbouiCOM.Item oItemCombo, string cTabla)
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox) oItemCombo.Specific;
            string sql = "SELECT T0.Code, T0.Name FROM [" + cTabla + "] T0 ORDER BY T0.Name ";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oItemCombo.DisplayDesc = true;
            oCombo.ExpandType = BoExpandType.et_DescriptionOnly;
        }


        public static void LLenoComboGenerico(ref SAPbouiCOM.Column  oColumCombo, string cTabla)
        {            
            string sql = "SELECT T0.Code, T0.Name FROM [" + cTabla + "] T0 ORDER BY T0.Name ";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oColumCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oColumCombo.DisplayDesc = true;
            oColumCombo.ExpandType = BoExpandType.et_DescriptionOnly;
        }


        public static SAPbouiCOM.Form BuscoFormLanzado(string TypeEx)
        {
            SAPbouiCOM.Form oFORMORET = null;
            for (int i = 0; i < Matriz.oGlobal.SBOApp.Forms.Count; i++)
            {
                if (Matriz.oGlobal.SBOApp.Forms.Item(i).TypeEx == TypeEx)
                {
                    oFORMORET = Matriz.oGlobal.SBOApp.Forms.GetForm(Matriz.oGlobal.SBOApp.Forms.Item(i).TypeEx, Matriz.oGlobal.SBOApp.Forms.Item(i).TypeCount);                                        
                    break;
                }
            }

            return oFORMORET;
        }

        public static bool LanzoMenuUserTable(string cTablaSinArroba)
        {
           bool lRetorno = false;
           SAPbouiCOM.Menus oMenus = Matriz.oGlobal.SBOApp.Menus.Item("51200").SubMenus;
           for (int i = 0; i <= oMenus.Count - 1; i++)
           {
               if (oMenus.Item(i).String.IndexOf(cTablaSinArroba) == 0)
              {
                 Matriz.oGlobal.SBOApp.ActivateMenuItem(oMenus.Item(i).UID);
                 lRetorno = true;
                 break;                 
               }
            }


           EXO_CleanCOM.CLiberaCOM.Menus(oMenus);
           return lRetorno;
        }

        public static string leerFichEmbebido(ref Type tipo, string fichero)
        {
            string result = "";
            try
            {
                Assembly assembly = tipo.Assembly;           
                StreamReader streamReader = new StreamReader(tipo.Assembly.GetManifestResourceStream( tipo.Namespace + "." + fichero));
                result = streamReader.ReadToEnd();
                result = result.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
                streamReader.Close();
            }
            catch (Exception expr_44)
            {
                //ProjectData.SetProjectError(expr_44);
                //ProjectData.ClearProjectError();
            }
            return result;
        }
        
        //public static string LeoQueryFich(string cNomFichLargo)
        //{
        //    string sql = "", cAux = "";                           
        //    System.IO.StreamReader Fichero = new System.IO.StreamReader(cNomFichLargo);
        //    while (Fichero.Peek() != -1)
        //    {
        //      cAux = Fichero.ReadLine();
        //      if (cAux.Length > 2 && cAux.Substring(0, 2) == "--") continue;

        //      sql += cAux.Replace("\t", " ") + " ";
        //    }
        //    Fichero.Close();
                        
        //    return sql;
        //}

        //public static string LeoQueryFich(string cNomQueryIncrustada, Type Tipo)
        //{
        //    //string cQuery = Matriz.oGlobal.Functions.leerEmbebido(ref Tipo, cNomQueryIncrustada);
        //    string cQuery = leerFormEmbebido(ref Tipo, cNomQueryIncrustada);
        //    cQuery = cQuery.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");

        //    return cQuery;
        //}

        //public static void ActualizarFormularioSAPXML(string ArchivoXml, SAPbouiCOM.Form oForm)
        //{
        //    System.Xml.XmlDocument oXMLDoc = new System.Xml.XmlDocument();
        //    oXMLDoc.Load(ArchivoXml);
            

        //    System.Xml.XmlNode oNode;
        //    oNode = oXMLDoc.SelectSingleNode("Application/forms/action/form/@uid");
        //    oNode.InnerText = oForm.UniqueID;

        //    string Xml = oXMLDoc.InnerXml.ToString();
        //    Matriz.oGlobal.SBOApp.LoadBatchActions(ref Xml);            
        //}


   
        public static double Evaluate(string expression)
        {
        //    var loDataTable = new System.Data.DataTable();
        //    var loDataColumn = new System.Data.DataColumn("Eval", typeof(double), expression);
        //    loDataTable.Columns.Add(loDataColumn);
        //    loDataTable.Rows.Add(0);
        //    return (double)(loDataTable.Rows[0]["Eval"]);

            return 0;
        }

        public static void DeshabilitoMenus(ref SAPbouiCOM.Form oForm)
        {
            oForm.EnableMenu("1281", false);
            oForm.EnableMenu("1282", false);
            oForm.EnableMenu("1290", false);
            oForm.EnableMenu("1288", false);
            oForm.EnableMenu("1289", false);
            oForm.EnableMenu("1291", false);
        }

    }

}

