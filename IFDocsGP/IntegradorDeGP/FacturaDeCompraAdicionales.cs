using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Data;
using System.Data.SqlClient;

using Comun;
using OfficeOpenXml;

namespace IntegradorDeGP
{
    class FacturaDeCompraAdicionales
    {
        public int iError;
        public string sMensaje;
        public XmlDocument _xDocXml;

        private ConexionDB _DatosConexionDB;
        private FacturaDeCompraPOP _facturaPop;
        private FacturaDeCompraPM _docPm;

        public FacturaDeCompraAdicionales(ConexionDB DatosConexionDB, FacturaDeCompraPOP facturaPop)
        {
            _DatosConexionDB = DatosConexionDB;
            _facturaPop = facturaPop;
        }

        public FacturaDeCompraAdicionales(ConexionDB DatosConexionDB, FacturaDeCompraPM facturaPm)
        {
            _DatosConexionDB = DatosConexionDB;
            _docPm = facturaPm;
        }

        /// <summary>
        /// Arma los datos adicionales de la factura de compra usando xml. (Necesita un schema)
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="filaXl"></param>
        /// <param name="param"></param>
        public void armaDatosAdicionales(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            try
            {
                iError = 0;

                _xDocXml = new XmlDocument();
                _xDocXml.LoadXml("<spEconn_nsacoa_gl00021></spEconn_nsacoa_gl00021>");
                //_xDocXml.PrependChild(_xDocXml.CreateXmlDeclaration("1.0", "utf-8", ""));

                XmlNode datosAdicionales = _xDocXml.GetElementsByTagName("spEconn_nsacoa_gl00021")[0];

                //XmlNode datosAdicionales = _xDocXml.CreateElement("datosAdicionales");
                //datosAdicionales.AppendChild(datosAdicionales);

                XmlElement VENDORID = _xDocXml.CreateElement("VENDORID");
                datosAdicionales.AppendChild(VENDORID);
                XmlElement DOCNUMBR = _xDocXml.CreateElement("DOCNUMBR");
                datosAdicionales.AppendChild(DOCNUMBR);
                XmlElement DOCTYPE = _xDocXml.CreateElement("DOCTYPE");
                datosAdicionales.AppendChild(DOCTYPE);

                XmlElement sDOCDATE = _xDocXml.CreateElement("sDOCDATE");
                datosAdicionales.AppendChild(sDOCDATE);
                XmlElement sDATERECD = _xDocXml.CreateElement("sDATERECD");
                datosAdicionales.AppendChild(sDATERECD);
                XmlElement nsa_tipo_comprob = _xDocXml.CreateElement("nsa_tipo_comprob");
                datosAdicionales.AppendChild(nsa_tipo_comprob);
                XmlElement nsa_cod_transac = _xDocXml.CreateElement("nsa_cod_transac");
                datosAdicionales.AppendChild(nsa_cod_transac);
                XmlElement nsa_autorizacion = _xDocXml.CreateElement("nsa_autorizacion");
                datosAdicionales.AppendChild(nsa_autorizacion);
                XmlElement nsa_cred_trib = _xDocXml.CreateElement("nsa_cred_trib");
                datosAdicionales.AppendChild(nsa_cred_trib);
                XmlElement nsa_cod_iva1 = _xDocXml.CreateElement("nsa_cod_iva1");
                datosAdicionales.AppendChild(nsa_cod_iva1);
                XmlElement nsa_cod_iva2 = _xDocXml.CreateElement("nsa_cod_iva2");
                datosAdicionales.AppendChild(nsa_cod_iva2);
                XmlElement nsaCoa_secuencial = _xDocXml.CreateElement("nsaCoa_secuencial");
                datosAdicionales.AppendChild(nsaCoa_secuencial);
                XmlElement nsa_serie = _xDocXml.CreateElement("nsa_serie");
                datosAdicionales.AppendChild(nsa_serie);
                XmlElement nsaCoa_date_nota = _xDocXml.CreateElement("nsaCoa_date_nota");
                datosAdicionales.AppendChild(nsaCoa_date_nota);
                XmlElement nsa_tipo_comprob_mod = _xDocXml.CreateElement("nsa_tipo_comprob_mod");
                datosAdicionales.AppendChild(nsa_tipo_comprob_mod);
                XmlElement nsa_sernota = _xDocXml.CreateElement("nsa_sernota");
                datosAdicionales.AppendChild(nsa_sernota);
                XmlElement nsacoa_secuencial_mod = _xDocXml.CreateElement("nsacoa_secuencial_mod");
                datosAdicionales.AppendChild(nsacoa_secuencial_mod);

                VENDORID.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCaVENDORID)].Value.ToString();
                DOCNUMBR.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCaVNDDOCNM)].Value.ToString();
                DOCTYPE.InnerText= "1";
                //DateTime fecha = Convert.ToDateTime(hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCareceiptdate)].Value, param.culture);
                //DateTime fecha = DateTime.Parse(hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCareceiptdate)].Value.ToString(), param.culture, System.Globalization.DateTimeStyles.AssumeLocal);
                long serialDate = long.Parse(hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCareceiptdate)].Value.ToString());
                DateTime fecha = DateTime.FromOADate(serialDate);

                sDOCDATE.InnerText = String.Format("{0:yyyyMMdd}", fecha);
                sDATERECD.InnerText = String.Format("{0:yyyyMMdd}", fecha);

                if (param.nsa_tipo_comprob.ToLower().Equals("na"))
                    nsa_tipo_comprob.InnerText = param.nsa_tipo_comprob_default;
                else
                    nsa_tipo_comprob.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString();

                if (param.nsa_cod_transac.ToLower().Equals("na"))
                    nsa_cod_transac.InnerText = param.nsa_cod_transac_default;
                else
                    nsa_cod_transac.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_transac)].Value.ToString();

                if (param.nsa_autorizacion.ToLower().Equals("na"))
                    nsa_autorizacion.InnerText = param.nsa_autorizacion_default;
                else
                    nsa_autorizacion.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_autorizacion)].Value.ToString();

                if (param.nsa_cred_trib.ToLower().Equals("na"))
                    nsa_cred_trib.InnerText = param.nsa_cred_trib_default;
                else
                    nsa_cred_trib.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cred_trib)].Value.ToString();

                if (param.nsa_cod_iva1.ToLower().Equals("na"))
                    nsa_cod_iva1.InnerText = param.nsa_cod_iva1_default;
                else
                    nsa_cod_iva1.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_iva1)].Value.ToString();

                if (param.nsa_cod_iva2.ToLower().Equals("na"))
                    nsa_cod_iva2.InnerText = param.nsa_cod_iva2_default;
                else
                    nsa_cod_iva2.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_iva2)].Value.ToString();

                nsaCoa_secuencial.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsaCoa_secuencial)].Value.ToString();
                nsa_serie.InnerText = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_serie)].Value.ToString();
                nsaCoa_date_nota.InnerText = "19000101";
                nsa_tipo_comprob_mod.InnerText = "0";
                nsa_sernota.InnerText = "0";
                nsacoa_secuencial_mod.InnerText = "0";

            }
            catch (Exception errorGral)
            {
                sMensaje = "Error al armar datos adicionales de factura. " + errorGral.Message + " [armaDatosAdicionales]";
                iError++;
            }

        }

        public void spIfcAgregaDistribucionContable(string poprctnm)
        {
            iError = 0;
            sMensaje = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(_DatosConexionDB.Elemento.ConnStr))
                {
                    using (SqlCommand cmd = new SqlCommand("dbo.spIfcAgregaDistribucionContable", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@POPRCTNM", SqlDbType.Char, 17);                //0
                        cmd.Parameters.Add("@O_iErrorState", SqlDbType.Int);                
                        cmd.Parameters.Add("@oErrString", SqlDbType.VarChar, 255);

                        cmd.Parameters[0].Value = poprctnm;
                        cmd.Parameters["@O_iErrorState"].Direction = ParameterDirection.Output;
                        cmd.Parameters["@oErrString"].Direction = ParameterDirection.Output;

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        iError = Convert.ToInt32(cmd.Parameters["@O_iErrorState"].Value);
                        sMensaje = cmd.Parameters["@oErrString"].Value.ToString();
                        conn.Close();
                    }
                }
            }
            catch (Exception errorGral)
            {
                sMensaje = "Trip codes no agregados al asiento contable debido a: " + errorGral.Message + " [spIfcAgregaDistribucionContable]";
                iError++;
            }
        }

        public void spEconn_nsacoa_gl00021(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            iError = 0;
            sMensaje = "";
            int campo = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(_DatosConexionDB.Elemento.ConnStr))
                {
                    using (SqlCommand cmd = new SqlCommand("dbo.spEconn_nsacoa_gl00021", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@VENDORID", SqlDbType.Char, 15);                //0
                        cmd.Parameters.Add("@DOCNUMBR", SqlDbType.Char, 21);                //1
                        cmd.Parameters.Add("@DOCTYPE", SqlDbType.SmallInt);                 //2
                        cmd.Parameters.Add("@sDOCDATE", SqlDbType.VarChar, 20);             //3
                        cmd.Parameters.Add("@sDATERECD", SqlDbType.VarChar, 20);            //4
                        cmd.Parameters.Add("@nsa_tipo_comprob", SqlDbType.Char, 3);         //5  tipo de comprobante: 01 factura dos dígitos
                        cmd.Parameters.Add("@nsa_cod_transac", SqlDbType.Char, 3);          //6  Adquisiciones Gravadas destinadas a operaciones: 01 Gravadas y/o de exportación, 02 Gravadas y/o de export. y a operaciones no gravadas, 03 No gravadas
                        cmd.Parameters.Add("@nsa_autorizacion", SqlDbType.Char, 11);        //7  No utilizado. Indicar na
                        cmd.Parameters.Add("@nsa_cred_trib", SqlDbType.Char, 11);           //8  Tipo de operación sujeta al sistema de detracciones: 01 Venta de bienes, 02 Retiro de bienes, 03 Traslado de bienes
                        cmd.Parameters.Add("@nsa_cod_iva1", SqlDbType.Char, 3);             //9  Si el comprobante es sujeto a retención indicar 01, sino 00. Sirve para el libro de compras electrónico
                        cmd.Parameters.Add("@nsa_cod_iva2", SqlDbType.Char, 3);             //10 Código de bien o servicio sujeto a detracción
                        cmd.Parameters.Add("@nsaCoa_secuencial", SqlDbType.Char, 7);        //11 Número de la factura
                        cmd.Parameters.Add("@nsa_serie", SqlDbType.Char, 7);                //12 Serie de la factura
                        cmd.Parameters.Add("@nsaCoa_date_nota", SqlDbType.VarChar, 20);     //13 No utilizado por la app. Van todos en cero por default
                        cmd.Parameters.Add("@nsa_tipo_comprob_mod", SqlDbType.Char, 3);     //14
                        cmd.Parameters.Add("@nsa_sernota", SqlDbType.Char, 7);              //15
                        cmd.Parameters.Add("@nsacoa_secuencial_mod", SqlDbType.Char, 7);    //16
                        cmd.Parameters.Add("@O_iErrorState", SqlDbType.Int);                //17
                        cmd.Parameters.Add("@oErrString", SqlDbType.VarChar, 255);

                        cmd.Parameters[0].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCaVENDORID)].Value.ToString();
                        cmd.Parameters[1].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.facturaPopCaVNDDOCNM)].Value.ToString();
                        cmd.Parameters[2].Value = "1";
                        cmd.Parameters[3].Value = String.Format("{0:yyyyMMdd}", _facturaPop.fechaFactura);
                        cmd.Parameters[4].Value = String.Format("{0:yyyyMMdd}", _facturaPop.fechaFactura);

                        campo = 5;
                        if (param.nsa_tipo_comprob.ToLower().Equals("na"))
                            cmd.Parameters[5].Value = param.nsa_tipo_comprob_default;
                        else
                            cmd.Parameters[5].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString().PadLeft(2, '0');

                        campo = 6;
                        if (param.nsa_cod_transac.ToLower().Equals("na"))
                            cmd.Parameters[6].Value = param.nsa_cod_transac_default;
                        else
                            cmd.Parameters[6].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_transac)].Value.ToString().PadLeft(2, '0');

                        campo = 7;
                        if (param.nsa_autorizacion.ToLower().Equals("na"))
                            cmd.Parameters[7].Value = param.nsa_autorizacion_default;
                        else
                            cmd.Parameters[7].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_autorizacion)].Value.ToString();

                        campo = 8;
                        if (param.nsa_cred_trib.ToLower().Equals("na"))
                            cmd.Parameters[8].Value = param.nsa_cred_trib_default;
                        else
                            cmd.Parameters[8].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cred_trib)].Value.ToString().PadLeft(2, '0');

                        campo = 9;
                        if (param.nsa_cod_iva1.ToLower().Equals("na"))
                            cmd.Parameters[9].Value = param.nsa_cod_iva1_default;
                        else
                            cmd.Parameters[9].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_iva1)].Value.ToString().PadLeft(2, '0');

                        campo = 10;
                        if (param.nsa_cod_iva2.ToLower().Equals("na"))
                            cmd.Parameters[10].Value = param.nsa_cod_iva2_default;
                        else
                        {
                            if (hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_iva2)].Value != null)
                                cmd.Parameters[10].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_cod_iva2)].Value.ToString().PadLeft(3, '0');
                            else
                                cmd.Parameters[10].Value = "";
                        }
                        campo = 11;
                        cmd.Parameters[11].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsaCoa_secuencial)].Value.ToString().Trim();
                        campo = 12;
                        cmd.Parameters[12].Value = hojaXl.Cells[filaXl, Convert.ToInt32(param.nsa_serie)].Value.ToString().Trim();
                        cmd.Parameters[13].Value = "19000101";
                        cmd.Parameters[14].Value = "0";
                        cmd.Parameters[15].Value = "0";
                        cmd.Parameters[16].Value = "0";
                        cmd.Parameters["@O_iErrorState"].Direction = ParameterDirection.Output;
                        cmd.Parameters["@oErrString"].Direction = ParameterDirection.Output;

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        iError = Convert.ToInt32(cmd.Parameters["@O_iErrorState"].Value);
                        sMensaje = cmd.Parameters["@oErrString"].Value.ToString();
                        conn.Close();
                    }
                }
            }
            catch (Exception errorGral)
            {
                sMensaje = "Datos adicionales de factura no integrados debido a: " + errorGral.Message + " [spEconn_nsacoa_gl00021]";
                if (campo == 5)
                    sMensaje += "Error en Tipo de comprobante";
                if (campo == 6)
                    sMensaje += "Error en Tipo de Adquisiciones Gravadas";
                if (campo == 7)
                    sMensaje += "No utilizado. Indicar na";
                if (campo == 8)
                    sMensaje += "Error en Tipo de operación sujeta al sistema de detracciones";
                if (campo == 9)
                    sMensaje += "Error en el indicador de comprobante sujeto a retención";
                if (campo == 10)
                    sMensaje += "Error en el código de bien o servicio sujeto a detracción";
                if (campo == 11)
                    sMensaje += "Error en el Número de la factura";
                if (campo == 12)
                    sMensaje += "Error en la Serie de la factura";
                iError++;
            }
        }

        /// <summary>
        /// Datos adicionales de facturas para el servicio de impuestos de Bolivia
        /// </summary>
        public void spIfc_AgregaTII_4001()
        {
            string numControl = _docPm.facturaPm.USRDEFND1;
            Int16 tipoComprob = 1;                  //No se incluye en libros

            if (_docPm.facturaPm.DOCTYPE == 1)      //factura
                tipoComprob = 2;                    //Sí se incluye en libros

            if (_docPm.facturaPm.USRDEFND1 == null || _docPm.facturaPm.USRDEFND1.Equals(""))
                numControl = "0";

            iError = 0;
            sMensaje = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(_DatosConexionDB.Elemento.ConnStr))
                {
                    using (SqlCommand cmd = new SqlCommand("dbo.spIfc_AgregaTII_4001", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@VCHRNMBR", SqlDbType.Char, 21);                //0
                        cmd.Parameters.Add("@DOCTYPE", SqlDbType.SmallInt);                 //1
                        cmd.Parameters.Add("@TII_Num_Autorizacion", SqlDbType.Char, 21);    //2
                        cmd.Parameters.Add("@TII_Num_Control", SqlDbType.Char, 21);         //3 
                        cmd.Parameters.Add("@VENDORID", SqlDbType.Char, 15);                //4
                        cmd.Parameters.Add("@DOCNUMBR", SqlDbType.Char, 21);                //5
                        cmd.Parameters.Add("@Tipo_Comprobante", SqlDbType.SmallInt);        //6  

                        cmd.Parameters.Add("@O_iErrorState", SqlDbType.Int);
                        cmd.Parameters.Add("@oErrString", SqlDbType.VarChar, 255);

                        cmd.Parameters[0].Value = _docPm.facturaPm.VCHNUMWK;
                        cmd.Parameters[1].Value = _docPm.facturaPm.DOCTYPE;
                        cmd.Parameters[2].Value = _docPm.facturaPm.USRDEFND2;
                        cmd.Parameters[3].Value = numControl;
                        cmd.Parameters[4].Value = _docPm.facturaPm.VENDORID;
                        cmd.Parameters[5].Value = _docPm.facturaPm.DOCNUMBR;
                        cmd.Parameters[6].Value = tipoComprob;

                        cmd.Parameters["@O_iErrorState"].Direction = ParameterDirection.Output;
                        cmd.Parameters["@oErrString"].Direction = ParameterDirection.Output;

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        iError = Convert.ToInt32(cmd.Parameters["@O_iErrorState"].Value);
                        sMensaje = cmd.Parameters["@oErrString"].Value.ToString();
                        conn.Close();
                    }
                }
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción al ingresar datos para el servicio de impuestos. " + errorGral.Message + " [spIfc_AgregaTII_4001]";
                iError++;
            }
        }
        /// <summary>
        /// Retenciones de factura para localización argentina
        /// </summary>
        public void spIfc_Nfret_gl10030()
        {
            iError = 0;
            sMensaje = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(_DatosConexionDB.Elemento.ConnStr))
                {
                    using (SqlCommand cmd = new SqlCommand("dbo.spIfc_Nfret_gl10030", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@VENDORID", SqlDbType.Char, 15);                //0
                        cmd.Parameters.Add("@DOCTYPE", SqlDbType.SmallInt);                 //1
                        cmd.Parameters.Add("@DOCNUMBR", SqlDbType.Char, 21);                //2
                        cmd.Parameters.Add("@VCHRNMBR", SqlDbType.Char, 21);                //3
                        cmd.Parameters.Add("@nfRET_plan_de_retencione", SqlDbType.Char, 21);    //4
                        cmd.Parameters.Add("@nfRET_Applied", SqlDbType.SmallInt);           //5

                        cmd.Parameters[0].Value = _docPm.facturaPm.VENDORID;
                        cmd.Parameters[1].Value = _docPm.facturaPm.DOCTYPE;
                        cmd.Parameters[2].Value = _docPm.facturaPm.DOCNUMBR;
                        cmd.Parameters[3].Value = _docPm.facturaPm.VCHNUMWK;
                        cmd.Parameters[4].Value = _docPm.facturaPm.USRDEFND2 == null ? String.Empty : _docPm.facturaPm.USRDEFND2;
                        cmd.Parameters[5].Value = 0;

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción al ingresar retenciones en la localización argentina. " + errorGral.Message + " [spIfc_Nfret_gl10030]";
                iError++;
                throw;
            }
        }

        /// <summary>
        /// Datos adicionales de facturas pm para localización argentina
        /// </summary>
        public void spIfc_awli_pm00400()
        {
            iError = 0;
            sMensaje = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(_DatosConexionDB.Elemento.ConnStr))
                {
                    using (SqlCommand cmd = new SqlCommand("dbo.spIfc_awli_pm00400", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@VENDORID", SqlDbType.Char, 15);                //0
                        cmd.Parameters.Add("@DOCTYPE", SqlDbType.SmallInt);                 //1
                        cmd.Parameters.Add("@DOCNUMBR", SqlDbType.Char, 21);                //2
                        cmd.Parameters.Add("@VCHRNMBR", SqlDbType.Char, 21);                //3

                        cmd.Parameters[0].Value = _docPm.facturaPm.VENDORID;
                        cmd.Parameters[1].Value = _docPm.facturaPm.DOCTYPE;
                        cmd.Parameters[2].Value = _docPm.facturaPm.DOCNUMBR;
                        cmd.Parameters[3].Value = _docPm.facturaPm.VCHNUMWK;

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción al ingresar datos adicionales de factura en la localización argentina. " + errorGral.Message + " [spIfc_awli_pm00400]";
                iError++;
                throw;
            }
        }

    }

}
