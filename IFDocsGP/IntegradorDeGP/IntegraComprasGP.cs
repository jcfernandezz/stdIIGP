using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Globalization;

using Comun;
using OfficeOpenXml;
using ManipulaArchivos;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;

namespace IntegradorDeGP
{
    public class IntegraComprasGP
    {
        public int iError;
        public string sMensaje;

        private ConexionDB _DatosConexionDB;
        private Parametros _Param;

        private XmlDocument _xDocXml;
        private string _sDocXml = "";
        private string sMensajeDocu = "";
        private int _filaNuevaFactura = 0;

        public delegate void LogHandler(int iAvance, string sMsj);
        public event LogHandler Progreso;

        /// <summary>
        /// Dispara el evento para actualizar la barra de progreso
        /// </summary>
        /// <param name="iProgreso"></param>
        public void OnProgreso(int iAvance, string sMsj)
        {
            if (Progreso != null) 
                Progreso(iAvance, sMsj);
        }

        public event LogHandler Actualiza;
        public void OnActualiza(int i,  string carpeta)
        {
            if (Actualiza != null)
                Actualiza(i, carpeta);
        }

        public IntegraComprasGP(ConexionDB DatosConexionDB)
        {
            this.iError = 0;
            _DatosConexionDB = DatosConexionDB;                                                //Lee la configuración y obtiene los datos de conexión.
            _Param = new Parametros(_DatosConexionDB.NombreArchivoParametros, DatosConexionDB.Elemento.Intercompany);

            if (_Param.iError != 0)
            {
                this.iError++;
                this.sMensaje = _Param.ultimoMensaje;
            }
        }

        //public Decimal getImpuestos(string id)
        //{
        //    tx00201 impuestos = new tx00201(DatosConexionDB.Elemento.ConnStr);
        //    try
        //    {
        //        if (impuestos.LoadByPrimaryKey(id))
        //        {
        //            return impuestos.TXDTLPCT;
        //        }
        //        else
        //            return 0;
        //    }
        //    catch
        //    {
        //        return 0;
        //    }
        //}

        /// <summary>
        /// Construye documento xml en un xmlDocument.
        /// </summary>
        /// <param name="eConnect"></param>
        public void serializa(eConnectType eConnect)
        {
            try
            {
                iError = 0;
                _sDocXml = "";
                _xDocXml = new XmlDocument();
                StringBuilder sbDocXml = new StringBuilder();

                XmlSerializer serializer = new XmlSerializer(eConnect.GetType());
                XmlWriterSettings sett = new XmlWriterSettings();
                sett.Encoding = new UTF8Encoding();  //UTF8Encoding.UTF8; // Encoding.UTF8;
                using (XmlWriter writer = XmlWriter.Create(sbDocXml, sett))
                {
                    serializer.Serialize(writer, eConnect);
                    _sDocXml = sbDocXml.ToString();
                    _xDocXml.LoadXml(_sDocXml);
                }
            }
            catch (Exception errorGral)
            {
                sMensaje = "Error al serializar el documento. " + errorGral.Message + " [Serializa]";
                iError++;
            }

        }

        /// <summary>
        /// Crea el xml de una factura pop a partir de una fila de datos en una hoja excel.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void integraFacturaPOP(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp)
        {
            this.iError = 0;
            eConnectType docEConnectIV = new eConnectType();
            eConnectType docEConnectPOP = new eConnectType();
            FacturaDeCompraPOP facturaPop = new FacturaDeCompraPOP(_DatosConexionDB, 1);

            try
            {
                facturaPop.preparaFacturaPOP(hojaXl, filaXl, sTimeStamp, _Param);
                _filaNuevaFactura = facturaPop.iniciaNuevaFacturaEn;
                this.sMensajeDocu = "Fila: " + filaXl.ToString() + " Número Doc: " + facturaPop.facturaPopCa.VNDDOCNM + " Proveedor: " + facturaPop.facturaPopCa.VENDORID;

                //Ingresa el artículo - proveedor a GP
                if (facturaPop.iError == 0)
                {
                    docEConnectIV.IVVendorItemType = facturaPop.myVendorItemType;
                    this.serializa(docEConnectIV);
                    if (this.iError == 0)
                        this.integraEntityXml();
                }
                else
                {
                    this.sMensaje = facturaPop.sMensaje;
                    this.iError++;
                }

                //Ingresa la factura a GP
                if (this.iError == 0 && facturaPop.iError == 0)
                {
                    
                    docEConnectPOP.POPReceivingsType = facturaPop.myFacturaPopType;
                    this.serializa(docEConnectPOP);
                    if (this.iError == 0)
                        this.integraTransactionXml();
                }

                //Si no hubo error y no es factura cogs, agrega datos para el servicio de impuestos
                FacturaDeCompraAdicionales adicionalesFactura = new FacturaDeCompraAdicionales(_DatosConexionDB, facturaPop);
                if (this.iError == 0 && facturaPop.iError == 0 && facturaPop.facturaPopCa.REFRENCE.Length > 1 && facturaPop.facturaPopCa.REFRENCE.Substring(0, 2) != "CO")
                {
                    adicionalesFactura.spEconn_nsacoa_gl00021(hojaXl, filaXl, _Param);
                    if (adicionalesFactura.iError != 0)
                    {
                        this.iError++;
                        this.sMensaje = adicionalesFactura.sMensaje;
                    }
                }

                //Si no hubo errores, agregar trip codes a los asientos contables
                if (this.iError == 0 && facturaPop.iError ==0 )
                {
                    adicionalesFactura.spIfcAgregaDistribucionContable(facturaPop.facturaPopCa.POPRCTNM);
                    if (adicionalesFactura.iError != 0)
                    {
                        this.iError++;
                        this.sMensaje = adicionalesFactura.sMensaje;
                    }
                }

            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción al preparar factura. " + eConnErr.Message + "[integraFacturaPOP]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + "[integraFacturaPOP]"; 
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida. " + errorGral.Message + " [integraFacturaPOP]";
                iError++;
            }
        }

        /// <summary>
        /// Crea el xml de una factura PM a partir de una fila de datos en una hoja excel. Si no existe el proveedor, lo crea.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void integraFacturaPMyProveedor(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp)
        {
            this.iError = 0;
            eConnectType docEConnectPM = new eConnectType();
            eConnectType docEConnectProv = new eConnectType();
            FacturaDeCompraPM factura = new FacturaDeCompraPM(_DatosConexionDB);
            Proveedor proveedor = new Proveedor(_DatosConexionDB);
            try
            {
                //Preparar nuevo proveedor en caso que sea factura y no exista en gp
                proveedor.preparaProveedorEconn(hojaXl, filaXl, _Param);
                if (proveedor.iError != 0)
                {
                    this.sMensaje = proveedor.sMensaje;
                    this.iError++;
                }

                //Prepara factura
                factura.preparaFacturaPM(hojaXl, filaXl, sTimeStamp, _Param);
                this._filaNuevaFactura = factura.iniciaNuevaFacturaEn;
                this.sMensajeDocu = "Fila: " + filaXl.ToString() + " Número Doc: " + factura.facturaPm.DOCNUMBR + " Proveedor: " + factura.facturaPm.VENDORID + " Monto: " + factura.facturaPm.PRCHAMNT.ToString();

                if (this.iError == 0 && factura.iError != 0)
                {
                    this.sMensaje = factura.sMensaje;
                    this.iError++;
                }

                //Ingresa nuevo proveedor
                if (this.iError == 0 && proveedor.arrVendorType != null && proveedor.arrVendorType.Count() > 0)
                {
                    docEConnectProv.PMVendorMasterType = proveedor.arrVendorType;
                    this.serializa(docEConnectProv);

                    if (this.iError == 0)
                        this.integraEntityXml();
                }

                //Ingresa la factura a GP
                if (this.iError == 0)
                {
                    docEConnectPM.PMTransactionType = factura.arrFacturaPmType;
                    this.serializa(docEConnectPM);

                    //debug!!!!
                    //this.iError++;
                    //sMensaje = _sDocXml;

                    if (this.iError == 0)
                        this.integraTransactionXml();
                }

                //Si es factura agrega datos para el servicio de impuestos
                if (this.iError == 0 && factura.facturaPm.DOCTYPE == 1)
                {
                    FacturaDeCompraAdicionales adicionalesFactura = new FacturaDeCompraAdicionales(_DatosConexionDB, factura);
                    adicionalesFactura.spIfc_AgregaTII_4001();
                    if (adicionalesFactura.iError != 0)
                    {
                        this.sMensaje = adicionalesFactura.sMensaje;
                        this.iError++;
                    }
                }
            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción al preparar factura. " + eConnErr.Message + "[integraFacturaPMyProveedor.integraFacturaPM]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + "[integraFacturaPMyProveedor.integraFacturaPM]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida. " + errorGral.Message + " [integraFacturaPMyProveedor.integraFacturaPM]";
                iError++;
            }
        }

        /// <summary>
        /// Crea el xml de una factura PM a partir de una fila de datos en una hoja excel.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void integraFacturaPM(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp)
        {
            this.iError = 0;
            eConnectType docEConnectPM = new eConnectType();
            eConnectType docEConnectProv = new eConnectType();
            FacturaDeCompraPM factura = new FacturaDeCompraPM(_DatosConexionDB);
            Proveedor proveedor = new Proveedor(_DatosConexionDB);
            try
            {
                //Prepara factura
                factura.preparaFacturaPM(hojaXl, filaXl, sTimeStamp, _Param);
                this._filaNuevaFactura = factura.iniciaNuevaFacturaEn;
                this.sMensajeDocu = "Fila: " + filaXl.ToString() + " Número Doc: " + factura.facturaPm.DOCNUMBR + " Proveedor: " + factura.facturaPm.VENDORID + " Monto: " + factura.facturaPm.PRCHAMNT.ToString();

                if (this.iError == 0 && factura.iError != 0)
                {
                    this.sMensaje = factura.sMensaje;
                    this.iError++;
                }

                //Ingresa la factura a GP
                if (this.iError == 0)
                {
                    docEConnectPM.PMTransactionType = factura.arrFacturaPmType;
                    this.serializa(docEConnectPM);

                    //debug!!!!
                    //this.iError++;
                    //sMensaje = _sDocXml;
                    
                    if (this.iError == 0)
                        this.integraTransactionXml();
                }

                if (_Param.FacturaPmLOCALIZACION.Equals("ARG"))
                {
                    //Si es factura agrega datos para el servicio de impuestos
                    FacturaDeCompraAdicionales adicionalesFactura = new FacturaDeCompraAdicionales(_DatosConexionDB, factura);
                    if (this.iError == 0 && factura.facturaPm.DOCTYPE == 1)
                    {
                        adicionalesFactura.spIfc_Nfret_gl10030();
                        adicionalesFactura.spIfc_awli_pm00400();
                    }

                }

            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción al preparar factura. " + eConnErr.Message + "[IntegraComprasGP.integraFacturaPM]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + "[IntegraComprasGP.integraFacturaPM]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción. " + errorGral.Message + " [IntegraComprasGP.integraFacturaPM]";
                iError++;
            }
        }

        /// <summary>
        /// Crea el xml de un pago manual PM a partir de una fila de datos en una hoja excel.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void integraPagoPM(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp)
        {
            this.iError = 0;
            eConnectType docEConnectPM = new eConnectType();
            PagoManualPM pago = new PagoManualPM(_DatosConexionDB);
            try
            {
                //Prepara pago
                pago.preparaPagoPM(hojaXl, filaXl, sTimeStamp, _Param);
                this._filaNuevaFactura = pago.iniciaNuevoDocEn;
                this.sMensajeDocu = "Fila: " + filaXl.ToString() + " Número Doc: " + pago.pagoPm.DOCNUMBR + " Proveedor: " + pago.pagoPm.VENDORID + " Monto: " + pago.pagoPm.DOCAMNT.ToString();

                if (this.iError == 0 && pago.iError != 0)
                {
                    this.sMensaje = pago.sMensaje;
                    this.iError++;
                }

                //Ingresa el pago a GP
                if (this.iError == 0)
                {
                    docEConnectPM.PMManualCheckType = pago.arrPagoPmType;
                    this.serializa(docEConnectPM);

                    //debug!!!!
                    //this.iError++;
                    //sMensaje = _sDocXml;

                    if (this.iError == 0)
                        this.integraTransactionXml();
                }

            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción al preparar el pago. " + eConnErr.Message + " " + eConnErr.TargetSite.ToString();
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + " " + ex.TargetSite.ToString();
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción. " + errorGral.Message + " " + errorGral.TargetSite.ToString();
                iError++;
            }
        }

        /// <summary>
        /// Integra un documento xml sDocXml a GP.
        /// </summary>
        public void integraEntityXml()
        {
            iError = 0;
            bool eConnResult;
            eConnectMethods eConnObject = new eConnectMethods();

            try
            {
                //El método que integra eConnect_EntryPoint retorna True si fue exitoso
                //eConnResult = eConnObject.eConnect_EntryPoint(_DatosConexionDB.Elemento.ConnStr, EnumTypes.ConnectionStringType.SqlClient, _sDocXml, EnumTypes.SchemaValidationType.None);
                eConnResult = eConnObject.CreateEntity(_DatosConexionDB.Elemento.ConnStr, _sDocXml);

                if (eConnResult)
                    sMensaje = "--> Integrado a GP";
                else
                {
                    iError++;
                    sMensaje = "Error desconocido al crear la entidad eConnect.";
                }
            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción eConnect: " + eConnErr.Message + "[IntegraComprasGP.integraEntityXml]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación: " + ex.Message + "[IntegraComprasGP.integraEntityXml]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida: " + errorGral.Message + "[IntegraComprasGP.integraEntityXml]";
                iError++;
            }
        }

        /// <summary>
        /// Integra un documento xml sDocXml a GP.
        /// </summary>
        public void integraTransactionXml()
        {
            iError = 0;
            string eConnResult;
            eConnectMethods eConnObject = new eConnectMethods();

            try
            {
                //El método que integra eConnect_EntryPoint retorna True si fue exitoso
                //eConnResult = eConnObject.eConnect_EntryPoint(_DatosConexionDB.Elemento.ConnStr, EnumTypes.ConnectionStringType.SqlClient, _sDocXml, EnumTypes.SchemaValidationType.None);
                eConnResult = eConnObject.CreateTransactionEntity(_DatosConexionDB.Elemento.ConnStr, _sDocXml);

                this.sMensaje = "--> Integrado a GP";
                //sMensaje += _sDocXml;

            }
            catch (eConnectException eConnErr)
            {
                this.sMensaje = "Excepción eConnect: " + eConnErr.Message + " " + eConnErr.TargetSite.ToString() + " [integraTransactionXml] "+ _sDocXml;
                this.iError++;
            }
            catch (ApplicationException ex)
            {
                this.sMensaje = "Excepción de aplicación: " + ex.Message + "[integraTransactionXml]"; 
                this.iError++;
            }
            catch (Exception errorGral)
            {
                this.sMensaje = "Excepción desconocida: " + errorGral.Message + "[integraTransactionXml]";
                this.iError++;
            }
        }

        /// <summary>
        /// Abre los archivos excel de una carpeta y los integra a GP.
        /// </summary>
        public void procesaCarpetaEnTrabajo(List<string> archivosSeleccionados)
        {
            try
            {
                this.iError = 0;
                DirectoryInfo enTrabajoDir = new DirectoryInfo(this._Param.rutaCarpeta.ToString() + "\\EnTrabajo");
                archivosExcel archivosEnTrabajo = new archivosExcel();

                foreach(string item in archivosSeleccionados)
                {
                    this.iError = 0;
                    string sTimeStamp = System.DateTime.Now.ToString("yyMMdd.HHmmss");
                    string sNombreArchivo = item;   

                    archivosEnTrabajo.abreArchivoExcel(enTrabajoDir.ToString(), sNombreArchivo);
                    ExcelWorksheet hojaXl = archivosEnTrabajo.paqueteExcel.Workbook.Worksheets.First(); 
                    if (archivosEnTrabajo.iError == 0)
                    {
                        int startRow = _Param.facturaPmFilaInicial;
                        int iTotal = hojaXl.Dimension.End.Row - startRow + 1;
                        int iFacturasIntegradas = 0;
                        int iFilasIntegradas = 0;
                        int iFacturaIniciaEn = 0;
                        int iAntesIntegradas = 0;
                        int _columnaMensajes = 20;

                        bool facturasPm = _DatosConexionDB.NombreArchivoParametros.Contains("facturaspm");
                        bool pagosPm = _DatosConexionDB.NombreArchivoParametros.Contains("pagospm");

                        if (facturasPm)
                            _columnaMensajes = _Param.facturaPmColumnaMensajes;
                        if (pagosPm)
                            _columnaMensajes = _Param.PagosPmColMensajes;


                        OnProgreso(1, "INICIANDO CARGA DE ARCHIVO "+ sNombreArchivo + "...");              //Notifica al suscriptor
                        if (startRow > 1)
                            hojaXl.Cells[startRow - 1, _columnaMensajes].Value = "Observaciones";

                        for (int rowNumber = startRow; rowNumber <= hojaXl.Dimension.End.Row; rowNumber++)
                        {
                            if (hojaXl.Cells[rowNumber, _columnaMensajes].Value == null ||
                                !hojaXl.Cells[rowNumber, _columnaMensajes].Value.ToString().Equals("Integrado a GP"))
                            {
                                if (facturasPm)
                                    integraFacturaPM(hojaXl, rowNumber, "F"+sTimeStamp);
                                else if (pagosPm)
                                    integraPagoPM(hojaXl, rowNumber, "P"+sTimeStamp);
                                else
                                {
                                    iError++;
                                    sMensaje = "No ha ingresado un nombre válido para el archivo de parámetros al iniciar la aplicación. " + _DatosConexionDB.NombreArchivoParametros;
                                }

                                //this.integraFacturaPOP(hojaXl, rowNumber, sTimeStamp);
                                iFacturaIniciaEn = rowNumber;
                                rowNumber = _filaNuevaFactura - 1;

                                if (this.iError == 0)
                                {
                                    iFacturasIntegradas++;
                                    for (int ind = iFacturaIniciaEn; ind <= rowNumber; ind++)
                                    {
                                        hojaXl.Cells[ind, _columnaMensajes].Value = "Integrado a GP";
                                        iFilasIntegradas++;
                                    }
                                }
                                else
                                    hojaXl.Cells[rowNumber, _columnaMensajes].Value = this.sMensaje;
                            }
                            else
                            {
                                iAntesIntegradas++;
                                this.sMensajeDocu = "Fila: " + rowNumber.ToString();
                                this.sMensaje = "anteriormente integrada." ;
                            }
                            OnProgreso(100 / iTotal, this.sMensajeDocu +" "+ this.sMensaje);
                        }
                        OnProgreso(100, "----------------------------------------------");
                        this.sMensaje = "INTEGRACION FINALIZADA";
                        OnProgreso(100, this.sMensaje);
                        OnProgreso(100, "Nuevos documentos integrados: " + iFacturasIntegradas.ToString());
                        OnProgreso(100, "Nuevas filas integradas: " + iFilasIntegradas.ToString());
                        OnProgreso(100, "Número de filas con error: " + (iTotal - iFilasIntegradas - iAntesIntegradas).ToString());
                        OnProgreso(100, "Número de filas anteriormente integradas: " + iAntesIntegradas.ToString());
                        OnProgreso(100, "Total de filas leídas: " + iTotal.ToString());
                        archivosEnTrabajo.paqueteExcel.Save();
                        archivosEnTrabajo.paqueteExcel.Dispose();
                        archivosEnTrabajo.mueveAFinalizado(sNombreArchivo, this._Param.rutaCarpeta.ToString(), sTimeStamp);

                        if (archivosEnTrabajo.iError != 0)
                            OnProgreso(100, archivosEnTrabajo.sMensaje);

                        OnActualiza(0, _Param.rutaCarpeta);
                    }
                    else
                        OnProgreso(0, archivosEnTrabajo.sMensaje);                 
                }
            }
            catch (Exception errorGral)
            {
                this.sMensaje = "Excepción encontrada al leer la carpeta En trabajo. " + errorGral.Message + " " + errorGral.TargetSite.ToString();
                iError++;
                OnProgreso(0, this.sMensaje);                                      
            }
        }
    }
}
