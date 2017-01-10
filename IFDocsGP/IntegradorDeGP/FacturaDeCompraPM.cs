using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Globalization;

using Comun;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using OfficeOpenXml;
using System.Diagnostics;

namespace IntegradorDeGP
{
    class FacturaDeCompraPM
    {
        public int iError = 0;
        public string sMensaje = "";
        public int iniciaNuevaFacturaEn = 0;
        public DateTime fechaFactura = Convert.ToDateTime("1/1/1900");
        private ConexionDB _DatosConexionDB;

        public taPMTransactionInsert facturaPm;
        private taPMDistribution_ItemsTaPMDistribution[] _distribucionPm;
        public PMTransactionType facturaPmType;
        public PMTransactionType[] arrFacturaPmType;
        List<taPMTransactionTaxInsert_ItemsTaPMTransactionTaxInsert> taxDetails = new List<taPMTransactionTaxInsert_ItemsTaPMTransactionTaxInsert>();

        public FacturaDeCompraPM(ConexionDB DatosConexionDB)
        {
            _DatosConexionDB = DatosConexionDB;
            facturaPm = new taPMTransactionInsert();
            facturaPmType = new PMTransactionType();
            _distribucionPm = new taPMDistribution_ItemsTaPMDistribution[2];
        }

        /// <summary>
        /// Obtiene plan de impuestos del proveedor
        /// </summary>
        /// <param name="vendorid"></param>
        /// <returns></returns>
        private string getDatosProveedor(string vendorid)
        {
            int n = 0;
            string taxschid = string.Empty;
            using (EntitiesGP gp = new EntitiesGP())
            {
                var c = gp.PM00200.Where(w => w.VENDORID.Equals(vendorid.Trim()))
                    .Select(s => new { taxschid = s.TAXSCHID });

                n = c.Count();
                foreach (var r in c)
                    taxschid = r.taxschid;
            }
            if (n == 0)
                throw new NullReferenceException("Proveedor inexistente " + vendorid);

            return taxschid;

        }

        private void armaDetalleImpuestos(String taxschid)
        {
            using (EntitiesGP gp = new EntitiesGP())
            {
                var detalleImpuestosCompras = gp.vwImpuestosPlanYDetalle.Where(w => w.TXDTLTYP.Equals(2) && w.taxschid.Equals(taxschid))
                    .Select(s => new { s.TAXDTLID, s.TXDTLPCT});

                foreach (var impuesto in detalleImpuestosCompras)
                {
                    taPMTransactionTaxInsert_ItemsTaPMTransactionTaxInsert item = new taPMTransactionTaxInsert_ItemsTaPMTransactionTaxInsert();

                    item.VENDORID = facturaPm.VENDORID;
                    item.VCHRNMBR = facturaPm.VCHNUMWK;
                    item.DOCTYPE = facturaPm.DOCTYPE;
                    item.BACHNUMB = facturaPm.BACHNUMB;
                    item.TAXDTLID = impuesto.TAXDTLID;

                    item.TAXAMNT = Decimal.Round((facturaPm.PRCHAMNT - facturaPm.TRDISAMT) * impuesto.TXDTLPCT/100, 2);
                    item.TDTTXPUR = facturaPm.PRCHAMNT - facturaPm.TRDISAMT;
                    item.TXDTTPUR = facturaPm.PRCHAMNT - facturaPm.TRDISAMT;

                    taxDetails.Add(item);
                }
            }

        }
        /// <summary>
        /// Arma factura PM en objeto econnect facturaPm.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="fila"></param>
        /// <param name="sTimeStamp"></param>
        /// <param name="param"></param>
        public void armaFacturaPmEconn(ExcelWorksheet hojaXl, int fila, string sTimeStamp, Parametros param)
        {
            //Stream outputFile = File.Create(@"C:\gpusuario\traceArmaFActuraPmEconn"+fila.ToString()+".txt");
            //TextWriterTraceListener textListener = new TextWriterTraceListener(outputFile);
            //TraceSource trace = new TraceSource("trSource", SourceLevels.All);
            try
            {
                //trace.Listeners.Clear();
                //trace.Listeners.Add(textListener);
                //trace.TraceInformation("arma factura pm econn");

                iError = 0;
                string esf = hojaXl.Cells[fila, param.facturaPmEsFactura].Value.ToString().Trim().ToUpper();
                short tipoDocumento = 0;
                GetNextDocNumbers nextDocNumber = new GetNextDocNumbers();

                facturaPm.VCHNUMWK = nextDocNumber.GetPMNextVoucherNumber(IncrementDecrement.Increment, _DatosConexionDB.Elemento.ConnStr);

                if (short.TryParse(esf, out tipoDocumento))
                {
                    facturaPm.DOCTYPE = tipoDocumento;
                    facturaPm.VENDORID = hojaXl.Cells[fila, param.facturaPmVENDORID].Value.ToString().Trim();
                    facturaPm.DOCNUMBR = hojaXl.Cells[fila, param.facturaPmDOCNUMBR].Value.ToString().Trim();
                }
                else
                    if (esf.Equals("SI"))
                {
                    facturaPm.DOCTYPE = 1;  //invoice
                    facturaPm.VENDORID = hojaXl.Cells[fila, param.facturaPmVENDORID].Value.ToString().Trim();
                    facturaPm.DOCNUMBR = hojaXl.Cells[fila, param.facturaPmDOCNUMBR].Value.ToString().Trim();
                }
                else
                {
                    facturaPm.DOCTYPE = 3;  //misc charge
                    facturaPm.VENDORID = param.facturaPmGenericVENDORID;
                    facturaPm.DOCNUMBR = param.facturaPmGenericVENDORID + "-" + hojaXl.Cells[fila, param.facturaPmDOCNUMBR].Value.ToString().Trim();
                }

                facturaPm.BACHNUMB = sTimeStamp;
                facturaPm.BatchCHEKBKID = param.facturaPmBatchCHEKBKID;

                if (param.FacturaPmLOCALIZACION.Equals("BOL"))
                {
                    if (hojaXl.Cells[fila, param.addCodigoControl].Value != null)
                        facturaPm.USRDEFND1 = hojaXl.Cells[fila, param.addCodigoControl].Value.ToString().Trim();

                    if (hojaXl.Cells[fila, param.addNumAutorizacion].Value != null)
                        facturaPm.USRDEFND2 = hojaXl.Cells[fila, param.addNumAutorizacion].Value.ToString().Trim();
                }

                if (hojaXl.Cells[fila, param.facturaPmTRXDSCRN].Value != null)
                    facturaPm.TRXDSCRN = hojaXl.Cells[fila, param.facturaPmTRXDSCRN].Value.ToString();

                if (hojaXl.Cells[fila, param.facturaPmRETENCION].Value != null)
                {
                    facturaPm.USRDEFND2 = hojaXl.Cells[fila, param.facturaPmRETENCION].Value.ToString();
                }

                facturaPm.DOCDATE = DateTime.Parse(hojaXl.Cells[fila, param.facturaPmDOCDATE].Value.ToString().Trim()).ToString(param.FormatoFecha);

                if (hojaXl.Cells[fila, param.facturaPmDUEDATE].Value != null)
                    facturaPm.DUEDATE = DateTime.Parse(hojaXl.Cells[fila, param.facturaPmDUEDATE].Value.ToString().Trim()).ToString(param.FormatoFecha);

                facturaPm.CURNCYID = hojaXl.Cells[fila, param.facturaPmCURNCYID].Value.ToString();

                facturaPm.PRCHAMNT = Decimal.Round(Convert.ToDecimal(hojaXl.Cells[fila, param.facturaPmPRCHAMNT].Value.ToString(), CultureInfo.InvariantCulture), 2);


                if (param.DistribucionPmAplica.Equals("SI"))
                    facturaPm.CREATEDIST = 0;               //no crea el asiento contable automáticamente
                else
                {   //armado manual del detalle de los impuestos. El asiento contable se calcula automáticamente
                    facturaPm.TAXSCHID = getDatosProveedor(facturaPm.VENDORID);
                    armaDetalleImpuestos(facturaPm.TAXSCHID);
                    facturaPm.TAXAMNT = taxDetails.Sum(t => t.TAXAMNT);
                }

                facturaPm.DOCAMNT = facturaPm.MSCCHAMT + facturaPm.PRCHAMNT + facturaPm.TAXAMNT + facturaPm.FRTAMNT - facturaPm.TRDISAMT;
                //facturaPm.DOCAMNT = facturaPm.PRCHAMNT;
                facturaPm.CHRGAMNT = facturaPm.DOCAMNT;

                if (hojaXl.Cells[fila, param.facturaPmPAGADO].Value != null && hojaXl.Cells[fila, param.facturaPmPAGADO].Value.ToString() == "SI")
                {
                    facturaPm.CASHAMNT = facturaPm.PRCHAMNT;
                    facturaPm.CAMCBKID = hojaXl.Cells[param.facturaPmrowCHEKBKID, param.facturaPmcolCHEKBKID].Value.ToString().ToUpper().Trim();
                    facturaPm.CDOCNMBR = "R" + facturaPm.DOCNUMBR;
                    facturaPm.CAMPMTNM = "R" + facturaPm.VCHNUMWK;
                    facturaPm.CAMTDATE = facturaPm.DOCDATE;
                }

            }

            catch (NullReferenceException nr)
            {
                sMensaje = nr.Message + " " + nr.TargetSite.ToString();
                iError++;
            }
            catch (FormatException fmt)
            {
                sMensaje = "Alguna de las columnas contiene un monto o fecha con formato incorrecto. " + fmt.Message + " [Excepción en FacturaDeCompraPM.armaFacturaPmEconn]";
                iError++;
            }
            catch (OverflowException ovr)
            {
                sMensaje = "Alguna de las columna contiene un número demasiado grande. " + ovr.Message + " [Excepción en FacturaDeCompraPM.armaFacturaPmEconn]";
                iError++;
            }
            catch (Exception errorGral)
            {
                string inner = errorGral.InnerException == null ? "" : errorGral.InnerException.Message;
                sMensaje = "Excepción desconocida al armar la factura o comprobante. " + errorGral.Message + ". " + inner + " [Excepción en FacturaDeCompraPM.armaFacturaPmEconn]";
                iError++;
            }
            //finally
            //{
            //    trace.Flush();
            //    trace.Close();
            //}

        }

        /// <summary>
        /// Arma la distribución contable de la factura PM en objeto econnect.
        /// </summary>
        /// <param name="param"></param>
        public void armaDistribucionPmEconn(Parametros param)
        {
            try
            {
                if (facturaPm.DOCTYPE <= 3)
                {
                    _distribucionPm[0] = new taPMDistribution_ItemsTaPMDistribution()
                    {
                        DOCTYPE = facturaPm.DOCTYPE,
                        VCHRNMBR = facturaPm.VCHNUMWK,
                        VENDORID = facturaPm.VENDORID,
                        DISTTYPE = 6,
                        DistRef = facturaPm.TRXDSCRN,
                        ACTNUMST = param.DistribucionPmCuentaDebe,
                        DEBITAMT = facturaPm.DOCAMNT,
                        CRDTAMNT = 0
                    };

                    _distribucionPm[1] = new taPMDistribution_ItemsTaPMDistribution()
                    {
                        DOCTYPE = facturaPm.DOCTYPE,
                        VCHRNMBR = facturaPm.VCHNUMWK,
                        VENDORID = facturaPm.VENDORID,
                        DISTTYPE = 2,
                        DistRef = facturaPm.TRXDSCRN,
                        ACTNUMST = param.DistribucionPmCuentaHaber,
                        DEBITAMT = 0,
                        CRDTAMNT = facturaPm.DOCAMNT
                    };
                }
                else
                {
                    _distribucionPm[0] = new taPMDistribution_ItemsTaPMDistribution()
                    {
                        DOCTYPE = facturaPm.DOCTYPE,
                        VCHRNMBR = facturaPm.VCHNUMWK,
                        VENDORID = facturaPm.VENDORID,
                        DISTTYPE = 2,
                        DistRef = facturaPm.TRXDSCRN,
                        ACTNUMST = param.DistribucionPmCuentaHaber,
                        DEBITAMT = facturaPm.DOCAMNT,
                        CRDTAMNT = 0
                    };

                    _distribucionPm[1] = new taPMDistribution_ItemsTaPMDistribution()
                    {
                        DOCTYPE = facturaPm.DOCTYPE,
                        VCHRNMBR = facturaPm.VCHNUMWK,
                        VENDORID = facturaPm.VENDORID,
                        DISTTYPE = 6,
                        DistRef = facturaPm.TRXDSCRN,
                        ACTNUMST = param.DistribucionPmCuentaDebe,
                        DEBITAMT = 0,
                        CRDTAMNT = facturaPm.DOCAMNT
                    };
                }

            }
            catch (Exception errorGral)
            {
                throw new ArgumentException(errorGral.Message, "param");
            }

        }

        /// <summary>
        /// Revisa datos de la factura.
        /// iDetalleFactura guarda la posición de la siguiente factura.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="filaXl"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public void validaDatosDeIngreso(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            iError = 0;
            int iDetalleFactura = filaXl;

            try
            {
                if (param.FacturaPmTIPORETENCION.Equals("USA") && hojaXl.Cells[iDetalleFactura, param.facturaPmRETENCION].Value != null)
                {
                    decimal tasa = 0;
                    if (Decimal.TryParse(hojaXl.Cells[iDetalleFactura, param.facturaPmRETENCION].Value.ToString(), out tasa))
                    {
                        if (tasa < 0)
                            throw new ArgumentException("La tasa de retención es cero o negativa. Ingrese un monto positivo. [FacturaDeCompraPM.validaDatosDeIngreso]", "Columna: " + param.facturaPmRETENCION.ToString());
                    }
                    else
                    {
                        throw new ArgumentException("La tasa de retención no es un número. Ingrese un número válido. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]", "Columna: " + param.facturaPmRETENCION.ToString());
                    }
                }

                if (hojaXl.Cells[iDetalleFactura, param.facturaPmPAGADO].Value != null && hojaXl.Cells[iDetalleFactura, param.facturaPmPAGADO].Value.ToString() == "SI")
                    if (hojaXl.Cells[param.facturaPmrowCHEKBKID, param.facturaPmcolCHEKBKID].Value == null)
                    {
                        this.sMensaje = "No existe caja. Ingrese la caja de la rendición [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                        this.iError++;
                    }

                DateTime.Parse(hojaXl.Cells[iDetalleFactura, param.facturaPmDOCDATE].Value.ToString().Trim()).ToString(param.FormatoFecha);

                //if (iError == 0 && !Utiles.ConvierteAFechaFmt(hojaXl.Cells[iDetalleFactura, param.facturaPmDOCDATE].Value.ToString().Trim(), out fechaFactura))
                //{
                //    this.sMensaje = "La fecha de la factura no tiene el formato correcto: dd/mm/aaaa. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                //    this.iError++;
                //}

                if (iError == 0 &&
                    (hojaXl.Cells[iDetalleFactura, param.facturaPmEsFactura].Value == null ||
                    hojaXl.Cells[iDetalleFactura, param.facturaPmEsFactura].Value.ToString().Equals("")))
                {
                    sMensaje = "No existe el tipo de documento en la columna " + param.facturaPmEsFactura.ToString() + ". Posibles valores: 1-factura, 3-cargo misc, SI-factura, NO-cargo misc [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                    iError++;
                }

                if (iError == 0 && hojaXl.Cells[iDetalleFactura, param.facturaPmDOCNUMBR].Value == null)
                {
                    sMensaje = "No existe número de factura o comprobante. Ingrese un número en la columna " + param.facturaPmDOCNUMBR.ToString() + " . [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                    iError++;
                }

                if (param.FacturaPmLOCALIZACION == "BOL")
                    if (iError == 0 &&
                        hojaXl.Cells[iDetalleFactura, param.facturaPmEsFactura].Value.ToString().ToUpper().Equals("SI") &&
                          (hojaXl.Cells[iDetalleFactura, param.addNumAutorizacion].Value == null ||
                           hojaXl.Cells[iDetalleFactura, param.addNumAutorizacion].Value.ToString().Equals("")))
                    {
                        sMensaje = "No existe número de autorización. Ingrese un número en la columna Número de autorización. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                        iError++;
                    }

                if (iError == 0 && hojaXl.Cells[iDetalleFactura, param.facturaPmCURNCYID].Value == null)
                {
                    sMensaje = "No existe moneda. Ingrese el id de moneda en la columna Moneda. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                    iError++;
                }

                if (iError == 0 && hojaXl.Cells[iDetalleFactura, param.facturaPmPRCHAMNT].Value == null)
                {
                    sMensaje = "El monto está vacío. Ingrese un monto en la columna Monto. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                    iError++;
                }

                if (iError == 0)
                {
                    try
                    {
                        decimal monto = Convert.ToDecimal(hojaXl.Cells[iDetalleFactura, param.facturaPmPRCHAMNT].Value.ToString());
                        if (monto <= 0)
                        {
                            sMensaje = "El monto es cero o negativo. Ingrese un monto positivo en la columna Monto. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                            iError++;
                        }
                    }
                    catch (Exception exConv)
                    {
                        sMensaje = "El monto no es un número. Ingrese un número válido: sin separador de miles, con punto decimal y con dos decimales; en la columna Monto. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]" + exConv.Message; ;
                        iError++;
                    }
                }
            }
            catch (ArgumentNullException an)
            {
                sMensaje = "Excepción debido a un argumento nulo. " + an.Message + "[validaDatosDeIngreso]";
                iError++;
            }
            catch (FormatException fe)
            {
                sMensaje = "Excepción al validar el formato de monto o fecha. " + fe.Message + "[validaDatosDeIngreso]";
                iError++;
            }
            catch (ArgumentException ae)
            {
                sMensaje = "Excepción al validar datos de ingreso. " + ae.Message + " " + ae.ParamName + "[validaDatosDeIngreso]";
                iError++;
            }
            catch (Exception exRevision)
            {
                sMensaje = "Excepción desconocida al validar datos de ingreso. " + exRevision.Message + " [validaDatosDeIngreso]";
                iError++;
            }
        }

        /// <summary>
        /// Crea el xml de una factura PM a partir de una fila de datos en una hoja excel.
        /// Requisito. Ejecute la validación de datos de ingreso del proveedor.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void preparaFacturaPM(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp, Parametros param)
        {
            iError = 0;
            sMensaje = "";
            try
            {
                //validar input
                iniciaNuevaFacturaEn = filaXl + 1;
                this.validaDatosDeIngreso(hojaXl, filaXl, param);
                if (this.iError != 0)
                    return;

                //armar objeto econnect
                this.armaFacturaPmEconn(hojaXl, filaXl, sTimeStamp, param);
                if (this.iError != 0)
                    return;
                this.facturaPmType.taPMTransactionInsert = this.facturaPm;

                if (param.DistribucionPmAplica.Equals("SI"))
                {
                    armaDistribucionPmEconn(param);
                    this.facturaPmType.taPMDistribution_Items = _distribucionPm;
                }
                else
                {   //armado manual del detalle de los impuestos. El asiento contable se calcula automáticamente
                    this.facturaPmType.taPMTransactionTaxInsert_Items = taxDetails.ToArray();
                }
                arrFacturaPmType = new PMTransactionType[] { this.facturaPmType };

            }
            catch (ArgumentException ae)
            {
                sMensaje = ae.Message + " Revise el archivo de configuración de la distribución contable [FacturaDeCompraPM.preparaFacturaPM]";
                iError++;

            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción de econnect al preparar factura. " + eConnErr.Message + "[Excepción en preparaFacturaPM]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + "[Excepción en preparaFacturaPM]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida. " + errorGral.Message + " [Excepción en preparaFacturaPM]";
                iError++;
            }
        }

    }
}
