using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

using Comun;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using OfficeOpenXml;

namespace IntegradorDeGP
{
    class FacturaDeCompraPOP
    {
        public int iError = 0;
        public string sMensaje = "";
        private int _iNumLineas;
        public taPopRcptHdrInsert facturaPopCa;
        public taPopRcptLineInsert_ItemsTaPopRcptLineInsert[] facturaPopDe;
        public taPopRcptLineTaxInsert_ItemsTaPopRcptLineTaxInsert[] facturaPopDeTaxes;
        public POPReceivingsType facturaPopType;
        public IVVendorItemType[] myVendorItemType;
        public POPReceivingsType[] myFacturaPopType;
        public int iniciaNuevaFacturaEn = 0;
        public DateTime fechaFactura = Convert.ToDateTime("1/1/1900");

        private decimal extdCost = 0;
        private decimal totExtdCost = 0;
        private int filaNegativa = -1;
        private int primeraFilaPositiva = -1;

        private ConexionDB _DatosConexionDB;

        public FacturaDeCompraPOP(ConexionDB DatosConexionDB, int iNumLineas)
        {
            _DatosConexionDB = DatosConexionDB;
            _iNumLineas = iNumLineas;
            facturaPopCa = new taPopRcptHdrInsert();
            facturaPopType = new POPReceivingsType();
        }
        /// <summary>
        /// Arma cabecera de factura en objeto econnect.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="fila"></param>
        /// <param name="sTimeStamp"></param>
        /// <param name="param"></param>
        public void armaFacturaCaEconn(ExcelWorksheet hojaXl, int fila, string sTimeStamp, Parametros param)
        {
            try
            {
                iError = 0;
                GetNextDocNumbers nextDocNumber = new GetNextDocNumbers();

                facturaPopCa.POPRCTNM = nextDocNumber.GetNextPOPReceiptNumber(IncrementDecrement.Increment, _DatosConexionDB.Elemento.ConnStr);
                facturaPopCa.POPTYPE = 3;                                                 //shipment/invoice
                facturaPopCa.VNDDOCNM = hojaXl.Cells[fila, Convert.ToInt32(param.nsa_serie)].Value.ToString().Trim() + "-" +
                                        hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaVNDDOCNM)].Value.ToString().Trim();

                if (!Utiles.ConvierteAFecha(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCareceiptdate)].Value.ToString(), out fechaFactura))
                {
                    this.sMensaje = "Excepción. La fecha de la factura no tiene el formato correcto: yyyyMMdd";
                    this.iError++;
                    return;
                }

                facturaPopCa.receiptdate = String.Format("{0:MM/dd/yyyy}", fechaFactura);
                facturaPopCa.BACHNUMB = sTimeStamp;
                facturaPopCa.VENDORID = hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaVENDORID)].Value.ToString();
                facturaPopCa.REFRENCE = hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaREFRENCE)].Value.ToString();
                facturaPopCa.CURNCYID = hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaCURNCYID)].Value.ToString();

                if (hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaXCHGRATE)].Value == null)
                    facturaPopCa.XCHGRATE = 0;
                else
                {
                    facturaPopCa.XCHGRATE = Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaXCHGRATE)].Value);
                    facturaPopCa.EXCHDATE = String.Format("{0:MM/dd/yyyy}", fechaFactura);
                    facturaPopCa.EXPNDATE = String.Format("{0:MM/dd/yyyy}", fechaFactura.AddDays(+60));
                }

                if (facturaPopCa.CURNCYID != param.functionalCurrency && facturaPopCa.XCHGRATE <= 0)
                {
                    this.sMensaje = "Excepción. La tasa de cambio no puede ser cero.";
                    this.iError++;
                    return;
                }

                facturaPopCa.TRDISAMTSpecified = true;
                facturaPopCa.TRDISAMT = 0;
                facturaPopCa.DISAVAMTSpecified = true;
                facturaPopCa.DISAVAMT = 0;
                facturaPopCa.USINGHEADERLEVELTAXES = 0;
                facturaPopCa.REFRENCE = hojaXl.Cells[fila, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString().Trim() + " " + facturaPopCa.VNDDOCNM;

                if (hojaXl.Cells[fila, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString().Trim().Length != 2)
                {
                    this.sMensaje = "Excepción. No tiene tipo de documento o es inválido. [armaFacturaCaEconn]";
                    this.iError++;
                }

            }
            catch (FormatException fmt)
            {
                sMensaje = "Excepción. La tasa de cambio tiene un formato incorrecto. " + fmt.Message + " [armaFacturaCaEconn]";
                iError++;
            }
            catch (OverflowException ovr)
            {
                sMensaje = "Excepción. El monto de la tasa de cambio es demasiado grande. " + ovr.Message + " [armaFacturaCaEconn]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción encontrada al armar cabecera de factura. " + errorGral.Message + " [armaFacturaCaEconn]";
                iError++;
            }

        }
        /// <summary>
        /// Arma una línea de la factura en objeto Econnect
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="fila"></param>
        /// <param name="sTimeStamp"></param>
        /// <param name="param"></param>
        /// <param name="linea"></param>
        public void armaFacturaLinEconn(ExcelWorksheet hojaXl, int fila, string sTimeStamp, Parametros param, int linea)
        {
            try
            {
                iError = 0;
                taPopRcptLineInsert_ItemsTaPopRcptLineInsert facturaPopLinea = new taPopRcptLineInsert_ItemsTaPopRcptLineInsert();

                facturaPopLinea.POPTYPE = facturaPopCa.POPTYPE;
                facturaPopLinea.POPRCTNM = facturaPopCa.POPRCTNM;
                facturaPopLinea.RCPTLNNM = (linea + 1) * 16000;
                facturaPopLinea.ITEMNMBR = param.defaultInventoryItem;
                facturaPopLinea.VNDITDSC = hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaREFRENCE)].Value.ToString().Trim();
                facturaPopLinea.ITEMDESC = hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaREFRENCE)].Value.ToString().Trim();
                facturaPopLinea.VENDORID = facturaPopCa.VENDORID;
                facturaPopLinea.VNDITNUM = param.defaultInventoryItem;
                facturaPopLinea.LOCNCODE = param.defaultDeposit;
                facturaPopLinea.QTYSHPPD = 1;
                facturaPopLinea.QTYINVCD = 1;
                facturaPopLinea.UNITCOSTSpecified = true;
                facturaPopLinea.UNITCOST = Decimal.Round(Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeBaseImponible)].Value) +
                                                        Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeMontoExento)].Value)  , 2) ;
                
                facturaPopLinea.Purchase_IV_Item_Taxable = 3;       //basar en proveedor
                facturaPopLinea.InventoryAccount = hojaXl.Cells[fila, Convert.ToInt32(param.accountNumst)].Value.ToString().Trim();

                this.facturaPopDe[linea] = facturaPopLinea;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Error al armar la línea de factura: " + linea.ToString() +" " + errorGral.Message + " [armaFacturaDeEconn]";
                iError++;
            }

        }
        /// <summary>
        /// Arma detalle de impuestos de la factura en objeto Econnect
        /// </summary>
        /// <param name="iLineaImpuesto"></param>
        public void armaFacturaLinTaxEconn(ExcelWorksheet hojaXl, int fila, int iRcpLinea, Parametros param)
        {
            if (hojaXl.Cells[fila, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString().Trim() == "CO")
                return;

            List<string> tipoDocs = new List<string>() { "01", "03", "12", "14" };         //factura, boleta, ticket, recibo por servicios públicos
            taPopRcptLineTaxInsert_ItemsTaPopRcptLineTaxInsert facturaPopLinTax;
            try
            {
                iError = 0;
                int numImpuestos = 1;

                //Igv y exento para factura, ticket y recibo por serv. públicos
                if (tipoDocs.Find(x => x == hojaXl.Cells[fila, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString().Trim()) != null)
                    numImpuestos = 2;

                facturaPopDeTaxes = new taPopRcptLineTaxInsert_ItemsTaPopRcptLineTaxInsert[numImpuestos];

                if (numImpuestos == 1 && hojaXl.Cells[fila, Convert.ToInt32(param.nsa_tipo_comprob)].Value.ToString().Trim() == "02")   //Recibo por honorarios
                {
                    facturaPopLinTax = new taPopRcptLineTaxInsert_ItemsTaPopRcptLineTaxInsert();
                    facturaPopLinTax.VENDORID = facturaPopCa.VENDORID;
                    facturaPopLinTax.POPRCTNM = facturaPopCa.POPRCTNM;
                    facturaPopLinTax.RCPTLNNM = this.facturaPopDe[iRcpLinea].RCPTLNNM;
                    facturaPopLinTax.TAXTYPE = 0;
                    facturaPopLinTax.TAXDTLID = param.detImpuestoRetencion4;
                    facturaPopLinTax.TAXPURCH = this.facturaPopDe[iRcpLinea].EXTDCOST;
                    facturaPopLinTax.TOTPURCH = this.facturaPopDe[iRcpLinea].EXTDCOST;
                    facturaPopLinTax.TAXAMNT = decimal.Round( this.facturaPopDe[iRcpLinea].EXTDCOST * param.detImpuestoRetencion4Ptj/100, 2);
                    this.facturaPopDeTaxes[0] = facturaPopLinTax;

                    this.facturaPopCa.TAXSCHID = param.planImpuestoRetencion4;
                    this.facturaPopCa.TAXAMNT += facturaPopLinTax.TAXAMNT;
                    this.facturaPopDe[iRcpLinea].TAXAMNT = facturaPopLinTax.TAXAMNT;
                }

                if (numImpuestos == 2)
                {
                    //igv
                    facturaPopLinTax = new taPopRcptLineTaxInsert_ItemsTaPopRcptLineTaxInsert();
                    facturaPopLinTax.VENDORID = facturaPopCa.VENDORID;
                    facturaPopLinTax.POPRCTNM = facturaPopCa.POPRCTNM;
                    facturaPopLinTax.RCPTLNNM = this.facturaPopDe[iRcpLinea].RCPTLNNM;
                    facturaPopLinTax.TAXTYPE = 0;
                    facturaPopLinTax.TAXDTLID = param.detImpuestoIgv;
                    if (hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeBaseImponible)].Value == null)
                    {
                        facturaPopLinTax.TAXPURCH = 0;
                        facturaPopLinTax.TOTPURCH = 0;
                        facturaPopLinTax.TAXAMNT = 0;
                    }
                    else
                    {
                        facturaPopLinTax.TAXPURCH = Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeBaseImponible)].Value);
                        facturaPopLinTax.TOTPURCH = Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeBaseImponible)].Value);
                        if (hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeImpuesto)].Value == null)
                            facturaPopLinTax.TAXAMNT = decimal.Round(facturaPopLinTax.TOTPURCH * param.detImpuestoIgvPtj / 100, 2);
                        else
                            facturaPopLinTax.TAXAMNT = decimal.Round(Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeImpuesto)].Value), 2, MidpointRounding.AwayFromZero);
                    }
                    this.facturaPopDeTaxes[0] = facturaPopLinTax;

                    this.facturaPopCa.TAXSCHID = param.planImpuestoIgv;
                    this.facturaPopCa.TAXAMNT += facturaPopLinTax.TAXAMNT;
                    this.facturaPopDe[iRcpLinea].TAXAMNT = facturaPopLinTax.TAXAMNT;

                    //exento
                    facturaPopLinTax = new taPopRcptLineTaxInsert_ItemsTaPopRcptLineTaxInsert();
                    facturaPopLinTax.VENDORID = facturaPopCa.VENDORID;
                    facturaPopLinTax.POPRCTNM = facturaPopCa.POPRCTNM;
                    facturaPopLinTax.RCPTLNNM = this.facturaPopDe[iRcpLinea].RCPTLNNM;
                    facturaPopLinTax.TAXTYPE = 0;
                    facturaPopLinTax.TAXDTLID = param.detImpuestoExento;

                    if (hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeMontoExento)].Value == null)
                    {
                        facturaPopLinTax.TAXPURCH = 0;
                        facturaPopLinTax.TOTPURCH = 0;
                    }
                    else
                    {
                        facturaPopLinTax.TAXPURCH = Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeMontoExento)].Value);
                        facturaPopLinTax.TOTPURCH = Convert.ToDecimal(hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopDeMontoExento)].Value);
                    }
                    this.facturaPopDeTaxes[1] = facturaPopLinTax;
                }
                
            }
            catch (Exception errorGral)
            {
                sMensaje = "Error al armar impuestos de factura. " + errorGral.Message + " [armaFacturaLinTaxEconn]";
                iError++;
            }

        }
        /// <summary>
        /// Revisa cuántas filas tiene la factura.
        /// iDetalleFactura guarda la posición de la siguiente factura.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="filaXl"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public int revisarDatosFactura(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            iError = 0;
            int iDetalleFactura = filaXl;
            extdCost = 0;
            totExtdCost = 0;
            filaNegativa = -1;
            primeraFilaPositiva = -1;

            try
            {
                string docnum = hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopCaVNDDOCNM)].Value.ToString();
                string serie = hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.nsa_serie)].Value.ToString();
                string prov = hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopCaVENDORID)].Value.ToString();
                //Revisar si hay negativos y sumar montos
                while ( hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.nsa_serie)].Value != null &&
                        hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopCaVNDDOCNM)].Value != null &&
                        hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopCaVENDORID)].Value != null &&
                        serie + docnum + prov == hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.nsa_serie)].Value.ToString() + hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopCaVNDDOCNM)].Value.ToString() + hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopCaVENDORID)].Value.ToString())
                {
                    extdCost = Convert.ToDecimal(hojaXl.Cells[iDetalleFactura, Convert.ToInt32(param.facturaPopDeEXTDCOST)].Value);
                    totExtdCost += extdCost;

                    if (extdCost < 0)
                        filaNegativa = iDetalleFactura;

                    if (extdCost > 0 && primeraFilaPositiva == -1)
                        primeraFilaPositiva = iDetalleFactura;

                    iDetalleFactura++;
                }
                return iDetalleFactura;
            }
            catch (Exception exRevision)
            {
                sMensaje = "Excepción al revisar datos de la factura. Es probable que no exista número de documento, serie o proveedor. " + exRevision.Message + "[revisarDatosFactura]";
                iError++;
                return iDetalleFactura+1;
            }

        }
        /// <summary>
        /// Crea el xml de una factura pop a partir de una fila de datos en una hoja excel.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void preparaFacturaPOP(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp, Parametros param)
        {
            iError = 0;
            sMensaje = "";
            ArticuloIV articuloIv = new ArticuloIV(_iNumLineas);
            try
            {
                iniciaNuevaFacturaEn = revisarDatosFactura(hojaXl, filaXl, param);
                if (this.iError != 0)
                    return;

                this.armaFacturaCaEconn(hojaXl, filaXl, sTimeStamp, param);
                if (this.iError != 0)
                    return;

                facturaPopDe = new taPopRcptLineInsert_ItemsTaPopRcptLineInsert[iniciaNuevaFacturaEn - filaXl];
                this.facturaPopCa.TAXAMNT = 0;

                //Armar la factura
                for (int filaFactura = 0; filaFactura < (iniciaNuevaFacturaEn - filaXl); filaFactura++ )
                {
                    extdCost = Convert.ToDecimal(hojaXl.Cells[filaXl + filaFactura, Convert.ToInt32(param.facturaPopDeEXTDCOST)].Value);
                    articuloIv.armaArtProvEconn(hojaXl, filaXl + filaFactura, param, filaFactura);
                    this.armaFacturaLinEconn(hojaXl, filaXl + filaFactura, sTimeStamp, param, filaFactura);

                    if (filaNegativa >= 0)                               //Si hay una fila negativa, cargar el primer monto positivo y el resto montos cero
                    {
                        if (primeraFilaPositiva == filaXl + filaFactura) //Cargar los impuestos del primer monto positivo
                            this.armaFacturaLinTaxEconn(hojaXl, filaXl, filaFactura, param); 
                        else
                        {
                            this.facturaPopDe[filaFactura].TAXAMNT = 0;
                            this.facturaPopDe[filaFactura].UNITCOST = 0;
                        }
                    }
                    else
                        this.armaFacturaLinTaxEconn(hojaXl, filaXl, filaFactura, param);   

                    if (extdCost <0)
                        this.facturaPopDe[filaFactura].BOLPRONUMBER = "-";

                }

                articuloIv.vendorItemType.taCreateItemVendors_Items = articuloIv.itemsVendors;
                myVendorItemType = new IVVendorItemType[] { articuloIv.vendorItemType };

                this.facturaPopType.taPopRcptHdrInsert = this.facturaPopCa;
                this.facturaPopType.taPopRcptLineInsert_Items = this.facturaPopDe;
                this.facturaPopType.taPopRcptLineTaxInsert_Items = this.facturaPopDeTaxes;

                myFacturaPopType = new POPReceivingsType[] { this.facturaPopType };
            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción al preparar factura. " + eConnErr.Message + "[preparaFacturaPOP]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + "[preparaFacturaPOP]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida. " + errorGral.Message + " [preparaFacturaPOP]";
                iError++;
            }
        }


        //public void Serializa()
        //{ 
        //    try
        //    {
        //        iError = 0;
        //        eConnectType eConnect = new eConnectType();
        //        POPReceivingsType facturaPopType = new POPReceivingsType();

        //        XmlSerializer serializer = new XmlSerializer(eConnect.GetType());

        //        facturaPopType.taPopRcptHdrInsert = this.facturaPopCa;

        //        facturaPopType.taPopRcptLineInsert_Items = this.facturaPopDe;

        //        POPReceivingsType [] myFacturaPopType = {facturaPopType};
        //        eConnect.POPReceivingsType = myFacturaPopType;

        //        XmlWriterSettings sett = new XmlWriterSettings();
        //        sett.Encoding = Encoding.UTF8;
        //        using (XmlWriter writer = XmlWriter.Create(this.sFacturaPopXml, sett))
        //        {
        //            serializer.Serialize(writer, eConnect);
        //        }
        //    }
        //    catch (Exception errorGral)
        //    {
        //        sMensaje = "Error al serializar la factura. " + errorGral.Message + " [Serializa]";
        //        iError++;
        //    }

        //}
    }
}
