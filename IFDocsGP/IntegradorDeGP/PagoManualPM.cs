using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Comun;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;

namespace IntegradorDeGP
{
    class PagoManualPM
    {
        public int iError = 0;
        public string sMensaje = "";
        public int iniciaNuevoDocEn = 0;
        public DateTime fechaFactura = Convert.ToDateTime("1/1/1900");
        private ConexionDB _DatosConexionDB;

        public taPMManualCheck pagoPm;
        //private taPMDistribution_ItemsTaPMDistribution[] _distribucionPm;
        public PMManualCheckType pagoPmType;
        public PMManualCheckType[] arrPagoPmType;
        //List<taPMTransactionTaxInsert_ItemsTaPMTransactionTaxInsert> taxDetails = new List<taPMTransactionTaxInsert_ItemsTaPMTransactionTaxInsert>();

        public PagoManualPM(ConexionDB DatosConexionDB)
        {
            _DatosConexionDB = DatosConexionDB;
            pagoPm = new taPMManualCheck();
            pagoPmType = new PMManualCheckType();
            //_distribucionPm = new taPMDistribution_ItemsTaPMDistribution[2];
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
            using (EntitiesGP gp = new EntitiesGP(_DatosConexionDB.Elemento.EntityConnStr))
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

        /// <summary>
        /// Arma pago PM en objeto econnect pagoPm.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="fila"></param>
        /// <param name="sTimeStamp"></param>
        /// <param name="param"></param>
        public void armaPagoPmEconn(ExcelWorksheet hojaXl, int fila, string sTimeStamp, Parametros param)
        {
            //Stream outputFile = File.Create(@"C:\gpusuario\traceArmaFActuraPmEconn" + fila.ToString() + ".txt");
            //TextWriterTraceListener textListener = new TextWriterTraceListener(outputFile);
            //TraceSource trace = new TraceSource("trSource", SourceLevels.All);
            try
            {
                //trace.Listeners.Clear();
                //trace.Listeners.Add(textListener);
                //trace.TraceInformation("arma factura pm econn");
                //trace.TraceInformation("fila: " + fila.ToString() + " col vendorid:" +param.facturaPmVENDORID);

                iError = 0;
                GetNextDocNumbers nextDocNumber = new GetNextDocNumbers();

                pagoPm.PMNTNMBR = nextDocNumber.GetPMNextVoucherNumber(IncrementDecrement.Increment, _DatosConexionDB.Elemento.ConnStr);

                pagoPm.PYENTTYP = short.Parse( hojaXl.Cells[fila, param.PagosPmPYENTTYP].Value.ToString().Trim());
                pagoPm.VENDORID = hojaXl.Cells[fila, param.PagosPmVENDORID].Value.ToString().Trim();
                pagoPm.DOCNUMBR = hojaXl.Cells[fila, param.PagosPmDOCNUMBR].Value.ToString().Trim();

                pagoPm.BACHNUMB = sTimeStamp;
                pagoPm.CHEKBKID = hojaXl.Cells[fila, param.PagosPmCHEKBKID].Value.ToString().Trim();

                if (hojaXl.Cells[fila, param.PagosPmTRXDSCRN].Value != null)
                    pagoPm.TRXDSCRN = hojaXl.Cells[fila, param.PagosPmTRXDSCRN].Value.ToString().Trim();

                pagoPm.DOCDATE = DateTime.Parse(hojaXl.Cells[fila, param.PagosPmDOCDATE].Value.ToString().Trim()).ToString(param.FormatoFecha);

                //pagoPm.CURNCYID = hojaXl.Cells[fila, param.facturaPmCURNCYID].Value.ToString();

                pagoPm.DOCAMNT = Decimal.Round(Convert.ToDecimal(hojaXl.Cells[fila, param.PagosPmDOCAMNT].Value.ToString(), CultureInfo.InvariantCulture), 2);

            }

            catch (NullReferenceException nr)
            {
                sMensaje = nr.Message + " " + nr.TargetSite.ToString();
                iError++;
            }
            catch (FormatException fmt)
            {
                sMensaje = "Alguna de las columnas contiene un monto o fecha con formato incorrecto. " + fmt.Message + " " + fmt.TargetSite.ToString();
                iError++;
            }
            catch (OverflowException ovr)
            {
                sMensaje = "Alguna de las columna contiene un número demasiado grande. " + ovr.Message + " " + ovr.TargetSite.ToString();
                iError++;
            }
            catch (Exception errorGral)
            {
                string inner = errorGral.InnerException == null ? "" : errorGral.InnerException.Message;
                sMensaje = "Excepción desconocida al armar el pago pm. " + errorGral.Message + ". " + inner + " " + errorGral.TargetSite.ToString();
                iError++;
            }
            //finally
            //{
            //    trace.Flush();
            //    trace.Close();
            //}

        }

        /// <summary>
        /// Revisa datos del pago.
        /// posiSiguienteDoc guarda la posición del siguiente doc.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="filaXl"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public void validaDatosDeIngreso(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            iError = 0;
            int posiSiguienteDoc = filaXl;

            try
            {
                DateTime.Parse(hojaXl.Cells[posiSiguienteDoc, param.PagosPmDOCDATE].Value.ToString().Trim()).ToString(param.FormatoFecha);

                if (iError == 0 &&
                    (hojaXl.Cells[posiSiguienteDoc, param.PagosPmPYENTTYP].Value == null ||
                    hojaXl.Cells[posiSiguienteDoc, param.PagosPmPYENTTYP].Value.ToString().Equals("")))
                {
                    sMensaje = "No existe el medio de pago en la columna " + param.PagosPmPYENTTYP.ToString() + ". Posibles valores: 0-cheque, 1-efectivo/transferencia [Excepción en PagoManualPM.validaDatosDeIngreso]";
                    iError++;
                }

                if (iError == 0 && hojaXl.Cells[posiSiguienteDoc, param.PagosPmDOCNUMBR].Value == null)
                {
                    sMensaje = "No existe número de cheque o efectivo/transferencia. Ingrese un número en la columna " + param.PagosPmDOCNUMBR.ToString() + " . [Excepción en PagoManualPM.validaDatosDeIngreso]";
                    iError++;
                }

                //if (iError == 0 && hojaXl.Cells[posiSiguienteDoc, param.facturaPmCURNCYID].Value == null)
                //{
                //    sMensaje = "No existe moneda. Ingrese el id de moneda en la columna Moneda. [Excepción en FacturaDeCompraPM.validaDatosDeIngreso]";
                //    iError++;
                //}

                if (iError == 0 && hojaXl.Cells[posiSiguienteDoc, param.PagosPmDOCAMNT].Value == null)
                {
                    sMensaje = "El monto está vacío. Ingrese un monto en la columna Monto. [Excepción en PagoManuelPM.validaDatosDeIngreso]";
                    iError++;
                }

                if (iError == 0)
                {
                    try
                    {
                        decimal monto = Convert.ToDecimal(hojaXl.Cells[posiSiguienteDoc, param.PagosPmDOCAMNT].Value.ToString());
                        if (monto <= 0)
                        {
                            sMensaje = "El monto es cero o negativo. Ingrese un monto positivo en la columna Monto. [Excepción en PagoManualPM.validaDatosDeIngreso]";
                            iError++;
                        }
                    }
                    catch (Exception exConv)
                    {
                        sMensaje = "El monto no es un número. Ingrese un número válido: sin separador de miles, con punto decimal y con dos decimales; en la columna Monto. [Excepción en PagoManualPM.validaDatosDeIngreso]" + exConv.Message; 
                        iError++;
                    }
                }
            }
            catch (ArgumentNullException an)
            {
                sMensaje = "Excepción debido a un argumento nulo. " + an.Message + " " + an.TargetSite.ToString();
                iError++;
            }
            catch (FormatException fe)
            {
                sMensaje = "Excepción al validar el formato de monto o fecha. " + fe.Message + " " + fe.TargetSite.ToString();
                iError++;
            }
            catch (ArgumentException ae)
            {
                sMensaje = "Excepción al validar datos de ingreso. " + ae.Message + " " + ae.ParamName + " " + ae.TargetSite.ToString();
                iError++;
            }
            catch (Exception exRevision)
            {
                sMensaje = "Excepción desconocida al validar datos de ingreso. " + exRevision.Message + " " + exRevision.TargetSite.ToString();
                iError++;
            }
        }

        /// <summary>
        /// Crea el xml de un pago manual PM a partir de una fila de datos en una hoja excel.
        /// Requisito. Ejecute la validación de datos de ingreso del proveedor.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void preparaPagoPM(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp, Parametros param)
        {
            iError = 0;
            sMensaje = "";
            try
            {
                //validar input
                iniciaNuevoDocEn = filaXl + 1;
                this.validaDatosDeIngreso(hojaXl, filaXl, param);
                if (this.iError != 0)
                    return;

                //armar objeto econnect
                this.armaPagoPmEconn(hojaXl, filaXl, sTimeStamp, param);
                if (this.iError != 0)
                    return;
                this.pagoPmType.taPMManualCheck = this.pagoPm;

                arrPagoPmType = new PMManualCheckType[] { this.pagoPmType };

            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción de econnect al preparar el pago PM. " + eConnErr.Message + " " + eConnErr.TargetSite.ToString();
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + " " + ex.TargetSite.ToString();
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida. " + errorGral.Message + " " + errorGral.TargetSite.ToString();
                iError++;
            }
        }
    }
}
