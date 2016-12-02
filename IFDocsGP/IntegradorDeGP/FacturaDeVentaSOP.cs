using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
using Comun;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using System.Globalization;
using System.IO;

namespace IntegradorDeGP
{
    public class FacturaDeVentaSOP
    {
        ConexionDB _DatosConexionDB;
        taSopHdrIvcInsert facturaSopCa;
        taSopLineIvcInsert_ItemsTaSopLineIvcInsert facturaSopDe;
        SOPTransactionType facturaSop;
        private int _iniciaNuevaFacturaEn;
        //TraceSource trace;
        //TextWriterTraceListener textListener;

        public SOPTransactionType FacturaSop
        {
            get
            {
                return facturaSop;
            }

            set
            {
                facturaSop = value;
            }
        }

        public int IniciaNuevaFacturaEn
        {
            get
            {
                return _iniciaNuevaFacturaEn;
            }

            set
            {
                _iniciaNuevaFacturaEn = value;
            }
        }

        public FacturaDeVentaSOP(ConexionDB DatosConexionDB)
        {
            //Stream outputFile = File.Create(@"C:\GPDocIntegration\traceInterfazGP.txt");
            //textListener = new TextWriterTraceListener(outputFile);
            //trace = new TraceSource("trSource", SourceLevels.All);
            //trace.Listeners.Clear();
            //trace.Listeners.Add(textListener);
            //trace.TraceInformation("integra factura sop");

            _DatosConexionDB = DatosConexionDB;
            facturaSopCa = new taSopHdrIvcInsert();
            facturaSopDe = new taSopLineIvcInsert_ItemsTaSopLineIvcInsert();
            facturaSop = new SOPTransactionType();
        }

        public void armaFacturaCaEconn(ExcelWorksheet hojaXl, int fila, string sTimeStamp, Parametros param)
        {
            try
            {
                String sopnumbe = hojaXl.Cells[fila, int.Parse(param.FacturaSopnumbe)].Value.ToString().Trim();
                String serie = sopnumbe.Substring(0, 1);

                facturaSopCa.BACHNUMB = sTimeStamp;
                facturaSopCa.SOPTYPE = 3;
                facturaSopCa.DOCID = "SERIE " + serie;
                facturaSopCa.SOPNUMBE = sopnumbe;
                facturaSopCa.DOCDATE = DateTime.Parse(hojaXl.Cells[fila, int.Parse(param.FacturaSopDocdate)].Value.ToString().Trim()).ToString(param.FormatoFecha);
                facturaSopCa.DUEDATE = DateTime.Parse(hojaXl.Cells[fila, int.Parse(param.FacturaSopDuedate)].Value.ToString().Trim()).ToString(param.FormatoFecha);

                String custnmbr = hojaXl.Cells[fila, int.Parse(param.FacturaSopTXRGNNUM)].Value == null ? "_enblanco" : hojaXl.Cells[fila, int.Parse(param.FacturaSopTXRGNNUM)].Value.ToString().Trim();
                facturaSopCa.CUSTNMBR = getCustomer(custnmbr);

                facturaSopCa.CREATETAXES = 1;   //1:crear impuestos automáticamente
                facturaSopCa.DEFPRICING = 0;    //0:se debe indicar el precio unitario
                facturaSopCa.REFRENCE = "Carga automática";

                facturaSopDe.SOPTYPE = facturaSopCa.SOPTYPE;
                facturaSopDe.SOPNUMBE = facturaSopCa.SOPNUMBE;
                facturaSopDe.CUSTNMBR = facturaSopCa.CUSTNMBR;
                facturaSopDe.DOCDATE = facturaSopCa.DOCDATE;
                facturaSopDe.ITEMNMBR = facturaSopCa.DOCID;
                facturaSopDe.QUANTITY = 1;
                facturaSopDe.DEFEXTPRICE = 1;   //1: calcular el precio extendido en base al precio unitario y la cantidad

                Decimal unitprice = 0;
                if (Decimal.TryParse(hojaXl.Cells[fila, int.Parse(param.FacturaSopUNITPRCE)].Value.ToString(), out unitprice))
                {
                    facturaSopCa.SUBTOTAL = Decimal.Round(unitprice, 2);
                    facturaSopCa.DOCAMNT = facturaSopCa.SUBTOTAL;
                    facturaSopDe.UNITPRCE = Decimal.Round(unitprice, 2);
                }
                else
                    throw new FormatException("El monto es incorrecto en la fila " + fila.ToString() + ", columna " + param.FacturaSopUNITPRCE + " [armaFacturaCaEconn]");
            }
            catch (FormatException fmt)
            {
                throw new FormatException("Formato incorrecto en la fila " + fila.ToString() + " [armaFacturaCaEconn]", fmt);
            }
            catch (OverflowException ovr)
            {
                throw new OverflowException("Monto demasiado grande en la fila " + fila.ToString() + " [armaFacturaCaEconn]", ovr);
            }
            catch (NullReferenceException)
            {
                throw;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            //finally
            //{
            //    trace.Flush();
            //    trace.Close();

            //}
        }

        private string getCustomer(string txrgnnum)
        {
            int n = 0;
            string cliente = string.Empty;
            using (GPEntities gp = new GPEntities())
                {
                    var c = gp.RM00101.Where(w => w.TXRGNNUM.Equals(txrgnnum.Trim()) && w.INACTIVE == 0)
                                    .Select(s => new { custnmbr = s.CUSTNMBR.Trim() });
                    n = c.Count();
                    foreach (var r in c)
                        cliente = r.custnmbr;
                }
            if (n==0)
                    throw new NullReferenceException("Cliente inexistente "+ txrgnnum);
            else if (n>1)
                    throw new InvalidOperationException("Cliente con Id de impuesto duplicado " + txrgnnum);

            return cliente;

        }

        public void preparaFacturaSOP(ExcelWorksheet hojaXl, int filaXl, string sTimeStamp, Parametros param)
        {
            //try
            //{
                //_iniciaNuevaFacturaEn = filaXl + 1;

                armaFacturaCaEconn(hojaXl, filaXl, sTimeStamp, param);

                facturaSop.taSopHdrIvcInsert = facturaSopCa;

                facturaSop.taSopLineIvcInsert_Items = new taSopLineIvcInsert_ItemsTaSopLineIvcInsert[] { facturaSopDe };

            //}
            //catch (eConnectException)
            //{
            //    throw;
            //}
            //catch (NullReferenceException)
            //{
            //    throw;
            //}
            //catch (InvalidOperationException)
            //{
            //    throw;
            //}
            //catch (Exception)
            //{
            //    throw;
            //}

        }
    }
}

