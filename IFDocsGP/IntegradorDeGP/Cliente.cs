using Comun;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegradorDeGP
{
    public class Cliente
    {
        //public int iError = 0;
        //public string sMensaje = "";
        private ConexionDB _DatosConexionDB;
        private int _colIdImpuestoCliente = 0;
        private int _colCUSTNAME = 0;
        private Parametros _param;
        private taUpdateCreateCustomerRcd _Customer;
        private RMCustomerMasterType _CustomerType;
        private RMCustomerMasterType[] _arrCustomerType;

        public RMCustomerMasterType[] ArrCustomerType
        {
            get
            {
                return _arrCustomerType;
            }

            set
            {
                _arrCustomerType = value;
            }
        }

        public Cliente(ConexionDB DatosConexionDB, Parametros param)
        {
            _DatosConexionDB = DatosConexionDB;
            _param = param;
            if (!int.TryParse(param.FacturaSopTXRGNNUM, out _colIdImpuestoCliente))
                throw new NullReferenceException("No ha definido la columna del Id de impuestos del cliente (facturaSopCa.TXRGNNUM). Revise el archivo de configuración de la aplicación. ");
            if (!int.TryParse(param.FacturaSopCUSTNAME, out _colCUSTNAME))
                throw new NullReferenceException("No ha definido la columna del nombre del cliente (facturaSopCa.CUSTNAME). Revise el archivo de configuración de la aplicación. ");
            
        }

        private bool existeIdImpuestoCliente(string txrgnnum)
        {
            int n = 0;
            string cliente = string.Empty;
            using (GPEntities gp = new GPEntities())
            {
                var c = gp.RM00101.Where(w => w.TXRGNNUM.Equals(txrgnnum.Trim()) && w.INACTIVE==0)
                                .Select(s => new { custnmbr = s.CUSTNMBR.Trim() });
                n = c.Count();
                foreach (var r in c)
                    cliente = r.custnmbr;
            }
            return (n != 0);
        }

        /// <summary>
        /// Revisa datos del cliente.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="filaXl"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public void validaDatosDeIngreso(ExcelWorksheet hojaXl, int filaXl)
        {
            if (hojaXl.Cells[filaXl, _colIdImpuestoCliente].Value == null || hojaXl.Cells[filaXl, _colIdImpuestoCliente].Value.ToString().Equals(""))
            {
               throw new NullReferenceException( "El ID de impuesto está en blanco.");
            }
            if (hojaXl.Cells[filaXl, _colCUSTNAME].Value == null || hojaXl.Cells[filaXl, _colCUSTNAME].Value.ToString().Equals(""))
            {
                throw new NullReferenceException( "El nombre del cliente está en blanco. Ingrese un nombre en la columna Nombre del cliente.");
            }

        }

        public void armaClienteEconn(ExcelWorksheet hojaXl, int fila)
        {
            try
            {
                _Customer = new taUpdateCreateCustomerRcd();
                _CustomerType = new RMCustomerMasterType();

                _Customer.CUSTNMBR = hojaXl.Cells[fila, _colIdImpuestoCliente].Value.ToString().Trim().Replace(".", String.Empty).Replace("-", String.Empty);
                _Customer.CUSTNAME = hojaXl.Cells[fila, _colCUSTNAME].Value.ToString().Trim();
                _Customer.CUSTCLAS = _param.ClienteDefaultCUSTCLAS;
                _Customer.ADRSCODE = "MAIN";

                if (_colIdImpuestoCliente>0)
                    _Customer.TXRGNNUM = hojaXl.Cells[fila, _colIdImpuestoCliente].Value.ToString().Trim();

                _Customer.UpdateIfExists = 0;
                _Customer.UseCustomerClass = 1;

                _CustomerType.taUpdateCreateCustomerRcd = _Customer;
                _arrCustomerType = new RMCustomerMasterType[] { _CustomerType};

            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        /// Crea el xml de un cliente a partir de una fila de datos en una hoja excel.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public bool preparaClienteEconn(ExcelWorksheet hojaXl, int filaXl)
        {
            bool integrar = false;
            try
            {
                validaDatosDeIngreso(hojaXl, filaXl);

                if (!existeIdImpuestoCliente(hojaXl.Cells[filaXl, _colIdImpuestoCliente].Value.ToString().Trim()))
                {
                    armaClienteEconn(hojaXl, filaXl);
                    integrar = true;
                }
                return integrar;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
