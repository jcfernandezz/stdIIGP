using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Comun;
using MyGeneration.dOOdads;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using OfficeOpenXml;

namespace IntegradorDeGP
{
    class Proveedor
    {
        public int iError = 0;
        public string sMensaje = "";
        private ConexionDB _DatosConexionDB;

        public taUpdateCreateVendorRcd Vendor;
        public PMVendorMasterType VendorType;
        public PMVendorMasterType[] arrVendorType;

        public Proveedor(ConexionDB DatosConexionDB)
        {
            _DatosConexionDB = DatosConexionDB;
        }

        private bool existeProveedor(string idProveedor)
        {
            this.iError = 0;
            this.sMensaje = "";
            vwIfcProveedores proveedores = new vwIfcProveedores(_DatosConexionDB.Elemento.ConnStr);
            proveedores.Where.Vendorid.Value = idProveedor;
            proveedores.Where.Vendorid.Operator = WhereParameter.Operand.Equal;
            try
            {
                if (proveedores.Query.Load())
                {
                    return true;
                }
                else
                    return false;
            }
            catch (Exception eProv)
            {
                this.iError++;
                this.sMensaje = "No puede acceder a la base de datos para revisar proveedores. Contacte al administrador. [Excepción en Proveedor.existeProveedor] " + eProv.Message;
                return false;
            }
        }

        /// <summary>
        /// Revisa datos del proveedor.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="filaXl"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public void validaDatosDeIngreso(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            iError = 0;

            try
            {
                if (hojaXl.Cells[filaXl, param.facturaPmEsFactura].Value == null ||
                    hojaXl.Cells[filaXl, param.facturaPmEsFactura].Value.ToString().Equals(""))
                {
                    sMensaje = "No existe el indicador Es Factura. Ingrese el valor SI/NO en la columna Es factura. [Excepción en Proveedor.validaDatosDeIngreso]";
                    iError++;
                }

                if (iError == 0 && hojaXl.Cells[filaXl, param.facturaPmEsFactura].Value.ToString().ToUpper().Equals("SI"))
                {
                    if (hojaXl.Cells[filaXl, param.facturaPmVENDORID].Value == null || hojaXl.Cells[filaXl, param.facturaPmVENDORID].Value.ToString().Equals(""))
                    {
                        sMensaje = "El NIT está en blanco. Ingrese el número de NIT en la columna NIT. [Excepción en Proveedor.validaDatosDeIngreso]";
                        iError++;
                    }
                    else
                    {
                        if (hojaXl.Cells[filaXl, param.facturaPmVENDORID].Value.ToString().Trim().Length <= 5)
                        {
                            sMensaje = "El NIT de proveedor está incompleto. Ingrese el número completo en la columna NIT. [Excepción en Proveedor.validaDatosDeIngreso]";
                            iError++;
                        }

                        if (iError == 0)
                        {
                            try
                            {
                                decimal monto = Convert.ToDecimal(hojaXl.Cells[filaXl, param.facturaPmVENDORID].Value.ToString());
                            }
                            catch (Exception exConv)
                            {
                                sMensaje = "El NIT de proveedor no es un número. Ingrese un número correcto en la columna NIT. [Excepción en Proveedor.validaDatosDeIngreso]" + exConv.Message;
                                iError++;
                            }
                        }
                    }
                }

            }
            catch (Exception exRevision)
            {
                sMensaje = "Excepción desconocida al validar datos del proveedor. " + exRevision.Message + " [Proveedor.validaDatosDeIngreso]";
                iError++;
            }

        }

        public void armaProveedorEconn(ExcelWorksheet hojaXl, int fila, Parametros param)
        {
            try
            {
                iError = 0;
                Vendor = new taUpdateCreateVendorRcd();
                VendorType = new PMVendorMasterType();

                Vendor.VENDORID = hojaXl.Cells[fila, param.facturaPmVENDORID].Value.ToString().Trim();
                Vendor.VENDNAME = hojaXl.Cells[fila, param.facturaPmVENDORNAME].Value.ToString().Trim();
                Vendor.VNDCLSID = param.defaultVNDCLSID;
                Vendor.VADDCDPR = "PRINCIPAL";
                Vendor.UpdateIfExists = 0;
                Vendor.UseVendorClass = 1;

                this.VendorType.taUpdateCreateVendorRcd = this.Vendor;
                this.arrVendorType = new PMVendorMasterType[] { this.VendorType };
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción al armar el proveedor. " + errorGral.Message + " [armaProveedorEconn]";
                iError++;
            }

        }

        /// <summary>
        /// Crea el xml de un proveedor a partir de una fila de datos en una hoja excel.
        /// </summary>
        /// <param name="hojaXl">Hoja excel</param>
        /// <param name="filaXl">Fila de la hoja excel a procesar</param>
        public void preparaProveedorEconn(ExcelWorksheet hojaXl, int filaXl, Parametros param)
        {
            iError = 0;
            sMensaje = "";
            try
            {
                //validar input
                this.validaDatosDeIngreso(hojaXl, filaXl, param);
                if (this.iError != 0 || hojaXl.Cells[filaXl, param.facturaPmEsFactura].Value.ToString().Trim().Equals("NO"))
                    return;

                //Si es factura y no existe el proveedor en GP, ingresar un nuevo proveedor 
                bool existeP = existeProveedor(hojaXl.Cells[filaXl, param.facturaPmVENDORID].Value.ToString().Trim());
                if (this.iError != 0 || existeP)
                    return;

                if (hojaXl.Cells[filaXl, param.facturaPmVENDORNAME].Value == null)
                {
                    sMensaje = "El nombre del proveedor está en blanco. Ingrese un nombre en la columna Nombre del proveedor. [Excepción en Proveedor.preparaProveedorEconn]";
                    iError++;
                    return;
                }

                //armar objeto econnect
                this.armaProveedorEconn(hojaXl, filaXl, param);
                if (this.iError != 0)
                    return;
            }
            catch (eConnectException eConnErr)
            {
                sMensaje = "Excepción de econnect al preparar el proveedor. " + eConnErr.Message + "[Excepción en Proveedor.preparaProveedorEconn]";
                iError++;
            }
            catch (ApplicationException ex)
            {
                sMensaje = "Excepción de aplicación. " + ex.Message + "[Excepción en Proveedor.preparaProveedorEconn]";
                iError++;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Excepción desconocida al preparar Proveedor econnect. " + errorGral.Message + " [Excepción en Proveedor.preparaProveedorEconn]";
                iError++;
            }
        }
    }
}
