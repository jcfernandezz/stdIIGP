using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Comun
{
    public struct PrmtrsReporte
    {
        private string _nombre;
        private string _tipo;

        public PrmtrsReporte(string nombre, string tipo)
        {
            this._nombre = nombre;
            this._tipo = tipo;
        }

        public string nombre { get { return _nombre; } }
        public string tipo { get { return _tipo; } }
    }
    public class Bds
    {
        private string _id;
        private string _nombre;

        public Bds(string id, string nombre)
        {
            _id = id;
            _nombre = nombre;
        }
        public string Id
        {
            get
            {
                return _id;
            }

            set
            {
                _id = value;
            }
        }

        public string Nombre
        {
            get
            {
                return _nombre;
            }

            set
            {
                _nombre = value;
            }
        }
    }

    public struct DireccionesEmail
    {
        private string _mailTo;
        private string _mailCC;
        private string _mailCCO;

        public DireccionesEmail(string mailTo, string mailCC, string mailCCO)
        {
            this._mailTo = mailTo;
            this._mailCC = mailCC;
            this._mailCCO = mailCCO;
        }

        public string mailTo { get { return _mailTo; } set { _mailTo = value; } }
        public string mailCC { get { return _mailCC; } set { _mailCC = value; } }
        public string mailCCO { get { return _mailCCO; } set { _mailCCO = value; } }
    }

    public class Parametros
    {
        public int iError = 0;
        public string ultimoMensaje = "";
        public string nombreArchivoParametros = "ParametrosIgp.xml";
        public string defaultInventoryItem = "MISCELLANEOUS";
        public string defaultDeposit = "MAIN";
        public string functionalCurrency = "BOB";
        public IFormatProvider culture = new System.Globalization.CultureInfo("quz-BO");
        private string _rutaCarpeta = "";
        private string _detImpuestoIgv = "";
        private string _detImpuestoExento = "";
        private string _detImpuestoRetencion4 = "";

        private string _detImpuestoExentoPtj = "0";
        private string _detImpuestoIgvPtj ="0";
        private string _detImpuestoRetencion4Ptj ="0";
        private string _planImpuestoIgv = "";
        private string _planImpuestoExento ="";
        private string _planImpuestoRetencion4 ="";

        private string _formatoFecha = "";
        public string FormatoFecha
        {
            get { return _formatoFecha; }
            set { _formatoFecha = value; }
        }

        private string _facturaPopCafilaInicial = "";
        private string _facturaPopCacolumnaMensajes = "";
        private string _facturaPopCaVNDDOCNM = "";
        private string _facturaPopCareceiptdate = "";
        private string _facturaPopCaVENDORID = "";
        private string _facturaPopCaREFRENCE = "";
        private string _facturaPopCaCURNCYID = "";
        private string _facturaPopCaXCHGRATE = "";
        private string _facturaPopDefilaInicial = "";
        private string _facturaPopDeBaseImponible = "";
        private string _facturaPopDeImpuesto = "";
        private string _facturaPopDeMontoExento = ""; 
        private string _facturaPopDeEXTDCOST = "";

        private string _facturaPmEsFactura = "";
        private string _facturaPmFilaInicial = "";
        private string _facturaPmColumnaMensajes = "";
        private string _facturaPmDOCNUMBR = "";
        private string _facturaPmDOCDATE = "";
        private string _facturaPmVENDORID = "";
        private string _facturaPmVENDORNAME = "";
        private string _facturaPmTRXDSCRN = "";
        private string _facturaPmCURNCYID = "";
        private string _facturaPmPRCHAMNT = "";
        private string _facturaPmDUEDATE = string.Empty;
        private string _facturaPmPAGADO = string.Empty;
        private string _facturaPmTIPORETENCION = string.Empty;
        private string _facturaPmRETENCION = string.Empty;
        private string _facturaPmGenericVENDORID = "";
        private string _facturaPmBatchCHEKBKID = "";
        private string _facturaPmrowCHEKBKID = "";
        private string _facturaPmcolCHEKBKID = "";

        private string _facturaPmLOCALIZACION;

        private string _addNumAutorizacion = "";
        private string _addCodigoControl = "";
        private string _distribucionPmAplica = string.Empty;
        private string _distribucionPmCuentaDebe = string.Empty;
        private string _distribucionPmCuentaHaber = string.Empty;

        private string _nsa_tipo_comprob = "";
        private string _nsa_tipo_comprob_default = "";
        private string _nsa_cod_transac = "";
        private string _nsa_cod_transac_default = "";
        private string _nsa_autorizacion = "";
        private string _nsa_autorizacion_default = "";
        private string _nsa_cred_trib = "";
        private string _nsa_cred_trib_default = "";
        private string _nsa_cod_iva1 = "";
        private string _nsa_cod_iva1_default = "";
        private string _nsa_cod_iva2 = "";
        private string _nsa_cod_iva2_default = "";
        private string _nsaCoa_secuencial = "";
        private string _nsa_serie = "";
        private string _nsaCoa_date_nota = "";
        private string _nsa_tipo_comprob_mod = "";
        private string _nsa_sernota = "";
        private string _nsacoa_secuencial_mod = "";
        private string _accountNumst = "";

        private string _defaultVNDCLSID = "";

        private string _servidor = "";
        private string _seguridadIntegrada = "0";
        private string _usuarioSql = "";
        private string _passwordSql = "";
        private List<Bds> _listaCompannia;

        private string _URLArchivoXSD = "";
        private string _emite = "0";
        private string _anula = "0";
        private string _imprime = "0";
        private string _publica = "0";
        private string _envia = "0";

        private string _facturaSopSopnumbe;
        private string _facturaSopDocdate;
        private string _facturaSopDuedate;
        private string _facturaSopTXRGNNUM;
        private string _facturaSopUNITPRCE;
        private string _facturaSopColumnaMensajes;
        private string _facturaSopFilaInicial;
        private string _facturaSopCUSTNAME;
        private string _clienteDefaultCUSTCLAS;

        public Parametros()
        {
            try
            {
                this.iError = 0;
                XmlDocument listaParametros = new XmlDocument();
                listaParametros.Load(new XmlTextReader(nombreArchivoParametros));
                XmlNodeList listaElementos = listaParametros.DocumentElement.ChildNodes;
                _listaCompannia = new List<Bds>();

                foreach (XmlNode n in listaElementos)
                {
                    if (n.Name.Equals("servidor"))
                        this._servidor = n.InnerXml;
                    if (n.Name.Equals("seguridadIntegrada"))
                        this._seguridadIntegrada = n.InnerXml;
                    if (n.Name.Equals("usuariosql"))
                        this._usuarioSql = n.InnerXml;
                    if (n.Name.Equals("passwordsql"))
                        this._passwordSql = n.InnerXml;
                    if (n.Name.Equals("compannia"))
                        _listaCompannia.Add(new Bds(n.Attributes["bd"].Value, n.Attributes["nombre"].Value));
                }
            }
            catch (Exception eprm)
            {
                iError++;
                ultimoMensaje = "No se pudo obtener acceso al servidor. Revise el archivo de configuración. [Parametros()]" + eprm.Message;
            }
        }

        public Parametros(string IdCompannia)
        {
            try
            {
                this.iError = 0;
                XmlDocument listaParametros = new XmlDocument();
                listaParametros.Load(new XmlTextReader(nombreArchivoParametros));
                XmlNode elemento = listaParametros.DocumentElement;

                _rutaCarpeta = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/rutaCarpeta/text()").Value;
                _formatoFecha = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/formatoFecha/text()").Value;
                //_detImpuestoIgv = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/detalleImpuestos/igv/text()").Value;
                //_detImpuestoExento = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/detalleImpuestos/exento/text()").Value;
                //_detImpuestoRetencion4 = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/detalleImpuestos/retencion4/text()").Value;
                //_detImpuestoIgvPtj = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/detalleImpuestos/igvPorcentaje/text()").Value;
                //_detImpuestoRetencion4Ptj = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/detalleImpuestos/retencion4Porcentaje/text()").Value;
                //_planImpuestoIgv = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/planImpuestos/igv/text()").Value;
                //_planImpuestoExento = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/planImpuestos/exento/text()").Value;
                //_planImpuestoRetencion4 = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/impuestos/planImpuestos/retencion4/text()").Value;

                //_facturaPopCafilaInicial = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/filaInicial/text()").Value;
                //_facturaPopCacolumnaMensajes = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/columnaMensajes/text()").Value;
                
                //_facturaPopCaVNDDOCNM = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/VNDDOCNM/text()").Value;
                //_facturaPopCareceiptdate = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/receiptdate/text()").Value;
                //_facturaPopCaVENDORID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/VENDORID/text()").Value;
                //_facturaPopCaREFRENCE = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/REFRENCE/text()").Value;
                //_facturaPopCaCURNCYID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/CURNCYID/text()").Value;
                //_facturaPopCaXCHGRATE = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCa/XCHGRATE/text()").Value;

                _facturaPmEsFactura = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/EsFactura/text()").Value;
                _facturaPmFilaInicial = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/FilaInicial/text()").Value;
                _facturaPmColumnaMensajes = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/ColumnaMensajes/text()").Value;
                _facturaPmDOCNUMBR = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/DOCNUMBR/text()").Value;
                _facturaPmDOCDATE = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/DOCDATE/text()").Value;
                _facturaPmVENDORID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/VENDORID/text()").Value;
                _facturaPmVENDORNAME = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/VENDORNAME/text()").Value;
                _facturaPmTRXDSCRN = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/TRXDSCRN/text()").Value;
                _facturaPmCURNCYID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/CURNCYID/text()").Value; 
                //private string _facturaPmXCHGRATE = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/FilaInicial/text()").Value; 
                _facturaPmPRCHAMNT = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/PRCHAMNT/text()").Value;
                _facturaPmDUEDATE = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/DUEDATE/text()").Value;
                _facturaPmPAGADO = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/PAGADO/text()").Value;
                _facturaPmTIPORETENCION = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPmAdicionales/TIPORETENCION/text()").Value;
                _facturaPmRETENCION = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPmAdicionales/RETENCION/text()").Value;
                _facturaPmGenericVENDORID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/GenericVENDORID/text()").Value;
                _facturaPmBatchCHEKBKID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/BatchCHEKBKID/text()").Value;
                _facturaPmrowCHEKBKID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/rowCHEKBKID/text()").Value;
                _facturaPmcolCHEKBKID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/colCHEKBKID/text()").Value;
                _facturaPmLOCALIZACION = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPmAdicionales/LOCALIZACION/text()").Value;
                _addNumAutorizacion = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPmAdicionales/NumAutorizacion/text()").Value;
                _addCodigoControl = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPmAdicionales/CodigoControl/text()").Value; 
                _distribucionPmAplica = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/DISTRIBUCION/APLICA/text()").Value;
                _distribucionPmCuentaDebe = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/DISTRIBUCION/CUENTADEBE/text()").Value;
                _distribucionPmCuentaHaber = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPm/DISTRIBUCION/CUENTAHABER/text()").Value;

                _clienteDefaultCUSTCLAS = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/Cliente/DefaultCUSTCLAS/text()").Value;
                _facturaSopSopnumbe = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/sopnumbe/text()").Value;
                _facturaSopDocdate = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/docdate/text()").Value;
                _facturaSopDuedate = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/duedate/text()").Value;
                _facturaSopTXRGNNUM = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/TXRGNNUM/text()").Value;
                _facturaSopCUSTNAME = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/CUSTNAME/text()").Value;
                _facturaSopUNITPRCE = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopDe/UNITPRCE/text()").Value;
                _facturaSopColumnaMensajes = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/columnaMensajes/text()").Value;
                _facturaSopFilaInicial = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaSopCa/filaInicial/text()").Value;

                //_nsa_tipo_comprob = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_tipo_comprob/text()").Value;
                //_nsa_tipo_comprob_default = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_tipo_comprob_default/text()").Value;
                //_nsa_cod_transac = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cod_transac/text()").Value;
                //_nsa_cod_transac_default = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cod_transac_default/text()").Value;
                //_nsa_autorizacion = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_autorizacion/text()").Value;
                //_nsa_autorizacion_default = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_autorizacion_default/text()").Value;
                //_nsa_cred_trib = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cred_trib/text()").Value;
                //_nsa_cred_trib_default = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cred_trib_default/text()").Value;
                //_nsa_cod_iva1 = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cod_iva1/text()").Value;
                //_nsa_cod_iva1_default = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cod_iva1_default/text()").Value;
                //_nsa_cod_iva2 = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cod_iva2/text()").Value;
                //_nsa_cod_iva2_default = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_cod_iva2_default/text()").Value;
                //_nsaCoa_secuencial = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsaCoa_secuencial/text()").Value;
                //_nsa_serie = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_serie/text()").Value;
                //_nsaCoa_date_nota = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsaCoa_date_nota/text()").Value;
                //_nsa_tipo_comprob_mod = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_tipo_comprob_mod/text()").Value;
                //_nsa_sernota = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsa_sernota/text()").Value;
                //_nsacoa_secuencial_mod = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopCaAdicionales/nsacoa_secuencial_mod/text()").Value;

                //_facturaPopDefilaInicial = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopDe/filaInicial/text()").Value;
                //_facturaPopDeBaseImponible = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopDe/baseImponible/text()").Value;
                //_facturaPopDeImpuesto = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopDe/impuesto/text()").Value;
                //_facturaPopDeMontoExento = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopDe/montoExento/text()").Value;
                //_facturaPopDeEXTDCOST = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopDe/EXTDCOST/text()").Value;
                //_accountNumst = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/facturaPopDe/accountNumst/text()").Value;

                _defaultVNDCLSID = elemento.SelectSingleNode("//compannia[@bd='" + IdCompannia + "']/proveedor/DefaultVNDCLSID/text()").Value;
            }
            catch (Exception eprm)
            {
                iError++;
                ultimoMensaje = "No se pudo obtener la configuración de la compañía. Revise el archivo de configuración. " + IdCompannia + ". [Parametros(Compañía)] " + eprm.Message;
            }
        }
    #region Parámetros de seguridad
        public string servidor
        {
            get { return _servidor; }
            set { _servidor = value; }
        }

        public bool seguridadIntegrada
        {
            get 
            { 
                return _seguridadIntegrada.Equals("1"); 
            }
            set 
            { 
                if (value)
                    _seguridadIntegrada = "1"; 
                else
                    _seguridadIntegrada = "0"; 
            }
        }

        public string usuarioSql
        {
            get { return _usuarioSql; }
            set{ _usuarioSql = value;}
        }

        public string passwordSql
        {
            get { return _passwordSql; }
            set {_passwordSql=value;}
        }
    #endregion

    #region Configuraciones generales de rendiciones de caja
        /// <summary>
        /// Ruta donde están las carpetas En trabajo y En proceso
        /// </summary>
        public string rutaCarpeta
        {
            get { return _rutaCarpeta; }
            set { _rutaCarpeta = value; }
        }

        public string detImpuestoIgv
        {
            get { return _detImpuestoIgv; }
            set { _detImpuestoIgv = value; }
        }
        public string detImpuestoExento
        {
            get { return _detImpuestoExento; }
            set { _detImpuestoExento = value; }
        }
        public string detImpuestoRetencion4
        {
            get { return _detImpuestoRetencion4; }
            set { _detImpuestoRetencion4 = value; }
        }

        public decimal detImpuestoIgvPtj
        {
            get { return Convert.ToDecimal(_detImpuestoIgvPtj); }
            set { _detImpuestoIgvPtj = Convert.ToString(value); }
        }
        public decimal detImpuestoExentoPtj
        {
            get { return Convert.ToDecimal(_detImpuestoExentoPtj); }
            set { _detImpuestoExentoPtj =  Convert.ToString(value); }
        }
        public decimal detImpuestoRetencion4Ptj
        {
            get { return Convert.ToDecimal(_detImpuestoRetencion4Ptj); }
            set { _detImpuestoRetencion4Ptj = Convert.ToString(value); }
        }

        public string planImpuestoIgv
        {
            get { return _planImpuestoIgv; }
            set { _planImpuestoIgv = value; }
        }
        public string planImpuestoExento
        {
            get { return _planImpuestoExento; }
            set { _planImpuestoExento = value; }
        }
        public string planImpuestoRetencion4
        {
            get { return _planImpuestoRetencion4; }
            set { _planImpuestoRetencion4 = value; }
        }
    #endregion

    #region Columnas del archivo excel: Campos de facturas POP

        public string facturaPopCafilaInicial 
        {
            get{ return _facturaPopCafilaInicial;}
            set{ _facturaPopCafilaInicial = value;}
        }

        public int facturaPopCacolumnaMensajes 
        {
            get { return Convert.ToInt32( _facturaPopCacolumnaMensajes); }
            set { _facturaPopCacolumnaMensajes = value.ToString(); }
        }
         
        public string facturaPopCaVNDDOCNM 
        {
            get { return _facturaPopCaVNDDOCNM; }
            set { _facturaPopCaVNDDOCNM = value; }
        }

        public string facturaPopCareceiptdate 
        {
            get { return _facturaPopCareceiptdate; }
            set { _facturaPopCareceiptdate = value; }
        }

        public string facturaPopCaVENDORID 
        {
            get { return _facturaPopCaVENDORID; }
            set { _facturaPopCaVENDORID = value; }
        }

        public string facturaPopCaREFRENCE 
        {
            get { return _facturaPopCaREFRENCE; }
            set { _facturaPopCaREFRENCE = value; }
        }

        public string facturaPopCaCURNCYID 
        {
            get { return _facturaPopCaCURNCYID; }
            set { _facturaPopCaCURNCYID = value; }
        }
        public string facturaPopCaXCHGRATE
        {
            get { return _facturaPopCaXCHGRATE; }
            set { _facturaPopCaXCHGRATE = value; }
        }
    #endregion

    #region Columnas del archivo excel: Campos de facturas PM

        public string facturaPmBatchCHEKBKID
        {
            get { return _facturaPmBatchCHEKBKID; }
            set { _facturaPmBatchCHEKBKID = value; }
        }

        public string facturaPmGenericVENDORID
        {
            get { return _facturaPmGenericVENDORID; }
            set { _facturaPmGenericVENDORID = value; }
        }

        public int facturaPmrowCHEKBKID
        {
            get { return Convert.ToInt32(_facturaPmrowCHEKBKID); }
            set { _facturaPmrowCHEKBKID = value.ToString(); }
        }

        public int facturaPmcolCHEKBKID
        {
            get { return Convert.ToInt32(_facturaPmcolCHEKBKID); }
            set { _facturaPmcolCHEKBKID = value.ToString(); }
        }

        public int facturaPmEsFactura
        {
            get { return Convert.ToInt32(_facturaPmEsFactura); }
            set { _facturaPmEsFactura = value.ToString(); }
        }

        public int facturaPmFilaInicial
        {
            get { return Convert.ToInt32(_facturaPmFilaInicial); }
            set { _facturaPmFilaInicial = value.ToString(); }
        }

        public int facturaPmColumnaMensajes
        {
            get { return Convert.ToInt32(_facturaPmColumnaMensajes); }
            set { _facturaPmColumnaMensajes = value.ToString(); }
        }

        public int facturaPmDOCNUMBR
        {
            get { return Convert.ToInt32(_facturaPmDOCNUMBR); }
            set { _facturaPmDOCNUMBR = value.ToString(); }
        }

        public int facturaPmDOCDATE
        {
            get { return Convert.ToInt32(_facturaPmDOCDATE); }
            set { _facturaPmDOCDATE = value.ToString(); }
        }

        public int facturaPmVENDORID
        {
            get { return Convert.ToInt32(_facturaPmVENDORID); }
            set { _facturaPmVENDORID = value.ToString(); }
        }

        public int facturaPmVENDORNAME
        {
            get { return Convert.ToInt32(_facturaPmVENDORNAME); }
            set { _facturaPmVENDORNAME = value.ToString(); }
        }

        public int facturaPmTRXDSCRN
        {
            get { return Convert.ToInt32(_facturaPmTRXDSCRN); }
            set { _facturaPmTRXDSCRN = value.ToString(); }
        }

        public int facturaPmCURNCYID
        {
            get { return Convert.ToInt32(_facturaPmCURNCYID); }
            set { _facturaPmCURNCYID = value.ToString(); }
        }

        public int facturaPmPRCHAMNT
        {
            get { return Convert.ToInt32(_facturaPmPRCHAMNT); }
            set { _facturaPmPRCHAMNT = value.ToString(); }
        }

        public int facturaPmDUEDATE
        {
            get { return Convert.ToInt32(_facturaPmDUEDATE); }
            set { _facturaPmDUEDATE = value.ToString(); }
        }
        public int facturaPmPAGADO
        {
            get { return Convert.ToInt32(_facturaPmPAGADO); }
            set { _facturaPmPAGADO = value.ToString(); }
        }

        public int facturaPmRETENCION
        {
            get {
                int colRetencion = 0;
                int.TryParse(_facturaPmRETENCION, out colRetencion);
                return colRetencion;
                }
            set { _facturaPmRETENCION = value.ToString(); }
        }


    #endregion

    #region Columnas del archivo excel: Campos de datos adicionales de la factura

        public string FacturaPmLOCALIZACION
        {
            get { return _facturaPmLOCALIZACION; }
            set { _facturaPmLOCALIZACION = value; }
        }
        public int addNumAutorizacion
        {
            get { return Convert.ToInt32(_addNumAutorizacion); }
            set { _addNumAutorizacion = value.ToString(); }
        }

        public int addCodigoControl
        {
            get { return Convert.ToInt32(_addCodigoControl); }
            set { _addCodigoControl = value.ToString(); }
        }
    #endregion

    #region Columnas del archivo excel: Campos de datos adicionales de facturas POP en Perú
        public string nsa_tipo_comprob
        { 
            get { return _nsa_tipo_comprob;}
            set { _nsa_tipo_comprob=value;}
        }
        public string nsa_tipo_comprob_default
        {
            get { return _nsa_tipo_comprob_default; }
            set { _nsa_tipo_comprob_default = value; }
        }
        public string nsa_cod_transac
        { 
            get { return _nsa_cod_transac;}
            set { _nsa_cod_transac=value;}
        }
        public string nsa_cod_transac_default
        {
            get { return _nsa_cod_transac_default; }
            set { _nsa_cod_transac_default = value; }
        }
        public string nsa_autorizacion
        { 
            get { return _nsa_autorizacion;}
            set { _nsa_autorizacion=value;}
        }
        public string nsa_autorizacion_default
        {
            get { return _nsa_autorizacion_default; }
            set { _nsa_autorizacion_default = value; }
        }
        public string nsa_cred_trib
        {
            get { return _nsa_cred_trib; }
            set { _nsa_cred_trib = value; }
        }
        public string nsa_cred_trib_default
        {
            get { return _nsa_cred_trib_default; }
            set { _nsa_cred_trib_default = value; }
        }

        public string nsa_cod_iva1
        { 
            get { return _nsa_cod_iva1;}
            set { _nsa_cod_iva1=value;}
        }
        public string nsa_cod_iva1_default
        {
            get { return _nsa_cod_iva1_default; }
            set { _nsa_cod_iva1_default = value; }
        }
        
        public string nsa_cod_iva2
        { 
            get { return _nsa_cod_iva2;}
            set { _nsa_cod_iva2=value;}
        }
        public string nsa_cod_iva2_default
        {
            get { return _nsa_cod_iva2_default; }
            set { _nsa_cod_iva2_default = value; }
        }
        public string nsaCoa_secuencial
        { 
            get { return _nsaCoa_secuencial;}
            set { _nsaCoa_secuencial=value;}
        }
        public string nsa_serie
        { 
            get { return _nsa_serie;}
            set { _nsa_serie=value;}
        }
        public string nsaCoa_date_nota
        { 
            get { return _nsaCoa_date_nota;}
            set { _nsaCoa_date_nota=value;}
        }
        public string nsa_tipo_comprob_mod
        { 
            get { return _nsa_tipo_comprob_mod;}
            set { _nsa_tipo_comprob_mod=value;}
        }
        public string nsa_sernota
        {
            get { return _nsa_sernota; }
            set { _nsa_sernota = value; }
        }
        public string nsacoa_secuencial_mod
        { 
            get { return _nsacoa_secuencial_mod;}
            set { _nsacoa_secuencial_mod = value; }
        }
    #endregion

    #region Columnas del archivo excel: Campos del detalle de facturas POP
        public string facturaPopDefilaInicial
        {
            get { return _facturaPopDefilaInicial; }
            set { _facturaPopDefilaInicial = value; }
        }

        public string facturaPopDeBaseImponible
        {
            get { return _facturaPopDeBaseImponible; }
            set { _facturaPopDeBaseImponible = value; }
        }

        public string facturaPopDeImpuesto
        {
            get { return _facturaPopDeImpuesto; }
            set { _facturaPopDeImpuesto = value; }
        }

        public string facturaPopDeMontoExento
        {
            get { return _facturaPopDeMontoExento; }
            set { _facturaPopDeMontoExento = value; }
        }

        public string facturaPopDeEXTDCOST
        {
            get { return _facturaPopDeEXTDCOST; }
            set { _facturaPopDeEXTDCOST = value; }
        }
        public string accountNumst
        {
            get { return _accountNumst; }
            set { _accountNumst = value; }
        }
    #endregion

    #region Datos del proveedor
        public string defaultVNDCLSID
        {
            get { return _defaultVNDCLSID; }
            set { _defaultVNDCLSID = value; }
        }
    #endregion

        public string URLArchivoXSD
        {
            get { return _URLArchivoXSD; }
            set { _URLArchivoXSD = value; }
        }

        public int intEstadoCompletado
        {
            get
            {
                return
                        Convert.ToInt32(_emite) +
                    2 * 0 +
                    4 * Convert.ToInt32(_imprime) +
                    8 * Convert.ToInt32(_publica) +
                    16 * Convert.ToInt32(_envia);
            }
        }

        public bool emite
        {
            get { return _emite.Equals("1"); }
        }

        public bool anula
        {
            get { return _anula.Equals("1"); }
        }

        public bool imprime
        {
            get { return _imprime.Equals("1"); }
        }

        public bool publica
        {
            get { return _publica.Equals("1"); }
        }

        public bool envia
        {
            get { return _envia.Equals("1"); }
        }

        public string DistribucionPmAplica
        {
            get
            {
                return _distribucionPmAplica;
            }

            set
            {
                _distribucionPmAplica = value;
            }
        }

        public string DistribucionPmCuentaDebe
        {
            get
            {
                return _distribucionPmCuentaDebe;
            }

            set
            {
                _distribucionPmCuentaDebe = value;
            }
        }

        public string DistribucionPmCuentaHaber
        {
            get
            {
                return _distribucionPmCuentaHaber;
            }

            set
            {
                _distribucionPmCuentaHaber = value;
            }
        }

        public List<Bds> ListaCompannia
        {
            get
            {
                return _listaCompannia;
            }

            set
            {
                _listaCompannia = value;
            }
        }

        public string FacturaPmTIPORETENCION
        {
            get
            {
                return _facturaPmTIPORETENCION;
            }

            set
            {
                _facturaPmTIPORETENCION = value;
            }
        }

        public string FacturaSopnumbe
        {
            get
            {
                return _facturaSopSopnumbe;
            }

            set
            {
                _facturaSopSopnumbe = value;
            }
        }

        public string FacturaSopDocdate
        {
            get
            {
                return _facturaSopDocdate;
            }

            set
            {
                _facturaSopDocdate = value;
            }
        }

        public string FacturaSopDuedate
        {
            get
            {
                return _facturaSopDuedate;
            }

            set
            {
                _facturaSopDuedate = value;
            }
        }

        public string FacturaSopTXRGNNUM
        {
            get
            {
                return _facturaSopTXRGNNUM;
            }

            set
            {
                _facturaSopTXRGNNUM = value;
            }
        }

        public string FacturaSopUNITPRCE
        {
            get
            {
                return _facturaSopUNITPRCE;
            }

            set
            {
                _facturaSopUNITPRCE = value;
            }
        }

        public int FacturaSopColumnaMensajes
        {
            get
            {
                return int.Parse(_facturaSopColumnaMensajes);
            }

            set
            {
                _facturaSopColumnaMensajes = value.ToString();
            }
        }

        public int FacturaSopFilaInicial
        {
            get
            {
                return int.Parse( _facturaSopFilaInicial);
            }

            set
            {
                _facturaSopFilaInicial = value.ToString();
            }
        }

        public string FacturaSopCUSTNAME
        {
            get
            {
                return _facturaSopCUSTNAME;
            }

            set
            {
                _facturaSopCUSTNAME = value;
            }
        }

        public string ClienteDefaultCUSTCLAS
        {
            get
            {
                return _clienteDefaultCUSTCLAS;
            }

            set
            {
                _clienteDefaultCUSTCLAS = value;
            }
        }
    }
}
