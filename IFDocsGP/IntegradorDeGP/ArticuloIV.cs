using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

using Comun;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;
using OfficeOpenXml;

namespace IntegradorDeGP
{
    class ArticuloIV
    {
        public int iError = 0;
        public string sMensaje = "";
        //public StringBuilder sItemVendorXml ;
        public taCreateItemVendors_ItemsTaCreateItemVendors [] itemsVendors ;
        public IVVendorItemType vendorItemType ;

        public ArticuloIV(int numArtsProvs)
        {
            //sItemVendorXml = new StringBuilder();
            itemsVendors = new taCreateItemVendors_ItemsTaCreateItemVendors[numArtsProvs];
            vendorItemType = new IVVendorItemType();
        }

        /// <summary>
        /// Arma un objeto artículo de proveedor.
        /// </summary>
        /// <param name="hojaXl"></param>
        /// <param name="fila"></param>
        /// <param name="param"></param>
        /// <param name="linea"></param>
        public void armaArtProvEconn(ExcelWorksheet hojaXl, int fila, Parametros param, int linea)
        {
            try
            {
                iError = 0;

                taCreateItemVendors_ItemsTaCreateItemVendors articuloIvVendor = new taCreateItemVendors_ItemsTaCreateItemVendors();
                articuloIvVendor.ITEMNMBR = param.defaultInventoryItem;
                articuloIvVendor.VENDORID = hojaXl.Cells[fila, Convert.ToInt32(param.facturaPopCaVENDORID)].Value.ToString();
                articuloIvVendor.VNDITNUM = param.defaultInventoryItem;
                articuloIvVendor.UpdateIfExists = 0;

                itemsVendors[linea] = articuloIvVendor;
            }
            catch (Exception errorGral)
            {
                sMensaje = "Error al armar artículos de proveedor. " + errorGral.Message + " [armaFacturaDeEconn]";
                iError++;
            }

        }

        //public void Serializa()
        //{ 
        //    try
        //    {
        //        iError = 0;
        //        eConnectType eConnect = new eConnectType();
        //        IVVendorItemType vendorItemType = new IVVendorItemType();

        //        XmlSerializer serializer = new XmlSerializer(eConnect.GetType());

        //        vendorItemType.taCreateItemVendors_Items = this.itemsVendors;

        //        IVVendorItemType[] myVendorItemType = { vendorItemType};
        //        eConnect.IVVendorItemType = myVendorItemType;

        //        XmlWriterSettings sett = new XmlWriterSettings();
        //        sett.Encoding = Encoding.UTF8;
        //        using (XmlWriter writer = XmlWriter.Create(this.sItemVendorXml, sett))
        //        {
        //            serializer.Serialize(writer, eConnect);
        //        }
        //    }
        //    catch (Exception errorGral)
        //    {
        //        sMensaje = "Error al serializar artículos de proveedor. " + errorGral.Message + " [Serializa]";
        //        iError++;
        //    }

        //}

    }
}
