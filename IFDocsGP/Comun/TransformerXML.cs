using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using System.IO;

namespace Comun
{
    public class TransformerXML
    {
        private Dictionary<string, XslCompiledTransform> transforms = new Dictionary<string, XslCompiledTransform>();
        public string cadenaOriginal = "";
        public int numErrores = 0;             // Validation Error Count
        public string mensajeError = "";        // Validation Error Message

        public XslCompiledTransform Load(string rutaArchivoXSLT)
        {
            numErrores = 0;
            mensajeError = "";
            XslCompiledTransform transform = null;
            try
            {
                if (!transforms.TryGetValue(rutaArchivoXSLT, out transform))
                {
                    transform = new XslCompiledTransform();
                    transform.Load(rutaArchivoXSLT);
                    transforms[rutaArchivoXSLT] = transform;
                }
                return transform;
            }
            catch (Exception lo)
            {
                mensajeError = "Error al inicializar la plantilla de transformación de XML. Verifique la ruta:"+ rutaArchivoXSLT+ " " + lo.Message;
                numErrores++;
                return transform;
            }

        }

        /// <summary>
        /// Transforma el xml a cadena original
        /// </summary>
        /// <param name="archivoXml">Archivo xml a transformar.</param>
        /// <param name="transformer">Objeto que aplica un xslt al archivo xml.</param>
        /// <returns>False cuando hay al menos un error</returns>
        public bool getCadenaOriginal(XmlDocument archivoXml, XslCompiledTransform transformer)
        {
            StringWriter writer = new StringWriter();
            mensajeError = "";
            numErrores=0;
            try
            {
                transformer.Transform(archivoXml, null, writer);
                cadenaOriginal = writer.ToString();
                return true;
            }
            catch (Exception eXsl)
            {
                mensajeError = "[getCadenaOriginal] Contacte al administrador. Error al generar la cadena original. " + eXsl.Message;
                numErrores++;
                return false;
            }
        }
    }
}
