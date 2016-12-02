using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Text;

namespace Comun
{

public class ValidadorXML
{
    public int numErrores = 0;             // Validation Error Count
    public string mensajeError = "";        // Validation Error Message
    private XmlSchemaSet sc;                // Esquema

    public ValidadorXML(Parametros prm)
    {
        // Create the XmlSchemaSet class.
        sc = new XmlSchemaSet();
        try
        {
            // Add the schema to the collection.
            sc.Add(null, prm.URLArchivoXSD);
        }
        catch
        {
            mensajeError = "No se encontró el esquema en el URL: " + prm.URLArchivoXSD;
            numErrores++;
        }
    }

    // Display any warnings or errors.
    private void ValidationCallBack(object sender, ValidationEventArgs args)
    {
        if (args.Severity == XmlSeverityType.Warning)
            mensajeError = "No se encontró el esquema. No se pudo validar el archivo xml. Verifique la configuración. " + args.Message;
            //Console.WriteLine("\tWarning: Matching schema not found.  No validation occurred." + args.Message);
        else
            mensajeError = args.Message;
            //Console.WriteLine("\tValidation error: " + args.Message);
        numErrores++;
    }

    public bool ValidarXSD(XmlDocument archivoXml)
    {
        numErrores = 0;
        mensajeError = "";
        XmlNodeReader nodeReader = new XmlNodeReader(archivoXml);

        // Set the validation settings.
        XmlReaderSettings settings = new XmlReaderSettings();
        settings.ValidationType = ValidationType.Schema;
        settings.Schemas = sc;
        settings.ValidationEventHandler += new ValidationEventHandler (ValidationCallBack);

        try
        {
            // Create the XmlReader object.
            XmlReader reader = XmlReader.Create(nodeReader, settings);
            // Parse the file. 
            while (reader.Read()) ;
        }
        catch (Exception eXsd)
        {
            mensajeError = "[ValidarXSD] Contacte al administrador. Error al abrir el documento XML. " + eXsd.Message;
            numErrores++;
        }
        return numErrores == 0;
    }
}



}
