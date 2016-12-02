using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.AccessControl;

namespace Comun
{
    public class Utiles
    {
        static public int numErr = 0;
        static public string msgErr = "";
        /// <summary>
        /// Devuelve los últimos caracteres a la derecha del Texto
        /// </summary>
        /// <param name="Texto">Texto a procesar</param>
        /// <param name="Cuantos">Número de caracteres a retornar</param>
        /// <returns>String Devuelve los últimos caracteres a la derecha del Texto</returns>
        static public string Derecha(string Texto, short Cuantos)
        {
            if (Texto.Length > Cuantos && Cuantos > 0)
                return Texto.Remove(0, Texto.Length - Cuantos);
            else
                return Texto;
        }

        static public string Derecha(string Texto, int Cuantos)
        {
            if (Texto.Length > Cuantos && Cuantos > 0)
            {
                return Texto.Remove(0, Texto.Length - Cuantos);
            }
            else
                return Texto;
        }

        static public string Izquierda(string Texto, int Cuantos)
        {
            if (Texto.Length > Cuantos && Cuantos > 0)
                return Texto.Substring(0, Cuantos);
            else
                return Texto;
        }

        static public void SetRule(string filePath, string account, FileSystemRights rights, AccessControlType controlType)
        {
            FileSecurity fSecurity = File.GetAccessControl(filePath);
            fSecurity.ResetAccessRule(new FileSystemAccessRule(account, rights, controlType));
            File.SetAccessControl(filePath, fSecurity);
        }

        static public string FormatoNombreArchivo(string prefijo, string nombre, int largo)
        {
            string pre = prefijo.Trim().Replace(" ", "").Replace("'", "").Replace("&", "").Replace("<", "").Replace(">", "").Replace("/", "").Replace(@"\", "").Replace(",", "").Replace(".", "").Replace(";", "").Replace("@", "");
            string nom = nombre.Trim().PadRight(largo, '_').Substring(0, largo - 1).Replace(" ", "").Replace("'", "").Replace("&", "").Replace("<", "").Replace(">", "").Replace("/", "").Replace(@"\", "").Replace(",", "").Replace(".", "").Replace(";", "").Replace("@", "");
            return pre + "_" + nom;
        }

        /// <summary>
        /// Devuelve TRUE si convierte un string a Datetime exitosamente
        /// </summary>
        /// <param name="sFecha">Debe estar en el formato yyyyMMdd</param>
        /// <param name="fecha">Devuelve la fecha</param>
        /// <returns></returns>
        static public bool ConvierteAFecha(string sFecha, out DateTime fecha)
        {
            fecha = Convert.ToDateTime("1/1/1900");
            try
            {
                int iYear = Convert.ToInt16( sFecha.Substring(0,4));
                int iMonth = Convert.ToInt16( sFecha.Substring(4,2));
                int iDay = Convert.ToInt16( sFecha.Substring(6,2));
                fecha = new DateTime(iYear, iMonth, iDay);

                return true;
            }
            catch
            {
                return false;
            }

        }

        //static public bool ConvierteAFecha(string sFecha, Parametros param, out DateTime fecha)
        //{
        //    long serialDate = 0;
        //    fecha = Convert.ToDateTime("1/1/1900");
        //    try
        //    {
        //        int iYear = Convert.ToDateTime(sFecha, param.culture).Year;
        //        int iMonth = Convert.ToDateTime(sFecha, param.culture).Month ;
        //        int iDay = Convert.ToDateTime(sFecha, param.culture).Day;
        //        fecha = new DateTime(iYear, iMonth, iDay);
        //        return true;
        //    }
        //    catch
        //    {
        //        try
        //        {
        //            serialDate = long.Parse(sFecha);
        //            fecha = DateTime.FromOADate(serialDate);
        //            return true;
        //        }
        //        catch 
        //        {
        //            return false;
        //        }
        //    }
        //}
        /// <summary>
        /// Devuelve TRUE si convierte un string a Datetime exitosamente. La fecha tiene separadores.
        /// </summary>
        /// <param name="sFecha">Debe estar en el formato dd/mm/yyyy. El separador debe ser /</param>
        /// <param name="fecha">Devuelve la fecha</param>
        /// <returns></returns>
        static public bool ConvierteAFechaFmt(string sFecha, out DateTime fecha)
        {
            fecha = Convert.ToDateTime("1/1/1900");
            try
            {
                int dayIdx = sFecha.IndexOf('/', 0);
                int monthIdx = sFecha.IndexOf('/', dayIdx+1);
                int yearIdx = sFecha.IndexOf(' ', monthIdx+1);
                if (yearIdx < 0)
                    yearIdx = sFecha.Length;
                
                int iYear = Convert.ToInt16(sFecha.Substring(monthIdx + 1, yearIdx - monthIdx -1));
                int iMonth = Convert.ToInt16(sFecha.Substring(dayIdx + 1, monthIdx - dayIdx - 1));
                int iDay = Convert.ToInt16(sFecha.Substring(0, dayIdx));
                fecha = new DateTime(iYear, iMonth, iDay);

                return true;
            }
            catch
            {
                return false;
            }

        }

    }
}
