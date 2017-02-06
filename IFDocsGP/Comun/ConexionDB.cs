using System;
using System.Collections.Generic;
using System.Text;
using Comun;
using System.Security.Principal;

namespace Comun
{
    public class ConexionDB
    {
        //public static string sUsuario = Microsoft.Dexterity.Applications.Dynamics.Globals.UserId.Value;
        //public static string sPassword = Microsoft.Dexterity.Applications.Dynamics.Globals.SqlPassword.Value;
        //public static string sIntercompany = Microsoft.Dexterity.Applications.Dynamics.Globals.IntercompanyId.Value;
        //public static string sSqlDSN = Microsoft.Dexterity.Applications.Dynamics.Globals.SqlDataSourceName;

        private string _Compannia = "";
        private string _Usuario = "";
        private string _Password = "";
        private string _Intercompany = "";
        private string _ServerAddress = "";
        private bool _IntegratedSecurity = false;
        private List<Bds> _listaCompannias;

        public ConexionAFuenteDatos Elemento = null;
        public string ultimoMensaje = "";

        public List<Bds> ListaCompannias
        {
            get
            {
                return _listaCompannias;
            }

            set
            {
                _listaCompannias = value;
            }
        }

        public ConexionDB ()
        {
            ListaCompannias = new List<Bds>();
            Parametros config = new Parametros();
            _ServerAddress = config.servidor;
            ListaCompannias = config.ListaCompannia;
            ultimoMensaje = config.ultimoMensaje;

            //Si la app no viene de GP usar seguridad integrada o un usuario sql (configurado en el archivo de inicio)
            if (_Usuario.Equals(string.Empty))
            {
                _IntegratedSecurity = config.seguridadIntegrada;
                _Intercompany = "Dynamics";

                if (_IntegratedSecurity)            //Usar seguridad integrada
                    _Usuario = WindowsIdentity.GetCurrent().Name.Trim();
                else
                {                                   //Usar un usuario sql con privilegios
                    _Usuario = config.usuarioSql;
                    _Password = config.passwordSql  ;
                }
            }

            Elemento = new ConexionAFuenteDatos(_Compannia, _Usuario, _Password, _Intercompany, _ServerAddress, _IntegratedSecurity);

        }
    }
}
