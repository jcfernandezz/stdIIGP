using System;
using System.Collections.Generic;
using System.Text;

namespace Comun
{
    public class ConexionAFuenteDatos
    {
        private string _Compannia = "";
        private string _Usuario = "";
        private string _Password = "";
        private string _Intercompany = "";
        private string _ServerAddress = "";
        private bool _IntegratedSecurity = false;
        private string _ConnStr = "";
        private string _ConnStrDyn = "";
        private string _EntityConnStr = string.Empty;

        public ConexionAFuenteDatos(string Compannia, string Usuario, string Password, string Intercompany, string ServerAddress, bool IntegratedSecurity)
        {
            this._Compannia = Compannia;
            this._Usuario = Usuario;
            this._Password = Password;
            this._Intercompany = Intercompany;
            this._ServerAddress = ServerAddress;
            this._IntegratedSecurity = IntegratedSecurity;
            this._ConnStr = ArmaConnStr(IntegratedSecurity);
            this._ConnStrDyn = ArmaConnStrDynamics(IntegratedSecurity);
        }

        private string ArmaConnStr (bool SeguridadIntegrada)
        {
            if (SeguridadIntegrada)
                return "Initial Catalog=" + _Intercompany + ";Data Source=" + _ServerAddress + ";Integrated Security=SSPI";
            else
                return "User ID=" + _Usuario + ";Password=" + _Password + ";Initial Catalog=" + _Intercompany + ";Data Source=" + _ServerAddress;
        }

        private string ArmaConnStrDynamics(bool SeguridadIntegrada)
        {
            if (SeguridadIntegrada)
                return "Initial Catalog=Dynamics;Data Source=" + _ServerAddress + ";Integrated Security=SSPI";
            else
                return _ConnStrDyn = "User ID=" + _Usuario + ";Password=" + _Password + ";Initial Catalog=Dynamics;Data Source=" + _ServerAddress;
        }

        private string ArmaEntityConnStr(bool SeguridadIntegrada)
        {
            if (SeguridadIntegrada)
                return "metadata=res://*/ModelGP.csdl|res://*/ModelGP.ssdl|res://*/ModelGP.msl;provider=System.Data.SqlClient;provider connection string=';data source=" + _ServerAddress + ";initial catalog=" + _Intercompany + ";integrated security=True;MultipleActiveResultSets=True;App=EntityFramework';";
            else
                return "";  //completar aquí cuando no es seg integrada
        }

        public string Compannia
        {
            get { return _Compannia; }
            set { _Compannia = value; }
        }

        public string Usuario
        {
            get { return _Usuario; }
            set { 
                _Usuario = value;
                _ConnStr = ArmaConnStr(_IntegratedSecurity);
                _ConnStrDyn = ArmaConnStrDynamics(IntegratedSecurity);
                _EntityConnStr = ArmaEntityConnStr(_IntegratedSecurity);
            }
        }

        public string Password
        {
            get { return _Password; }
            set { 
                _Password = value;
                _ConnStr = ArmaConnStr(_IntegratedSecurity);
                _ConnStrDyn = ArmaConnStrDynamics(IntegratedSecurity);
                _EntityConnStr = ArmaEntityConnStr(_IntegratedSecurity);
            }
        }

        public string Intercompany
        {
            get { return _Intercompany; }
            set {
                _Intercompany = value;
                _ConnStr = ArmaConnStr(_IntegratedSecurity);
                _EntityConnStr = ArmaEntityConnStr(_IntegratedSecurity);
            }
        }

        public string ServerAddress
        {
            get { return _ServerAddress; }
            set { 
                _ServerAddress = value;
                _ConnStr = ArmaConnStr(_IntegratedSecurity);
                _ConnStrDyn = ArmaConnStrDynamics(IntegratedSecurity);
                _EntityConnStr = ArmaEntityConnStr(_IntegratedSecurity);
            }
        }

        public bool IntegratedSecurity
        {
            get { return _IntegratedSecurity; }
        }

        public string ConnStr
        {
            get { return _ConnStr; }
            //set { _ConnStr  = value; }
        }

        public string ConnStrDyn
        {
            get { return _ConnStrDyn; }
        }

        public string EntityConnStr
        {
            get
            {
                return _EntityConnStr;
            }

        }
    }
}
