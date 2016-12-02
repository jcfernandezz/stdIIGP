using System;
using System.Collections.Generic;
using System.Text;

namespace Comun
{
    public class SopDocument
    {
        private string _idDoc;
        private short _soptype;
        private string _sopnumbe;

        public SopDocument(string idDoc, short soptype, string sopnumbe)
        {
            this._idDoc = idDoc;
            this._soptype = soptype;
            this._sopnumbe = sopnumbe;
        }

        public string idDoc
        {
            get { return _idDoc; }
            set { _idDoc = value; }
        }
        public short soptype
        {
            get { return _soptype; }
            set { _soptype = value; } 
        }
        public string sopnumbe
        {
            get { return _sopnumbe; }
            set { _sopnumbe = value; }
        }
    }
}
