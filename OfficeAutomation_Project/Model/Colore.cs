using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAutomation_Project.Model
{
    class Colore
    {
        private String _strColore;
        private int _valColore;

        public Colore(string strColore, int valColore)
        {
            this._strColore = strColore;
            this._valColore = valColore;
        }

        public string StrColore { get => _strColore; set => _strColore = value; }
        public int ValColore { get => _valColore; set => _valColore = value; }
    }
}
