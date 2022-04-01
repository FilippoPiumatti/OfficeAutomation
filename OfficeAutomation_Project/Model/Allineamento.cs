using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAutomation_Project.Model
{
    class Allineamento
    {
        private String _strAllineamento;
        private int _valAllineamento;

        public Allineamento(string strAllineamento, int valAllienamento)
        {
             this._strAllineamento = strAllineamento;
             this._valAllineamento = valAllienamento;
        }

        public string StrAllineamento { get => _strAllineamento; set => _strAllineamento = value; }
        public int ValAllineamento { get => _valAllineamento; set => _valAllineamento = value; }
    }
}
