using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAutomation_Project.Model
{
    class Sottolineatura
    {
        private String _strSottolineatura;
        private int _valSottolineatura;

        public Sottolineatura(string strSottolineatura, int valSottolineatura)
        {
            this._strSottolineatura = strSottolineatura;
            this._valSottolineatura = valSottolineatura;
        }

        public string StrSottolineatura { get => _strSottolineatura; set => _strSottolineatura = value; }
        public int ValSottolineatura { get => _valSottolineatura; set => _valSottolineatura = value; }
    }
}
