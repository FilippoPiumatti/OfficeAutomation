using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using OfficeAutomation_Project.Model;

namespace OfficeAutomation_Project.Controller
{
    class SottolineaturaController
    {
        public List<Sottolineatura> GetSottolineaturas()
        {
            List<Sottolineatura> _lstSottolineatura = new List<Sottolineatura>();

            foreach (WdUnderline unLine in Enum.GetValues(typeof(WdUnderline)))
            {
                _lstSottolineatura.Add(new Sottolineatura(unLine.ToString(), (int)unLine));
            }

            return _lstSottolineatura;
        }
    }
}
