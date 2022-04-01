using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using OfficeAutomation_Project.Model;

namespace OfficeAutomation_Project.Controller
{
    class AllineamentoController
    {
        public List<Allineamento> GetAllineamenti()
        {
            List<Allineamento> _lstAllineamento = new List<Allineamento>();

            foreach (WdParagraphAlignment al in Enum.GetValues(typeof(WdParagraphAlignment)))
            {
                _lstAllineamento.Add(new Allineamento(al.ToString(),(int)al));
            }

            return _lstAllineamento;
        }
    }
}
