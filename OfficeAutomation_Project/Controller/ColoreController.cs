using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using OfficeAutomation_Project.Model;

namespace OfficeAutomation_Project.Controller
{
    class ColoreController
    {

       public List<Colore> GetColores()
       {
            List<Colore> _lstColore = new List<Colore>();

            foreach (WdColor color in Enum.GetValues(typeof(WdColor)))
            {
                _lstColore.Add(new Colore(color.ToString() , (int)color));
            }

            return _lstColore;
       }
    }
}
