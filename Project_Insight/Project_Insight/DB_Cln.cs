using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Cln
    {
        public string Grupo { get; set; }
        public DateTime Fecha { get; set; }

        public DB_Cln(string _Grupo, string _Fecha)
        {
            Grupo = _Grupo;
            if (_Fecha.Contains('/'))
            {
                Fecha = Convert.ToDateTime(_Fecha);
            }
        }
    }
}
