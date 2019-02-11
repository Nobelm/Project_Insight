using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class AC_Mes
    {
        public AC_Sem Semana1 = new AC_Sem();
        public AC_Sem Semana2 = new AC_Sem();
        public AC_Sem Semana3 = new AC_Sem();
        public AC_Sem Semana4 = new AC_Sem();
        public AC_Sem Semana5 = new AC_Sem();
    }

    public class AC_Sem
    {
        public string Cp_Aseo_VyM  { get; set; }
        public string Cp_Aseo_RP   { get; set; }
        public string Vym_Cap      { get; set; }
        public string Vym_Izq      { get; set; }
        public string Vym_Der      { get; set; }
        public string Rp_Cap       { get; set; }
        public string Rp_Izq       { get; set; }
        public string Rp_Der       { get; set; }
    }
}
