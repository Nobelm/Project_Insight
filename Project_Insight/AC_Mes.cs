﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    static class AC_Mes
    {
        public static AC_Sem Semana1 = new AC_Sem();
        public static AC_Sem Semana2 = new AC_Sem();
        public static AC_Sem Semana3 = new AC_Sem();
        public static AC_Sem Semana4 = new AC_Sem();
        public static AC_Sem Semana5 = new AC_Sem();
    }

    class AC_Sem
    {
        public string Vym_Izq { get; set; }
        public string Vym_Der { get; set; }
        public string Vym_Cap { get; set; }
        public string Rp_Izq { get; set; }
        public string Rp_Der { get; set; }
        public string Rp_Cap { get; set; }
    }
}
