﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Gnr
    {
        public string Nombre { get; set; }
        public string Acom { get; set; }
        public string Lec_RP { get; set; }
        public string Libro_L { get; set; }
        public string Oracion_VyM { get; set; }

        public DB_Gnr(string _Nombre, string _Acom, string _Lec_RP, string _Lec_VyM, string _Ora_VyM)
        {
            Nombre  = _Nombre;
            Acom    = _Acom;
            Lec_RP  = _Lec_RP;
            Libro_L = _Lec_VyM;
            Oracion_VyM = _Ora_VyM;
        }
    }
}
