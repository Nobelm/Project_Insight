﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Mns
    {
        public string Nombre     { get; set; }
        public DateTime Capitan  { get; set; }
        public DateTime Acom     { get; set; }
        public DateTime Pres_RP  { get; set; }
        public DateTime Lec_RP   { get; set; }
        public DateTime Ora_RP   { get; set; }
        public DateTime Cpt_Aseo { get; set; }

        public DB_Mns(string _Nombre, string _Capitan, string _Acom, string _Pres_RP, string _Lec_RP, string _Ora_RP, string _Cpt_Aseo)
        {
            Nombre   = _Nombre;
            Capitan  = GetDate(_Capitan);
            Acom     = GetDate(_Acom);
            Pres_RP  = GetDate(_Pres_RP);
            Lec_RP   = GetDate(_Lec_RP);
            Ora_RP   = GetDate(_Ora_RP);
            Cpt_Aseo = GetDate(_Cpt_Aseo);
        }

        public DateTime GetDate(string Str)
        {
            DateTime date = new DateTime(2019, 01, 01);
            if (Str.Contains('/'))
            {
                date = Convert.ToDateTime(Str);
            }
            return date;
        }
    }
}
