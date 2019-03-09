using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Eld
    {
        public string Nombre     { get; set; }
        public string Capitan  { get; set; }
        public string Pres_RP  { get; set; }
        public string Lec_RP   { get; set; }
        public string Ora_RP   { get; set; }
        public string Atalaya  { get; set; }
        public string Cpt_Aseo { get; set; }


        public DB_Eld(string _Nombre, string _Capitan, string _Pres_RP, string _Lec_RP, string _Ora_RP, string _Atalaya, string _Cpt_Aseo)
        {
            Nombre   = _Nombre;
            Capitan  = _Capitan;
            Pres_RP  = _Pres_RP;
            Lec_RP   = _Lec_RP;
            Ora_RP   = _Ora_RP;
            Atalaya  = _Atalaya;
            Cpt_Aseo = _Cpt_Aseo;
        }
    }
}
