using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Mns
    {
        public string Nombre     { get; set; }
        public string Capitan  { get; set; }
        public string Acom     { get; set; }
        public string Pres_RP  { get; set; }
        public string Lec_RP   { get; set; }
        public string Ora_RP   { get; set; }
        public string Cpt_Aseo { get; set; }

        public DB_Mns(string _Nombre, string _Capitan, string _Acom, string _Pres_RP, string _Lec_RP, string _Ora_RP, string _Cpt_Aseo)
        {
            Nombre   = _Nombre;
            Capitan  = _Capitan;
            Acom     = _Acom;
            Pres_RP  = _Pres_RP;
            Lec_RP   = _Lec_RP;
            Ora_RP   = _Ora_RP;
            Cpt_Aseo = _Cpt_Aseo;
        }
    }
}
