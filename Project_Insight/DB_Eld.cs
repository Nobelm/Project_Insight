using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Eld
    {
        public string Nombre { get; set; }
        public DateTime Capitan { get; set; }
        public DateTime Pres_RP { get; set; }
        public DateTime Lec_RP { get; set; }
        public DateTime Ora_RP { get; set; }
        public DateTime Atalaya { get; set; }


        public DB_Eld(string _Nombre, string _Capitan, string _Pres_RP, string _Lec_RP, string _Ora_RP, string _Atalaya)
        {
            Nombre = _Nombre;
            if (_Capitan.Contains('/'))
            {
                Capitan = Convert.ToDateTime(_Capitan);
            }
            if (_Pres_RP.Contains('/'))
            {
                Pres_RP = Convert.ToDateTime(_Pres_RP);
            }
            if (_Lec_RP.Contains('/'))
            {
                Lec_RP = Convert.ToDateTime(_Lec_RP);
            }
            if (_Ora_RP.Contains('/'))
            {
                Ora_RP = Convert.ToDateTime(_Ora_RP);
            }
            if (_Atalaya.Contains('/'))
            {
                Atalaya = Convert.ToDateTime(_Atalaya);
            }
        }
    }
}
