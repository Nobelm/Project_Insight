using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class DB_Gnr
    {
        public string Nombre { get; set; }
        public DateTime Acom { get; set; }
        public DateTime Lec_RP { get; set; }
        public DateTime Ora_VyM { get; set; }
        public DateTime Lec_VyM { get; set; }

        public DB_Gnr(string _Nombre, string _Acom, string _Lec_RP, string _Ora_VyM, string _Lec_VyM)
        {
            Nombre = _Nombre;

            if (_Acom.Contains('/'))
            {
                Acom = Convert.ToDateTime(_Acom);
            }
            if (_Lec_RP.Contains('/'))
            {
                Lec_RP = Convert.ToDateTime(_Lec_RP);
            }
            if (_Ora_VyM.Contains('/'))
            {
                Ora_VyM = Convert.ToDateTime(_Ora_VyM);
            }
            if (_Lec_VyM.Contains('/'))
            {
                Lec_VyM = Convert.ToDateTime(_Lec_VyM);
            }
        }
    }
}
