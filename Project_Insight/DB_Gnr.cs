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

        public DB_Gnr(string _Nombre, string _Acom, string _Lec_RP, string _Lec_VyM, string _Ora_VyM)
        {
            Nombre  = _Nombre;
            Acom    = GetDate(_Acom);
            Lec_RP  = GetDate(_Lec_RP);
            Lec_VyM = GetDate(_Lec_VyM);
            Ora_VyM = GetDate(_Ora_VyM);
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
