using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class RP_Mes
    {
        public RP_Sem Semana1 = new RP_Sem();
        public RP_Sem Semana2 = new RP_Sem();
        public RP_Sem Semana3 = new RP_Sem();
        public RP_Sem Semana4 = new RP_Sem();
        public RP_Sem Semana5 = new RP_Sem();
    }

    public class RP_Sem
    {
        public string Fecha { get; set; }
        public string Titulo { get; set; }
        public string Presidente { get; set; }
        public string Congregacion { get; set; }
        public string Discursante { get; set; }
        public string Titulo_Atalaya { get; set; }
        public string Conductor { get; set; }
        public string Lector { get; set; }
        public string Oracion { get; set; }
        public string Discu_Sal { get; set; }
        public string Ttl_Sal { get; set; }
        public string Cong_Sal { get; set; }

        public void Autofill()
        {
            string[] str_name = new string[3];

        }
    }
}
