using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class VyM_Mes
    {
        public VyM_Sem Semana1 = new VyM_Sem();
        public VyM_Sem Semana2 = new VyM_Sem();
        public VyM_Sem Semana3 = new VyM_Sem();
        public VyM_Sem Semana4 = new VyM_Sem();
        public VyM_Sem Semana5 = new VyM_Sem();
    }

    public class VyM_Sem
    {
        public string Fecha { get; set; }
        public string Sem_Biblia { get; set; }
        public string Presidente { get; set; }
        public string Discurso { get; set; }
        public string Discurso_A { get; set; } //Asignado
        public string Perlas { get; set; }
        public string Lectura { get; set; }
        public string SMM1 { get; set; }
        public string SMM1_A { get; set; }
        public string SMM2 { get; set; }
        public string SMM2_A { get; set; }
        public string SMM3 { get; set; }
        public string SMM3_A { get; set; }
        public string SMM4 { get; set; }
        public string SMM4_A { get; set; }
        public string NVC1 { get; set; }
        public string NVC1_A { get; set; }
        public string NVC2 { get; set; }
        public string NVC2_A { get; set; }
        public string Libro_A { get; set; }
        public string Libro_L { get; set; } //Lector
        public string Oracion { get; set; }
    }
}
