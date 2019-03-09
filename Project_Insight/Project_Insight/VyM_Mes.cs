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

        /*Autofill for Libro_L and Oracion*/
        public void AutoFill()
        {
            string[] str_name = new string[3];
            if ((Libro_L == null) || (Libro_L == ""))
            {
                str_name = DB_Form.Get_VyM_Assigned(nameof(Libro_L));
                Libro_L = Select_member(str_name);
            }
            else if ((Oracion == null) || (Oracion == ""))
            {
                str_name = DB_Form.Get_VyM_Assigned(nameof(Oracion));
                Oracion = Select_member(str_name);
            }
        }

        private string Select_member(string[] str_name)
        {
            string name = "";
            object[] members = { Presidente, Discurso_A, Perlas, Lectura, SMM1_A, SMM2_A, SMM3_A, SMM4_A, NVC1_A, NVC2_A, Libro_L, Oracion };
            for (int i = 0; i < 3; i++)
            {
                bool compare = false;
                foreach (var item in members)
                {
                    compare = str_name[i].Equals(item);
                    if (compare)
                    {
                        break;
                    }
                }
                if (!compare)
                {
                    name = str_name[i];
                    break;
                }                
            }
            return name;
        }
    }

}
