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
        public string SMM1_B { get; set; }
        public string SMM2 { get; set; }
        public string SMM2_A { get; set; }
        public string SMM2_B { get; set; }
        public string SMM3 { get; set; }
        public string SMM3_A { get; set; }
        public string SMM3_B { get; set; }
        public string SMM4 { get; set; }
        public string SMM4_A { get; set; }
        public string SMM4_B { get; set; }
        public string NVC1 { get; set; }
        public string NVC1_A { get; set; }
        public string NVC2 { get; set; }
        public string NVC2_A { get; set; }
        public string Libro_A { get; set; }
        public string Libro_L { get; set; } //Lector
        public string Oracion { get; set; }

        public void AutoFill()
        {
            List<Person> People = new List<Person>();
            List<string> Asignee = new List<string>
            {
                Lectura,
                SMM1_A,
                SMM1_B,
                SMM2_A,
                SMM2_B,
                SMM3_A,
                SMM3_B,
                SMM4_A,
                SMM4_B
            };

            Libro_L = Asignee_Handler(Libro_L, ref Asignee, "Libro_L");
            Oracion = Asignee_Handler(Oracion, ref Asignee, "Oracion");
        }

        public string Asignee_Handler(string Field, ref List<string> Asignee, string iterator)
        {
            string Final_Asigned = "";
            List<Person> People = new List<Person>();
            if ((Field == null) || (Field == ""))
            {
                switch (iterator)
                {
                    case "Libro_L":
                        {
                            foreach (DB_Gnr item in DB_Form.Generals)
                            {
                                Person ps = new Person
                                {
                                    Name = item.Nombre,
                                    Date = item.Lec_VyM
                                };
                                People.Add(ps);
                            }
                            break;
                        }
                    case "Oracion":
                        {
                            foreach (DB_Gnr item in DB_Form.Generals)
                            {
                                Person ps = new Person
                                {
                                    Name = item.Nombre,
                                    Date = item.Ora_VyM,
                                };
                                People.Add(ps);
                            }
                            break;
                        }
                }

                People.Sort(delegate (Person ps1, Person ps2)
                {
                    return DateTime.Compare(ps1.Date, ps2.Date);
                });
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee.Contains(People[i].Name))
                    {
                        Final_Asigned = People[i].Name;
                        Asignee.Add(Final_Asigned);
                        break;
                    }
                }
            }
            else
            {
                Final_Asigned = Field;
            }
            People.Clear();
            return Final_Asigned;
        }

        public class Person
        {
            public string Name;
            public DateTime Date;
        }
    }
}
