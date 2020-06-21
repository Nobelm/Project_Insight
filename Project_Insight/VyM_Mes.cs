using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

/*Developed by AGR-Systems Science and Tech Division*/
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
        public DateTime Fecha { get; set; }
        public string Sem_Biblia { get; set; }
        public string Presidente { get; set; }
        public string Consejero_Aux { get; set; }
        public string Discurso { get; set; }
        public string Discurso_A { get; set; } //Asignado
        public string Perlas { get; set; }
        public string Lectura { get; set; }
        public string Lectura_A { get; set; }
        public string Lectura_B { get; set; }
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
        public string Libro_Titulo { get; set; }
        public string Libro_A { get; set; }
        public string Libro_L { get; set; } //Lector
        public string Oracion { get; set; }
        public bool HW_Data { get; set; }
        public bool Conv_Week { get; set; }
        public bool Vst_Week { get; set; }
        public short Num_of_Week { get; set; }
        public void Clear()
        {
           // Fecha = "";
            Sem_Biblia = "";
            Presidente = "";
            Consejero_Aux = "";
            Discurso = "";
            Discurso_A = "";
            Perlas = "";
            Lectura_A = "";
            Lectura_B = "";
            SMM1 = "";
            SMM1_A = "";
            SMM1_B = "";
            SMM2 = "";
            SMM2_A = "";
            SMM2_B = "";
            SMM3 = "";
            SMM3_A = "";
            SMM3_B = "";
            SMM4 = "";
            SMM4_A = "";
            SMM4_B = "";
            NVC1 = "";
            NVC1_A = "";
            NVC2 = "";
            NVC2_A = "";
            Libro_A = "";
            Libro_L = "";
            Oracion = "";
        }

        public void AutoFill()
        {
            List<string> Asignee = Get_Asignee_List();

            Libro_L = Asignee_Handler(Libro_L, ref Asignee, "Libro_L");
            Oracion = Asignee_Handler(Oracion, ref Asignee, "Oracion");
        }

        public List<string> Get_Asignee_List()
        {
            List<string> Asignee = new List<string>
            {
                Presidente,
                Consejero_Aux,
                Discurso_A,
                Perlas,
                Lectura_A,
                Lectura_B,
                SMM1_A,
                SMM1_B,
                SMM2_A,
                SMM2_B,
                SMM3_A,
                SMM3_B,
                SMM4_A,
                SMM4_B,
                NVC1_A,
                NVC2_A,
                Libro_A,
                Libro_L,
                Oracion,
            };

            return Asignee;
        }

        private string Asignee_Handler(string Field, ref List<string> Asignee, string iterator)
        {
            string Final_Asigned = "";
            List<Person> People = new List<Person>();
            if ((Field == null) || (Field == ""))
            {
                switch (iterator)
                {
                    case "Libro_L":
                        {
                            for (int i = 0; i < Main_Form.Male_List.Count; i++)
                            {                                                           //Custom rule: Elders not read on VyM
                                if (Main_Form.Male_List[i].Lector.Contains('/') && Main_Form.Male_List[i].Male_Type != Main_Form.Male_Type.Anciano)
                                {
                                    Person ps = new Person
                                    {
                                        Name = Main_Form.Male_List[i].Name,
                                        Date = Convert.ToDateTime(Main_Form.Male_List[i].Lector)
                                    };
                                    People.Add(ps);
                                }
                            }
                            break;
                        }
                    case "Oracion":
                        {
                            for (int i = 0; i < Main_Form.Male_List.Count; i++)
                            {
                                if (Main_Form.Male_List[i].Oracion.Contains('/'))
                                {
                                    Person ps = new Person
                                    {
                                        Name = Main_Form.Male_List[i].Name,
                                        Date = Convert.ToDateTime(Main_Form.Male_List[i].Oracion)
                                    };
                                    People.Add(ps);
                                }
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

        private class Person
        {
            public string Name;
            public DateTime Date;
        }

        public void Save_Heavensward_Info(VyM_Sem sem)
        {
            if (Thread.CurrentThread.Name == "Heavensward")
            {
                Sem_Biblia = sem.Sem_Biblia;
                Discurso = sem.Discurso;
                Lectura = sem.Lectura;
                SMM1 = sem.SMM1;
                SMM2 = sem.SMM2;
                SMM3 = sem.SMM3;
                SMM4 = sem.SMM4;
                NVC1 = sem.NVC1;
                NVC2 = sem.NVC2;
                HW_Data = sem.HW_Data;
            }
        }
    }
}
