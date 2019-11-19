using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

/*Developed by AGR-Systems Science and Tech Division*/
namespace Project_Insight
{
    public class AC_Mes
    {
        public AC_Sem Semana1 = new AC_Sem();
        public AC_Sem Semana2 = new AC_Sem();
        public AC_Sem Semana3 = new AC_Sem();
        public AC_Sem Semana4 = new AC_Sem();
        public AC_Sem Semana5 = new AC_Sem();
    }

    public class AC_Sem
    {
        public DateTime Fecha_VyM { get; set; }
        public DateTime Fecha_RP  { get; set; }
        public string Aseo      { get; set; }
        public string Vym_Cap   { get; set; }
        public string Vym_Izq   { get; set; }
        public string Vym_Der   { get; set; }
        public string Rp_Cap    { get; set; }
        public string Rp_Izq    { get; set; }
        public string Rp_Der    { get; set; }
        public bool Conv_Week   { get; set; }
        public bool Vst_Week    { get; set; }
        public short Num_of_Week { get; set; }

         public List<string> Get_Asignee_List()
        {
            List<string> Asignee = new List<string>
            {
                Vym_Cap,
                Vym_Izq,
                Vym_Der,
                Rp_Cap,
                Rp_Izq,
                Rp_Der
            };
            return Asignee;
        }

        public void AutoFill()
        {
            while (Persistence.attending_persistance)
            {
                Thread.Sleep(100);
            }
            List<Person> People = new List<Person>();
            List<string> Asignee_vym = new List<string>();
            List<string> Asignee_rp = new List<string>();
            VyM_Sem sem_vym;
            RP_Sem sem_rp;
            switch (Num_of_Week)
            {
                case 1:
                    {
                        sem_vym = Main_Form.VyM_mes.Semana1;
                        sem_rp = Main_Form.RP_mes.Semana1;
                        break;
                    }
                case 2:
                    {
                        sem_vym = Main_Form.VyM_mes.Semana2;
                        sem_rp = Main_Form.RP_mes.Semana2;
                        break;
                    }
                case 3:
                    {
                        sem_vym = Main_Form.VyM_mes.Semana3;
                        sem_rp = Main_Form.RP_mes.Semana3;
                        break;
                    }
                case 4:
                    {
                        sem_vym = Main_Form.VyM_mes.Semana4;
                        sem_rp = Main_Form.RP_mes.Semana4;
                        break;
                    }
                default:
                    {
                        sem_vym = Main_Form.VyM_mes.Semana5;
                        sem_rp = Main_Form.RP_mes.Semana5;
                        break;
                    }
            }

            Asignee_vym.Add(sem_vym.Presidente);
            Asignee_vym.Add(sem_vym.Consejero_Aux);
            Asignee_vym.Add(sem_vym.Discurso_A);
            Asignee_vym.Add(sem_vym.Perlas);
            Asignee_vym.Add(sem_vym.Lectura_A);
            Asignee_vym.Add(sem_vym.Lectura_B);
            Asignee_vym.Add(sem_vym.SMM1_A);
            Asignee_vym.Add(sem_vym.SMM1_B);
            Asignee_vym.Add(sem_vym.SMM2_A);
            Asignee_vym.Add(sem_vym.SMM2_B);
            Asignee_vym.Add(sem_vym.SMM3_A);
            Asignee_vym.Add(sem_vym.SMM3_B);
            Asignee_vym.Add(sem_vym.SMM4_A);
            Asignee_vym.Add(sem_vym.SMM4_B);
            Asignee_vym.Add(sem_vym.NVC1_A);
            Asignee_vym.Add(sem_vym.NVC2_A);
            Asignee_vym.Add(sem_vym.Libro_A);
            Asignee_vym.Add(sem_vym.Libro_L);
            Asignee_vym.Add(sem_vym.Oracion);

            Asignee_rp.Add(sem_rp.Presidente);
            Asignee_rp.Add(sem_rp.Discursante);
            Asignee_rp.Add(sem_rp.Conductor);
            Asignee_rp.Add(sem_rp.Lector);
            Asignee_rp.Add(sem_rp.Oracion);

            /*------------------------ VyM Handler --------------------------*/

            if ((Vym_Cap == null) || (Vym_Cap == ""))
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Capitan.Contains('/'))
                    {
                        Person ps = new Person
                        {
                            Name = Main_Form.Male_List[i].Name,
                            Date = Convert.ToDateTime(Main_Form.Male_List[i].Capitan)
                        };
                        People.Add(ps);
                    }
                }
                People.Sort(delegate (Person ps1, Person ps2)
                {
                    return DateTime.Compare(ps1.Date, ps2.Date);
                });
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee_vym.Contains(People[i].Name))
                    {
                        Vym_Cap = People[i].Name;
                        Asignee_vym.Add(Vym_Cap);
                        Asignee_rp.Add(Vym_Cap);  //In order to not repeat same Cap in RP
                        break;
                    }
                }
            }
            People.Clear();

            for (int i = 0; i < Main_Form.Male_List.Count; i++)
            {
                if (Main_Form.Male_List[i].Acomodador.Contains('/'))
                {
                    Person ps = new Person
                    {
                        Name = Main_Form.Male_List[i].Name,
                        Date = Convert.ToDateTime(Main_Form.Male_List[i].Acomodador)
                    };
                    People.Add(ps);
                }
            }
            People.Sort(delegate (Person ps1, Person ps2)
            {
                return DateTime.Compare(ps1.Date, ps2.Date);
            });

            if ((Vym_Izq == null) || (Vym_Izq == ""))
            {
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee_vym.Contains(People[i].Name))
                    {
                        Vym_Izq = People[i].Name;
                        Asignee_vym.Add(Vym_Izq);
                        Asignee_rp.Add(Vym_Izq);
                        break;
                    }
                }
            }

            if ((Vym_Der == null) || (Vym_Der == ""))
            {
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee_vym.Contains(People[i].Name))
                    {
                        Vym_Der = People[i].Name;
                        Asignee_vym.Add(Vym_Der);
                        Asignee_rp.Add(Vym_Der);
                        break;
                    }
                }
            }
            People.Clear();

            /*------------------------ RP Handler --------------------------*/

            if ((Rp_Cap == null) || (Rp_Cap == ""))
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Capitan.Contains('/'))
                    {
                        Person ps = new Person
                        {
                            Name = Main_Form.Male_List[i].Name,
                            Date = Convert.ToDateTime(Main_Form.Male_List[i].Capitan)
                        };
                        People.Add(ps);
                    }
                }
                People.Sort(delegate (Person ps1, Person ps2)
                {
                    return DateTime.Compare(ps1.Date, ps2.Date);
                });
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee_rp.Contains(People[i].Name))
                    {
                        Rp_Cap = People[i].Name;
                        Asignee_rp.Add(Rp_Cap);
                        break;
                    }
                }
            }
            People.Clear();


            for (int i = 0; i < Main_Form.Male_List.Count; i++)
            {
                if (Main_Form.Male_List[i].Acomodador.Contains('/'))
                {
                    Person ps = new Person
                    {
                        Name = Main_Form.Male_List[i].Name,
                        Date = Convert.ToDateTime(Main_Form.Male_List[i].Acomodador)
                    };
                    People.Add(ps);
                }
            }
            People.Sort(delegate (Person ps1, Person ps2)
            {
                return DateTime.Compare(ps1.Date, ps2.Date);
            });

            if ((Rp_Izq == null) || (Rp_Izq == ""))
            {
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee_rp.Contains(People[i].Name))
                    {
                        Rp_Izq = People[i].Name;
                        Asignee_rp.Add(Rp_Izq);
                        break;
                    }
                }
            }

            if ((Rp_Der == null) || (Rp_Der == ""))
            {
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee_rp.Contains(People[i].Name))
                    {
                        Rp_Der = People[i].Name;
                        Asignee_rp.Add(Rp_Der);
                        break;
                    }
                }
            }
            People.Clear();
            Asignee_rp.Clear();
            Asignee_vym.Clear();

        }

        public class Person
        {
            public string Name;
            public DateTime Date;
        }
    }
}
