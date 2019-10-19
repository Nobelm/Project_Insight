using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*Developed by AGR-Systems Science and Tech Division*/
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
        public bool HW_Data { get; set; }
        public short Num_of_Week { get; set; }
        public void Clear()
        {
            Fecha = "";
            Titulo = "";
            Presidente = "";
            Congregacion = "";
            Discursante = "";
            Titulo_Atalaya = "";
            Conductor = "";
            Lector = "";
            Oracion = "";
            Discu_Sal = "";
            Ttl_Sal = "";
            Cong_Sal = "";
        }

        public void AutoFill()
        {
            List<string> Asignee = Get_Asignee_List();
            int last_week = 4;
            if (Main_Form.week_five_exist)
            {
                last_week = 5;
            }

            List<Person> People = new List<Person>();

            if ((Presidente == null) || (Presidente == ""))
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {                                                           
                    if (Main_Form.Male_List[i].Pres_RP.Contains('/'))
                    {
                        Person ps = new Person
                        {
                            Name = Main_Form.Male_List[i].Name,
                            Date = Convert.ToDateTime(Main_Form.Male_List[i].Pres_RP)
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
                    if (!Asignee.Contains(People[i].Name))
                    {
                        Presidente = People[i].Name;
                        Asignee.Add(Presidente);
                        break;
                    }
                }
            }
            People.Clear();

            if ((Lector == null) || (Lector == ""))
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Lector.Contains('/'))
                    {
                        Person ps = new Person
                        {
                            Name = Main_Form.Male_List[i].Name,
                            Date = Convert.ToDateTime(Main_Form.Male_List[i].Lector)
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
                    if (!Asignee.Contains(People[i].Name))
                    {
                        Lector = People[i].Name;
                        Asignee.Add(Lector);
                        break;
                    }
                }
            }
            People.Clear();

            if ((Oracion == null) || (Oracion == ""))
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
                People.Sort(delegate (Person ps1, Person ps2)
                {
                    return DateTime.Compare(ps1.Date, ps2.Date);
                });
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee.Contains(People[i].Name))
                    {
                        Oracion = People[i].Name;
                        Asignee.Add(Oracion);
                        break;
                    }
                }
            }
            People.Clear();

            if ((Conductor == null) || (Conductor == "") && (Num_of_Week == last_week))
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Atalaya.Contains('/'))
                    {
                        Person ps = new Person
                        {
                            Name = Main_Form.Male_List[i].Name,
                            Date = Convert.ToDateTime(Main_Form.Male_List[i].Atalaya)
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
                    if (!Asignee.Contains(People[i].Name))
                    {
                        Conductor = People[i].Name;
                        Asignee.Add(Conductor);
                        break;
                    }
                }
            }
            People.Clear();
        }

        public List<string> Get_Asignee_List()
        {
            List<string> Asignee = new List<string>
            {
                Presidente,
                Discursante,
                Conductor,
                Lector,
                Oracion,
                Discu_Sal
            };
            return Asignee;
        }

        public class Person
        {
            public string Name;
            public DateTime Date;
        }
    }
}
