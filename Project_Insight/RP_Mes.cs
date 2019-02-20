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

        public void AutoFill()
        {
            List<Person> People = new List<Person>();
            List<string> Asignee = new List<string>();

            if ((Presidente == null) || (Presidente == ""))
            {
                foreach (DB_Eld item in DB_Form.Elders)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Pres_RP,
                    };
                    People.Add(ps);
                }
                foreach (DB_Mns item in DB_Form.Ministerials)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Pres_RP,
                    };
                    People.Add(ps);
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
                foreach (DB_Eld item in DB_Form.Elders)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Lec_RP,
                    };
                    People.Add(ps);
                }
                foreach (DB_Mns item in DB_Form.Ministerials)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Lec_RP,
                    };
                    People.Add(ps);
                }
                foreach (DB_Gnr item in DB_Form.Generals)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Lec_RP,
                    };
                    People.Add(ps);
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
                foreach (DB_Eld item in DB_Form.Elders)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Ora_RP,
                    };
                    People.Add(ps);
                }
                foreach (DB_Mns item in DB_Form.Ministerials)
                {
                    Person ps = new Person
                    {
                        Name = item.Nombre,
                        Date = item.Ora_RP,
                    };
                    People.Add(ps);
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
        }

        public class Person
        {
            public string Name;
            public DateTime Date;
        }
    }
}
