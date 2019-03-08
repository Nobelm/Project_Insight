using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public string Cp_Aseo_VyM  { get; set; }
        public string Cp_Aseo_RP   { get; set; }
        public string Vym_Cap      { get; set; }
        public string Vym_Izq      { get; set; }
        public string Vym_Der      { get; set; }
        public string Rp_Cap       { get; set; }
        public string Rp_Izq       { get; set; }
        public string Rp_Der       { get; set; }

        public void AutoFill(int sem)
        {
            List<Person> People = new List<Person>();
            List<string> Asignee_vym = new List<string>();
            List<string> Asignee_rp = new List<string>();
            VyM_Sem sem_vym;
            RP_Sem sem_rp;
            switch (sem)
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
            Asignee_vym.Add(sem_vym.Discurso_A);
            Asignee_vym.Add(sem_vym.Perlas);
            Asignee_vym.Add(sem_vym.Lectura);
            Asignee_vym.Add(sem_vym.SMM1_A);
            Asignee_vym.Add(sem_vym.SMM2_A);
            Asignee_vym.Add(sem_vym.SMM3_A);
            Asignee_vym.Add(sem_vym.SMM4_A);
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

            /*Seeting for multi type fill*/
            int higher_count_eld_mins = (DB_Form.Elders.Count > DB_Form.Ministerials.Count) ? DB_Form.Elders.Count : DB_Form.Ministerials.Count;
            int higher_count_all = (higher_count_eld_mins > DB_Form.Generals.Count) ? higher_count_eld_mins : DB_Form.Generals.Count;
            for (int i = 0; i < higher_count_eld_mins; i++)
            {

            }

            foreach (DB_Eld item in DB_Form.Elders)
            {
                Person ps = new Person
                {
                    Name = item.Nombre,
                    Date = item.Cpt_Aseo,
                };
                People.Add(ps);
            }
            foreach (DB_Mns item in DB_Form.Ministerials)
            {
                Person ps = new Person
                {
                    Name = item.Nombre,
                    Date = item.Cpt_Aseo,
                };
                People.Add(ps);
            }
            People.Sort(delegate (Person ps1, Person ps2)
            {
                return DateTime.Compare(ps1.Date, ps2.Date);
            });
            bool asigned_vym = false, asigned_rp = false;
            for (int i = 0; i < People.Count; i++)
            {
                if (!Asignee_vym.Contains(People[i].Name) && !asigned_vym)
                {
                    Cp_Aseo_VyM = People[i].Name;
                    Asignee_vym.Add(Cp_Aseo_VyM);
                    asigned_vym = true;
                }
                else if (!Asignee_rp.Contains(People[i].Name) && !asigned_rp)
                {
                    Cp_Aseo_RP = People[i].Name;
                    Asignee_rp.Add(Cp_Aseo_RP);
                    asigned_rp = true;
                }
                if (asigned_vym && asigned_rp)
                {
                    break;
                }
            }
            People.Clear();
            /*----------*/
            foreach (DB_Eld item in DB_Form.Elders)
            {
                Person ps = new Person
                {
                    Name = item.Nombre,
                    Date = item.Capitan,
                };
                People.Add(ps);
            }
            foreach (DB_Mns item in DB_Form.Ministerials)
            {
                Person ps = new Person
                {
                    Name = item.Nombre,
                    Date = item.Capitan,
                };
                People.Add(ps);
            }
            People.Sort(delegate (Person ps1, Person ps2)
            {
                return DateTime.Compare(ps1.Date, ps2.Date);
            });
            asigned_vym = false;
            asigned_rp = false;
            for (int i = 0; i < People.Count; i++)
            {
                if (!Asignee_vym.Contains(People[i].Name) && !asigned_vym)
                {
                    Vym_Cap = People[i].Name;
                    Asignee_vym.Add(Vym_Cap);
                    asigned_vym = true;
                }
                else if (!Asignee_rp.Contains(People[i].Name) && !asigned_rp)
                {
                    Rp_Cap = People[i].Name;
                    Asignee_rp.Add(Rp_Cap);
                    asigned_rp = true;
                }
                if (asigned_vym && asigned_rp)
                {
                    break;
                }
            }
            People.Clear();

            /*----------*/
            foreach (DB_Gnr item in DB_Form.Generals)
            {
                Person ps = new Person
                {
                    Name = item.Nombre,
                    Date = item.Acom,
                };
                People.Add(ps);
            }
            foreach (DB_Mns item in DB_Form.Ministerials)
            {
                Person ps = new Person
                {
                    Name = item.Nombre,
                    Date = item.Acom,
                };
                People.Add(ps);
            }
            People.Sort(delegate (Person ps1, Person ps2)
            {
                return DateTime.Compare(ps1.Date, ps2.Date);
            });
            asigned_vym = false;
            asigned_rp = false;
            bool asigned_vym_2 = false, asigned_rp_2 = false;
            for (int i = 0; i < People.Count - 1; i++)
            {
                if (!Asignee_vym.Contains(People[i].Name) && !asigned_vym)
                {
                    Vym_Izq = People[i].Name;
                    Asignee_vym.Add(Vym_Izq);
                    asigned_vym = true;
                }
                else if (!Asignee_rp.Contains(People[i].Name) && !asigned_rp)
                {
                    Rp_Izq = People[i].Name;
                    Asignee_rp.Add(Rp_Izq);
                    asigned_rp = true;
                }
                else if (!Asignee_vym.Contains(People[i].Name) && !asigned_vym_2)
                {
                    Vym_Der = People[i].Name;
                    Asignee_vym.Add(Vym_Der);
                    asigned_vym_2 = true;
                }
                else if (!Asignee_rp.Contains(People[i].Name) && !asigned_rp_2)
                {
                    Rp_Der = People[i].Name;
                    Asignee_rp.Add(Rp_Der);
                    asigned_rp_2 = true;
                }
                if (asigned_vym && asigned_rp && asigned_vym_2 && asigned_rp_2)
                {
                    break;
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
