using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    public class Insight_Month
    {
        public Insight_Sem Semana1 = new Insight_Sem();
        public Insight_Sem Semana2 = new Insight_Sem();
        public Insight_Sem Semana3 = new Insight_Sem();
        public Insight_Sem Semana4 = new Insight_Sem();
        public Insight_Sem Semana5 = new Insight_Sem();
    }

    public class Insight_Sem
    {
        //VyM
        public DateTime Fecha_VyM { get; set; }
        public DateTime Fecha_RP { get; set; }
        public string Cancion_VyM_1 { get; set; }
        public string Cancion_VyM_2 { get; set; }
        public string Cancion_VyM_3 { get; set; }
        public string Sem_Biblia { get; set; }
        public string Presidente_VyM { get; set; }
        public string Consejero_Aux { get; set; }
        public string Discurso_VyM { get; set; }
        public string Discurso_VyM_A { get; set; } //Asignado
        public string Perlas { get; set; }
        public string Lectura_Biblia { get; set; }
        public string Lectura_Biblia_A { get; set; }
        public string Lectura_Biblia_B { get; set; }
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
        public string Libro_Conductor { get; set; }
        public string Libro_Lector { get; set; } //Lector
        public string Oracion_End_VyM { get; set; }
        //RP
        public string Cancion_RP_1 { get; set; }
        public string Cancion_RP_2 { get; set; }
        public string Cancion_RP_3 { get; set; }
        public string Presidente_RP { get; set; }
        public string Titulo_Discurso_RP { get; set; }
        public string Congregacion_RP { get; set; }
        public string Discursante_RP { get; set; }
        public string Titulo_Atalaya { get; set; }
        public string Conductor_Atalaya { get; set; }
        public string Lector_Atalaya { get; set; }
        public string Oracion_End_RP { get; set; }
        public string Discu_Sal { get; set; }
        public string Ttl_Sal { get; set; }
        public string Cong_Sal { get; set; }
        //AC
        public string Aseo { get; set; }
        public string Vym_Cap { get; set; }
        public string Vym_Izq { get; set; }
        public string Vym_Der { get; set; }
        public string Rp_Cap { get; set; }
        public string Rp_Izq { get; set; }
        public string Rp_Der { get; set; }
        //Settings
        public bool HW_Data { get; set; }
        public Main_Form.Special_Meeting_Type Special_VyM_Meeting { get; set; }
        public Main_Form.Special_Meeting_Type Special_RP_Meeting { get; set; }
        public string Special_VyM_Meeting_Info { get; set; }
        public string Special_RP_Meeting_Info { get; set; }
        public short Num_of_Week { get; set; }
        public bool Overwatch_Aprobal { get; set; }
        public void Clear(Main_Form.Clear_Insight clear_Insight)
        {
            // Fecha = "";
            if (clear_Insight == Main_Form.Clear_Insight.Clear_VyM || clear_Insight == Main_Form.Clear_Insight.Clear_Full)
            {
                Cancion_VyM_1 = "";
                Cancion_VyM_2 = "";
                Cancion_VyM_3 = "";
                Sem_Biblia = "";
                Presidente_VyM = "";
                Consejero_Aux = "";
                Discurso_VyM = "";
                Discurso_VyM_A = "";
                Perlas = "";
                Lectura_Biblia = "";
                Lectura_Biblia_A = "";
                Lectura_Biblia_B = "";
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
                Libro_Conductor = "";
                Libro_Lector = "";
                Oracion_End_VyM = "";
            }
            if (clear_Insight == Main_Form.Clear_Insight.Clear_RP || clear_Insight == Main_Form.Clear_Insight.Clear_Full)
            {
                Cancion_RP_1 = "";
                Cancion_RP_2 = "";
                Cancion_RP_3 = "";
                Presidente_RP = "";
                Titulo_Discurso_RP = "";
                Congregacion_RP = "";
                Discursante_RP = "";
                Titulo_Atalaya = "";
                Conductor_Atalaya = "";
                Lector_Atalaya = "";
                Oracion_End_RP = "";
                Discu_Sal = "";
                Ttl_Sal = "";
                Cong_Sal = "";
            }
            if (clear_Insight == Main_Form.Clear_Insight.Clear_Ac || clear_Insight == Main_Form.Clear_Insight.Clear_Full)
            {
                Aseo = "";
                Vym_Cap = "";
                Vym_Izq = "";
                Vym_Der = "";
                Rp_Cap = "";
                Rp_Izq = "";
                Rp_Der = "";
            }
            HW_Data = false;
            Special_VyM_Meeting = Main_Form.Special_Meeting_Type.Non_status;
            Special_RP_Meeting = Main_Form.Special_Meeting_Type.Non_status;
            Special_VyM_Meeting_Info = "";
            Special_RP_Meeting_Info = "";
        }

        public void Save_Heavensward_Info(Insight_Sem sem)
        {
            Cancion_VyM_1 = sem.Cancion_VyM_1;
            Cancion_VyM_2 = sem.Cancion_VyM_2;
            Cancion_VyM_3 = sem.Cancion_VyM_3;
            Sem_Biblia = sem.Sem_Biblia;
            Discurso_VyM = sem.Discurso_VyM;
            Lectura_Biblia = sem.Lectura_Biblia;
            SMM1 = sem.SMM1;
            SMM2 = sem.SMM2;
            SMM3 = sem.SMM3;
            SMM4 = sem.SMM4;
            Cancion_RP_2 = sem.Cancion_RP_2;
            Cancion_RP_3 = sem.Cancion_RP_3;
            NVC1 = sem.NVC1;
            NVC2 = sem.NVC2;
            Titulo_Atalaya = sem.Titulo_Atalaya;
            HW_Data = sem.HW_Data;
        }
        public void AutoFill()
        {
            Asignee_VyM_Handler(); 
            Asignee_RP_Handler();
        }

        public List<string> Get_Asignee_VyM_List()
        {
            List<string> Asignee = new List<string>
            {
                Presidente_VyM,
                Consejero_Aux,
                Discurso_VyM_A,
                Perlas,
                Lectura_Biblia_A,
                Lectura_Biblia_B,
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
                Libro_Conductor,
                Libro_Lector,
                Oracion_End_VyM,
                Vym_Cap,
                Vym_Izq,
                Vym_Der
            };

            return Asignee;
        }
        public List<string> Get_Asignee_RP_List()
        {
            List<string> Asignee = new List<string>
            {
                Libro_Lector,
                Oracion_End_VyM,
                Presidente_RP,
                Discursante_RP,
                Conductor_Atalaya,
                Lector_Atalaya,
                Oracion_End_RP,
                Discu_Sal,
                Rp_Cap,
                Rp_Izq,
                Rp_Der,
                Vym_Cap,
                Vym_Izq,
                Vym_Der
            };

            return Asignee;
        }

        protected void Asignee_VyM_Handler()
        {
            List<string> Asignee = Get_Asignee_VyM_List();
            List<Person> People = new List<Person>();
            if (!(Special_VyM_Meeting_Info != null) || !Special_VyM_Meeting_Info.Contains("Visita"))
            {
                //Lector Libro
                if ((Libro_Lector == null) || (Libro_Lector == ""))
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
                }

                People.Sort(delegate (Person ps1, Person ps2)
                {
                    return DateTime.Compare(ps1.Date, ps2.Date);
                });
                for (int i = 0; i < People.Count; i++)
                {
                    if (!Asignee.Contains(People[i].Name))
                    {
                        Libro_Lector = People[i].Name;
                        Asignee.Add(Libro_Lector);
                        break;
                    }
                }
                People.Clear();
            }
            //Oracion Final VyM
            if ((Oracion_End_VyM == null) || (Oracion_End_VyM == ""))
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
                        Oracion_End_VyM = People[i].Name;
                        Asignee.Add(Oracion_End_VyM);
                        break;
                    }
                }
            }
            People.Clear();
            if (Special_VyM_Meeting != Main_Form.Special_Meeting_Type.Conv_type)
            {
                //Acomodadores
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
                        if (!Asignee.Contains(People[i].Name))
                        {
                            Vym_Cap = People[i].Name;
                            Asignee.Add(Vym_Cap);
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
                        if (!Asignee.Contains(People[i].Name))
                        {
                            Vym_Izq = People[i].Name;
                            Asignee.Add(Vym_Izq);
                            break;
                        }
                    }
                }

                if ((Vym_Der == null) || (Vym_Der == ""))
                {
                    for (int i = 0; i < People.Count; i++)
                    {
                        if (!Asignee.Contains(People[i].Name))
                        {
                            Vym_Der = People[i].Name;
                            Asignee.Add(Vym_Der);
                            break;
                        }
                    }
                }
            }
            People.Clear();
        }

        protected void Asignee_RP_Handler()
        {

            List<string> Asignee = Get_Asignee_RP_List();
            int last_week = 4;
            if (Main_Form.week_five_exist)
            {
                last_week = 5;
            }
            bool read_lector = true;
            List<Person> People = new List<Person>();

            if ((Presidente_RP == null) || (Presidente_RP == ""))
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
                        Presidente_RP = People[i].Name;
                        Asignee.Add(Presidente_RP);
                        break;
                    }
                }
            }
            People.Clear();
            if (Special_RP_Meeting_Info != null && Special_RP_Meeting_Info.Contains("Visita"))
            {
                read_lector = false;
            }

            if (read_lector)
            {
                if ((Lector_Atalaya == null) || (Lector_Atalaya == ""))
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
                            Lector_Atalaya = People[i].Name;
                            Asignee.Add(Lector_Atalaya);
                            break;
                        }
                    }
                }
                People.Clear();
            }

            if ((Oracion_End_RP == null) || (Oracion_End_RP == ""))
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
                        Oracion_End_RP = People[i].Name;
                        Asignee.Add(Oracion_End_RP);
                        break;
                    }
                }
            }
            People.Clear();

            if ((Conductor_Atalaya == null) || (Conductor_Atalaya == "") && (Num_of_Week == last_week))
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
                        Conductor_Atalaya = People[i].Name;
                        Asignee.Add(Conductor_Atalaya);
                        break;
                    }
                }
            }
            People.Clear();
            //Acomodadores
            if (Special_RP_Meeting != Main_Form.Special_Meeting_Type.Conv_type)
            {
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
                        if (!Asignee.Contains(People[i].Name))
                        {
                            Rp_Cap = People[i].Name;
                            Asignee.Add(Rp_Cap);
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
                        if (!Asignee.Contains(People[i].Name))
                        {
                            Rp_Izq = People[i].Name;
                            Asignee.Add(Rp_Izq);
                            break;
                        }
                    }
                }

                if ((Rp_Der == null) || (Rp_Der == ""))
                {
                    for (int i = 0; i < People.Count; i++)
                    {
                        if (!Asignee.Contains(People[i].Name))
                        {
                            Rp_Der = People[i].Name;
                            Asignee.Add(Rp_Der);
                            break;
                        }
                    }
                }
            }
            People.Clear();
        }

        private class Person
        {
             public string Name;
             public DateTime Date;
        }

    }
}


                                                                          