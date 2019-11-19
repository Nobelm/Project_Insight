using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net;
using System.IO;

/*Developed by AGR-Systems Science and Tech Division*/
namespace Project_Insight
{
    public class Heavensward
    {
        private static int tdb_attend = 0;
        private static int smm_attend = 0;
        private static int nvc_attend = 0;
        private static int current_meeting = 99; // 0 IG, 1 TDLB, 2 SMM, 3 NVC
        private static string month = "";
        private const string find_bible_week = "Guía de actividades";
        private const string find_hidden_perls = "Busquemos perlas escondidas";
        private const string find_book_study = "Estudio bíblico";
        private const string find_watchtower = "Artículo de estudio";
        public static string[] Bible_Books;
        public static string[] Smm_keys = { "sm_11", "sm_21", "sm_31", "sm_41" };
        public static string[] Nvc_keys = { "nv_11", "nv_21" };
        private static bool pending_break = false;
        private static bool break_reader = false;
        public static bool Hw_inProgress = false;
        private static bool Hw_oracle_inProgress = false;
        private static bool month_found = false;
        private static bool WT_found = false;
        private static VyM_Mes VyM_mes = new VyM_Mes();
        private static VyM_Sem Aux_VyM_Sem = new VyM_Sem();
        private static RP_Mes RP_mes = new RP_Mes();
        private static RP_Sem Aux_RP_Sem = new RP_Sem();
        public static List<HW_request> HW_Requests_List = new List<HW_request>();
        public static List<HW_Oracle_Request> HW_Oracle_Requests_List = new List<HW_Oracle_Request>();
        public static Timer Heavensward_Timer;
        private static bool Initial_Check = false;
        //public static string AGR_Pass = "AdvancedGeneralResearch";
        public static List<string> Assignment_VyM_List = new List<string>
            {
               "Presidente",
               "Consejero_Aux",
               "Discurso_A",
               "Perlas",
               "Lectura_A",
               "Lectura_B",
               "SMM1_A",
               "SMM1_B",
               "SMM2_A",
               "SMM2_B",
               "SMM3_A",
               "SMM3_B",
               "SMM4_A",
               "SMM4_B",
               "NVC1_A",
               "NVC2_A",
               "Libro_A",
               "Libro_L",
               "Oracion",
            };
        public static List<string> Assignment_RP_List = new List<string>
            {
               "Presidente",
               "Discursante",
               "Conductor",
               "Lector",
               "Oracion",
               "Discu_Sal"
            };

        public static List<string> Assignment_AC_List = new List<string>
        {
               "Vym_Cap",
               "Vym_Izq",
               "Vym_Der",
               "Rp_Cap",
               "Rp_Izq",
               "Rp_Der"
        };

        public static void Start_Heavensward()
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Heavensward";
            }
            Heavensward_Timer = new Timer(
                new TimerCallback(HW_Timer_Tick),
                null,
                1000, //Time which pass after its creation in ms
                1000  //Period
                );
            string raw = Properties.Resources.Bible_Books;
            Bible_Books = raw.Split('\n');
            for (int i = 0; i < Bible_Books.Length -1; i++)
            {
                Bible_Books[i] = Bible_Books[i].Remove(Bible_Books[i].Length - 1);
            }
            VyM_mes.Semana1.Num_of_Week = 1;
            VyM_mes.Semana2.Num_of_Week = 2;
            VyM_mes.Semana3.Num_of_Week = 3;
            VyM_mes.Semana4.Num_of_Week = 4;
            VyM_mes.Semana5.Num_of_Week = 5;
            RP_mes.Semana1.Num_of_Week = 1;
            RP_mes.Semana2.Num_of_Week = 2;
            RP_mes.Semana3.Num_of_Week = 3;
            RP_mes.Semana4.Num_of_Week = 4;
            RP_mes.Semana5.Num_of_Week = 5;
        }

        private static void HW_Timer_Tick(object sender)
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Heavensward";
            }
            if (!Initial_Check)
            {
                Initial_Check = true;
                Initial_Heavensward_Check();
            }
            if ((HW_Requests_List.Count > 0) && !Hw_inProgress)
            {
                Access_Heaven(HW_Requests_List[0]);
            }
            if ((HW_Oracle_Requests_List.Count > 0) && !Hw_oracle_inProgress)
            {
                Heavensward_Oracle_Handler(HW_Oracle_Requests_List[0]);
            }
        }

        public class HW_request
        {
            public DateTime date;
            public int tab;
            public int week;
        }

        public class HW_Oracle_Request
        {
            public VyM_Sem hw_oracle_vym;
            public RP_Sem hw_oracle_rp;
            public AC_Sem hw_oracle_ac;
        }
        
        public static void HW_Bridge(DateTime inc_date, int tab, int week)
        {
            HW_request request = new HW_request
            {
                date = inc_date,
                tab = tab,
                week = week
            };
            HW_Requests_List.Add(request);
        }

        private static void Initial_Heavensward_Check()
        {
            try
            {
                using (var client = new WebClient())
                using (client.OpenRead("http://clients3.google.com/generate_204"))
                {
                    Main_Form.Notify("Initial Check: Internet Connection: Granted!");
                }
            }
            catch
            {
                Main_Form.Warn("Initial Check: Internet Connection: Denied!");
            }
        }

        /*-------------------- Attending request -------------------- */

        private static void Access_Heaven(HW_request info)
        {
            Hw_inProgress = true;
            string fecha = info.date.ToString("yyyy/MM/dd");
            string url = "https://wol.jw.org/es/wol/dt/r4/lp-s/" + fecha;
            month = info.date.ToString("MMMM");
            Main_Form.Notify("Gather information from wol.jw.org for " + fecha);
            /*Aux_RP_Sem.Clear();
            Aux_VyM_Sem.Clear();*/
            switch (info.week)
            {
                case 1:
                    {
                        Aux_VyM_Sem = VyM_mes.Semana1;
                        Aux_RP_Sem = RP_mes.Semana1;
                        break;
                    }
                case 2:
                    {
                        Aux_VyM_Sem = VyM_mes.Semana2;
                        Aux_RP_Sem = RP_mes.Semana2;
                        break;
                    }
                case 3:
                    {
                        Aux_VyM_Sem = VyM_mes.Semana3;
                        Aux_RP_Sem = RP_mes.Semana3;
                        break;
                    }
                case 4:
                    {
                        Aux_VyM_Sem = VyM_mes.Semana4;
                        Aux_RP_Sem = RP_mes.Semana4;
                        break;
                    }
                case 5:
                    {
                        Aux_VyM_Sem = VyM_mes.Semana5;
                        Aux_RP_Sem = RP_mes.Semana5;
                        break;
                    }
            }
            if (Aux_VyM_Sem.HW_Data && Aux_RP_Sem.HW_Data)
            {
                Insert_previously_Saved(info.tab);
                HW_Requests_List.RemoveAt(0);
                Hw_inProgress = false;
            }
            else
            {
                try
                {
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string raw;// = reader.ReadToEnd(); //Test
                        CleanUp();
                        Main_Form.Notify("Connection Successfull. Reading info");
                        while ((raw = reader.ReadLine()) != null)
                        {
                            if (break_reader)
                            {
                                break;
                            }
                            if (info.tab == 0 || info.tab == 1)
                            {
                                String_Handler(raw);
                            }
                        }
                        Main_Form.Notify("Information successfully recieved from wol.jw.org");
                    }
                }
                catch
                {
                    Main_Form.Warn("Unable to connect to wol.jw.org\n");
                    Insert_previously_Saved(info.tab);
                }
                HW_Requests_List.RemoveAt(0);
                Hw_inProgress = false;
            }
        }

        private static void Insert_previously_Saved(int tab)
        {
            if ((Aux_VyM_Sem.HW_Data == true) && (tab == 0))
            {
                Main_Form.Notify("Insert previously saved data for VyM");
                Main_Form.Hw_requested_info request = new Main_Form.Hw_requested_info
                {
                    vym_sem = Aux_VyM_Sem,
                };
                Main_Form.HW_request.Add(request);
            }
            else
                if ((Aux_RP_Sem.HW_Data == true) && (tab == 1))
            {
                Main_Form.Notify("Insert previously saved data for RP");
                Main_Form.Hw_requested_info request = new Main_Form.Hw_requested_info
                {
                    rp_sem = Aux_RP_Sem
                };
                Main_Form.HW_request.Add(request);
            }
            else
            {
                Main_Form.Warn("No info available");
            }
        }

        private static void CleanUp()
        {
            tdb_attend = 0;
            smm_attend = 0;
            nvc_attend = 0;
            pending_break = false;
            break_reader = false;
            current_meeting = 99;
            month_found = false;
            WT_found = false;
        }

        //Extract the usefull information from string to vym
        private static void String_Handler(string str)
        {
            Get_current_meeting(str);
            switch (current_meeting)
            {
                case 0:
                    {
                        if(month_found)
                        {
                            Aux_VyM_Sem.Sem_Biblia = Analyze_string(str);
                            current_meeting = 99;
                        }
                        if (str.Contains(month))
                        {
                            month_found = true;
                        }
                        break;
                    }
                case 1:
                    {
                        if(str.Contains("<li>"))
                        {
                            string aux = Analyze_string(str);
                            if (aux.Contains("mins."))
                            {
                                if (tdb_attend == 0)
                                {
                                    Aux_VyM_Sem.Discurso = aux;
                                    tdb_attend++;
                                }
                                else if(!aux.Contains(find_hidden_perls))
                                {
                                    Aux_VyM_Sem.Lectura = aux;
                                    current_meeting = 99;
                                }

                            }
                        }
                        break;
                    }
                case 2:
                    {
                        if (str.Contains("<li>"))
                        {
                            string aux = Analyze_string(str);
                            if(aux.Contains("mins."))
                            {
                                switch (smm_attend)
                                {
                                    case 0:
                                        {
                                            Aux_VyM_Sem.SMM1 = aux;
                                            break;
                                        }
                                    case 1:
                                        {
                                            Aux_VyM_Sem.SMM2 = aux;
                                            break;
                                        }
                                    case 2:
                                        {
                                            Aux_VyM_Sem.SMM3 = aux;
                                            break;
                                        }
                                    case 3:
                                        {
                                            Aux_VyM_Sem.SMM4 = aux;
                                            break;
                                        }
                                }
                                smm_attend++;
                                if (smm_attend >= 4)
                                {
                                    smm_attend = Smm_keys.Length - 1;
                                }
                            }
                        }
                        break;
                    }
                case 3:
                    {
                        if(str.Contains(find_book_study))
                        {
                            current_meeting = 99;
                            //break_reader = true;
                            break;
                        }
                        if (str.Contains("<li>"))
                        {
                            string aux = Analyze_string(str);
                            if (aux.Contains("mins."))
                            {
                                if (nvc_attend == 0)
                                {
                                    Aux_VyM_Sem.NVC1 = aux;
                                    nvc_attend++;
                                }
                                else
                                {
                                    Aux_VyM_Sem.NVC2 = aux;
                                }
                            }
                        }
                        break;
                    }
                case 4:
                    {
                        String_rp_handler(str);
                        break;
                    }
            }
            if (break_reader)
            {
                Save_VyM_Information();
                Main_Form.Hw_requested_info request = new Main_Form.Hw_requested_info
                {
                    vym_sem = Aux_VyM_Sem,
                };
                Main_Form.HW_request.Add(request);
            }
        }

        private static void Get_current_meeting(string str)
        {
            if (str.Contains(find_bible_week))
            {
                current_meeting = 0;
            }
            else if (str.Contains("TESOROS DE LA BIBLIA"))
            {
                current_meeting = 1;
            }
            else if (str.Contains("SEAMOS MEJORES MAESTROS"))
            {
                current_meeting = 2;
            }
            else if (str.Contains("NUESTRA VIDA CRISTIANA"))
            {
                current_meeting = 3;
            }
            else if (str.Contains(find_watchtower))
            {
                current_meeting = 4;
            }
        }
        private static void String_rp_handler(string str)
        {
            if(WT_found)
            {
                string final_value = Analyze_string(str);
                final_value = final_value.Substring(2); //String have number page at the beginning 
                if (final_value != "")
                {
                    Aux_RP_Sem.Titulo_Atalaya = final_value;
                    Main_Form.Hw_requested_info request = new Main_Form.Hw_requested_info
                    {
                        rp_sem = Aux_RP_Sem
                    };
                    Main_Form.HW_request.Add(request);
                    Save_RP_Information();
                }
                break_reader = true;
            }
            if (str.Contains(find_watchtower))
            {
                WT_found = true;
            }
        }

        private static string Analyze_string(string str)
        {
            var array = str.ToCharArray();
            bool open = false;
            string retval = "";
            pending_break = false;
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i].Equals('<'))
                {
                    open = false;
                }
                else if (array[i].Equals('>'))
                {
                    open = true;
                }
                else if (open)
                {
                    retval += array[i].ToString();
                }
                if (current_meeting == 3)
                {
                    if (array[i].Equals(')') && pending_break)
                    {
                        break;
                    }
                    if (retval.Contains("mins"))
                    {
                        pending_break = true;
                    }
                }
            }
            retval = retval.Replace("  ", "");
            return retval;
        }

        private static void Save_VyM_Information()
        {
            switch (Aux_VyM_Sem.Num_of_Week)
            {
                case 1:
                    {
                        VyM_mes.Semana1 = Aux_VyM_Sem;
                        VyM_mes.Semana1.HW_Data = true;
                        break;
                    }
                case 2:
                    {
                        VyM_mes.Semana2 = Aux_VyM_Sem;
                        VyM_mes.Semana2.HW_Data = true;
                        break;
                    }
                case 3:
                    {
                        VyM_mes.Semana3 = Aux_VyM_Sem;
                        VyM_mes.Semana3.HW_Data = true;
                        break;
                    }
                case 4:
                    {
                        VyM_mes.Semana4 = Aux_VyM_Sem;
                        VyM_mes.Semana4.HW_Data = true;
                        break;
                    }
                case 5:
                    {
                        VyM_mes.Semana5 = Aux_VyM_Sem;
                        VyM_mes.Semana5.HW_Data = true;
                        break;
                    }
            }
        }

        private static void Save_RP_Information()
        {
            switch (Aux_RP_Sem.Num_of_Week)
            {
                case 1:
                    {
                        RP_mes.Semana1 = Aux_RP_Sem;
                        RP_mes.Semana1.HW_Data = true;
                        break;
                    }
                case 2:
                    {
                        RP_mes.Semana2 = Aux_RP_Sem;
                        RP_mes.Semana2.HW_Data = true;
                        break;
                    }
                case 3:
                    {
                        RP_mes.Semana3 = Aux_RP_Sem;
                        RP_mes.Semana3.HW_Data = true;
                        break;
                    }
                case 4:
                    {
                        RP_mes.Semana4 = Aux_RP_Sem;
                        RP_mes.Semana4.HW_Data = true;
                        break;
                    }
                case 5:
                    {
                        RP_mes.Semana5 = Aux_RP_Sem;
                        RP_mes.Semana5.HW_Data = true;
                        break;
                    }
            }
        }

        /*-------------------------- Heavensward Oracle Handler -------------------------------*/

        private static void Heavensward_Oracle_Handler(HW_Oracle_Request request)
        {
            Hw_oracle_inProgress = true;
            string asignee = "Noel Belin";
            bool asignee_found = false;
            string asig = "";
            string date = "";
            //Main_Form.Notify("Search for Noel Belin in week: " + request.hw_oracle_vym.Num_of_Week.ToString());
            if (request.hw_oracle_vym != null)
            {
                List<string> VyM_Asignee = request.hw_oracle_vym.Get_Asignee_List();
                if (VyM_Asignee.Count > 0)
                {
                    for (int i = 0; i < VyM_Asignee.Count; i++)
                    {
                        if (VyM_Asignee[i].Contains(asignee))
                        {
                            Main_Form.Notify("Noel Belin Found in " + Assignment_VyM_List[i] + ", date " + request.hw_oracle_vym.Fecha.ToString("dddd, dd MMMM, yyyy"));
                            asig = Assignment_VyM_List[i];
                            date = request.hw_oracle_vym.Fecha.ToString("dddd, dd MMMM"); ;
                            asignee_found = true;
                        }
                    }
                }
            }
            else if (request.hw_oracle_rp != null)
            {
                List<string> RP_Asignee = request.hw_oracle_rp.Get_Asignee_List();
                if (RP_Asignee.Count > 0)
                {
                    for (int i = 0; i < RP_Asignee.Count; i++)
                    {
                        if (RP_Asignee[i].Contains(asignee))
                        {
                            Main_Form.Notify("Noel Belin Found in " + Assignment_RP_List[i] + ", date " + request.hw_oracle_rp.Fecha.ToString("dddd, dd MMMM, yyyy"));
                            asig = Assignment_RP_List[i];
                            date = request.hw_oracle_rp.Fecha.ToString("dddd, dd MMMM");
                            asignee_found = true;
                        }
                    }
                }
            }
            else if (request.hw_oracle_ac != null)
            {
                List<string> AC_Asignee = request.hw_oracle_ac.Get_Asignee_List();
                if (AC_Asignee.Count > 0)
                {
                    for (int i = 0; i < AC_Asignee.Count; i++)
                    {
                        if (AC_Asignee[i].Contains(asignee))
                        {
                            string fecha;
                            if (i < 3)
                            {
                                fecha = request.hw_oracle_ac.Fecha_VyM.ToString("dddd, dd MMMM, yyyy");
                            }
                            else
                            {
                                fecha = request.hw_oracle_ac.Fecha_RP.ToString("dddd, dd MMMM, yyyy");
                            }
                            Main_Form.Notify("Noel Belin Found in " + Assignment_AC_List[i] + ", date " + fecha);
                            asig = Assignment_AC_List[i];
                            date = fecha;
                            asignee_found = true;
                        }
                    }
                }
            }

            if (asignee_found)
            {
                date = date.Replace(" ", "+");
                try
                {
                    Main_Form.Notify("Accesing to Heavensward Oracle Network");
                    string url_local = "https://us-central1-agr-connected-services.cloudfunctions.net/Heavensward_Oracle";
                    string modifier1 = "?rw=write";
                    string modifier2 = "&asig=" + asig;
                    string modifier3 = "&date=" + date;
                    /*using (var client = new WebClient())
                    using (client.OpenRead(url_local))
                    {
                        Main_Form.Notify("Welcome to AGR Connected Services!");
                    }*/
                    HttpWebRequest Oracle_Request = (HttpWebRequest)WebRequest.Create(url_local + modifier1 + modifier2 + modifier3);
                    using (HttpWebResponse response = (HttpWebResponse)Oracle_Request.GetResponse())
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string raw = reader.ReadToEnd();
                        Main_Form.Notify(raw);
                    }
                }
                catch
                {
                    Main_Form.Warn("Connection: Denied!");
                }
            }
            HW_Oracle_Requests_List.RemoveAt(0);
            Hw_oracle_inProgress = false;
        }
    }
}
