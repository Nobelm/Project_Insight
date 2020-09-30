using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Threading;
using System.Security.Cryptography.X509Certificates;

namespace Project_Insight
{
    public class Overwatch
    {
        public class Overwatch_Object
        {
            public string Name { get; set; }
            public int Assignments { get; set; }
            public int Filled { get; set; }
            public string  Percentaje { get; set; }
            public Main_Form.Special_Meeting_Type Special_VyM_Meeting_Type { get; set; }
            public Main_Form.Special_Meeting_Type Special_RP_Meeting_Type { get; set; }
        }

        public static Overwatch_Object Semana1 = new Overwatch_Object();
        public static Overwatch_Object Semana2 = new Overwatch_Object();
        public static Overwatch_Object Semana3 = new Overwatch_Object();
        public static Overwatch_Object Semana4 = new Overwatch_Object();
        public static Overwatch_Object Semana5 = new Overwatch_Object();
        private static bool Initial_Check = false;
        private static bool Attending_OW_Request = false;
        private static bool Objects_Created = false;
        public static bool OW_Request = false;
        private static int Total_Num_weeks = 4;
        private static int assignments = 0;
        private static int filled = 0;
        private static float percentaje;

        public static void Start_Overwatch()
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Overwatch";
                Thread.CurrentThread.Priority = ThreadPriority.Lowest;
            }
            Overwatch_Thread_Handler();
        }

        public static void Overwatch_Thread_Handler()
        {
            while (true)
            {
                if (OW_Request && !Attending_OW_Request && Objects_Created)
                {
                    OW_Request = false;
                    Attending_OW_Request = true;
                    Overwatch_Control();
                }

                if (!Initial_Check)
                {
                    Initial_Check = true;
                    Main_Form.Notify("Initial Check: Overwatch reporting");
                }
                if (!Objects_Created && Main_Form.UI_running)
                {
                    Create_Overwatch_Objs();
                    Objects_Created = true;
                }
                Thread.Sleep(2000);
            }
        }

        public static void Overwatch_Control()
        {
            Insight_Sem Overwatch_Sem = new Insight_Sem();
            for (int i = 1; i <= Total_Num_weeks; i++)
            {
                switch (i)
                {
                    case 1:
                        {
                            Overwatch_Sem = Main_Form.Insight_month.Semana1;
                            Analyze_Sem(Overwatch_Sem);
                            break;
                        }
                    case 2:
                        {
                            Overwatch_Sem = Main_Form.Insight_month.Semana2;
                            Analyze_Sem(Overwatch_Sem);
                            break;
                        }
                    case 3:
                        {
                            Overwatch_Sem = Main_Form.Insight_month.Semana3;
                            Analyze_Sem(Overwatch_Sem);
                            break;
                        }
                    case 4:
                        {
                            Overwatch_Sem = Main_Form.Insight_month.Semana4;
                            Analyze_Sem(Overwatch_Sem);
                            break;
                        }
                    default:
                        {
                            Overwatch_Sem = Main_Form.Insight_month.Semana5;
                            Analyze_Sem(Overwatch_Sem);
                            break;
                        }
                }
            }
            Main_Form.Pending_Overwatch_Refresh = true;
            Attending_OW_Request = false;
        }

        private static void Analyze_Sem(Insight_Sem sem)
        {
            assignments = 0;
            filled = 0;
            percentaje = 0;
            if (sem.Special_VyM_Meeting != Main_Form.Special_Meeting_Type.Conv_type)
            {
                assignments++;
                Check_Null_Object(sem.Presidente_VyM);
                assignments++;
                Check_Null_Object(sem.Discurso_VyM_A);
                assignments++;
                Check_Null_Object(sem.Perlas);
                assignments++;
                Check_Null_Object(sem.Lectura_Biblia_A);
                assignments++;
                Check_Null_Object(sem.SMM1_A);
                assignments++;
                Check_Null_Object(sem.SMM2_A);
                if (sem.SMM3 != null)
                {
                    assignments++;
                    Check_Null_Object(sem.SMM3_A);
                    if (Main_Form.Room_B_enabled)
                    {
                        assignments++;
                        Check_Null_Object(sem.SMM3_B);
                    }
                }
                if (sem.SMM4 != null)
                {
                    assignments++;
                    Check_Null_Object(sem.SMM4_A);
                    if (Main_Form.Room_B_enabled)
                    {
                        assignments++;
                        Check_Null_Object(sem.SMM4_B);
                    }
                }
                if (Main_Form.Room_B_enabled)
                {
                    assignments++;
                    Check_Null_Object(sem.Consejero_Aux);
                    assignments++;
                    Check_Null_Object(sem.Lectura_Biblia_B);
                    if ((sem.SMM1 != null) && (sem.SMM1.Length > 4))
                    {
                        string str = sem.SMM1.Substring(0, 6);
                        if (!(str.Contains("Video")) || !(str.Contains("Seamos")))
                        {
                            assignments++;
                            Check_Null_Object(sem.SMM1_B);
                        }
                    }
                    assignments++;
                    Check_Null_Object(sem.SMM2_B);
                }
                assignments++;
                Check_Null_Object(sem.NVC1_A);
                if (sem.NVC2 != null)
                {
                    assignments++;
                    Check_Null_Object(sem.NVC2_A);
                }
                assignments++;
                Check_Null_Object(sem.Libro_Conductor);
                if (sem.Special_VyM_Meeting != Main_Form.Special_Meeting_Type.Visit_type)
                {
                    assignments++;
                    Check_Null_Object(sem.Libro_Lector);
                }
                assignments++;
                Check_Null_Object(sem.Oracion_End_VyM);
                assignments++;
                Check_Null_Object(sem.Aseo);
                assignments++;
                Check_Null_Object(sem.Vym_Cap);
                assignments++;
                Check_Null_Object(sem.Vym_Der);
                assignments++;
                Check_Null_Object(sem.Vym_Izq);
            }
            if (sem.Special_VyM_Meeting != Main_Form.Special_Meeting_Type.Conv_type)
            {
                assignments++;
                Check_Null_Object(sem.Presidente_RP);
                if (sem.Titulo_Discurso_RP != null)
                {
                    assignments++;
                    Check_Null_Object(sem.Discursante_RP);
                    assignments++;
                    Check_Null_Object(sem.Congregacion_RP);
                }
                assignments++;
                Check_Null_Object(sem.Titulo_Atalaya);
                assignments++;
                Check_Null_Object(sem.Conductor_Atalaya);
                if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
                {

                }
            }
            if (assignments > 0)
            {
                percentaje = (filled * 100) / assignments;
            }
            else
            {
                percentaje = 100;
            }
            if (Main_Form.Overwatch_Information_List.Count > 0)
            {
                Main_Form.Overwatch_Information_List[sem.Num_of_Week - 1].Assignments = assignments;
                Main_Form.Overwatch_Information_List[sem.Num_of_Week - 1].Filled = filled;
                Main_Form.Overwatch_Information_List[sem.Num_of_Week - 1].Percentaje = percentaje.ToString() + "%";
                Main_Form.Overwatch_Information_List[sem.Num_of_Week - 1].Special_VyM_Meeting_Type = sem.Special_VyM_Meeting;
                Main_Form.Overwatch_Information_List[sem.Num_of_Week - 1].Special_RP_Meeting_Type = sem.Special_RP_Meeting;
            }
        }

        private static void Check_Null_Object(string str)
        {
            if (str != null)
            {
                filled++;
            }
        }

        public static void Create_Overwatch_Objs()
        {
            Main_Form.Notify("Create Overwatch Objects");
            Semana1.Name = "Semana 1";
            Semana1.Percentaje = "0%";
            Semana1.Special_VyM_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana1.Special_RP_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana2.Name = "Semana 2";
            Semana2.Percentaje = "0%";
            Semana2.Special_VyM_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana2.Special_RP_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana3.Name = "Semana 3";
            Semana3.Percentaje = "0%";
            Semana3.Special_VyM_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana3.Special_RP_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana4.Name = "Semana 4";
            Semana4.Percentaje = "0%";
            Semana4.Special_VyM_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana4.Special_RP_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana5.Name = "Semana 5";
            Semana5.Percentaje = "0%";
            Semana5.Special_VyM_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Semana5.Special_RP_Meeting_Type = Main_Form.Special_Meeting_Type.Non_status;
            Main_Form.Overwatch_Information_List.Add(Semana1);
            Main_Form.Overwatch_Information_List.Add(Semana2);
            Main_Form.Overwatch_Information_List.Add(Semana3);
            Main_Form.Overwatch_Information_List.Add(Semana4);
            if (Main_Form.week_five_exist)
            {
                Total_Num_weeks = 5;
                Main_Form.Overwatch_Information_List.Add(Semana5);
            }
            Main_Form.Pending_Overwatch_Refresh = true;
        }
    }
}
