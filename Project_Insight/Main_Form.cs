using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Threading;
using System.Runtime.CompilerServices;
using System.IO;
using System.Collections;
using System.Diagnostics;


/*Developed by AGR-Systems Science and Tech Division*/
namespace Project_Insight
{
    public partial class Main_Form : Form
    {
        public enum P
        {
            Executor,
            Oracle,
            DarkTemplar,
            FenixDragoon,
            HunterKiller,
            FenixZealot,
            Selendis,
            Hybrid,
            Artanis
        };
        public enum Male_Type
        {
            Anciano,
            Ministerial,
            Publicador
        };
        public enum Male_State
        {
            Non_status,
            Allowed,
            Blocked
        };
        private Excel.Application objApp;
        private Excel.Workbook objBooks = null;
        private Excel.Sheets objSheets;
        private Excel.Worksheet Sheet_VyM;
        private Excel.Worksheet Sheet_RP;
        private Excel.Worksheet Sheet_AC;
        private Excel.Range range_1;
        private Excel.Range range_2;
        private Excel.Range range_3;
        public static short iterator_stack = 0;
        public static short m_semana = 1;
        public static short presenter_RP = 3;
        public static short presenter_AC = 6;
        public static short Conv_Wk = 0;
        public static short Vst_Wk = 0;
        public static int A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7, H = 8;
        public static int actual_presenter = 10;
        public static int m_dia = 1;
        public static int m_mes = 1;
        public static int m_año = DateTime.Today.Year;
        public static int date_checksum = 0;
        public static int command_iterator = 0;
        public static int loading_delta = 1;
        public static int loading = 0;
        public static int current_tab = 0;
        public static int Generals_Count = 0, Ministerials_Count = 0, Elders_Count = 0, Males_Count = 0;
        public const int DPI = 96;
        public const int Constant = 72;
        public static bool excel_ready = false;
        public static bool busy_trace = false;
        public static bool pending_trace = false;
        public static bool week_five_exist = false;
        public static bool UI_running = false;
        public static bool is_new_instance = false;
        public static bool is_room_B_enabled = false;
        public static bool Write_config_status = false;
        public static bool Save_as_pdf = false;
        public static bool Ac_same_all_week = false;
        public static bool Save_thread_is_running = false;
        public static bool Pending_refresh_status_grids = false;
        public static bool Heavensward_month_in_progress = false;
        public static bool month_found = false;
        public static bool Male_List_filled = false;
        public static bool Autocomplete_aux_status = true;
        private static bool Initial_Check = false;
        private static bool Main_Allowed = false;
        private static bool Covert_Ops = false;
        private static bool Edit_Rule = true;
        public static DayOfWeek VyM_Day;
        public static DayOfWeek RP_Day;
        public static DateTime start_time = new DateTime(DateTime.Today.Year, 1, 1, 7, 00, 00);
        public static DateTime date;
        public static DateTime VyM_horary = new DateTime(DateTime.Today.Year, 1, 1, 7, 00, 00);
        public static DateTime RP_horary = new DateTime(DateTime.Today.Year, 1, 1, 7, 00, 00);
        public static DateTime[,] meetings_days = new DateTime[5, 2];
        private object[,] cellValue_1 = null;
        private object[,] cellValue_2 = null;
        private object[,] cellValue_3 = null;
        public static string File_Path = Application.StartupPath + "\\\\Programs.xlsx";
        public static string Path_DB = Application.StartupPath + "\\\\DB.xlsx";
        public static string Config_Path = "";
        public static string aux_command;
        public static string Conv_Name = "";
        public static string Cong_Name = "";
        public static string Previous_Male_Type = "";
        public static string[] str_stack = new string[50];
        public static string[] Command_history = new string[10];
        public static string[] Months = new string[] { "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre" };
        //public DB_Form DB_Form = new DB_Form();
        public static VyM_Mes VyM_mes = new VyM_Mes();
        public static RP_Mes RP_mes = new RP_Mes();
        public static AC_Mes AC_mes = new AC_Mes();
        public static IDictionary<string, object> Dict_vym = new Dictionary<string, object>();
        public static IDictionary<string, object> Dict_rp = new Dictionary<string, object>();
        public static IDictionary<string, object> Dict_ac = new Dictionary<string, object>();
        StreamReader Reader_config;
        StreamWriter Writer_config;
        public static List<Trace> Info_trace = new List<Trace>();
        public static List<Hw_requested_info> HW_request = new List<Hw_requested_info>();
        public static List<string> Autocomplete_Males_List = new List<string>();
        public static BindingList<Males> Male_List = new BindingList<Males>();
        public static Males Rule_Elders = new Males();
        public static Males Rule_Ministerials = new Males();
        public static Males Rule_Generals = new Males();


        /*----------------System Functions------------------*/
        public Main_Form()
        {
            InitializeComponent();
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Main";
                Thread.CurrentThread.Priority = ThreadPriority.Highest;
            }
            Main_Timer.Enabled = true;
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            Notify("Project Insight 2.0");
            Notify("UI up and ready \n\n ----- Welcome back Hierarch! -----\n");
            Presenter(P.Executor);
            Autocomplete_dictionary();
            txt_Command.Focus();
            Config_Control(true);
            Run_Threads();
        }

        private void Run_Threads()
        {
            Thread DB_thread = new Thread(() => DB_Handler.Start_DataBase());
            DB_thread.Start();

            Thread heavensward = new Thread(() => Heavensward.Start_Heavensward());
            heavensward.Start();

        }

        public static void Set_weeks()
        {
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
            AC_mes.Semana1.Num_of_Week = 1;
            AC_mes.Semana2.Num_of_Week = 2;
            AC_mes.Semana3.Num_of_Week = 3;
            AC_mes.Semana4.Num_of_Week = 4;
            AC_mes.Semana5.Num_of_Week = 5;
        }

        public void Autocomplete_dictionary()
        {
            Dict_vym.Add("ig_01", txt_Lec_Sem);
            Dict_vym.Add("ig_02", txt_Pres);
            Dict_vym.Add("ig_03", txt_ConAux);
            Dict_vym.Add("tb_01", txt_TdlB_1);
            Dict_vym.Add("tb_02", txt_TdlB_A1);
            Dict_vym.Add("tb_03", txt_TdlB_A2);
            Dict_vym.Add("tb_04", txt_TdlB_3);
            Dict_vym.Add("tb_05", txt_TdlB_A3);
            Dict_vym.Add("tb_06", txt_TdlB_B3);
            Dict_vym.Add("sm_11", txt_SMM1);
            Dict_vym.Add("sm_12", txt_SMM_A1);
            Dict_vym.Add("sm_13", txt_SMM_B1);
            Dict_vym.Add("sm_21", txt_SMM2);
            Dict_vym.Add("sm_22", txt_SMM_A2);
            Dict_vym.Add("sm_23", txt_SMM_B2);
            Dict_vym.Add("sm_31", txt_SMM3);
            Dict_vym.Add("sm_32", txt_SMM_A3);
            Dict_vym.Add("sm_33", txt_SMM_B3);
            Dict_vym.Add("sm_41", txt_SMM4);
            Dict_vym.Add("sm_42", txt_SMM_A4);
            Dict_vym.Add("sm_43", txt_SMM_B4);
            Dict_vym.Add("nv_11", txt_NVC1);
            Dict_vym.Add("nv_12", txt_NVC_A1);
            Dict_vym.Add("nv_21", txt_NVC2);
            Dict_vym.Add("nv_22", txt_NVC_A2);
            Dict_vym.Add("nv_31", txt_NVC_A3);
            Dict_vym.Add("nv_40", txt_NVC_A4);
            Dict_vym.Add("nv_50", txt_Ora2VyM);

            Dict_rp.Add("rp_01", txt_PresRP);
            Dict_rp.Add("rp_02", txt_RP_Speech);
            Dict_rp.Add("rp_03", txt_RP_Disc);
            Dict_rp.Add("rp_04", txt_RP_Cong);
            Dict_rp.Add("rp_05", txt_Title_Atly);
            Dict_rp.Add("rp_06", txt_Con_Atly);
            Dict_rp.Add("rp_07", txt_Lect_Atly);
            Dict_rp.Add("rp_08", txt_OraRP);
            Dict_rp.Add("rp_09", txt_Sal_Disc);
            Dict_rp.Add("rp_10", txt_Sal_Title);
            Dict_rp.Add("rp_11", txt_Sal_Cong);

            Dict_ac.Add("ac_11", txt_Aseo_1);
            Dict_ac.Add("ac_12", txt_Cap_L_1);
            Dict_ac.Add("ac_13", txt_AC1_L_1);
            Dict_ac.Add("ac_14", txt_AC2_L_1);
            Dict_ac.Add("ac_16", txt_Cap_S_1);
            Dict_ac.Add("ac_17", txt_AC1_S_1);
            Dict_ac.Add("ac_18", txt_AC2_S_1);
            Dict_ac.Add("ac_21", txt_Aseo_2);
            Dict_ac.Add("ac_22", txt_Cap_L_2);
            Dict_ac.Add("ac_23", txt_AC1_L_2);
            Dict_ac.Add("ac_24", txt_AC2_L_2);
            Dict_ac.Add("ac_26", txt_Cap_S_2);
            Dict_ac.Add("ac_27", txt_AC1_S_2);
            Dict_ac.Add("ac_28", txt_AC2_S_2);
            Dict_ac.Add("ac_31", txt_Aseo_3);
            Dict_ac.Add("ac_32", txt_Cap_L_3);
            Dict_ac.Add("ac_33", txt_AC1_L_3);
            Dict_ac.Add("ac_34", txt_AC2_L_3);
            Dict_ac.Add("ac_36", txt_Cap_S_3);
            Dict_ac.Add("ac_37", txt_AC1_S_3);
            Dict_ac.Add("ac_38", txt_AC2_S_3);
            Dict_ac.Add("ac_41", txt_Aseo_4);
            Dict_ac.Add("ac_42", txt_Cap_L_4);
            Dict_ac.Add("ac_43", txt_AC1_L_4);
            Dict_ac.Add("ac_44", txt_AC2_L_4);
            Dict_ac.Add("ac_46", txt_Cap_S_4);
            Dict_ac.Add("ac_47", txt_AC1_S_4);
            Dict_ac.Add("ac_48", txt_AC2_S_4);
            Dict_ac.Add("ac_51", txt_Aseo_5);
            Dict_ac.Add("ac_52", txt_Cap_L_5);
            Dict_ac.Add("ac_53", txt_AC1_L_5);
            Dict_ac.Add("ac_54", txt_AC2_L_5);
            Dict_ac.Add("ac_56", txt_Cap_S_5);
            Dict_ac.Add("ac_57", txt_AC1_S_5);
            Dict_ac.Add("ac_58", txt_AC2_S_5);
        }

        public void Main_Form_FormClosed(object sender, FormClosedEventArgs e)
        {

            DB_Handler.Close_DB();
            if (excel_ready)
            {
                excel_ready = false;
                objBooks.Close(0);
                objApp.Quit();
                Marshal.ReleaseComObject(Sheet_VyM);
                Marshal.ReleaseComObject(Sheet_RP);
                Marshal.ReleaseComObject(Sheet_AC);
                Marshal.ReleaseComObject(objBooks);
                Marshal.ReleaseComObject(objApp);
            }
            Application.Exit();
        }

        /*--------------------------------------- Traces and UI functions ---------------------------------------*/
        public async void Presenter(P ID_Presenter)
        {
            if (!Covert_Ops)
            {
                if (actual_presenter != (int)ID_Presenter)
                {
                    actual_presenter = (int)ID_Presenter;
                    picPresenter.Image = Project_Insight.Properties.Resources.Noise;
                    await Task.Delay(300);
                    switch (ID_Presenter)
                    {
                        case P.Executor:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.Executor;
                                break;
                            }
                        case P.FenixZealot:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.FenixZealot;
                                break;
                            }
                        case P.FenixDragoon:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.FenixDragoon;
                                break;
                            }
                        case P.Selendis:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.Selendis;
                                break;
                            }
                        case P.Oracle:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.Oracle;
                                break;
                            }
                        case P.DarkTemplar:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.DarkTemplar;
                                break;
                            }
                        case P.HunterKiller:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.HunterKiller;
                                break;
                            }
                        case P.Hybrid:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.Hybrid;
                                break;
                            }
                        case P.Artanis:
                            {
                                picPresenter.Image = Project_Insight.Properties.Resources.Artanis;
                                break;
                            }
                    }
                }
            }
            else
            {
                picPresenter.Image = Project_Insight.Properties.Resources.Noise;
            }
        }

        public static void Notify(string data)
        {
            if (!Heavensward_month_in_progress)
            {
                Trace trace = new Trace
                {
                    Current_Thread = Thread.CurrentThread.Name,
                    Info = data,
                    Type = 1
                };
                Info_trace.Add(trace);
            }
        }

        public static void Warn(string data)
        {
            Trace trace = new Trace
            {
                Current_Thread = Thread.CurrentThread.Name,
                Info = data,
                Type = 2
            };
            Info_trace.Add(trace);
        }

        public async void Process_Trace(Trace trace)
        {
            if (!busy_trace)
            {
                try
                {
                    busy_trace = true;
                    var array = ("[" + trace.Current_Thread + "]: " + trace.Info).ToCharArray();
                    if (trace.Type == 1)
                    {
                        Log_txtBx.SelectionColor = Color.White;
                    }
                    else
                    {
                        Log_txtBx.SelectionColor = Color.Red;
                    }
                    for (int i = 0; i < array.Length; i++)
                    {
                            Log_txtBx.AppendText(array[i].ToString());
                            await Task.Delay(1);
                            if (array[i] == '\n')
                            {
                                Log_txtBx.ScrollToCaret();
                            }
                    }
                    Log_txtBx.AppendText("\n");
                    Log_txtBx.SelectionStart = Log_txtBx.Text.Length;
                    Log_txtBx.ScrollToCaret();
                    Info_trace.RemoveAt(0);
                    busy_trace = false;
                }
                catch { }
            }
        }

        private void Main_Timer_Tick(object sender, EventArgs e)
        {
            if ((Info_trace.Count > 0) && !busy_trace)
            {
                Process_Trace(Info_trace[0]);
            }
            if (Save_thread_is_running)
            {
                LoadingBarHandler();
            }
            if (HW_request.Count > 0)
            {
                Heavensward_request_handler();
            }
            if (Pending_refresh_status_grids && Male_List_filled)
            {
                Refresh_Males_Grid();
            }
            if (!Initial_Check)
            {
                Initial_Check = true;
                Initial_Main_Check();
            }
        }

        public class Trace
        {
            public string Current_Thread;
            public string Info;
            public short  Type;
        }

        private static void Initial_Main_Check()
        {
            if (File.Exists(File_Path))
            {
                Notify("Initial Check: Main File exist");
                Main_Allowed = true;
            }
            else
            {
                Warn("Initial Check: Main file missing");
                Warn("Disabling Main features");
                Main_Allowed = false;
            }
        }
        
        /*--------------------------------------- Command functions ---------------------------------------*/

        private void Process_txt_Command(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txt_Command.Text != "")
                {
                    string Str = txt_Command.Text;
                    string cmd = Str;
                    string sup = "";
                    int index = 0;
                    if (cmd.Length > 4)
                    {
                        index = cmd.IndexOf(" ");
                        if (index >= 0)
                        {
                            cmd = cmd.Substring(0, index);
                            sup = Str.Substring(index + 1);
                        }
                    }
                    cmd = cmd.ToLower();
                    Str = cmd + " " + sup;
                    Save_command(txt_Command.Text);
                    command_iterator = 0;
                    if (!Write_config_status)
                    {
                        switch (cmd)
                        {
                            case "new":
                                {
                                    if (Main_Allowed)
                                    {
                                        sup = sup.ToLower();
                                        bool month_found = false;
                                        for (int i = 0; i <= Months.Length - 1; i++)
                                        {
                                            if (Months[i].Contains(sup))
                                            {
                                                m_mes = i + 1;
                                                month_found = true;
                                                Notify("New file for month " + Months[i]);
                                                Set_weeks();
                                                DB_Handler.DB_Requests_List.Add(DB_Handler.DB_Request.read);
                                                break;
                                            }
                                        }
                                        if (month_found)
                                        {
                                            New_Instance();
                                        }
                                        else
                                        {
                                            Warn("Invalid Parameters");
                                        }
                                    }
                                    else
                                    {
                                        Warn("Main Function Disabled!");
                                    }
                                    break;
                                }
                            case "open":
                                {
                                    Known_Instance();
                                    Set_weeks();
                                    DB_Handler.DB_Requests_List.Add(DB_Handler.DB_Request.read);
                                    break;
                                }
                            case "save":
                                {
                                    if (UI_running)
                                    {
                                        sup = sup.ToLower();
                                        if (sup.Contains("vym"))
                                        {
                                            Thread Save_thread = new Thread(() => Process_save(1));
                                            Save_thread.Start();
                                        }
                                        else if (sup.Contains("rp"))
                                        {
                                            Thread Save_thread = new Thread(() => Process_save(2));
                                            Save_thread.Start();
                                        }
                                        else if (sup.Contains("ac"))
                                        {
                                            Thread Save_thread = new Thread(() => Process_save(3));
                                            Save_thread.Start();
                                        }
                                        else if (sup.Contains("all"))
                                        {
                                            Thread Save_thread = new Thread(() => Process_save(4));
                                            Save_thread.Start();
                                        }
                                        else if (sup.Contains("db"))
                                        {
                                            DB_Handler.DB_Requests_List.Add(DB_Handler.DB_Request.write);
                                        }
                                        else
                                        {
                                            Warn("Unable to get selected save");
                                        }

                                        if (sup.Contains("pdf"))
                                        {
                                            Save_as_pdf = true;
                                        }
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                                /*Tab Section*/
                            case "vym":
                                {
                                    if (UI_running)
                                    {
                                        tab_Control.SelectedIndex = 0;
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "rp":
                                {
                                    if (UI_running)
                                    {
                                        tab_Control.SelectedIndex = 1;
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "ac":
                                {
                                    if (UI_running)
                                    {
                                        tab_Control.SelectedIndex = 2;
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "stat":
                                {
                                    if (UI_running)
                                    {
                                        tab_Control.SelectedIndex = 3;
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "db":
                                {
                                    if (UI_running)
                                    {
                                        tab_Control.SelectedIndex = 4;
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "week":
                                {
                                    if (UI_running)
                                    {
                                        if (int.TryParse(sup, out int wk))
                                        {
                                            if ((wk != m_semana) && (wk > 0))
                                            {
                                                if ((wk == 5) && (!week_five_exist))
                                                {
                                                    Warn("Selected month [" + meetings_days[0,0].ToString("MMMM") + "] doesn't have 5 weeks");
                                                    Notify("Current week is [" + m_semana.ToString() + "]");
                                                }
                                                else
                                                {
                                                    Pre_save_info();
                                                    m_semana = (short)wk;
                                                    Week_Handler();
                                                }
                                            }
                                        }
                                        else
                                        {
                                            Warn("Unable to get selected week");
                                        }
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "conv":
                                {
                                    if (UI_running)
                                    {
                                        //sup = sup.ToLower();
                                        if (sup.Contains("false"))
                                        {
                                            Conv_Wk = 0;
                                        }
                                        else
                                        {
                                            Conv_Name = "Asamblea de los Testigos de Jehová: " + sup;
                                            Conv_Wk = m_semana;
                                        }
                                        Notify("Current week [" + m_semana.ToString() + "] setting as Convention [" + (Conv_Wk > 0 ? "True" : "False") + "]");
                                        Alert_Label_VyM.Text = "Semana de Asamblea!";
                                        Alert_Label_VyM.Visible = true;
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "vst":
                                {
                                    if (UI_running)
                                    {
                                        sup = sup.ToLower();
                                        if (sup.Contains("true"))
                                        {
                                            Vst_Wk = m_semana;
                                            Notify("Marked current week [" + m_semana + "] as Visit");
                                            Alert_Label_VyM.Text = "Semana de la Visita del Superintendente de Circuito";
                                            Alert_Label_VyM.Visible = true;
                                            txt_NVC3.Enabled = true;
                                        }
                                        else if (sup.Contains("false") && Vst_Wk > 0)
                                        {
                                            Notify("Removed mark week [" + Vst_Wk + "] as Visit");
                                            Vst_Wk = 0;
                                        }
                                        else
                                        {
                                            Warn("Command not recognized or not supported");
                                        }
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "afil":
                                {
                                    if (UI_running)
                                    {
                                        AutoFill_Handler();
                                    }
                                    break;
                                }
                            case "cfg":
                                {
                                    sup = sup.ToLower();
                                    if (sup == "read")
                                    {
                                        Notify("Reading config status\n");
                                        Notify("Congregation Name: " + Cong_Name);
                                        Notify("Room B : " + (is_room_B_enabled ? "Enabled" : "Disabled"));
                                        Notify("VyM horary : " + VyM_Day + " " + VyM_horary.ToString("HH:mm"));
                                        Notify("RP horary : " + RP_Day + " " + RP_horary.ToString("HH:mm"));
                                        Notify("AC all week : " + (Ac_same_all_week ? "Enabled": "Disabled") + "\n");
                                    }
                                    else if (sup == "write")
                                    {
                                        Notify("Entering Write_Config mode");
                                        Write_config_status = true;
                                    }
                                    else
                                    {
                                        Warn("Invalid command");
                                    }
                                    break;
                                }
                            case "help":
                                {
                                    Notify("TBD");
                                    //Notify(Project_Insight.Properties.Resources.Commands_Helper);
                                    break;
                                }
                            case "hw":
                                {
                                    if (UI_running)
                                    {
                                        sup = sup.ToLower();
                                        if (sup.Contains("all"))
                                        {
                                            Notify("Extracting all month info from Heavensward");
                                            Heavensward_All_Info();
                                        }
                                        else
                                        {
                                            if (current_tab == 0)
                                            {
                                                Heavensward.HW_Bridge(meetings_days[m_semana - 1, 0], current_tab, m_semana);
                                            }
                                            else if (current_tab == 1)
                                            {
                                                Heavensward.HW_Bridge(meetings_days[m_semana - 1, 1], current_tab, m_semana);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                            case "cov":
                                {
                                    if (!UI_running)
                                    {
                                        Notify("Covert Operations");
                                        Covert_Ops = true;
                                    }
                                    break;
                                }
                            case "test":
                                {
                                    Heavensward.HW_Oracle_Request request = new Heavensward.HW_Oracle_Request();
                                    request.hw_oracle_vym = VyM_mes.Semana1;
                                    Heavensward.HW_Oracle_Requests_List.Add(request);
                                    break;
                                }
                            default:
                                {
                                    if (UI_running)
                                    {
                                        switch (current_tab)
                                        {
                                            case 0:
                                                {

                                                    if (Dict_vym.ContainsKey(cmd))
                                                    {
                                                        TextBox txt = (TextBox)Dict_vym[cmd];
                                                        txt.Text = sup;
                                                        txt.BackColor = Color.White;
                                                        Notify("Writing in [" + txt.Name + "]");
                                                    }
                                                    else
                                                    {
                                                        Warn("Unrecognized command: " + cmd);
                                                    }
                                                    break;
                                                }
                                            case 1:
                                                {
                                                    if (Dict_rp.ContainsKey(cmd))
                                                    {
                                                        TextBox txt = (TextBox)Dict_rp[cmd];
                                                        txt.Text = sup;
                                                        txt.BackColor = Color.White;
                                                        Notify("Writing in [" + txt.Name + "]");
                                                        if (cmd.Equals("rp_02") || cmd.Equals("rp_10"))
                                                        {
                                                            txt.Text = Get_Speech(sup);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        Warn("Unrecognized command: " + cmd);
                                                    }
                                                    break;
                                                }
                                            case 2:
                                                {
                                                    if (Dict_ac.ContainsKey(cmd))
                                                    {
                                                        TextBox txt = (TextBox)Dict_ac[cmd];
                                                        txt.Text = sup;
                                                        txt.BackColor = Color.White;
                                                        Notify("Writing in [" + txt.Name + "]");
                                                    }
                                                    else
                                                    {
                                                        Warn("Unrecognized command: " + cmd);
                                                    }
                                                    break;
                                                }
                                        }
                                    }
                                    else
                                    {
                                        Warn("Need to create a new instance or open an existing program");
                                    }
                                    break;
                                }
                        }
                    }
                    else
                    {
                        /*Switch to attend writing request to cfg file*/
                        switch (cmd)
                        {
                            case "cong":
                                {
                                    Cong_Name = sup;
                                    Notify("Seeting Congregation Name to: " + Cong_Name);
                                    break;
                                }
                            case "roomb":
                                {
                                    sup = sup.ToLower();
                                    if (sup.Contains("false"))
                                    {
                                        is_room_B_enabled = false;
                                    }
                                    else if (sup.Contains("true"))
                                    {
                                        is_room_B_enabled = true;
                                    }
                                    else
                                    {
                                        Warn("Wrong condition, only boolean status");
                                    }
                                    Notify("Room B available status: " + (is_room_B_enabled ? "true" : "false"));
                                    break;
                                }
                            case "vymh":
                                {
                                    sup = sup.ToLower();
                                    if (sup.Contains(":"))
                                    {
                                        int dot = sup.IndexOf(":");
                                        string aux_sup = "";
                                        aux_sup = sup.Remove(dot);
                                        if (!int.TryParse(aux_sup, out int Hr))
                                        {
                                            Warn("Unable to get hour");
                                        }
                                        aux_sup = sup.Remove(0, 3);
                                        if (!int.TryParse(aux_sup, out int Mn))
                                        {
                                            Warn("Unable to get minutes");
                                        }
                                        if (Hr > 0 && Mn >= 0)
                                        {
                                            VyM_horary = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, Hr, Mn, 0);
                                            Notify("Seeting Schedule for VyM Meeting at: " + VyM_horary.ToString("HH:mm"));
                                        }
                                    }
                                    break;
                                }
                            case "rph":
                                {
                                    sup = sup.ToLower();
                                    if (sup.Contains(":"))
                                    {
                                        int dot = sup.IndexOf(":");
                                        string aux_sup = "";
                                        aux_sup = sup.Remove(dot);
                                        if (!int.TryParse(aux_sup, out int Hr))
                                        {
                                            Warn("Unable to get hour");
                                        }
                                        aux_sup = sup.Remove(0, 3);
                                        if (!int.TryParse(aux_sup, out int Mn))
                                        {
                                            Warn("Unable to get minutes");
                                        }
                                        if (Hr > 0 && Mn >= 0)
                                        {
                                            RP_horary = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, Hr, Mn, 0);
                                            Notify("Seeting Schedule for RP Meeting at: " + VyM_horary.ToString("HH:mm"));
                                        }
                                    }
                                    break;
                                }
                            case "vymd":
                                {
                                    sup = sup.ToLower();
                                    VyM_Day = GetDayOfWeek(sup);
                                    Notify("VyM day meeting set in: " + VyM_Day.ToString());
                                    break;
                                }
                            case "rpd":
                                {
                                    sup = sup.ToLower();
                                    RP_Day = GetDayOfWeek(sup);
                                    Notify("RP day meeting set in: " + RP_Day.ToString());
                                    break;
                                }
                            case "ac":
                                {
                                    sup = sup.ToLower();
                                    if (sup.Contains("false"))
                                    {
                                        Ac_same_all_week = false;
                                    }
                                    else if (sup.Contains("true"))
                                    {
                                        Ac_same_all_week = true;
                                    }
                                    else
                                    {
                                        Warn("Wrong condition, only boolean status");
                                    }
                                    Notify("AC same all week status: " + (Ac_same_all_week ? "true" : "false"));
                                    break;
                                }
                            case "exit":
                                {
                                    Write_config_status = false;
                                    Notify("Exiting Write_Config mode");
                                    Config_Control(false);
                                    Get_Meetings();
                                    Week_Handler();
                                    break;
                                }
                            default:
                                {
                                    Warn("Unexpected command while Write_Config_Status is true");
                                    break;
                                }
                        }
                    }
                    txt_Command.Text = "";
                    txt_Command.Focus();
                }
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (command_iterator < Command_history.Length - 1)
                {
                    command_iterator++;
                    if (Command_history[command_iterator] != null)
                    {
                        txt_Command.Text = Command_history[command_iterator];
                        txt_Command.SelectionStart = txt_Command.Text.Length;
                    }
                    else
                    {
                        command_iterator--;
                    }
                }
            }
            else if (e.KeyCode == Keys.Down)
            {
                if (command_iterator > 0)
                {
                    command_iterator--;
                    if (Command_history[command_iterator] != null)
                    {
                        txt_Command.Text = Command_history[command_iterator];
                        txt_Command.SelectionStart = txt_Command.Text.Length;
                    }
                    else if (command_iterator == 0)
                    {
                        txt_Command.Text = "";
                    }
                }
            }
        }

        private void Txt_Command_TextChanged(object sender, EventArgs e)
        {
            string Str = txt_Command.Text.ToLower();
            string cmd = Str;
            int index;
            if (Str.Length > 4)
            {
                index = cmd.IndexOf(" ");
                if (index >= 0)
                {
                    cmd = cmd.Substring(0, index);
                }
            }
            switch (current_tab)
            {
                case 0:
                    {
                        if (Dict_vym.ContainsKey(cmd))
                        {
                            Change_Presenter(cmd);
                            TextBox txt = (TextBox)Dict_vym[cmd];
                            txt.BackColor = Color.OrangeRed;
                            if ((cmd != aux_command) && (aux_command != null) && Dict_vym.ContainsKey(aux_command))
                            {
                                TextBox txt_aux = (TextBox)Dict_vym[aux_command];
                                txt_aux.BackColor = Color.White;
                            }
                            aux_command = cmd;
                        }
                        else if (aux_command != null)
                        {
                            if (Dict_vym.ContainsKey(aux_command))
                            {
                                TextBox txt = (TextBox)Dict_vym[aux_command];
                                txt.BackColor = Color.White;
                            }
                        }
                        break;
                    }
                case 1:
                    {
                        if (Dict_rp.ContainsKey(cmd))
                        {
                            Change_Presenter(cmd);
                            TextBox txt = (TextBox)Dict_rp[cmd];
                            txt.BackColor = Color.OrangeRed;
                            if ((cmd != aux_command) && (aux_command != null) && Dict_rp.ContainsKey(aux_command))
                            {
                                TextBox txt_aux = (TextBox)Dict_rp[aux_command];
                                txt_aux.BackColor = Color.White;
                            }
                            aux_command = cmd;
                        }
                        else if (aux_command != null)
                        {
                            if (Dict_rp.ContainsKey(aux_command))
                            {
                                TextBox txt = (TextBox)Dict_rp[aux_command];
                                txt.BackColor = Color.White;
                            }
                        }
                        break;
                    }
                case 2:
                    {
                        if (Dict_ac.ContainsKey(cmd))
                        {
                            Change_Presenter(cmd);
                            TextBox txt = (TextBox)Dict_ac[cmd];
                            txt.BackColor = Color.OrangeRed;
                            if ((cmd != aux_command) && (aux_command != null) && Dict_ac.ContainsKey(aux_command))
                            {
                                TextBox txt_aux = (TextBox)Dict_ac[aux_command];
                                txt_aux.BackColor = Color.White;
                            }
                            aux_command = cmd;
                        }
                        else if (aux_command != null)
                        {
                            if (Dict_ac.ContainsKey(aux_command))
                            {
                                TextBox txt = (TextBox)Dict_ac[aux_command];
                                txt.BackColor = Color.White;
                            }
                        }
                        break;
                    }
            }
        }

        public void Change_Presenter(string cmd)
        {
            if (cmd.Contains("ig") || cmd.Contains("tb"))
            {
                Presenter(P.Executor);
            }
            else if (cmd.Contains("sm"))
            {
                Presenter(P.Oracle);
            }
            else if (cmd.Contains("nv"))
            {
                Presenter(P.DarkTemplar);
            }
        }

        public void Save_command(string cmd)
        {
            for (int i = Command_history.Length - 1; i >= 2; i--)
            {
                Command_history[i] = Command_history[i - 1];
            }
            Command_history[1] = cmd;
        }

        /*--------------------------------------- Excel file Control ---------------------------------------*/

        /*New excel file*/
        public void New_Instance()
        {
            is_new_instance = true;
            Get_Meetings();
            Week_Handler();
            var autocomplete = new AutoCompleteStringCollection();
            autocomplete.AddRange(Dict_vym.Keys.ToArray());
            txt_Command.AutoCompleteCustomSource = autocomplete;
            UI_running = true;
            Time_Handler();
            Notify("Project Insight Ready Executor Nobelm!");
            tab_Control.Enabled = true;
        }

        /*Open config file*/
        public void Config_Control(bool read)
        {
            if (read)
            {
                Config_Path = Application.StartupPath + "\\\\Project_Insight_Config.txt";
                int len = File.ReadAllLines(Config_Path).Length;
                string[] data = new string[len];
                Reader_config = new StreamReader(Config_Path);

                for (int i = 0; i < len; i++)
                {
                    data[i] = Reader_config.ReadLine();
                }
                Reader_config.Close();
                Cong_Name = data[0];
                bool.TryParse(data[1], out is_room_B_enabled);
                if (is_room_B_enabled)
                {
                    txt_SMM_B1.Visible = true;
                    txt_SMM_B2.Visible = true;
                    txt_SMM_B3.Visible = true;
                    txt_SMM_B4.Visible = true;
                }
                else
                {
                    txt_SMM_B1.Visible = false;
                    txt_SMM_B2.Visible = false;
                    txt_SMM_B3.Visible = false;
                    txt_SMM_B4.Visible = false;
                }
                VyM_horary = Convert.ToDateTime(data[2]);
                RP_horary = Convert.ToDateTime(data[3]);
                VyM_Day = GetDayOfWeek(data[4]);
                RP_Day = GetDayOfWeek(data[5]);
                bool.TryParse(data[6], out Ac_same_all_week);
                Rule_Elders.Name        = data[7];
                Rule_Elders.male_type   = Male_Type.Anciano;
                Rule_Elders.Atalaya     = data[8];
                Rule_Elders.Capitan     = data[9];
                Rule_Elders.Acomodador  = data[10];
                Rule_Elders.Lector      = data[11];
                Rule_Elders.Pres_RP     = data[12];
                Rule_Elders.Oracion     = data[13];
                Rule_Ministerials.Name  = data[14];
                Rule_Ministerials.male_type  = Male_Type.Ministerial;
                Rule_Ministerials.Atalaya    = data[15];
                Rule_Ministerials.Capitan    = data[16];
                Rule_Ministerials.Acomodador = data[17];
                Rule_Ministerials.Lector     = data[18];
                Rule_Ministerials.Pres_RP    = data[19];
                Rule_Ministerials.Oracion    = data[20];
                Rule_Generals.Name       = data[21];
                Rule_Generals.male_type  = Male_Type.Publicador;
                Rule_Generals.Atalaya    = data[22];
                Rule_Generals.Capitan    = data[23];
                Rule_Generals.Acomodador = data[24];
                Rule_Generals.Lector     = data[25];
                Rule_Generals.Pres_RP    = data[26];
                Rule_Generals.Oracion    = data[27];
            }
            else
            {
                Writer_config = new StreamWriter(Config_Path);
                Writer_config.WriteLine(Cong_Name);
                Writer_config.WriteLine(is_room_B_enabled.ToString());
                Writer_config.WriteLine(VyM_horary.ToString("HH:mm"));
                Writer_config.WriteLine(RP_horary.ToString("HH:mm"));
                Writer_config.WriteLine(VyM_Day);
                Writer_config.WriteLine(RP_Day);
                Writer_config.WriteLine(Ac_same_all_week.ToString());
                Writer_config.WriteLine(Rule_Elders.Name);

                Writer_config.Close();
                Config_Control(true);
            }
        }

        public DayOfWeek GetDayOfWeek(string Str)
        {
            string[] Days = new string[] { "sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday" };
            DayOfWeek day = DayOfWeek.Sunday;
            Str = Str.ToLower();
            bool day_found = false;
            for (int i = 0; i < 7; i++)
            {
                if (Str.Contains(Days[i]))
                {
                    day = (DayOfWeek)i;
                    day_found = true;
                    break;
                }
            }
            if (!day_found)
            {
                Warn("Unable to get day, default day as [" + day.ToString() + "]");
            }
            return day;
        }

        /*Open an existing excel program*/
        public void Known_Instance()
        {
            OpenFileDialog openExcel = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls",
                FileName = "",
                Title = "Load Excel File"
            }; 
            if (DialogResult.OK == openExcel.ShowDialog())
            {
                if (null != openExcel.FileName)
                {
                    File_Path = openExcel.FileName;
                    Opening_Excel(File_Path);
                    VyM_Handler(false);
                    RP_Handler(false);
                    AC_Handler(false);
                    Get_Meetings();
                    Week_Handler();
                    if (excel_ready)
                    {
                        excel_ready = false;
                        objBooks.Close(0);
                        objApp.Quit();
                    }
                    UI_running = true;
                    Main_Allowed = true;
                }
            }
            else
            {
                Warn("File not loaded");
            }

            is_new_instance = false;
            tab_Control.Enabled = true;
            var autocomplete = new AutoCompleteStringCollection();
            autocomplete.AddRange(Dict_vym.Keys.ToArray());
            txt_Command.AutoCompleteCustomSource = autocomplete;
            Time_Handler();
        }

        public void Opening_Excel(string path)
        {
            objApp = new Excel.Application();
            objBooks = objApp.Workbooks.Open(path, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);

            objSheets = objBooks.Worksheets;
            Sheet_VyM = (Excel.Worksheet)objSheets.get_Item(1);
            range_1 = Sheet_VyM.get_Range("A1", "H137");
            cellValue_1 = (object[,])range_1.get_Value();
            excel_ready = false;

            if ((cellValue_1[53, 1] != null) && (cellValue_1[53, 1].ToString() == "S-140 AGR-Technologies"))
            {
                Notify("File decoded correctly");

                Sheet_RP = (Excel.Worksheet)objSheets.get_Item(2);
                range_2 = Sheet_RP.get_Range("A1", "H70");
                cellValue_2 = (object[,])range_2.get_Value();

                Sheet_AC = (Excel.Worksheet)objSheets.get_Item(3);
                range_3 = Sheet_AC.get_Range("A1", "H70");
                cellValue_3 = (object[,])range_3.get_Value();

                Notify("Opening excel Main file");
                excel_ready = true;
            }
            else
            {
                Warn("Invalid file");
            }
        }

        /*Set font size to (x min.)*/
        public void Set_Font(Excel.Range cell)
        {
            string Str = cell.Text;
            int index = Str.IndexOf("mins");
            if (index > 4)
            {
                cell.Characters[index - 3, Str.Length].Font.Size = 8;
            }
        }

        /*Add DateTime of the month meetings in the array*/
        public void Get_Meetings()
        {
            Notify("Getting meetings for month [" + Months[m_mes - 1].ToString() + "]");
            int week = -1;
            bool week_found = false;
            DateTime day = new DateTime(m_año, m_mes, 1);
            do
            {
                if (day.DayOfWeek == DayOfWeek.Monday)
                {
                    if (day.Month != m_mes)
                    {
                        break;
                    }
                    week_found = true;
                }
                if (week_found && (day.DayOfWeek == VyM_Day))
                {
                    week++;
                    meetings_days[week, 0] = day;
                }
                else if (week_found && (day.DayOfWeek == RP_Day))
                {
                    meetings_days[week, 1] = day;
                    week_found = false;
                }
                day = day.AddDays(1);

            } while (true);
            if (week == 4)
            {
                week_five_exist = true;
            }
            //ToDo save all dates in objects!
            VyM_mes.Semana1.Fecha = meetings_days[0, 0].ToString("dddd, dd MMMM");
            VyM_mes.Semana2.Fecha = meetings_days[1, 0].ToString("dddd, dd MMMM");
            VyM_mes.Semana3.Fecha = meetings_days[2, 0].ToString("dddd, dd MMMM");
            VyM_mes.Semana4.Fecha = meetings_days[3, 0].ToString("dddd, dd MMMM");
            if (week_five_exist)
            {
                VyM_mes.Semana5.Fecha = meetings_days[4, 0].ToString("dddd, dd MMMM");
                RP_mes.Semana5.Fecha  = meetings_days[4, 1].ToString("dddd, dd MMMM");
                AC_grpbx_wk5.Visible = true;
            }
            else
            {
                AC_grpbx_wk5.Visible = false;
            }
            RP_mes.Semana1.Fecha = meetings_days[0, 1].ToString("dddd, dd MMMM");
            RP_mes.Semana2.Fecha = meetings_days[1, 1].ToString("dddd, dd MMMM");
            RP_mes.Semana3.Fecha = meetings_days[2, 1].ToString("dddd, dd MMMM");
            RP_mes.Semana4.Fecha = meetings_days[3, 1].ToString("dddd, dd MMMM");
        }

        public void Process_save(int save)
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "SaveThread";
                //Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            }
            Save_thread_is_running = true;
            Notify("Executing Save Thread");
            string FileName = meetings_days[0, 0].ToString("MMMM");
            loading = 1;
            Opening_Excel(File_Path);
            loading += 4;
            Pre_save_info();
            if (save == 4)
            {
                loading_delta = 3;
            }
            else
            {
                loading_delta = 1;
            }
            loading += 5;
            if ((save == 1) || (save == 4))
            {
                VyM_Handler(true);
                FileName += "_VyM";
            }
            if ((save == 2) || (save == 4))
            {
                RP_Handler(true);
                FileName += "_RP";
            }
            if ((save == 3) || (save == 4))
            {
                AC_Handler(true);
                FileName += "_AC";
            }
            Notify("FileName: " + FileName);
            loading = 80;
            if (is_new_instance)
            {
                string createfolder = "c:\\Project_Insight";
                System.IO.Directory.CreateDirectory(createfolder);
                objApp.DisplayAlerts = false;
                objBooks.SaveAs(createfolder + "\\" + FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

                //objBooks.SaveAs(createfolder + "\\" + FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                File_Path = createfolder + "\\" + FileName;
                if (Save_as_pdf)
                {
                    objBooks.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, File_Path);
                    Notify("Saving as Pdf");
                    Save_as_pdf = false;
                }
                Notify("Saved path: " + File_Path);
            }
            else
            {
                objBooks.Save();
            }
            loading = 90;
            if (excel_ready)
            {
                excel_ready = false;
                objBooks.Close(0);
                objApp.Quit();
            }
            DB_Handler.DB_Requests_List.Add(DB_Handler.DB_Request.write);
            Notify("Save successful!");
            loading = 100;
        }

        public async void LoadingBarHandler()
        {
            if (!LoadingBar.Visible)
            {
                LoadingBar.Visible = true;
                txt_Command.Enabled = false;
                LoadingBar.Value = 0;
            }
            if (LoadingBar.Value != loading)
            {
                for (int i = LoadingBar.Value; i < loading; i++)
                {
                    LoadingBar.PerformStep();
                    await Task.Delay(5);
                }
            }
            if (100 == loading)
            {
                LoadingBar.Visible = false;
                txt_Command.Enabled = true;
                txt_Command.Focus();
                Save_thread_is_running = false;
            }
        }

        /*--------------------------------------- Meeting handlers ---------------------------------------*/

        public void VyM_Handler(bool save)
        {
            if (save)
            {
                VyM_Save_Week(VyM_mes.Semana1);
                loading += (15/loading_delta);
                VyM_Save_Week(VyM_mes.Semana2);
                loading += (15/loading_delta);
                VyM_Save_Week(VyM_mes.Semana3);
                loading += (15 / loading_delta);
                VyM_Save_Week(VyM_mes.Semana4);
                loading += (15 / loading_delta);
                if (week_five_exist)
                {
                    VyM_Save_Week(VyM_mes.Semana5);
                }
                loading += (15 / loading_delta);
            }
            else
            {
                for (short sm = 1; sm <= 5; sm++) //cycle to read all 5 weeks
                {
                    VyM_Read_Week(sm);
                }
            }
        }

        public void VyM_Save_Week(VyM_Sem sem)
        {
            short num_sem = sem.Num_of_Week;
            short primary_cell = Get_vym_cell(num_sem);
            short increment_smm = 0;
            bool increment_nvc = false;
            Sheet_VyM.PageSetup.LeftHeader = "&16&B" + Cong_Name;
            if (num_sem == Vst_Wk)
            {
                Excel.Range range;
                Sheet_VyM.Cells[primary_cell, A] = "Visita del Superintendente de Circuito";
                range = Sheet_VyM.get_Range("A" + primary_cell.ToString());
                range.Interior.Color = Color.Orange;
            }
            if (((num_sem + 10) % 2) != 0)
            {
                CultureInfo spanish = new CultureInfo("es-MX");
                Sheet_VyM.Cells[primary_cell, G] = meetings_days[(num_sem-1), 0].ToString("dddd", spanish) + " " + VyM_horary.ToString("hh:mm tt"); 
            }
            Excel.Range aux_range;
            primary_cell++;
            Sheet_VyM.Cells[primary_cell, A] = sem.Fecha.ToUpper();
            if (num_sem != Conv_Wk)
            {
                if ((sem.Sem_Biblia != null) && (sem.Sem_Biblia != ""))
                {
                    Notify("Saving VyM Week: " + sem.Num_of_Week.ToString());
                    string a = "A", g = "G";
                    Sheet_VyM.Cells[primary_cell, D] = sem.Sem_Biblia.ToUpper();
                    Sheet_VyM.Cells[primary_cell, G] = sem.Presidente;
                    Sheet_VyM.Cells[primary_cell + 1, G] = sem.Consejero_Aux;
                    if (!sem.Discurso.Contains("min"))
                    {
                        sem.Discurso += "(10 mins.)";
                    }
                    Sheet_VyM.Cells[primary_cell + 6, C] = sem.Discurso;
                    Set_Font(Sheet_VyM.get_Range("C" + (primary_cell + 6)));
                    Sheet_VyM.Cells[primary_cell + 6, G] = sem.Discurso_A;
                    Sheet_VyM.Cells[primary_cell + 7, G] = sem.Perlas;
                    Sheet_VyM.Cells[primary_cell + 8, C] = sem.Lectura;
                    Set_Font(Sheet_VyM.get_Range("C" + (primary_cell + 8)));
                    Sheet_VyM.Cells[primary_cell + 8, G] = sem.Lectura_A;

                    Sheet_VyM.Cells[primary_cell + 11, C] = sem.SMM1;
                    Set_Font(Sheet_VyM.get_Range("C" + (primary_cell + 11)));
                    Sheet_VyM.Cells[primary_cell + 11, G] = sem.SMM1_A;
                    Sheet_VyM.Cells[primary_cell + 12, C] = sem.SMM2;
                    Set_Font(Sheet_VyM.get_Range("C" + (primary_cell + 12)));
                    Sheet_VyM.Cells[primary_cell + 12, G] = sem.SMM2_A;
                    if ((sem.SMM3 != null) && (sem.SMM3 != ""))
                    {
                        Sheet_VyM.Cells[primary_cell + 13, C] = sem.SMM3;
                        Set_Font(Sheet_VyM.get_Range("C" + (primary_cell + 13)));
                        Sheet_VyM.Cells[primary_cell + 13, G] = sem.SMM3_A;
                        if (is_room_B_enabled)
                        {
                            Sheet_VyM.Cells[primary_cell + 13, F] = sem.SMM3_B;
                        }
                    }
                    else
                    {
                        aux_range = Sheet_VyM.get_Range(a + (primary_cell + 13).ToString(), g + (primary_cell + 13).ToString());
                        aux_range.RowHeight = 3.75; // pixel = points * DPI / 72; DPI = 96  Set in 5 pixels if assignation is not filled; points = pixels / DPI * 72
                        aux_range.Cells.Clear();
                        increment_smm++;
                    }
                    if ((sem.SMM4 != null) && (sem.SMM4 != ""))
                    {
                        aux_range = Sheet_VyM.get_Range(a + (primary_cell + 14).ToString(), g + (primary_cell + 14).ToString());
                        aux_range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        aux_range.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        aux_range.Cells.WrapText = true;
                        aux_range.Characters.Font.Size = 9;
                        Sheet_VyM.Cells[primary_cell + 14, C] = sem.SMM4;
                        Set_Font(Sheet_VyM.get_Range("C" + (primary_cell + 14)));
                        Sheet_VyM.Cells[primary_cell + 14, G] = sem.SMM4_A;
                        if (is_room_B_enabled)
                        {
                            Sheet_VyM.Cells[primary_cell + 14, F] = sem.SMM4_B;
                        }
                    }
                    else
                    {
                        aux_range = Sheet_VyM.get_Range(a + (primary_cell + 14).ToString(), g + (primary_cell + 14).ToString());
                        aux_range.RowHeight = 3.75; // pixel = points * DPI / 72; DPI = 96  Set in 5 pixels if assignation is not filled; points = pixels / DPI * 72
                        aux_range.Cells.Clear();
                        increment_smm++;
                    }
                    if (is_room_B_enabled)
                    {
                        Sheet_VyM.Cells[primary_cell + 8,  F] = sem.Lectura_B;
                        Sheet_VyM.Cells[primary_cell + 10, F] = "Sala auxiliar";
                        Sheet_VyM.Cells[primary_cell + 11, F] = sem.SMM1_B;
                        Sheet_VyM.Cells[primary_cell + 12, F] = sem.SMM2_B;
                    }
                    Sheet_VyM.Cells[primary_cell + 17, C] = sem.NVC1;
                    Sheet_VyM.Cells[primary_cell + 17, G] = sem.NVC1_A;
                    if ((sem.NVC2 != null) && (sem.NVC2 != ""))
                    {
                        Sheet_VyM.Cells[primary_cell + 18, C] = sem.NVC2;
                        Sheet_VyM.Cells[primary_cell + 18, G] = sem.NVC2_A;
                    }
                    else
                    {
                        aux_range = Sheet_VyM.get_Range(a + (primary_cell + 18).ToString(), g + (primary_cell + 18).ToString());
                        aux_range.RowHeight = 3.75; // pixel = points * PDI / 72; PDI = 96  Set in 5 pixels if assignation is not filled; points = pixels / DPI * 72
                        aux_range.Cells.Clear();
                        increment_nvc = true;
                    }
                    if (sem.Num_of_Week == Vst_Wk)
                    {
                        Excel.Range range;
                        range = Sheet_VyM.get_Range("F" + (primary_cell + 19).ToString(), g +(primary_cell + 20).ToString());
                        range.Clear();
                        range = Sheet_VyM.get_Range("G" + (primary_cell + 19).ToString(), g + (primary_cell + 20).ToString());
                        range.Cells.Merge();
                        range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.WrapText = true;
                        range.Characters.Font.Size = 9;
                        Sheet_VyM.Cells[primary_cell + 19, C] = sem.Libro_Titulo;
                    }
                    Sheet_VyM.Cells[primary_cell + 19, G] = sem.Libro_A;
                    Sheet_VyM.Cells[primary_cell + 20, G] = sem.Libro_L;
                    Sheet_VyM.Cells[primary_cell + 22, G] = sem.Oracion;
                    DB_Handler.Persistence_VyM(sem, meetings_days[num_sem - 1, 0]);
                    //Heavensward Oracle set request

                    string[] Time_data = Get_time_from_week(sem);
                    Sheet_VyM.Cells[primary_cell + 2, A]  = Time_data[0];
                    Sheet_VyM.Cells[primary_cell + 3, A]  = Time_data[1];
                    Sheet_VyM.Cells[primary_cell + 6, A]  = Time_data[2];
                    Sheet_VyM.Cells[primary_cell + 7, A]  = Time_data[3];
                    Sheet_VyM.Cells[primary_cell + 8, A]  = Time_data[4];
                    Sheet_VyM.Cells[primary_cell + 11, A] = Time_data[5];
                    Sheet_VyM.Cells[primary_cell + 12, A] = Time_data[6];
                    Sheet_VyM.Cells[primary_cell + 13, A] = Time_data[7];
                    Sheet_VyM.Cells[primary_cell + 14, A] = Time_data[8];
                    Sheet_VyM.Cells[primary_cell + 16, A] = Time_data[9];
                    Sheet_VyM.Cells[primary_cell + 17, A] = Time_data[10];
                    Sheet_VyM.Cells[primary_cell + 18, A] = Time_data[11];
                    Sheet_VyM.Cells[primary_cell + 19, A] = Time_data[12];
                    Sheet_VyM.Cells[primary_cell + 21, A] = Time_data[13];
                    Sheet_VyM.Cells[primary_cell + 22, A] = Time_data[14];


                    // original = 15, increments = 23
                    aux_range = Sheet_VyM.get_Range(a + (primary_cell + 23).ToString(), g + (primary_cell + 23).ToString());
                    //double height = aux_range.RowHeight; //11.25 at 15 pixels
                    aux_range.RowHeight = 11.25;
                    if (increment_smm > 0)
                    {
                        aux_range.RowHeight += 17.25 * increment_smm;
                    }
                    if (increment_nvc)
                    {
                        aux_range.RowHeight += 10.5;
                    }
                }
            }
            else
            {
                Convention_Handler(1);
            }
        }

        public void VyM_Read_Week(short num_sem)
        {
            VyM_Sem sem = new VyM_Sem();
            short primary_cell = Get_vym_cell(num_sem);
            Get_month_from_Excel(cellValue_1[primary_cell + 1, A]);
            if (Check_null_string(cellValue_1[primary_cell, A]).Contains("Visita"))
            {
                Vst_Wk = num_sem;
                Notify("Marked current week [" + num_sem + "] as Visit");
                Alert_Label_VyM.Text = "Semana de la Visita del Superintendente de Circuito";
                Alert_Label_VyM.Visible = true;
            }
            else if (Check_null_string(cellValue_1[primary_cell + 10, A]).Contains("Asamblea"))
            {
                Conv_Wk = num_sem;
                Notify("Current week [" + num_sem.ToString() + "] setting as Convention [" + (Conv_Wk > 0 ? "True" : "False") + "]");
                Conv_Name = Check_null_string(cellValue_1[primary_cell + 10, A]);
                Alert_Label_VyM.Text = "Semana de Asamblea!";
                Alert_Label_VyM.Visible = true;
            }
            sem.Sem_Biblia    = Check_null_string(cellValue_1[primary_cell + 1, D]);
            sem.Presidente    = Check_null_string(cellValue_1[primary_cell + 1, G]);
            sem.Consejero_Aux = Check_null_string(cellValue_1[primary_cell + 2, G]);
            sem.Discurso      = Check_null_string(cellValue_1[primary_cell + 7, C]);
            sem.Discurso_A    = Check_null_string(cellValue_1[primary_cell + 7, G]);
            sem.Perlas        = Check_null_string(cellValue_1[primary_cell + 8, G]);
            sem.Lectura       = Check_null_string(cellValue_1[primary_cell + 9, C]);
            sem.Lectura_A     = Check_null_string(cellValue_1[primary_cell + 9, G]);
            sem.Lectura_B     = Check_null_string(cellValue_1[primary_cell + 9, F]);
            sem.SMM1          = Check_null_string(cellValue_1[primary_cell + 12, C]);
            sem.SMM1_A        = Check_null_string(cellValue_1[primary_cell + 12, G]);
            sem.SMM1_B        = Check_null_string(cellValue_1[primary_cell + 12, F]);
            sem.SMM2          = Check_null_string(cellValue_1[primary_cell + 13, C]);
            sem.SMM2_A        = Check_null_string(cellValue_1[primary_cell + 13, G]);
            sem.SMM2_B        = Check_null_string(cellValue_1[primary_cell + 13, F]);
            sem.SMM3          = Check_null_string(cellValue_1[primary_cell + 14, C]);
            sem.SMM3_A        = Check_null_string(cellValue_1[primary_cell + 14, G]);
            sem.SMM3_B        = Check_null_string(cellValue_1[primary_cell + 14, F]);
            sem.SMM4          = Check_null_string(cellValue_1[primary_cell + 15, C]);
            sem.SMM4_A        = Check_null_string(cellValue_1[primary_cell + 15, G]);
            sem.SMM4_B        = Check_null_string(cellValue_1[primary_cell + 15, F]);
            sem.NVC1          = Check_null_string(cellValue_1[primary_cell + 18, C]);
            sem.NVC1_A        = Check_null_string(cellValue_1[primary_cell + 18, G]);
            sem.NVC2          = Check_null_string(cellValue_1[primary_cell + 19, C]);
            sem.NVC2_A        = Check_null_string(cellValue_1[primary_cell + 19, G]);
            sem.Libro_Titulo  = Check_null_string(cellValue_1[primary_cell + 20, C]);
            sem.Libro_A       = Check_null_string(cellValue_1[primary_cell + 20, G]);
            sem.Libro_L       = Check_null_string(cellValue_1[primary_cell + 21, G]);
            sem.Oracion       = Check_null_string(cellValue_1[primary_cell + 23, G]);
            sem.Num_of_Week   = num_sem;

            switch (num_sem)
            {
                case 1:
                    {
                        VyM_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        VyM_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        VyM_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        VyM_mes.Semana4 = sem;
                        break;
                    }
                case 5:
                    {
                        VyM_mes.Semana5 = sem;
                        break;
                    }
            }
        }

        public void RP_Handler(bool save)
        {
            if (save)
            {
                RP_Save_Week(RP_mes.Semana1);
                loading += (15 / loading_delta);
                RP_Save_Week(RP_mes.Semana2);
                loading += (15 / loading_delta);
                RP_Save_Week(RP_mes.Semana3);
                loading += (15 / loading_delta);
                RP_Save_Week(RP_mes.Semana4);
                loading += (15 / loading_delta);
                if (week_five_exist)
                {
                    RP_Save_Week(RP_mes.Semana5);
                }
                loading += (15 / loading_delta);
            }
            else
            {
                for (short sm = 1; sm <= 5; sm++) //cycle to read all 5 weeks
                {
                    RP_Read_Week(sm);
                }
            }
        }

        public void RP_Save_Week(RP_Sem sem)
        {
            short num_sem = sem.Num_of_Week;
            short primary_cell = Get_rp_cell(num_sem);
            Sheet_RP.PageSetup.LeftHeader = "&16&B" + Cong_Name;
            if (num_sem == Vst_Wk)
            {
                Excel.Range range;
                Sheet_RP.Cells[primary_cell, A] = "Visita del Superintendente de Circuito";
                range = Sheet_RP.get_Range("A" + primary_cell.ToString());
                range.Interior.Color = Color.Orange;
            }
            primary_cell++;
            Sheet_RP.Cells[primary_cell, C] = sem.Fecha.ToUpper();
            if (num_sem != Conv_Wk)
            {
                if (sem.Presidente != null)
                {
                    Notify("Saving RP Week: " + sem.Num_of_Week.ToString());
                    Sheet_RP.Cells[primary_cell + 1, H]  = sem.Presidente;
                    Sheet_RP.Cells[primary_cell + 2, D]  = sem.Titulo;
                    if (sem.Titulo.ToLower().Contains("pendiente"))
                    {
                        Sheet_RP.Cells[primary_cell + 2, D].Font.Italic = true; // D = 4
                    }
                    Sheet_RP.Cells[primary_cell + 2, H]  = sem.Discursante;
                    Sheet_RP.Cells[primary_cell + 3, E]  = sem.Congregacion;
                    Sheet_RP.Cells[primary_cell + 6, D]  = sem.Titulo_Atalaya;
                    Sheet_RP.Cells[primary_cell + 5, H]  = sem.Conductor;
                    Sheet_RP.Cells[primary_cell + 6, H]  = sem.Lector;
                    Sheet_RP.Cells[primary_cell + 7, H]  = sem.Oracion;
                    Sheet_RP.Cells[primary_cell + 10, C] = sem.Discu_Sal;
                    Sheet_RP.Cells[primary_cell + 10, E] = sem.Ttl_Sal;
                    Sheet_RP.Cells[primary_cell + 10, H] = sem.Cong_Sal;
                    DB_Handler.Persistence_RP(sem, meetings_days[num_sem-1, 1]);
                    //Heavensward Oracle set request
                }
            }
            else
            {
                Convention_Handler(2);
            }
        }

        public void RP_Read_Week(short num_sem)
        {
            RP_Sem sem = new RP_Sem();
            short primary_cell = Get_rp_cell(num_sem);
            primary_cell++;
            sem.Presidente      = Check_null_string(cellValue_2[primary_cell + 1, H]);
            sem.Titulo          = Check_null_string(cellValue_2[primary_cell + 2, D]);
            sem.Discursante     = Check_null_string(cellValue_2[primary_cell + 2, H]);
            sem.Congregacion    = Check_null_string(cellValue_2[primary_cell + 3, E]);
            sem.Titulo_Atalaya  = Check_null_string(cellValue_2[primary_cell + 6, D]);
            sem.Conductor       = Check_null_string(cellValue_2[primary_cell + 5, H]);
            sem.Lector          = Check_null_string(cellValue_2[primary_cell + 6, H]);
            sem.Oracion         = Check_null_string(cellValue_2[primary_cell + 7, H]);
            sem.Discu_Sal       = Check_null_string(cellValue_2[primary_cell + 10, C]);
            sem.Ttl_Sal         = Check_null_string(cellValue_2[primary_cell + 10, E]);
            sem.Cong_Sal        = Check_null_string(cellValue_2[primary_cell + 10, H]);
            sem.Num_of_Week     = num_sem;
            switch (num_sem)
            {
                case 1:
                    {
                        RP_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        RP_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        RP_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        RP_mes.Semana4 = sem;
                        break;
                    }
                case 5:
                    {
                        RP_mes.Semana5 = sem;
                        break;
                    }
            }
        }

        public void AC_Handler(bool save) 
        {
            if (save)
            {
                AC_Save_Week(AC_mes.Semana1);
                loading += (15 / loading_delta);
                AC_Save_Week(AC_mes.Semana2);
                loading += (15 / loading_delta);
                AC_Save_Week(AC_mes.Semana3);
                loading += (15 / loading_delta);
                AC_Save_Week(AC_mes.Semana4);
                loading += (15 / loading_delta);
                if (week_five_exist)
                {
                    AC_Save_Week(AC_mes.Semana5);
                }
                else
                {
                    Excel.Range aux_range = Sheet_AC.get_Range("A" + 37, "E" + 41);
                    aux_range.Cells.UnMerge();
                    aux_range.Cells.Clear();
                }
                loading += (15 / loading_delta);
            }
            else
            {
                for (short sm = 1; sm <= 5; sm++) //cycle to read all 5 weeks
                {
                    AC_Read_Week(sm);
                }
            }
        }

        public void AC_Save_Week(AC_Sem sem)
        {
            short num_sem = sem.Num_of_Week;
            short primary_cell = Get_ac_cell(num_sem);
            Sheet_AC.PageSetup.LeftHeader = "&16&B" + Cong_Name;
            if (num_sem == Vst_Wk)
            {
                Excel.Range range;
                Sheet_AC.Cells[primary_cell, A] = "Visita del Superintendente de Circuito";
                range = Sheet_AC.get_Range("A" + primary_cell.ToString());
                range.Interior.Color = Color.Orange;
            }
            primary_cell++;
            Notify("Saving AC Week: " + sem.Num_of_Week.ToString());
            if (Ac_same_all_week)
            {
                string Date = "Semana del " + meetings_days[num_sem - 1, 0].ToString("dddd dd") + " de " + meetings_days[num_sem - 1, 0].ToString("MMMM") 
                    + " al " + meetings_days[num_sem - 1, 1].ToString("dddd dd") + " de " + meetings_days[num_sem - 1, 1].ToString("MMMM");
                string b = "B", c = "C", d = "D", e = "E";
                Excel.Range aux_range = Sheet_AC.get_Range(b + primary_cell.ToString(), d + primary_cell.ToString());
                aux_range.Cells.Merge();
                Sheet_AC.Cells[primary_cell, B] = Date;
                if (num_sem != Conv_Wk)
                {
                    if (sem.Vym_Cap != null)
                    {
                        Sheet_AC.Cells[primary_cell + 2, D] = "";
                        Sheet_AC.Cells[primary_cell + 3, D] = "";
                        aux_range = Sheet_AC.get_Range(d + (primary_cell + 3).ToString(), d + (primary_cell + 4).ToString());
                        aux_range.Cells.UnMerge();
                        aux_range = Sheet_AC.get_Range(c + (primary_cell + 2).ToString(), e + (primary_cell + 2).ToString());
                        aux_range.Cells.Merge();
                        aux_range = Sheet_AC.get_Range(c + (primary_cell + 3).ToString(), e + (primary_cell + 3).ToString());
                        aux_range.Cells.Merge();
                        aux_range = Sheet_AC.get_Range(c + (primary_cell + 4).ToString(), e + (primary_cell + 4).ToString());
                        aux_range.Cells.Merge();
                        Sheet_AC.Cells[primary_cell + 2, A] = sem.Aseo;
                        Sheet_AC.Cells[primary_cell + 2, C] = sem.Vym_Cap;
                        Sheet_AC.Cells[primary_cell + 3, C] = sem.Vym_Izq;
                        Sheet_AC.Cells[primary_cell + 4, C] = sem.Vym_Der;
                        DB_Handler.Persistence_AC(sem, meetings_days[num_sem - 1, 0], meetings_days[num_sem - 1, 1]);
                        //Heavensward Oracle set request
                    }
                }
                else
                {
                    Convention_Handler(3);
                }
            }
            else
            {
                Sheet_AC.Cells[primary_cell, B] = meetings_days[num_sem - 1, 0].ToString("dddd, dd MMMM");
                Sheet_AC.Cells[primary_cell, D] = meetings_days[num_sem - 1, 1].ToString("dddd, dd MMMM");
                if (num_sem != Conv_Wk)
                {
                    if (sem.Vym_Cap != null)
                    {
                        Sheet_AC.Cells[primary_cell + 2, A] = sem.Aseo;
                        Sheet_AC.Cells[primary_cell + 2, C] = sem.Vym_Cap;
                        Sheet_AC.Cells[primary_cell + 3, C] = sem.Vym_Izq;
                        Sheet_AC.Cells[primary_cell + 4, C] = sem.Vym_Der;
                        Sheet_AC.Cells[primary_cell + 2, E] = sem.Rp_Cap;
                        Sheet_AC.Cells[primary_cell + 3, E] = sem.Rp_Izq;
                        Sheet_AC.Cells[primary_cell + 4, E] = sem.Rp_Der;
                        DB_Handler.Persistence_AC(sem, meetings_days[num_sem - 1, 0], meetings_days[num_sem - 1, 1]);
                        //Heavensward Oracle set request
                    }
                }
                else
                {
                    Convention_Handler(3);
                }
            }
        }

        public void AC_Read_Week(short num_sem)
        {
            AC_Sem sem = new AC_Sem();
            short primary_cell = Get_ac_cell(num_sem);
            primary_cell++;
            sem.Aseo        = Check_null_string(cellValue_3[primary_cell + 2, A]);
            sem.Vym_Cap     = Check_null_string(cellValue_3[primary_cell + 2, C]);
            sem.Vym_Izq     = Check_null_string(cellValue_3[primary_cell + 3, C]);
            sem.Vym_Der     = Check_null_string(cellValue_3[primary_cell + 4, C]);
            sem.Rp_Cap      = Check_null_string(cellValue_3[primary_cell + 2, E]);
            sem.Rp_Izq      = Check_null_string(cellValue_3[primary_cell + 3, E]);
            sem.Rp_Der      = Check_null_string(cellValue_3[primary_cell + 4, E]);
            sem.Num_of_Week = num_sem;
            switch (num_sem)
            {
                case 1:
                    {
                        AC_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        AC_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        AC_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        AC_mes.Semana4 = sem;
                        break;
                    }
                case 5:
                    {
                        AC_mes.Semana5 = sem;
                        break;
                    }
            }
        }

        /*Function to set proper format to program when a Convention ocurrs*/
        public void Convention_Handler(short program)
        {
            Excel.Range range;
            string a = "A", g = "G", e = "E", h = "H", f = "F";
            switch (program)
            {
                case 1: //VyM
                    {
                        short cell = Get_vym_cell(Conv_Wk);
                        range = Sheet_VyM.get_Range(f + (cell + 1).ToString(), g + (cell + 2).ToString());
                        range.Cells.UnMerge();
                        range.Cells.Clear();
                        range = Sheet_VyM.get_Range(a+(cell+3).ToString(), g+(cell+23).ToString());
                        range.Cells.UnMerge();
                        range.Cells.Clear();
                        range = Sheet_VyM.get_Range(a + (cell + 10).ToString(), g + (cell + 15).ToString());
                        range.Cells.Merge();
                        range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Characters.Font.Size = 16;
                        range.Interior.Color = Color.Orange;
                        Sheet_VyM.Cells[cell + 10, A] = Conv_Name;
                        Sheet_VyM.Cells[cell, F] = "";
                        Sheet_VyM.Cells[cell + 1, F] = "";
                        break;
                    }
                case 2: //RP
                    {
                        short cell = Get_rp_cell(Conv_Wk);
                        cell++;
                        range = Sheet_RP.get_Range(a + (cell + 1).ToString(), h + (cell + 15).ToString());
                        range.Cells.UnMerge();
                        range.Cells.Clear();
                        range = Sheet_RP.get_Range(a + (cell + 2).ToString(), h + (cell + 7).ToString());
                        range.Cells.Merge();
                        range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Characters.Font.Size = 16;
                        range.Interior.Color = Color.Orange;
                        Sheet_RP.Cells[cell + 2, A] = Conv_Name;
                        break;
                    }
                case 3: //AC
                    {
                        short cell = Get_ac_cell(Conv_Wk);
                        cell++;
                        range = Sheet_AC.get_Range(a + (cell + 1).ToString(), e + (cell + 4).ToString());
                        range.Cells.UnMerge();
                        range.Cells.Clear();
                        range = Sheet_AC.get_Range(a + (cell + 1).ToString(), e + (cell + 4).ToString());
                        range.Cells.Merge();
                        range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Characters.Font.Size = 16;
                        range.Interior.Color = Color.Orange;
                        Sheet_AC.Cells[cell + 1, A] = Conv_Name;
                        break;
                    }
            }
        }

        /*--------------------------------------- Support functions to set/read strings ---------------------------------------*/

        private void Get_month_from_Excel(object cellvalue)
        {
            if (cellvalue != null && !month_found)
            {
                //bool month_found = false;
                for (int i = 0; i <= Months.Length - 1; i++)
                {
                    if (cellvalue.ToString().ToLower().Contains(Months[i]))
                    {
                        m_mes = i + 1;
                        month_found = true;
                        break;
                    }
                }
                if (month_found)
                {
                    Notify("Month set in [" + m_mes.ToString() + "]");
                }
                else
                {
                    m_mes = DateTime.Today.Month;
                    Warn("Month not found in first week, seeting today's month [" + m_mes.ToString() + "]");
                }
            }
        }

        public string Check_null_string(object cellvalue)
        {
            if (cellvalue == null)
            {
                cellvalue = "";
            }
            return cellvalue.ToString();
        }

        public short Get_vym_cell(short num_sem)
        {
            short cell = 0;
            switch (num_sem - 1)
            {
                case 0:
                    {
                        cell = 2;
                        break;
                    }
                case 1:
                    {
                        cell = 27;
                        break;
                    }
                case 2:
                    {
                        cell = 55;
                        break;
                    }
                case 3:
                    {
                        cell = 80;
                        break;
                    }
                case 4:
                    {
                        cell = 108;
                        break;
                    }
            }          
            return cell;
        }

        public short Get_rp_cell(short num_sem)
        {
            short cell = 0;
            switch (num_sem - 1)
            {
                case 0:
                    {
                        cell = 3;
                        break;
                    }
                case 1:
                    {
                        cell = 16;
                        break;
                    }
                case 2:
                    {
                        cell = 29;
                        break;
                    }
                case 3:
                    {
                        cell = 42;
                        break;
                    }
                case 4:
                    {
                        cell = 58;
                        break;
                    }
            }
            return cell;
        }

        public short Get_ac_cell(short num_sem)
        {
            short cell = 0;
            switch (num_sem - 1)
            {
                case 0:
                    {
                        cell = 4;
                        break;
                    }
                case 1:
                    {
                        cell = 12;
                        break;
                    }
                case 2:
                    {
                        cell = 20;
                        break;
                    }
                case 3:
                    {
                        cell = 28;
                        break;
                    }
                case 4:
                    {
                        cell = 36;
                        break;
                    }
            }
            return cell;
        }
        private void General_Info_Enter(object sender, EventArgs e)
        {
            Presenter(P.Executor);
            Notify("Overview");
        }

        private void Tesoros_Biblia_Enter(object sender, EventArgs e)
        {
            Presenter(P.DarkTemplar);
            Notify("Section 'Tesoros de la Biblia'");
        }

        private void Seamos_Maestros_Enter(object sender, EventArgs e)
        {
            Presenter(P.Selendis);
            Notify("Section 'Seamos Mejores Maestros'");
        }

        private void Nuestra_Vida_Enter(object sender, EventArgs e)
        {
            Presenter(P.DarkTemplar);
            Notify("Section 'Nuestra Vida Cristiana'");
        }

        private void Tab_Control_SelectedIndexChanged(object sender, EventArgs e)
        {
            var autocomplete = new AutoCompleteStringCollection();
            if (current_tab != tab_Control.SelectedIndex)
            {
                Pre_save_info();
            }
            switch (tab_Control.SelectedIndex)
            {
                case 0:
                    {
                        Presenter(P.Executor);
                        current_tab = 0;
                        autocomplete.AddRange(Dict_vym.Keys.ToArray());
                        Notify("\"Vida y Ministerio\" meeting");
                        break;
                    }
                case 1:
                    {
                        presenter_RP++;
                        if (presenter_RP > 5)
                        {
                            presenter_RP = 3;
                        }
                        Presenter((P)presenter_RP);
                        current_tab = 1;
                        autocomplete.AddRange(Dict_rp.Keys.ToArray());
                        Notify("\"Reunion Publica y Analisis de La Atalaya\" meeting");
                        break;
                    }
                case 2:
                    {
                        presenter_AC++;
                        if (presenter_AC > 8)
                        {
                            presenter_AC = 6;
                        }
                        Presenter((P)presenter_AC);
                        current_tab = 2;
                        autocomplete.AddRange(Dict_ac.Keys.ToArray());
                        Notify("\"Acomodadores\" Section");
                        break;
                    }
                case 3:
                    {
                        Pending_refresh_status_grids = true;
                        Presenter(P.HunterKiller);
                        current_tab = 3;
                        Notify("\"Male Status\" Section");
                        Rules_cmbx.SelectedIndex = 0;
                        break;
                    }
            }
            Week_Handler();
            txt_Command.AutoCompleteCustomSource = null;
            txt_Command.AutoCompleteCustomSource = autocomplete;
        }

        /*Get in real time the time used per asignment*/
        private void Txt_TextChanged(object sender, EventArgs e)
        {
            TextBox txbx = (TextBox)sender;
            VyM_Sem sem;
            switch (m_semana)
            {
                case 1:
                    {
                        sem = VyM_mes.Semana1;
                        break;
                    }
                case 2:
                    {
                        sem = VyM_mes.Semana2;
                        break;
                    }
                case 3:
                    {
                        sem = VyM_mes.Semana3;
                        break;
                    }
                case 4:
                    {
                        sem = VyM_mes.Semana4;
                        break;
                    }
                default:
                    {
                        sem = VyM_mes.Semana5;
                        break;
                    }
            }
            switch (txbx.Name)
            {
                case "txt_SMM1":
                    {
                        sem.SMM1 = txbx.Text;
                        break;
                    }
                case "txt_SMM2":
                    {
                        sem.SMM2 = txbx.Text;
                        break;
                    }
                case "txt_SMM3":
                    {
                        sem.SMM3 = txbx.Text;
                        break;
                    }
                case "txt_SMM4":
                    {
                        sem.SMM4 = txbx.Text;
                        break;
                    }
                case "txt_NVC1":
                    {
                        sem.NVC1 = txbx.Text;
                        break;
                    }
                case "txt_NVC2":
                    {
                        sem.NVC2 = txbx.Text;
                        break;
                    }
            }
            switch (m_semana)
            {
                case 1:
                    {
                        VyM_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        VyM_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        VyM_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        VyM_mes.Semana4 = sem;
                        break;
                    }
                default:
                    {
                        VyM_mes.Semana5 = sem;
                        break;
                    }
            }
            Time_Handler();
        }

        public void Time_Handler()
        {
            string[] Time_data = new string[15];
            switch (m_semana)
            {
                case 1:
                    {
                        Time_data = Get_time_from_week(VyM_mes.Semana1);
                        break;
                    }
                case 2:
                    {
                        Time_data = Get_time_from_week(VyM_mes.Semana2);
                        break;
                    }
                case 3:
                    {
                        Time_data = Get_time_from_week(VyM_mes.Semana3);
                        break;
                    }
                case 4:
                    {
                        Time_data = Get_time_from_week(VyM_mes.Semana4);
                        break;
                    }
                default:
                    {
                        Time_data = Get_time_from_week(VyM_mes.Semana5);
                        break;
                    }
            }
            time_0.Text = Time_data[0];
            time_1.Text = Time_data[1];
            time_2.Text = Time_data[2];
            time_3.Text = Time_data[3];
            time_4.Text = Time_data[4];
            time_5.Text = Time_data[5];
            time_6.Text = Time_data[6];
            time_7.Text = Time_data[7];
            time_8.Text = Time_data[8];
            time_9.Text = Time_data[9];
            time_10.Text = Time_data[10];
            time_11.Text = Time_data[11];
            time_12.Text = Time_data[12];
            time_13.Text = Time_data[13];
            time_14.Text = Time_data[14];
        }

        /*ToDo low priority set Dictionary txtbx name and time_X name*/
        private string[] Get_time_from_week(VyM_Sem sem)
        {
            string[] Time_data = new string[15];
            DateTime Aux_dateTime = new DateTime(2019, 1, 1, 7, 00, 00);

            Time_data[0] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(5);
            Time_data[1] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(3);
            Time_data[2] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(10);
            Time_data[3] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(8);
            Time_data[4] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(5 + 1); //adjusting to real time
            Time_data[5] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(sem.SMM1) + 1); //adjusting to real time
            Time_data[6] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(sem.SMM2) + 1); //adjusting to real time
            if ((sem.SMM3 == null) || (sem.SMM3 == ""))
            {
                Time_data[7] = "";
            }
            else
            {
                Time_data[7] = Aux_dateTime.ToString("HH:mm");
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(sem.SMM3) + 1); //adjusting to real time
            }
            if ((sem.SMM4 == null) || (sem.SMM4 == ""))
            {
                Time_data[8] = " ";
            }
            else
            {
                Time_data[8] = Aux_dateTime.ToString("HH:mm");
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(sem.SMM4) + 1); //adjusting to real time
            }
            Time_data[9] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(5);
            Time_data[10] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(sem.NVC1));
            if ((sem.NVC2 == null) || (sem.NVC2 == ""))
            {
                Time_data[11] = " ";
            }
            else
            {
                Time_data[11] = Aux_dateTime.ToString("HH:mm");
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(sem.NVC2));
            }
            Aux_dateTime = Aux_dateTime.AddMinutes(1); //adjusting to real time
            Time_data[12] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(30);
            Time_data[13] = Aux_dateTime.ToString("HH:mm");
            Aux_dateTime = Aux_dateTime.AddMinutes(3);
            Time_data[14] = Aux_dateTime.ToString("HH:mm");

            return Time_data;
        }

        public int Get_time_from_string(string Str)
        {
            int time = 0;
            if (Str != null)
            {
                Str = Str.ToLower();
                string min = "mins.";
                string number = "";
                //var array = Str.ToCharArray();
                if (Str.Contains(min))
                {
                    int index = Str.IndexOf(min);
                    number = Str.Substring(index - 3, 2);
                    try
                    {
                        if (number.Contains('('))
                        {
                            number = number.Remove(0, 1);
                        }
                        time = int.Parse(number);
                    }
                    catch
                    {
                        Warn("Must be numbers");
                        time = 0;
                    }
                }
                if (Str.Contains("video"))
                {
                    time--;
                }
            }
            return time;
        }

        private void Set_date()
        {
            int checksum_aux = m_año + m_mes + m_dia;
            lbl_Month.Text = "Mes: " + meetings_days[m_semana-1, 0].ToString("MMMM");
            if (checksum_aux != date_checksum)
            {
                date_checksum = checksum_aux;
                if ((m_año != 0) && (m_mes != 0) && (m_dia != 0))
                {
                    date = new DateTime(m_año, m_mes, m_dia);
                    m_calendar.SetDate(date);
                    Notify("Date Set in: [" + m_dia.ToString() + "/" + m_mes.ToString() + "/" + m_año.ToString() + "]");
                }
            }
        }

        /*Autofill handler*/
        public void AutoFill_Handler()
        {
            Notify("Executing AutoFill_Handler");
            Pre_save_info();
            switch (current_tab)
            {
                case 0:
                    {
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    VyM_mes.Semana1.AutoFill();
                                    DB_Handler.Persistence_VyM(VyM_mes.Semana1, meetings_days[0, 0]);
                                    break;
                                }
                            case 2:
                                {
                                    VyM_mes.Semana2.AutoFill();
                                    DB_Handler.Persistence_VyM(VyM_mes.Semana2, meetings_days[1, 0]);
                                    break;
                                }
                            case 3:
                                {
                                    VyM_mes.Semana3.AutoFill();
                                    DB_Handler.Persistence_VyM(VyM_mes.Semana3, meetings_days[2, 0]);
                                    break;
                                }
                            case 4:
                                {
                                    VyM_mes.Semana4.AutoFill();
                                    DB_Handler.Persistence_VyM(VyM_mes.Semana4, meetings_days[3, 0]);
                                    break;
                                }
                            case 5:
                                {
                                    VyM_mes.Semana5.AutoFill();
                                    DB_Handler.Persistence_VyM(VyM_mes.Semana5, meetings_days[4, 0]);
                                    break;
                                }
                        }
                        break;
                    }
                case 1:
                    {
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    RP_mes.Semana1.AutoFill();
                                    DB_Handler.Persistence_RP(RP_mes.Semana1, meetings_days[0, 1]);
                                    break;
                                }
                            case 2:
                                {
                                    RP_mes.Semana2.AutoFill();
                                    DB_Handler.Persistence_RP(RP_mes.Semana2, meetings_days[1, 1]);
                                    break;
                                }
                            case 3:
                                {
                                    RP_mes.Semana3.AutoFill();
                                    DB_Handler.Persistence_RP(RP_mes.Semana3, meetings_days[2, 1]);
                                    break;
                                }
                            case 4:
                                {
                                    RP_mes.Semana4.AutoFill();
                                    DB_Handler.Persistence_RP(RP_mes.Semana4, meetings_days[3, 1]);
                                    break;
                                }
                            case 5:
                                {
                                    RP_mes.Semana5.AutoFill();
                                    DB_Handler.Persistence_RP(RP_mes.Semana5, meetings_days[4, 1]);
                                    break;
                                }
                        }
                        break;
                    }

                case 2:
                    {
                        AC_mes.Semana1.AutoFill();
                        DB_Handler.Persistence_AC(AC_mes.Semana1, meetings_days[0, 0], meetings_days[0, 1]);
                        AC_mes.Semana2.AutoFill();
                        DB_Handler.Persistence_AC(AC_mes.Semana2, meetings_days[1, 0], meetings_days[1, 1]);
                        AC_mes.Semana3.AutoFill();
                        DB_Handler.Persistence_AC(AC_mes.Semana3, meetings_days[2, 0], meetings_days[2, 1]);
                        AC_mes.Semana4.AutoFill();
                        DB_Handler.Persistence_AC(AC_mes.Semana4, meetings_days[3, 0], meetings_days[3, 1]);
                        if (week_five_exist)
                        {
                            AC_mes.Semana5.AutoFill();
                            DB_Handler.Persistence_AC(AC_mes.Semana5, meetings_days[4, 0], meetings_days[4, 1]);
                        }
                        break;
                    }
            }
            Week_Handler();
        }

        /*Reader for speech number*/
        public string Get_Speech(string str)
        {
            string retval = str;
            if (int.TryParse(str, out int num))
            {
                string Speech_Path = Application.StartupPath + "\\\\Speech_List.txt";
                int len = File.ReadAllLines(Speech_Path).Length;
                string[] data = new string[len];

                StreamReader Reader_Speech = new StreamReader(Speech_Path);
                for (int i = 0; i < len; i++)
                {
                    data[i] = Reader_Speech.ReadLine();
                }
                Reader_Speech.Close();
                num--;
                if ((num < len) && (num >= 0))
                {
                    retval = data[num];
                    Notify("Selected speech number: " + (num++).ToString());
                }
                else
                {
                    Warn("Num outside the bounds");
                }
            }
            else
            {
                Warn("Unable to get selected speech");
            }
            return retval;
        }
        
        /*--------------------------------------- Week Handlers  ---------------------------------------*/

        /*Function so set local variables' info into form*/
        public void Week_Handler()
        {
            int lun = 0;
            lbl_Week.Text = "Semana: " + m_semana.ToString();
            if (Vst_Wk == m_semana)
            {
                Warn("Week [" + m_semana.ToString() + "] selected as Visit");
                Alert_Label_VyM.Text = "Semana de la Visita del Superintendente de Circuito";
                Alert_Label_VyM.Visible = true;
                Alert_Label_RP.Text = "Semana de la Visita del Superintendente de Circuito";
                Alert_Label_RP.Visible = true;
                txt_NVC3.Enabled = true;
                txt_NVC_A4.Enabled = false;
            }
            else if (Conv_Wk == m_semana)
            {
                Warn("Week [" + m_semana.ToString() + "] selected as Convention");
                Alert_Label_VyM.Text = Conv_Name;
                Alert_Label_VyM.Visible = true;
                Alert_Label_RP.Text = Conv_Name;
                Alert_Label_RP.Visible = true;
            }
            else
            {
                Alert_Label_VyM.Visible = false;
                Alert_Label_RP.Visible = false;
                txt_NVC3.Enabled = false;
                txt_NVC_A4.Enabled = true;
            }
            switch (current_tab)
            {
                case 0:
                    {
                        lbl_DateVyM.Text = meetings_days[m_semana - 1, 0].ToString("dddd, dd MMMM");
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    VyM_Week_Handler(VyM_mes.Semana1);
                                    break;
                                }
                            case 2:
                                {
                                    VyM_Week_Handler(VyM_mes.Semana2);
                                    break;
                                }
                            case 3:
                                {
                                    VyM_Week_Handler(VyM_mes.Semana3);
                                    break;
                                }
                            case 4:
                                {
                                    VyM_Week_Handler(VyM_mes.Semana4);
                                    break;
                                }
                            case 5:
                                {
                                    VyM_Week_Handler(VyM_mes.Semana5);
                                    break;
                                }
                        }
                        lun = 0;
                        break;
                    }
                case 1:
                    {
                        lbl_DateRP.Text = meetings_days[m_semana - 1, 1].ToString("dddd, dd MMMM");
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    RP_Week_Handler(RP_mes.Semana1);
                                    break;
                                }
                            case 2:
                                {
                                    RP_Week_Handler(RP_mes.Semana2);
                                    break;
                                }
                            case 3:
                                {
                                    RP_Week_Handler(RP_mes.Semana3);
                                    break;
                                }
                            case 4:
                                {
                                    RP_Week_Handler(RP_mes.Semana4);
                                    break;
                                }
                            case 5:
                                {
                                    RP_Week_Handler(RP_mes.Semana5);
                                    break;
                                }
                        }
                        Random rnd = new Random();
                        int rnd_pr = rnd.Next(3, 5);
                        Presenter((P)rnd_pr);
                        lun = 1;
                        break;
                    }
                case 2:
                    {
                        lbl_Dia_L_1.Text  = "Dia: " + meetings_days[0, 0].ToString("dddd, dd MMMM");
                        lbl_Dia_S_1.Text  = "Dia: " + meetings_days[0, 1].ToString("dddd, dd MMMM");
                        txt_Aseo_1.Text   = AC_mes.Semana1.Aseo;
                        txt_Cap_L_1.Text  = AC_mes.Semana1.Vym_Cap;
                        txt_AC1_L_1.Text  = AC_mes.Semana1.Vym_Izq;
                        txt_AC2_L_1.Text  = AC_mes.Semana1.Vym_Der;
                        txt_Cap_S_1.Text  = AC_mes.Semana1.Rp_Cap;
                        txt_AC1_S_1.Text  = AC_mes.Semana1.Rp_Izq;
                        txt_AC2_S_1.Text  = AC_mes.Semana1.Rp_Der;

                        lbl_Dia_L_2.Text  = "Dia: " + meetings_days[1, 0].ToString("dddd, dd MMMM");
                        lbl_Dia_S_2.Text  = "Dia: " + meetings_days[1, 1].ToString("dddd, dd MMMM");
                        txt_Aseo_2.Text   = AC_mes.Semana2.Aseo;
                        txt_Cap_L_2.Text  = AC_mes.Semana2.Vym_Cap;
                        txt_AC1_L_2.Text  = AC_mes.Semana2.Vym_Izq;
                        txt_AC2_L_2.Text  = AC_mes.Semana2.Vym_Der;
                        txt_Cap_S_2.Text  = AC_mes.Semana2.Rp_Cap;
                        txt_AC1_S_2.Text  = AC_mes.Semana2.Rp_Izq;
                        txt_AC2_S_2.Text  = AC_mes.Semana2.Rp_Der;

                        lbl_Dia_L_3.Text  = "Dia: " + meetings_days[2, 0].ToString("dddd, dd MMMM");
                        lbl_Dia_S_3.Text  = "Dia: " + meetings_days[2, 1].ToString("dddd, dd MMMM");
                        txt_Aseo_3.Text   = AC_mes.Semana3.Aseo;
                        txt_Cap_L_3.Text  = AC_mes.Semana3.Vym_Cap;
                        txt_AC1_L_3.Text  = AC_mes.Semana3.Vym_Izq;
                        txt_AC2_L_3.Text  = AC_mes.Semana3.Vym_Der;
                        txt_Cap_S_3.Text  = AC_mes.Semana3.Rp_Cap;
                        txt_AC1_S_3.Text  = AC_mes.Semana3.Rp_Izq;
                        txt_AC2_S_3.Text  = AC_mes.Semana3.Rp_Der;

                        lbl_Dia_L_4.Text  = "Dia: " + meetings_days[3, 0].ToString("dddd, dd MMMM");
                        lbl_Dia_S_4.Text  = "Dia: " + meetings_days[3, 1].ToString("dddd, dd MMMM");
                        txt_Aseo_4.Text   = AC_mes.Semana4.Aseo;
                        txt_Cap_L_4.Text  = AC_mes.Semana4.Vym_Cap;
                        txt_AC1_L_4.Text  = AC_mes.Semana4.Vym_Izq;
                        txt_AC2_L_4.Text  = AC_mes.Semana4.Vym_Der;
                        txt_Cap_S_4.Text  = AC_mes.Semana4.Rp_Cap;
                        txt_AC1_S_4.Text  = AC_mes.Semana4.Rp_Izq;
                        txt_AC2_S_4.Text  = AC_mes.Semana4.Rp_Der;

                        lbl_Dia_L_5.Text  = "Dia: " + meetings_days[4, 0].ToString("dddd, dd MMMM");
                        lbl_Dia_S_5.Text  = "Dia: " + meetings_days[4, 1].ToString("dddd, dd MMMM");
                        txt_Aseo_5.Text = AC_mes.Semana5.Aseo;
                        txt_Cap_L_5.Text  = AC_mes.Semana5.Vym_Cap;
                        txt_AC1_L_5.Text  = AC_mes.Semana5.Vym_Izq;
                        txt_AC2_L_5.Text  = AC_mes.Semana5.Vym_Der;
                        txt_Cap_S_5.Text  = AC_mes.Semana5.Rp_Cap;
                        txt_AC1_S_5.Text  = AC_mes.Semana5.Rp_Izq;
                        txt_AC2_S_5.Text  = AC_mes.Semana5.Rp_Der;
                        break;
                    }
            }
            //meetings_days
            Notify("Seeting info for week [" + m_semana.ToString() + "]");
            m_dia = meetings_days[m_semana - 1, lun].Day;
            m_mes = meetings_days[m_semana - 1, lun].Month;
            m_año = meetings_days[m_semana - 1, lun].Year;
            Set_date();
        }

        public void VyM_Week_Handler(VyM_Sem sem)
        {
            txt_Lec_Sem.Text = sem.Sem_Biblia;
            txt_Pres.Text    = sem.Presidente;
            txt_ConAux.Text  = sem.Consejero_Aux;
            txt_TdlB_1.Text  = sem.Discurso;
            txt_TdlB_A1.Text = sem.Discurso_A;
            txt_TdlB_A2.Text = sem.Perlas;
            txt_TdlB_3.Text  = sem.Lectura;
            txt_TdlB_A3.Text = sem.Lectura_A;
            txt_TdlB_B3.Text = sem.Lectura_B;
            txt_SMM1.Text    = sem.SMM1;
            txt_SMM_A1.Text  = sem.SMM1_A;
            txt_SMM_B1.Text  = sem.SMM1_B;
            txt_SMM2.Text    = sem.SMM2;
            txt_SMM_A2.Text  = sem.SMM2_A;
            txt_SMM_B2.Text  = sem.SMM2_B;
            txt_SMM3.Text    = sem.SMM3;
            txt_SMM_A3.Text  = sem.SMM3_A;
            txt_SMM_B3.Text  = sem.SMM3_B;
            txt_SMM4.Text    = sem.SMM4;
            txt_SMM_A4.Text  = sem.SMM4_A;
            txt_SMM_B4.Text  = sem.SMM4_B;
            txt_NVC1.Text    = sem.NVC1;
            txt_NVC_A1.Text  = sem.NVC1_A;
            txt_NVC2.Text    = sem.NVC2;
            txt_NVC_A2.Text  = sem.NVC2_A;
            txt_NVC_A3.Text  = sem.Libro_A;
            txt_NVC_A4.Text  = sem.Libro_L;
            txt_Ora2VyM.Text = sem.Oracion;

            if (sem.Num_of_Week == Vst_Wk)
            {
                txt_NVC3.Text = sem.Libro_Titulo;
            }
            else
            {
                txt_NVC3.Text = "Estudio biblico de congregacion (30 min.)";
            }
        }

        public void RP_Week_Handler(RP_Sem sem)
        {
            txt_RP_Speech.Text  = sem.Titulo;
            txt_PresRP.Text     = sem.Presidente;
            txt_RP_Cong.Text    = sem.Congregacion;
            txt_RP_Disc.Text    = sem.Discursante;
            txt_Title_Atly.Text = sem.Titulo_Atalaya;
            txt_Con_Atly.Text   = sem.Conductor;
            txt_Lect_Atly.Text  = sem.Lector;
            txt_OraRP.Text      = sem.Oracion;
            txt_Sal_Disc.Text   = sem.Discu_Sal;
            txt_Sal_Title.Text  = sem.Ttl_Sal;
            txt_Sal_Cong.Text   = sem.Cong_Sal;
        }
    

        /*Function to save txtbx info in local variables*/
        public void Pre_save_info()
        {
            Notify("Saving info into local variables");
            switch (current_tab)
            {
                case 0:
                    {
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    VyM_mes.Semana1 = VyM_Set_Week(VyM_mes.Semana1.Num_of_Week);
                                    break;
                                }
                            case 2:
                                {
                                    VyM_mes.Semana2 = VyM_Set_Week(VyM_mes.Semana2.Num_of_Week);
                                    break;
                                }
                            case 3:
                                {
                                    VyM_mes.Semana3 = VyM_Set_Week(VyM_mes.Semana3.Num_of_Week);
                                    break;
                                }
                            case 4:
                                {
                                    VyM_mes.Semana4 = VyM_Set_Week(VyM_mes.Semana4.Num_of_Week);
                                    break;
                                }
                            case 5:
                                {
                                    VyM_mes.Semana5 = VyM_Set_Week(VyM_mes.Semana5.Num_of_Week);
                                    break;
                                }
                        }
                        break;
                    }
                case 1:
                    {
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    RP_mes.Semana1 = RP_Set_Week(RP_mes.Semana1.Num_of_Week);
                                    break;
                                }
                            case 2:
                                {
                                    RP_mes.Semana2 = RP_Set_Week(RP_mes.Semana2.Num_of_Week);
                                    break;
                                }
                            case 3:
                                {
                                    RP_mes.Semana3 = RP_Set_Week(RP_mes.Semana3.Num_of_Week);
                                    break;
                                }
                            case 4:
                                {
                                    RP_mes.Semana4 = RP_Set_Week(RP_mes.Semana4.Num_of_Week);
                                    break;
                                }
                            case 5:
                                {
                                    RP_mes.Semana5 = RP_Set_Week(RP_mes.Semana5.Num_of_Week);
                                    break;
                                }
                        }
                        break;
                    }
                case 2:
                    {
                        AC_mes.Semana1.Aseo     = txt_Aseo_1.Text;
                        AC_mes.Semana1.Vym_Cap  = txt_Cap_L_1.Text;
                        AC_mes.Semana1.Vym_Izq  = txt_AC1_L_1.Text;
                        AC_mes.Semana1.Vym_Der  = txt_AC2_L_1.Text;
                        AC_mes.Semana1.Rp_Cap   = txt_Cap_S_1.Text;
                        AC_mes.Semana1.Rp_Izq   = txt_AC1_S_1.Text;
                        AC_mes.Semana1.Rp_Der   = txt_AC2_S_1.Text;

                        AC_mes.Semana2.Aseo     = txt_Aseo_2.Text;
                        AC_mes.Semana2.Vym_Cap  = txt_Cap_L_2.Text;
                        AC_mes.Semana2.Vym_Izq  = txt_AC1_L_2.Text;
                        AC_mes.Semana2.Vym_Der  = txt_AC2_L_2.Text;
                        AC_mes.Semana2.Rp_Cap   = txt_Cap_S_2.Text;
                        AC_mes.Semana2.Rp_Izq   = txt_AC1_S_2.Text;
                        AC_mes.Semana2.Rp_Der   = txt_AC2_S_2.Text;

                        AC_mes.Semana3.Aseo     = txt_Aseo_3.Text;
                        AC_mes.Semana3.Vym_Cap  = txt_Cap_L_3.Text;
                        AC_mes.Semana3.Vym_Izq  = txt_AC1_L_3.Text;
                        AC_mes.Semana3.Vym_Der  = txt_AC2_L_3.Text;
                        AC_mes.Semana3.Rp_Cap   = txt_Cap_S_3.Text;
                        AC_mes.Semana3.Rp_Izq   = txt_AC1_S_3.Text;
                        AC_mes.Semana3.Rp_Der   = txt_AC2_S_3.Text;

                        AC_mes.Semana4.Aseo     = txt_Aseo_4.Text;
                        AC_mes.Semana4.Vym_Cap  = txt_Cap_L_4.Text;
                        AC_mes.Semana4.Vym_Izq  = txt_AC1_L_4.Text;
                        AC_mes.Semana4.Vym_Der  = txt_AC2_L_4.Text;
                        AC_mes.Semana4.Rp_Cap   = txt_Cap_S_4.Text;
                        AC_mes.Semana4.Rp_Izq   = txt_AC1_S_4.Text;
                        AC_mes.Semana4.Rp_Der   = txt_AC2_S_4.Text;

                        AC_mes.Semana5.Aseo     = txt_Aseo_5.Text;
                        AC_mes.Semana5.Vym_Cap  = txt_Cap_L_5.Text;
                        AC_mes.Semana5.Vym_Izq  = txt_AC1_L_5.Text;
                        AC_mes.Semana5.Vym_Der  = txt_AC2_L_5.Text;
                        AC_mes.Semana5.Rp_Cap   = txt_Cap_S_5.Text;
                        AC_mes.Semana5.Rp_Izq   = txt_AC1_S_5.Text;
                        AC_mes.Semana5.Rp_Der   = txt_AC2_S_5.Text;
                        break;
                    }
            }
        }

        public VyM_Sem VyM_Set_Week(int num_week)
        {
            VyM_Sem sem = new VyM_Sem
            {
                Fecha         = lbl_DateVyM.Text,
                Sem_Biblia    = txt_Lec_Sem.Text,
                Presidente    = txt_Pres.Text,
                Consejero_Aux = txt_ConAux.Text,
                Discurso      = txt_TdlB_1.Text,
                Discurso_A    = txt_TdlB_A1.Text,
                Perlas        = txt_TdlB_A2.Text,
                Lectura       = txt_TdlB_3.Text,
                Lectura_A     = txt_TdlB_A3.Text,
                Lectura_B     = txt_TdlB_B3.Text,
                SMM1          = txt_SMM1.Text,
                SMM1_A        = txt_SMM_A1.Text,
                SMM1_B        = txt_SMM_B1.Text,
                SMM2          = txt_SMM2.Text,
                SMM2_A        = txt_SMM_A2.Text,
                SMM2_B        = txt_SMM_B2.Text,
                SMM3          = txt_SMM3.Text,
                SMM3_A        = txt_SMM_A3.Text,
                SMM3_B        = txt_SMM_B3.Text,
                SMM4          = txt_SMM4.Text,
                SMM4_A        = txt_SMM_A4.Text,
                SMM4_B        = txt_SMM_B4.Text,
                NVC1          = txt_NVC1.Text,
                NVC1_A        = txt_NVC_A1.Text,
                NVC2          = txt_NVC2.Text,
                NVC2_A        = txt_NVC_A2.Text,
                Libro_A       = txt_NVC_A3.Text,
                Libro_L       = txt_NVC_A4.Text,
                Oracion       = txt_Ora2VyM.Text,
                Num_of_Week   = (short)num_week,
            };
            if (Vst_Wk == num_week)
            {
                sem.Libro_Titulo = txt_NVC3.Text;
            }
            else
            {
                sem.Libro_Titulo = "Estudio biblico de congregacion (30 min.)";
            }
            return sem;
        }

        public RP_Sem RP_Set_Week(int num_week)
        {
            RP_Sem sem = new RP_Sem
            {
                Fecha          = lbl_DateRP.Text,
                Titulo         = txt_RP_Speech.Text,
                Presidente     = txt_PresRP.Text,
                Congregacion   = txt_RP_Cong.Text,
                Discursante    = txt_RP_Disc.Text,
                Titulo_Atalaya = txt_Title_Atly.Text,
                Conductor      = txt_Con_Atly.Text,
                Lector         = txt_Lect_Atly.Text,
                Oracion        = txt_OraRP.Text,
                Discu_Sal      = txt_Sal_Disc.Text,
                Ttl_Sal        = txt_Sal_Title.Text,
                Cong_Sal       = txt_Sal_Cong.Text,
                Num_of_Week    = (short)num_week
            };
            return sem;
        }


        /*--------------------------------------- Heavensward Handlers  ---------------------------------------*/

        public void Heavensward_request_handler()
        {
            if (HW_request[0].vym_sem != null)
            {
                Pre_save_info();
                switch (HW_request[0].vym_sem.Num_of_Week)
                {
                    case 1:
                        {
                            VyM_mes.Semana1 = Set_VyM_week_from_HW(VyM_mes.Semana1);
                            break;
                        }
                    case 2:
                        {
                            VyM_mes.Semana2 = Set_VyM_week_from_HW(VyM_mes.Semana2);
                            break;
                        }
                    case 3:
                        {
                            VyM_mes.Semana3 = Set_VyM_week_from_HW(VyM_mes.Semana3);
                            break;
                        }
                    case 4:
                        {
                            VyM_mes.Semana4 = Set_VyM_week_from_HW(VyM_mes.Semana4);
                            break;
                        }
                    case 5:
                        {
                            VyM_mes.Semana5 = Set_VyM_week_from_HW(VyM_mes.Semana5);
                            break;
                        }
                }
                Week_Handler();
            }
            else if (HW_request[0].rp_sem != null)
            {
                Pre_save_info();
                switch (HW_request[0].rp_sem.Num_of_Week)
                {
                    case 1:
                        {
                            RP_mes.Semana1 = Set_RP_week_from_HW(RP_mes.Semana1);
                            break;
                        }
                    case 2:
                        {
                            RP_mes.Semana2 = Set_RP_week_from_HW(RP_mes.Semana2);
                            break;
                        }
                    case 3:
                        {
                            RP_mes.Semana3 = Set_RP_week_from_HW(RP_mes.Semana3);
                            break;
                        }
                    case 4:
                        {
                            RP_mes.Semana4 = Set_RP_week_from_HW(RP_mes.Semana4);
                            break;
                        }
                    case 5:
                        {
                            RP_mes.Semana5 = Set_RP_week_from_HW(RP_mes.Semana5);
                            break;
                        }
                }
                Week_Handler();
            }
            HW_request.RemoveAt(0);
        }

        public static VyM_Sem Set_VyM_week_from_HW(VyM_Sem sem)
        {
            sem.Sem_Biblia = HW_request[0].vym_sem.Sem_Biblia;
            sem.Discurso   = HW_request[0].vym_sem.Discurso;
            sem.Lectura    = HW_request[0].vym_sem.Lectura;
            sem.SMM1       = HW_request[0].vym_sem.SMM1;
            sem.SMM2       = HW_request[0].vym_sem.SMM2;
            sem.SMM3       = HW_request[0].vym_sem.SMM3;
            sem.SMM4       = HW_request[0].vym_sem.SMM4;
            sem.NVC1       = HW_request[0].vym_sem.NVC1;
            sem.NVC2       = HW_request[0].vym_sem.NVC2;
            sem.HW_Data    = HW_request[0].vym_sem.HW_Data;
            return sem;
        }

        public static RP_Sem Set_RP_week_from_HW(RP_Sem sem)
        {
            sem.Titulo_Atalaya = HW_request[0].rp_sem.Titulo_Atalaya;
            return sem;
        }

        public class Hw_requested_info
        {
            public VyM_Sem vym_sem;
            public RP_Sem rp_sem;
        };

        public void Heavensward_All_Info()
        {
            Notify("Heavensward Month call");
            int max_sem = 4;
            if (week_five_exist)
            {
                max_sem = 5;
            }
            for (int sem = 1; sem <= max_sem; sem++)
            {
                Heavensward.HW_Bridge(meetings_days[sem - 1, 0], current_tab, sem);
            }
            Notify("Month information successfull");
        }

        /*--------------------------------------- State Database Handlers  ---------------------------------------*/

        public void Refresh_Males_Grid()
        {
            Male_Status_GridView.DataSource = Male_List;
            Male_Status_GridView.Refresh();
            Male_Status_GridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            Pending_refresh_status_grids = false;
            Set_Color_Result_DataGrid();
        }

        public void Set_Color_Result_DataGrid()
        {
            foreach (DataGridViewRow row in Male_Status_GridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        string str = cell.Value.ToString();
                        if (str.Contains('/'))
                        {
                            cell.Style.BackColor = Color.LightGreen;
                        }
                        else
                        {
                            switch (str)
                            {
                                case "Blocked":
                                    {
                                        cell.Style.BackColor = Color.LightCoral;
                                        break;
                                    }
                                case "Non_Status":
                                    {
                                        cell.Style.BackColor = Color.LightGray;
                                        break;
                                    }
                                case "Anciano":
                                    {
                                        cell.Style.BackColor = Color.LightSkyBlue;
                                        break;
                                    }
                                case "Ministerial":
                                    {
                                        cell.Style.BackColor = Color.LightSeaGreen;
                                        break;
                                    }
                                case "Publicador":
                                    {
                                        cell.Style.BackColor = Color.LightYellow;
                                        break;
                                    }
                                default:
                                    {
                                        cell.Style.BackColor = Color.White;
                                        break;
                                    }
                            }
                        }
                    }
                }
            }
        }

        private void Male_Status_GridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DataGridViewCell cell = Male_Status_GridView.CurrentCell;
            Previous_Male_Type = "";
            if (cell.ColumnIndex == 7 && cell.Value != null)
            {
                Previous_Male_Type = cell.Value.ToString();
            }
        }


        private void Male_Status_GridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = Male_Status_GridView.CurrentCell;
            if (cell.ColumnIndex == 0)
            {
                if ((Male_List.Count > Males_Count) && (cell.Value != null))  
                {
                    Notify("New Male added: " + cell.Value.ToString());
                    Male_List[cell.RowIndex].male_type = Male_Type.Publicador;
                    Change_Male_type(cell.RowIndex, Male_Type.Publicador);
                }
                else
                {
                    if (cell.Value == null || cell.Value.ToString() == "")
                    {
                        Notify("Male removed: " + Male_List[cell.RowIndex].Name);
                        Previous_Male_Type = Male_List[cell.RowIndex].male_type.ToString();
                        Remove_Males_Count();
                        Males_Count = Elders_Count + Ministerials_Count + Generals_Count;
                        Male_List.RemoveAt(cell.RowIndex);
                    }
                    else
                    {
                        Notify("Modified Male name as: " + cell.Value.ToString());
                    }
                }
            }
            else if (cell.ColumnIndex == 7 && cell.Value != null)
            {
                string cell_str = cell.Value.ToString();
                switch (cell_str)
                {
                    case "Anciano":
                        {
                            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightSkyBlue;
                            if (!Previous_Male_Type.Equals(cell_str))
                            {
                                Notify("Status of " + Male_List[cell.RowIndex].Name + " changed to \"Anciano\"");
                                Change_Male_type(cell.RowIndex, Male_Type.Anciano);
                            }
                            break;
                        }
                    case "Ministerial":
                        {
                            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightSeaGreen;
                            if (!Previous_Male_Type.Equals(cell_str))
                            {
                                Notify("Status of " + Male_List[cell.RowIndex].Name + " changed to \"Ministerial\"");
                                Change_Male_type(cell.RowIndex, Male_Type.Ministerial);
                            }
                            break;
                        }
                    case "Publicador":
                        {
                            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightYellow;
                            if (!Previous_Male_Type.Equals(cell_str))
                            {
                                Notify("Status of " + Male_List[cell.RowIndex].Name + " changed to \"Ministerial\"");
                                Change_Male_type(cell.RowIndex, Male_Type.Publicador);
                            }
                            break;
                        }
                    default:
                        {
                            Warn("Unrecognized value, changing to default state: \"Publicador\"");
                            Male_Status_GridView.CurrentCell.Value = Male_Type.Publicador;
                            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightGray;
                            break;
                        }
                }
            }
            else if(cell.Value != null)
            {
                //string cell_Value = cell.Value.ToString();
                switch (cell.Value.ToString())
                {
                    case "Non_status":
                        {
                            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightGray;
                            Notify("Status of " + Male_List[cell.RowIndex].Name + " changed successfully to \"Non_status\"");
                            break;
                        }
                    case "Blocked":
                        {

                            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightCoral;
                            Notify("Status of " + Male_List[cell.RowIndex].Name + " changed successfully to \"Blocked\"");
                            break;
                        }
                    default:
                        {
                            if (DateTime.TryParse(cell.Value.ToString(), out DateTime date))
                            {
                                Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightGreen;
                                Notify("DateTime updated Successful");
                            }
                            else
                            {
                                Warn("Unrecognized value, changing status of " + Male_List[cell.RowIndex].Name + " to default state: \"Non_status\"");
                                Male_Status_GridView.CurrentCell.Value = Male_State.Non_status;
                                Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightGray;
                            }
                            break;
                        }
                }
            }
        }

        public static void Change_Male_type(int index, Male_Type new_male_type)
        {
            Males male = Male_List[index];
            Male_List.RemoveAt(index);
            Remove_Males_Count();
            switch (new_male_type)
            {
                case Male_Type.Anciano:
                    {
                        male = DB_Handler.Set_Status(Rule_Elders, male);
                        Male_List.Insert(Elders_Count, male);
                        Elders_Count++;
                        break;
                    }
                case Male_Type.Ministerial:
                    {
                        male = DB_Handler.Set_Status(Rule_Ministerials, male);
                        Male_List.Insert(Elders_Count + Ministerials_Count, male);
                        Ministerials_Count++;
                        break;
                    }
                case Male_Type.Publicador:
                    {
                        male = DB_Handler.Set_Status(Rule_Generals, male);
                        Male_List.Insert(Elders_Count + Ministerials_Count + Generals_Count, male);
                        Generals_Count++;
                        break;
                    }
            }
            Males_Count = Elders_Count + Ministerials_Count + Generals_Count;
            Previous_Male_Type = "";
            Pending_refresh_status_grids = true;
        }

        public static void Remove_Males_Count()
        {
            if (Previous_Male_Type != "")
            {
                switch (Previous_Male_Type)
                {
                    case "Anciano":
                        {
                            Elders_Count--;
                            break;
                        }
                    case "Ministerial":
                        {
                            Ministerials_Count--;
                            break;
                        }
                    case "Publicador":
                        {
                            Generals_Count--;
                            break;
                        }
                }
                Previous_Male_Type = "";
            }
        }

        protected void Male_Status_GridView_Data_Error(object sender, DataGridViewDataErrorEventArgs e)
        {
            Warn("Unrecognized value, changing to default state: Non_status");
            Male_Status_GridView.CurrentCell.Value = Male_State.Non_status;
            Male_Status_GridView.CurrentCell.Style.BackColor = Color.LightGray;
            e.ThrowException = false;
            e.Cancel = false;
        }

        private void Rules_cmbx_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (Rules_cmbx.SelectedIndex)
            {
                case 0:
                    {
                        Set_Rule_Status(Chk_Atalaya, Rule_Elders.Atalaya);
                        Set_Rule_Status(Chk_Capitan, Rule_Elders.Capitan);
                        Set_Rule_Status(Chk_Acomodador, Rule_Elders.Acomodador);
                        Set_Rule_Status(Chk_Lector, Rule_Elders.Lector);
                        Set_Rule_Status(Chk_PresRp, Rule_Elders.Pres_RP);
                        Set_Rule_Status(Chk_Oracion, Rule_Elders.Oracion);
                        break;
                    }
                case 1:
                    {
                        Set_Rule_Status(Chk_Atalaya, Rule_Ministerials.Atalaya);
                        Set_Rule_Status(Chk_Capitan, Rule_Ministerials.Capitan);
                        Set_Rule_Status(Chk_Acomodador, Rule_Ministerials.Acomodador);
                        Set_Rule_Status(Chk_Lector, Rule_Ministerials.Lector);
                        Set_Rule_Status(Chk_PresRp, Rule_Ministerials.Pres_RP);
                        Set_Rule_Status(Chk_Oracion, Rule_Ministerials.Oracion);
                        break;
                    }
                case 2:
                    {
                        Set_Rule_Status(Chk_Atalaya, Rule_Generals.Atalaya);
                        Set_Rule_Status(Chk_Capitan, Rule_Generals.Capitan);
                        Set_Rule_Status(Chk_Acomodador, Rule_Generals.Acomodador);
                        Set_Rule_Status(Chk_Lector, Rule_Generals.Lector);
                        Set_Rule_Status(Chk_PresRp, Rule_Generals.Pres_RP);
                        Set_Rule_Status(Chk_Oracion, Rule_Generals.Oracion);
                        break;
                    }
            }
        }

        private void Set_Rule_Status(CheckBox checkBox, string str)
        {
            checkBox.Checked = false;
            if (str.Equals("Allowed"))
            {
                checkBox.Checked = true;
            }
        }

        private void Checkbox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox Chkbx = (CheckBox)sender;
            if (Chkbx.Checked)
            {
                Chkbx.BackColor = Color.LightGreen;
            }
            else
            {
                Chkbx.BackColor = Color.LightCoral;
            }
        }

        private void Edit_Rule_Btn_Click(object sender, EventArgs e)
        {
            if (Edit_Rule)
            {
                Notify("Enabled Edit Rules");
                Edit_Rule_Btn.Text = "Save";
                Edit_Rule = false;
                Chk_Atalaya.Enabled = true;
                Chk_Capitan.Enabled = true;
                Chk_Acomodador.Enabled = true;
                Chk_Lector.Enabled = true;
                Chk_PresRp.Enabled = true;
                Chk_Oracion.Enabled = true;
            }
            else
            {
                Notify("Saving and Re-Run Rules");
                Edit_Rule_Btn.Text = "Edit";
                Edit_Rule = true;
                Chk_Atalaya.Enabled = false;
                Chk_Capitan.Enabled = false;
                Chk_Acomodador.Enabled = false;
                Chk_Lector.Enabled = false;
                Chk_PresRp.Enabled = false;
                Chk_Oracion.Enabled = false;
                switch (Rules_cmbx.SelectedIndex)
                {
                    case 0:
                        {
                            Rule_Elders = Run_Rules(Rule_Elders);
                            break;
                        }
                    case 1:
                        {
                            Rule_Ministerials = Run_Rules(Rule_Ministerials);
                            break;
                        }
                    case 2:
                        {
                            Rule_Generals = Run_Rules(Rule_Generals);
                            break;
                        }
                }
            }
        }

        public Males Run_Rules(Males local_rule)
        {
            bool rerun_rules = false;
            if(!Chk_Atalaya.Checked && (local_rule.Atalaya == "Allowed"))
            {
                rerun_rules = true;
            }
            return local_rule;
        }
    }
}
