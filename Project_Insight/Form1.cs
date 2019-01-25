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

/*Developed by AGR-Systems Science and Tech Division*/

namespace Project_Insight
{
    public partial class Main_Form : Form
    {
        public enum p
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

        public static int A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7, H = 8;
        public static int actual_presenter = 10;
        private Excel.Application objApp;
        private Excel.Workbook objBooks = null;
        private Excel.Sheets objSheets;
        private Excel.Worksheet Sheet_VyM;
        private Excel.Worksheet Sheet_RP;
        private Excel.Worksheet Sheet_AC;
        private Excel.Range range_1;
        private Excel.Range range_2;
        private Excel.Range range_3;
        public static bool excel_ready = false;
        private DateTime start_time = new DateTime(DateTime.Today.Year, 1, 1, 7, 00, 00);
        private DateTime date;
        private object[,] cellValue_1 = null;
        private object[,] cellValue_2 = null;
        private object[,] cellValue_3 = null;
        public static object[,] cellValue_4 = null;
        public static string[] str_stack = new string[50];
        public static int[] int_stack = new int[50];
        public static bool busy_trace = false;
        public static bool pending_trace = false;
        public static short iterator_stack = 0;
        //public static string message_form2 = null;
        //public static bool pending_refresh_DB = false;
        public static int m_dia = 1;
        public static int m_mes = 1;
        public static int m_año = DateTime.Today.Year;
        public static short m_semana = 1;
        public static DateTime[,] meetings_days = new DateTime[5, 2];
        public static string[] guard_cbx_names = new string[10];
        public static int date_checksum = 0;
        public static string[] Command_history = new string[10];
        public static string[] month = new string[] { "ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic" };
        public static int command_iterator = 0;
        DB_Form DB_Form = new DB_Form();
        //public static bool DB_form_show = false;
        public static string Path = "";
        public static bool is_new_instance = false;
        public static VyM_Mes VyM_mes = new VyM_Mes();
        public static RP_Mes RP_mes = new RP_Mes();
        public static AC_Mes AC_mes = new AC_Mes();
        IDictionary<string, object> Dict_vym = new Dictionary<string, object>();
        IDictionary<string, object> Dict_rp = new Dictionary<string, object>();
        IDictionary<string, object> Dict_ac = new Dictionary<string, object>();
        public static int tab_meeting = 0;
        public static string aux_command;
        //public static bool selected_txt = false;
        public static int loading_delta = 1;
        public static int loading = 0;
        public static bool week_five_exist = false;
        public static bool UI_running = false;
        public static short Conv_Wk = 0;
        public static string Conv_Name = "";



        public Main_Form()
        {
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            InitializeComponent();
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            Notify("Project Insight 2.0");
            Notify("UI up and ready \nWelcome back Hierarch!");
            Presenter(p.Executor);
            Autocomplete_dictionary();
            txt_Command.Focus();
        }

        public void Autocomplete_dictionary()
        {
            Dict_vym.Add("ig_01", txt_Date);
            Dict_vym.Add("ig_02", txt_Pres);
            Dict_vym.Add("tb_01", txt_TdlB_1);
            Dict_vym.Add("tb_02", txt_TdlB_A1);
            Dict_vym.Add("tb_03", txt_TdlB_A2);
            Dict_vym.Add("tb_04", txt_TdlB_A3);
            Dict_vym.Add("sm_11", txt_SMM1);
            Dict_vym.Add("sm_12", txt_SMM_A1);
            Dict_vym.Add("sm_21", txt_SMM2);
            Dict_vym.Add("sm_22", txt_SMM_A2);
            Dict_vym.Add("sm_31", txt_SMM3);
            Dict_vym.Add("sm_32", txt_SMM_A3);
            Dict_vym.Add("sm_41", txt_SMM4);
            Dict_vym.Add("sm_42", txt_SMM_A4);
            Dict_vym.Add("nv_11", txt_NVC1);
            Dict_vym.Add("nv_12", txt_NVC_A1);
            Dict_vym.Add("nv_21", txt_NVC2);
            Dict_vym.Add("nv_22", txt_NVC_A2);
            Dict_vym.Add("nv_30", txt_NVC_A3);
            Dict_vym.Add("nv_40", txt_NVC_A4);
            Dict_vym.Add("nv_50", txt_Ora2VyM);

            //Dict_rp.Add("rp_00", lbl_DateRP);
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
            Dict_ac.Add("ac_15", txt_Cap_S_1);
            Dict_ac.Add("ac_16", txt_AC1_S_1);
            Dict_ac.Add("ac_17", txt_AC2_S_1);
            Dict_ac.Add("ac_21", txt_Aseo_2);
            Dict_ac.Add("ac_22", txt_Cap_L_2);
            Dict_ac.Add("ac_23", txt_AC1_L_2);
            Dict_ac.Add("ac_24", txt_AC2_L_2);
            Dict_ac.Add("ac_25", txt_Cap_S_2);
            Dict_ac.Add("ac_26", txt_AC1_S_2);
            Dict_ac.Add("ac_27", txt_AC2_S_2);
            Dict_ac.Add("ac_31", txt_Aseo_3);
            Dict_ac.Add("ac_32", txt_Cap_L_3);
            Dict_ac.Add("ac_33", txt_AC1_L_3);
            Dict_ac.Add("ac_34", txt_AC2_L_3);
            Dict_ac.Add("ac_35", txt_Cap_S_3);
            Dict_ac.Add("ac_36", txt_AC1_S_3);
            Dict_ac.Add("ac_37", txt_AC2_S_3);
            Dict_ac.Add("ac_41", txt_Aseo_4);
            Dict_ac.Add("ac_42", txt_Cap_L_4);
            Dict_ac.Add("ac_43", txt_AC1_L_4);
            Dict_ac.Add("ac_44", txt_AC2_L_4);
            Dict_ac.Add("ac_45", txt_Cap_S_4);
            Dict_ac.Add("ac_46", txt_AC1_S_4);
            Dict_ac.Add("ac_47", txt_AC2_S_4);
            Dict_ac.Add("ac_51", txt_Aseo_5);
            Dict_ac.Add("ac_52", txt_Cap_L_5);
            Dict_ac.Add("ac_53", txt_AC1_L_5);
            Dict_ac.Add("ac_54", txt_AC2_L_5);
            Dict_ac.Add("ac_55", txt_Cap_S_5);
            Dict_ac.Add("ac_56", txt_AC1_S_5);
            Dict_ac.Add("ac_57", txt_AC2_S_5);
        }

        public void Main_Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (excel_ready)
            {
                excel_ready = false;
                objBooks.Close(0);
                objApp.Quit();
                Marshal.ReleaseComObject(Sheet_VyM);
                Marshal.ReleaseComObject(objBooks);
                Marshal.ReleaseComObject(objApp);
            }
            Application.Exit();
        }

        /*--------------------------------------- Traces and UI functions ---------------------------------------*/
        public async void Presenter(p ID_Presenter)
        {
            if (actual_presenter != (int)ID_Presenter)
            {
                actual_presenter = (int)ID_Presenter;
                picPresenter.Image = Project_Insight.Properties.Resources.Noise;
                await Task.Delay(300);
                switch (ID_Presenter)
                {
                    case p.Executor:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.Executor;
                            break;
                        }
                    case p.FenixZealot:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.FenixZealot;
                            break;
                        }
                    case p.FenixDragoon:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.FenixDragoon;
                            break;
                        }
                    case p.Selendis:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.Selendis;
                            break;
                        }
                    case p.Oracle:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.Oracle;
                            break;
                        }
                    case p.DarkTemplar:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.DarkTemplar;
                            break;
                        }
                    case p.HunterKiller:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.HunterKiller;
                            break;
                        }
                    case p.Hybrid:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.Hybrid;
                            break;
                        }
                    case p.Artanis:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.Artanis;
                            break;
                        }
                }
            }
        }

        public async void Notify(string data, [CallerLineNumber] int lineNumber = 0)
        {
            if (!busy_trace)
            {
                busy_trace = true;
                var array = data.ToCharArray();
                log_txtBx.SelectionColor = Color.Orange;
                //log_txtBx.AppendText("L-" + lineNumber + ": ");
                for (int i = 0; i <= array.Length - 1; i++)
                {
                    log_txtBx.AppendText(array[i].ToString());
                    await Task.Delay(5);
                }
                log_txtBx.AppendText("\n");
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                log_txtBx.ScrollToCaret();
                busy_trace = false;
                if (pending_trace)
                {
                    String_stack("", false, 1, lineNumber);
                }
            }
            else
            {
                String_stack(data, true, 1, lineNumber);
            }
        }

        public async void Warn(string data, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            if (!busy_trace)
            {
                busy_trace = true;
                var array = data.ToCharArray();
                log_txtBx.SelectionColor = Color.Red;
                log_txtBx.AppendText("L-" + lineNumber + ": ");
                if (caller != "String_stack")
                {
                    log_txtBx.AppendText("(" + caller + ") ");
                }
                for (int i = 0; i <= array.Length - 1; i++)
                {
                    log_txtBx.AppendText(array[i].ToString());
                    await Task.Delay(5);
                }
                log_txtBx.AppendText("\n");
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                log_txtBx.ScrollToCaret();
                busy_trace = false;
                if (pending_trace)
                {
                    String_stack("", false, 2, lineNumber);
                }
            }
            else
            {
                String_stack("(" + caller + ") " + data, true, 2, lineNumber);
            }
        }

        public async void Loading_Trace()
        {
            string aux = "";
            int delay = 40;
            while (busy_trace)
            {
                await Task.Delay(500);
            }
            busy_trace = true;
            log_txtBx.SelectionColor = Color.Orange;
            log_txtBx.AppendText("\nLoading:  ");
            aux = log_txtBx.Text;
            log_txtBx.Text = "";
            log_txtBx.SelectionStart = log_txtBx.Text.Length;
            log_txtBx.ScrollToCaret();
            aux += loading.ToString() + " % ...-";
            log_txtBx.Text = aux;
            log_txtBx.SelectionStart = log_txtBx.Text.Length;
            log_txtBx.ScrollToCaret();
            aux = aux.Substring(0, aux.Length - 7 - loading.ToString().Length);
            while (loading < 100)
            {
                aux += loading.ToString() + " % ...\\";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                aux = aux.Substring(0, aux.Length - 7 - loading.ToString().Length);
                await Task.Delay(delay);
                aux += loading.ToString() + " % ...|";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                aux = aux.Substring(0, aux.Length - 7 - loading.ToString().Length);
                await Task.Delay(delay);
                aux += loading.ToString() + " % .../";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                aux = aux.Substring(0, aux.Length - 7 - loading.ToString().Length);
                await Task.Delay(delay);
                aux += loading.ToString() + " % ...-";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                aux = aux.Substring(0, aux.Length - 7 - loading.ToString().Length);
                await Task.Delay(delay);
            }
            aux += loading.ToString() + " % ...\\";
            log_txtBx.Text = aux;
            log_txtBx.SelectionStart = log_txtBx.Text.Length;
            log_txtBx.ScrollToCaret();
            await Task.Delay(delay);
            log_txtBx.AppendText("\n");
            log_txtBx.SelectionColor = Color.Green;
            log_txtBx.AppendText("Complete!");
            log_txtBx.AppendText("\n");
            await Task.Delay(delay);
            log_txtBx.SelectionStart = log_txtBx.Text.Length;
            log_txtBx.ScrollToCaret();
            busy_trace = false;
            if (pending_trace)
            {
                String_stack("", false, 1, 0);
            }
        }

        public void String_stack(string data, bool save, int trace, int line)
        {
            if (save)
            {
                switch (trace)
                {
                    case 1:
                        {
                            data += "1";
                            break;
                        }
                    case 2:
                        {
                            data += "2";
                            break;
                        }
                }
                str_stack[iterator_stack] = data;
                int_stack[iterator_stack] = line;
                pending_trace = true;
                iterator_stack++;
            }
            else
            {
                int notify_warn = int.Parse(str_stack[0].Substring(str_stack[0].Length - 1));
                str_stack[0] = str_stack[0].Substring(0, str_stack[0].Length - 1);
                if (notify_warn == 1)
                {
                    Notify(str_stack[0], int_stack[0]);
                }
                else if (notify_warn == 2)
                {
                    Warn(str_stack[0], int_stack[0]);
                }
                for (int i = 1; i <= str_stack.Length - 1; i++)
                {
                    str_stack[i - 1] = str_stack[i];
                    int_stack[i - 1] = int_stack[i];
                }
                iterator_stack--;
                if (iterator_stack == 0)
                {
                    pending_trace = false;
                }
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
                    if (Str.Length > 4)
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
                    Notify("Executing [" + Str + "] command");
                    Save_command(Str);
                    command_iterator = 0;
                    switch (cmd)
                    {
                        case "new":
                            {
                                bool month_found = false;
                                for (int i = 0; i <= month.Length - 1; i++)
                                {
                                    if (sup.Contains(month[i]))
                                    {
                                        m_mes = i + 1;
                                        month_found = true;
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
                                break;
                            }
                        case "open":
                            {
                                Known_Instance();
                                break;
                            }
                        case "exit":
                            {
                                Main_Form_FormClosed(this, null);
                                break;
                            }
                        case "save":
                            {
                                if (UI_running)
                                {
                                    if (int.TryParse(sup, out int sv))
                                    {
                                        Thread Save_thread = new Thread(() => Process_save(sv));
                                        Save_thread.Start();
                                        Loading_Trace();
                                    }
                                }
                                else
                                {
                                    Warn("Need to create a new instance or open an existing program");
                                }
                                break;
                            }
                        case "tab":
                            {
                                if (UI_running)
                                {
                                    if (int.TryParse(sup, out int tab))
                                    {
                                        tab_Control.SelectedIndex = tab;
                                    }
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
                                                Warn("Selected month [" + m_mes.ToString() + "] doesn't have 5 weeks");
                                            }
                                            else
                                            {
                                                Pre_save_info();
                                                m_semana = (short)wk;
                                                Week_Handler();
                                            }
                                        }
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
                                    if (sup.Contains("false"))
                                    {
                                        Conv_Wk = 0;
                                    }
                                    else 
                                    {
                                        Conv_Name = sup;
                                        Conv_Wk = m_semana;
                                    }
                                    Notify("Current week [" + m_semana.ToString() + "] setting as Convention [" + (Conv_Wk > 0 ? "True" : "False") + "]");
                                }
                                else
                                {
                                    Warn("Need to create a new instance or open an existing program");
                                }
                                break;
                            }
                        case "db":
                            {
                                DB_Form.Show();
                                break;
                            }
                        case "test":
                            {
                                break;
                            }
                        default:
                            {
                                if (UI_running)
                                {
                                    switch (tab_meeting)
                                    {
                                        case 0:
                                            {

                                                if (Dict_vym.ContainsKey(cmd))
                                                {
                                                    TextBox txt = (TextBox)Dict_vym[cmd];
                                                    txt.Text = sup;
                                                    txt.BackColor = Color.White;
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

        private void txt_Command_TextChanged(object sender, EventArgs e)
        {
            string Str = txt_Command.Text.ToLower();
            string cmd = Str;
            int index = 0;
            if (Str.Length > 4)
            {
                index = cmd.IndexOf(" ");
                if (index >= 0)
                {
                    cmd = cmd.Substring(0, index);
                }
            }
            switch (tab_meeting)
            {
                case 0:
                    {
                        if (Dict_vym.ContainsKey(cmd))
                        {
                            Change_Presenter(cmd);
                            TextBox txt = (TextBox)Dict_vym[cmd];
                            txt.BackColor = Color.OrangeRed;
                            if ((cmd != aux_command) && (aux_command != null))
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
                            if ((cmd != aux_command) && (aux_command != null))
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
                            if ((cmd != aux_command) && (aux_command != null))
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
                Presenter(p.Executor);
            }
            else if (cmd.Contains("sm"))
            {
                Presenter(p.Oracle);
            }
            else if (cmd.Contains("nv"))
            {
                Presenter(p.DarkTemplar);
            }
            else if (cmd.Contains("rp"))
            {
                Random rnd = new Random();
                int rnd_pr = rnd.Next(3, 5);
                Presenter((p)rnd_pr);
            }
            else if (cmd.Contains("ac"))
            {
                Random rnd = new Random();
                int rnd_pr = rnd.Next(6, 8);
                Presenter((p)rnd_pr);
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
            Path = Application.StartupPath + "\\\\Programs.xlsx";
            Notify("Running new instance");
            is_new_instance = true;
            //tab_Control.Enabled = true;
            Get_Meetings();
            Week_Handler();
            var autocomplete = new AutoCompleteStringCollection();
            autocomplete.AddRange(Dict_vym.Keys.ToArray());
            txt_Command.AutoCompleteCustomSource = autocomplete;
            UI_running = true;
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
                    Path = openExcel.FileName;
                    Opening_Excel(Path);
                    VyM_Handler(false);
                    Get_Meetings();
                    Week_Handler();
                    if (excel_ready)
                    {
                        excel_ready = false;
                        objBooks.Close(0);
                        objApp.Quit();
                    }
                    UI_running = true;
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
        }

        public void Opening_Excel(string path)
        {
            objApp = new Excel.Application();
            objBooks = (Excel.Workbook)objApp.Workbooks.Open(path, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);

            objSheets = objBooks.Worksheets;
            Sheet_VyM = (Excel.Worksheet)objSheets.get_Item(1);
            range_1 = Sheet_VyM.get_Range("A1", "H137");
            cellValue_1 = (System.Object[,])range_1.get_Value();
            excel_ready = false;

            if ((cellValue_1[53, 1] != null) && (cellValue_1[53, 1].ToString() == "S-140 AGR-Technologies"))
            {
                Notify("File decoded correctly");

                Sheet_RP = (Excel.Worksheet)objSheets.get_Item(2);
                range_2 = Sheet_RP.get_Range("A1", "H70");
                cellValue_2 = (System.Object[,])range_2.get_Value();

                Sheet_AC = (Excel.Worksheet)objSheets.get_Item(3);
                range_3 = Sheet_AC.get_Range("A1", "H70");
                cellValue_3 = (System.Object[,])range_3.get_Value();

                Notify("Opening path: " + path);
                excel_ready = true;
            }
            else
            {
                Warn("Invalid file");
            }
        }

        private void Process_clear()
        {
            Warn("Clear all!");
        }

        /*@ToDo Set font size to (x min.)*/
        public void Set_Font(Excel.Range cell)
        {
            string Str = cell.Text;
            var array = Str.ToCharArray();
            for (int i = Str.Length-1; i >= 0; i--)
            {
                if (array[i] == '(')
                {
                    cell.Characters[i, Str.Length].Font.Size = 9;
                    break;
                }
            }
        }
             
        /*ToDo Reset all information in form*/
        private void Process_restore()
        {
            Notify("Restore info");
        }

        /*Add DateTime of the month meetings in the array*/
        public void Get_Meetings()
        {
            Notify("Getting meetings for month [" + month[m_mes - 1].ToString() + "]");
            int days = DateTime.DaysInMonth(2018, m_mes);
            int i = -1;
            int check = 0;
            int aux_m = 0;
            int aux_y = m_año;
            for (short d = 1; d <= days; d++)
            {
                if (new DateTime(m_año, m_mes, d).DayOfWeek == DayOfWeek.Monday)
                {
                    i++;
                    meetings_days[i, 0] = new DateTime(m_año, m_mes, d);
                    check++;
                }
                if ((new DateTime(m_año, m_mes, d).DayOfWeek == DayOfWeek.Saturday) && (i>=0))
                {                    
                    meetings_days[i, 1] = new DateTime(m_año, m_mes, d);
                    check++;
                }
            }
            /*Handler to check Saturdays in another month*/
            if (check % 2 != 0)
            {
                if ((i == 3) || (i == 4)) //Week 4 or 5
                {
                    if (i == 4)
                    {
                        week_five_exist = true;
                    }
                    else
                    {
                        week_five_exist = false;
                    }
                    aux_m = m_mes + 1;
                    if (aux_m > 12)
                    {
                        aux_m = 1;
                        aux_y++;
                    }
                    for (int d = 1; d <= days; d++)
                    {
                        if ((new DateTime(aux_y, aux_m, d).DayOfWeek == DayOfWeek.Saturday) && (i >= 0))
                        {
                            meetings_days[i, 1] = new DateTime(aux_y, aux_m, d);
                            break;
                        }
                    }
                }
            }
            //ToDo save all dates in objects!
            VyM_mes.Semana1.Fecha = meetings_days[0, 0].ToString("dddd, dd MMMM");
            VyM_mes.Semana2.Fecha = meetings_days[1, 0].ToString("dddd, dd MMMM");
            VyM_mes.Semana3.Fecha = meetings_days[2, 0].ToString("dddd, dd MMMM");
            VyM_mes.Semana4.Fecha = meetings_days[3, 0].ToString("dddd, dd MMMM");
            if (week_five_exist)
            {
                VyM_mes.Semana5.Fecha = meetings_days[4, 0].ToString("dddd, dd MMMM");
            }
            RP_mes.Semana1.Fecha = meetings_days[0, 1].ToString("dddd, dd MMMM");
            RP_mes.Semana2.Fecha = meetings_days[1, 1].ToString("dddd, dd MMMM");
            RP_mes.Semana3.Fecha = meetings_days[2, 1].ToString("dddd, dd MMMM");
            RP_mes.Semana4.Fecha = meetings_days[3, 1].ToString("dddd, dd MMMM");
            if (week_five_exist)
            {
                RP_mes.Semana5.Fecha = meetings_days[4, 1].ToString("dddd, dd MMMM");
            }
        }

        private void Process_save(int save)
        {
            string FileName = meetings_days[0, 0].ToString("MMMM");
            loading = 1;
            Opening_Excel(Path);
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
                objBooks.SaveAs(createfolder + "\\" + FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                Path = createfolder + "\\" + FileName;
                //PDF Implementation TBD
                //objBooks.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Path);
                Notify("Saved path: " + Path);
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
            Notify("Saved date: [" + m_dia.ToString() + "-" + m_mes.ToString() + "-" + m_año.ToString() + "]");
            //Check_time(this, null);
            loading = 100;
        }

        /*--------------------------------------- Meeting handlers ---------------------------------------*/

        public void VyM_Handler(bool save)
        {
            if (save)
            {
                VyM_Save_Week(VyM_mes.Semana1, 1);
                loading += (15/loading_delta);
                VyM_Save_Week(VyM_mes.Semana2, 2);
                loading += (15/loading_delta);
                VyM_Save_Week(VyM_mes.Semana3, 3);
                loading += (15 / loading_delta);
                VyM_Save_Week(VyM_mes.Semana4, 4);
                loading += (15 / loading_delta);
                if (week_five_exist)
                {
                    VyM_Save_Week(VyM_mes.Semana5, 5);
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

        public void VyM_Save_Week(VyM_Sem sem, short num_sem)
        {
            short primary_cell = Get_vym_cell(num_sem);
            Sheet_VyM.Cells[primary_cell, A] = sem.Fecha.ToUpper();
            if (num_sem != Conv_Wk)
            {
                if (sem.Sem_Biblia != null)
                {
                    Sheet_VyM.Cells[primary_cell, D] = sem.Sem_Biblia.ToUpper();
                    Sheet_VyM.Cells[primary_cell, G] = sem.Presidente;
                    Sheet_VyM.Cells[primary_cell + 6, C] = sem.Discurso;
                    Sheet_VyM.Cells[primary_cell + 6, G] = sem.Discurso_A;
                    Sheet_VyM.Cells[primary_cell + 7, G] = sem.Perlas;
                    Sheet_VyM.Cells[primary_cell + 8, G] = sem.Lectura;
                    Sheet_VyM.Cells[primary_cell + 11, C] = sem.SMM1;
                    Sheet_VyM.Cells[primary_cell + 11, G] = sem.SMM1_A;
                    Sheet_VyM.Cells[primary_cell + 12, C] = sem.SMM2;
                    Sheet_VyM.Cells[primary_cell + 12, G] = sem.SMM2_A;
                    Sheet_VyM.Cells[primary_cell + 13, C] = sem.SMM3;
                    Sheet_VyM.Cells[primary_cell + 13, G] = sem.SMM3_A;
                    /*@ToDo Implement SMM4  and SMM4_A*/
                    if ((sem.SMM4 != null) && (sem.SMM4_A != null))
                    {
                        Sheet_VyM.Cells[primary_cell + 14, C] = sem.SMM4;
                        Sheet_VyM.Cells[primary_cell + 14, G] = sem.SMM4_A;
                        primary_cell++;
                    }
                    Sheet_VyM.Cells[primary_cell + 17, C] = sem.NVC1;
                    Sheet_VyM.Cells[primary_cell + 17, G] = sem.NVC1_A;
                    Sheet_VyM.Cells[primary_cell + 18, C] = sem.NVC2;
                    Sheet_VyM.Cells[primary_cell + 18, G] = sem.NVC2_A;
                    Sheet_VyM.Cells[primary_cell + 19, G] = sem.Libro_A;
                    Sheet_VyM.Cells[primary_cell + 20, G] = sem.Libro_L;
                    Sheet_VyM.Cells[primary_cell + 22, G] = sem.Oracion;
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
            Get_month_from_Excel(cellValue_1[primary_cell, A]);
            sem.Sem_Biblia  = Check_null_string(cellValue_1[primary_cell, D]);
            sem.Presidente  = Check_null_string(cellValue_1[primary_cell, G]);
            sem.Discurso    = Check_null_string(cellValue_1[primary_cell + 6, C]);
            sem.Discurso_A  = Check_null_string(cellValue_1[primary_cell + 6, G]);
            sem.Perlas      = Check_null_string(cellValue_1[primary_cell + 7, G]);
            sem.Lectura     = Check_null_string(cellValue_1[primary_cell + 8, G]);
            sem.SMM1        = Check_null_string(cellValue_1[primary_cell + 11, C]);
            sem.SMM1_A      = Check_null_string(cellValue_1[primary_cell + 11, G]);
            sem.SMM2        = Check_null_string(cellValue_1[primary_cell + 12, C]);
            sem.SMM2_A      = Check_null_string(cellValue_1[primary_cell + 12, G]);
            sem.SMM3        = Check_null_string(cellValue_1[primary_cell + 13, C]);
            sem.SMM3_A      = Check_null_string(cellValue_1[primary_cell + 13, G]);
            /*@ToDo Implement SMM4  and SMM4_A*/
            sem.NVC1        = Check_null_string(cellValue_1[primary_cell + 17, C]);
            sem.NVC1_A      = Check_null_string(cellValue_1[primary_cell + 17, G]);
            sem.NVC2        = Check_null_string(cellValue_1[primary_cell + 18, C]);
            sem.NVC2_A      = Check_null_string(cellValue_1[primary_cell + 18, G]);
            sem.Libro_A     = Check_null_string(cellValue_1[primary_cell + 19, G]);
            sem.Libro_L     = Check_null_string(cellValue_1[primary_cell + 20, G]);
            sem.Oracion     = Check_null_string(cellValue_1[primary_cell + 22, G]);

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
                RP_Save_Week(RP_mes.Semana1, 1);
                loading += (15 / loading_delta);
                RP_Save_Week(RP_mes.Semana2, 2);
                loading += (15 / loading_delta);
                RP_Save_Week(RP_mes.Semana3, 3);
                loading += (15 / loading_delta);
                RP_Save_Week(RP_mes.Semana4, 4);
                loading += (15 / loading_delta);
                if (week_five_exist)
                {
                    RP_Save_Week(RP_mes.Semana5, 5);
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

        public void RP_Save_Week(RP_Sem sem, short num_sem)
        {
            short primary_cell = Get_rp_cell(num_sem);
            Sheet_RP.Cells[primary_cell, C] = sem.Fecha.ToUpper();
            if (num_sem != Conv_Wk)
            {
                if (sem.Presidente != null)
                {
                    Sheet_RP.Cells[primary_cell + 1, H] = sem.Presidente;
                    Sheet_RP.Cells[primary_cell + 2, D] = sem.Titulo;
                    Sheet_RP.Cells[primary_cell + 2, H] = sem.Discursante;
                    Sheet_RP.Cells[primary_cell + 3, E] = sem.Congregacion;
                    Sheet_RP.Cells[primary_cell + 6, D] = sem.Titulo_Atalaya;
                    Sheet_RP.Cells[primary_cell + 5, H] = sem.Conductor;
                    Sheet_RP.Cells[primary_cell + 7, H] = sem.Lector;
                    Sheet_RP.Cells[primary_cell + 8, H] = sem.Oracion;
                    Sheet_RP.Cells[primary_cell + 10, C] = sem.Discu_Sal;
                    Sheet_RP.Cells[primary_cell + 10, E] = sem.Ttl_Sal;
                    Sheet_RP.Cells[primary_cell + 10, H] = sem.Cong_Sal;
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
            sem.Presidente      = Check_null_string(cellValue_2[primary_cell + 1, H]);
            sem.Titulo          = Check_null_string(cellValue_2[primary_cell + 2, D]);
            sem.Discursante     = Check_null_string(cellValue_2[primary_cell + 2, H]);
            sem.Congregacion    = Check_null_string(cellValue_2[primary_cell + 3, E]);
            sem.Titulo_Atalaya  = Check_null_string(cellValue_2[primary_cell + 6, D]);
            sem.Conductor       = Check_null_string(cellValue_2[primary_cell + 5, H]);
            sem.Lector          = Check_null_string(cellValue_2[primary_cell + 7, H]);
            sem.Oracion         = Check_null_string(cellValue_2[primary_cell + 8, H]);
            sem.Discu_Sal       = Check_null_string(cellValue_2[primary_cell + 10, C]);
            sem.Ttl_Sal         = Check_null_string(cellValue_2[primary_cell + 10, E]);
            sem.Cong_Sal        = Check_null_string(cellValue_2[primary_cell + 10, H]);
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
                AC_Save_Week(AC_mes.Semana1, 1);
                loading += (15 / loading_delta);
                AC_Save_Week(AC_mes.Semana2, 2);
                loading += (15 / loading_delta);
                AC_Save_Week(AC_mes.Semana3, 3);
                loading += (15 / loading_delta);
                AC_Save_Week(AC_mes.Semana4, 4);
                loading += (15 / loading_delta);
                if (week_five_exist)
                {
                    AC_Save_Week(AC_mes.Semana5, 5);
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

        public void AC_Save_Week(AC_Sem sem, short num_sem)
        {
            short primary_cell = Get_ac_cell(num_sem);
            Sheet_AC.Cells[primary_cell, B] = meetings_days[num_sem - 1, 0].ToString("dddd, dd MMMM");
            Sheet_AC.Cells[primary_cell, D] = meetings_days[num_sem - 1, 1].ToString("dddd, dd MMMM");
            if (num_sem != Conv_Wk)
            {
                if (sem.Vym_Cap != null)
                {
                    Sheet_AC.Cells[primary_cell + 1, C] = sem.Vym_Cap;
                    Sheet_AC.Cells[primary_cell + 1, E] = sem.Rp_Cap;
                    Sheet_AC.Cells[primary_cell + 2, A] = sem.Cap_Aseo;
                    Sheet_AC.Cells[primary_cell + 2, C] = sem.Vym_Der;
                    Sheet_AC.Cells[primary_cell + 2, E] = sem.Rp_Der;
                    Sheet_AC.Cells[primary_cell + 3, C] = sem.Vym_Izq;
                    Sheet_AC.Cells[primary_cell + 3, E] = sem.Rp_Cap;
                }
            }
            else
            {
                Convention_Handler(3);
            }
        }

        public void AC_Read_Week(short num_sem)
        {
            AC_Sem sem = new AC_Sem();
            short primary_cell = Get_ac_cell(num_sem);
            sem.Vym_Cap     = Check_null_string(cellValue_3[primary_cell + 1, C]);
            sem.Rp_Cap      = Check_null_string(cellValue_3[primary_cell + 1, E]);
            sem.Cap_Aseo    = Check_null_string(cellValue_3[primary_cell + 2, A]);
            sem.Vym_Der     = Check_null_string(cellValue_3[primary_cell + 2, C]);
            sem.Rp_Der      = Check_null_string(cellValue_3[primary_cell + 2, E]);
            sem.Vym_Izq     = Check_null_string(cellValue_3[primary_cell + 3, C]);
            sem.Rp_Izq      = Check_null_string(cellValue_3[primary_cell + 3, E]);
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

        /*ToDo Handler to modify Convention week*/
        public void Convention_Handler(short program)
        {
            Excel.Range range;
            string a = "A", g = "G", e = "E", h = "H";
            switch (program)
            {
                case 1: //VyM
                    {
                        short cell = Get_vym_cell(Conv_Wk);
                        range = Sheet_VyM.get_Range(a+(cell+2).ToString(), g+(cell+22).ToString());
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
                        range = Sheet_RP.get_Range(a + (cell + 1).ToString(), h + (cell + 15).ToString());
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
                        range = Sheet_AC.get_Range(a + (cell + 1).ToString(), e + (cell + 3).ToString());
                        range.Cells.Clear();
                        range = Sheet_AC.get_Range(a + (cell + 1).ToString(), e + (cell + 3).ToString());
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

        /*--------------------------------------- Auxiliar functions to set/read strings ---------------------------------------*/

        private void Get_month_from_Excel(object cellvalue)
        {
            if (cellvalue != null)
            {
                bool month_found = false;
                for (int i = 0; i <= month.Length - 1; i++)
                {
                    if (cellvalue.ToString().ToLower().Contains(month[i]))
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
                        cell = 3;
                        break;
                    }
                case 1:
                    {
                        cell = 28;
                        break;
                    }
                case 2:
                    {
                        cell = 56;
                        break;
                    }
                case 3:
                    {
                        cell = 81;
                        break;
                    }
                case 4:
                    {
                        cell = 109;
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
                        cell = 4;
                        break;
                    }
                case 1:
                    {
                        cell = 17;
                        break;
                    }
                case 2:
                    {
                        cell = 30;
                        break;
                    }
                case 3:
                    {
                        cell = 43;
                        break;
                    }
                case 4:
                    {
                        cell = 59;
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
                        cell = 5;
                        break;
                    }
                case 1:
                    {
                        cell = 13;
                        break;
                    }
                case 2:
                    {
                        cell = 21;
                        break;
                    }
                case 3:
                    {
                        cell = 29;
                        break;
                    }
                case 4:
                    {
                        cell = 37;
                        break;
                    }
            }
            return cell;
        }
        private void General_Info_Enter(object sender, EventArgs e)
        {
            Presenter(p.Executor);
            Notify("Overview");
        }

        private void Tesoros_Biblia_Enter(object sender, EventArgs e)
        {
            Presenter(p.DarkTemplar);
            Notify("Section 'Tesoros de la Biblia'");
        }

        private void Seamos_Maestros_Enter(object sender, EventArgs e)
        {
            Presenter(p.Selendis);
            Notify("Section 'Seamos Mejores Maestros'");
        }

        private void Nuestra_Vida_Enter(object sender, EventArgs e)
        {
            //Presenter(p.Fenix);
            Notify("Section 'Nuestra Vida Cristiana'");
        }

        private void tab_Control_SelectedIndexChanged(object sender, EventArgs e)
        {
            var autocomplete = new AutoCompleteStringCollection();
            if (tab_meeting != tab_Control.SelectedIndex)
            {
                Pre_save_info();
            }
            if (tab_Control.SelectedIndex == 0)
            {
                Presenter(p.Executor);
                tab_meeting = 0;
                autocomplete.AddRange(Dict_vym.Keys.ToArray());
                Notify("\"Vida y Ministerio\" meeting");
            }
            else if (tab_Control.SelectedIndex == 1)
            {
                Random rnd = new Random();
                int rnd_pr = rnd.Next(3, 5);
                Presenter((p)rnd_pr);
                tab_meeting = 1;
                autocomplete.AddRange(Dict_rp.Keys.ToArray());
                Notify("\"Reunion Publica y Analisis de La Atalaya\" meeting");
                
            }
            else
            {
                Random rnd = new Random();
                int rnd_pr = rnd.Next(6, 8);
                Presenter((p)rnd_pr);
                tab_meeting = 2;
                autocomplete.AddRange(Dict_ac.Keys.ToArray());
                Notify("\"Acomodadores\" Section");
            }
            Week_Handler();
            txt_Command.AutoCompleteCustomSource = autocomplete;
        }

        private void Txt_TextChanged(object sender, EventArgs e)
        {
            Save_time_from_string();
        }

        private void Save_time_from_string()
        {
            DateTime Aux_dateTime = new DateTime(2018, 1, 1, 7, 00, 00);

            time_0.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(5);
            time_1.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(3);
            time_2.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(10);
            time_3.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(8);
            time_4.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(5 + 1); //adjusting to real time
            time_5.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(txt_SMM1.Text) + 1); //adjusting to real time
            time_6.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(txt_SMM2.Text) + 1); //adjusting to real time
            time_7.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(txt_SMM3.Text) + 1); //adjusting to real time
            if ((txt_SMM4.Text == null) || (txt_SMM4.Text == ""))
            {
                time7_5.Text = " ";
            }
            else
            {
                time7_5.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(txt_SMM4.Text) + 1); //adjusting to real time
            }      
            time_8.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(3);
            time_9.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(txt_NVC1.Text));
            if ((txt_NVC2.Text == null) || (txt_NVC2.Text == ""))
            {
                time_10.Text = " ";
            }
            else
            {
                time_10.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(txt_NVC2.Text));
            }
            Aux_dateTime = Aux_dateTime.AddMinutes(1); //adjusting to real time
            time_11.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(30);
            time_12.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
            Aux_dateTime = Aux_dateTime.AddMinutes(3);
            time_13.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
        }

        public int Get_time_from_string(string Str)
        {
            int time = 0;
            if (Str != null)
            {
                Str = Str.ToLower();
                string min = "mins.";
                string number = "";
                var array = Str.ToCharArray();
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
            int checksum_aux = 0;
            checksum_aux = m_año + m_mes + m_dia;
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

        /*private void Open_DB()
        {
            if (!DB_form_show)
            {
                DB_form_show = true;
                //timer_Form2.Enabled = true;
                DB_Form.Show();
            }
        }

        private void timer_Form2_Tick(object sender, EventArgs e)
        {
            if (message_form2 != null)
            {
                Notify(message_form2);
                message_form2 = null;
            }
            if (!DB_form_show)
            {
                timer_Form2.Enabled = false;
            }
        }*/

        /*--------------------------------------- Week Handlers  ---------------------------------------*/

        /*Function so set local variables' info into form*/
        public void Week_Handler()
        {
            int lun = 0;
            lbl_Week.Text = "Semana: " + m_semana.ToString();
            if (Conv_Wk == m_semana)
            {
                Warn("Week [" + m_semana.ToString() + "] selected as Convention Week");
            }
            switch (tab_meeting)
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
                        lun = 1;
                        break;
                    }
                case 2:
                    {
                        txt_Aseo_1.Text = AC_mes.Semana1.Cap_Aseo;
                        txt_Cap_L_1.Text = AC_mes.Semana1.Vym_Cap;
                        txt_AC1_L_1.Text = AC_mes.Semana1.Vym_Izq;
                        txt_AC2_L_1.Text = AC_mes.Semana1.Vym_Der;
                        txt_Cap_S_1.Text = AC_mes.Semana1.Rp_Cap;
                        txt_AC1_S_1.Text = AC_mes.Semana1.Rp_Izq;
                        txt_AC2_S_1.Text = AC_mes.Semana1.Rp_Der;

                        txt_Aseo_2.Text = AC_mes.Semana2.Cap_Aseo;
                        txt_Cap_L_2.Text = AC_mes.Semana2.Vym_Cap;
                        txt_AC1_L_2.Text = AC_mes.Semana2.Vym_Izq;
                        txt_AC2_L_2.Text = AC_mes.Semana2.Vym_Der;
                        txt_Cap_S_2.Text = AC_mes.Semana2.Rp_Cap;
                        txt_AC1_S_2.Text = AC_mes.Semana2.Rp_Izq;
                        txt_AC2_S_2.Text = AC_mes.Semana2.Rp_Der;

                        txt_Aseo_3.Text = AC_mes.Semana3.Cap_Aseo;
                        txt_Cap_L_3.Text = AC_mes.Semana3.Vym_Cap;
                        txt_AC1_L_3.Text = AC_mes.Semana3.Vym_Izq;
                        txt_AC2_L_3.Text = AC_mes.Semana3.Vym_Der;
                        txt_Cap_S_3.Text = AC_mes.Semana3.Rp_Cap;
                        txt_AC1_S_3.Text = AC_mes.Semana3.Rp_Izq;
                        txt_AC2_S_3.Text = AC_mes.Semana3.Rp_Der;

                        txt_Aseo_4.Text = AC_mes.Semana4.Cap_Aseo;
                        txt_Cap_L_4.Text = AC_mes.Semana4.Vym_Cap;
                        txt_AC1_L_4.Text = AC_mes.Semana4.Vym_Izq;
                        txt_AC2_L_4.Text = AC_mes.Semana4.Vym_Der;
                        txt_Cap_S_4.Text = AC_mes.Semana4.Rp_Cap;
                        txt_AC1_S_4.Text = AC_mes.Semana4.Rp_Izq;
                        txt_AC2_S_4.Text = AC_mes.Semana4.Rp_Der;

                        txt_Aseo_5.Text = AC_mes.Semana5.Cap_Aseo;
                        txt_Cap_L_5.Text = AC_mes.Semana5.Vym_Cap;
                        txt_AC1_L_5.Text = AC_mes.Semana5.Vym_Izq;
                        txt_AC2_L_5.Text = AC_mes.Semana5.Vym_Der;
                        txt_Cap_S_5.Text = AC_mes.Semana5.Rp_Cap;
                        txt_AC1_S_5.Text = AC_mes.Semana5.Rp_Izq;
                        txt_AC2_S_5.Text = AC_mes.Semana5.Rp_Der;
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
            txt_Date.Text    = sem.Sem_Biblia;
            txt_Pres.Text    = sem.Presidente;
            txt_TdlB_1.Text  = sem.Discurso;
            txt_TdlB_A1.Text = sem.Discurso_A;
            txt_TdlB_A2.Text = sem.Perlas;
            txt_TdlB_A3.Text = sem.Lectura;
            txt_SMM1.Text    = sem.SMM1;
            txt_SMM_A1.Text  = sem.SMM1_A;
            txt_SMM2.Text    = sem.SMM2;
            txt_SMM_A2.Text  = sem.SMM2_A;
            txt_SMM3.Text    = sem.SMM3;
            txt_SMM_A3.Text  = sem.SMM3_A;
            txt_NVC1.Text    = sem.NVC1;
            txt_NVC_A1.Text  = sem.NVC1_A;
            txt_NVC2.Text    = sem.NVC2;
            txt_NVC2.Text    = sem.NVC2_A;
            txt_NVC_A3.Text  = sem.Libro_A;
            txt_NVC_A4.Text  = sem.Libro_L;
            txt_Ora2VyM.Text = sem.Oracion;
        }

        public void RP_Week_Handler(RP_Sem sem)
        {
            txt_RP_Speech.Text  = sem.Titulo;
            txt_PresRP.Text     = sem.Presidente;
            txt_RP_Disc.Text    = sem.Congregacion;
            txt_RP_Cong.Text    = sem.Discursante;
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
            switch (tab_meeting)
            {
                case 0:
                    {
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    VyM_mes.Semana1 = VyM_Set_Week();
                                    break;
                                }
                            case 2:
                                {
                                    VyM_mes.Semana2 = VyM_Set_Week();
                                    break;
                                }
                            case 3:
                                {
                                    VyM_mes.Semana3 = VyM_Set_Week();
                                    break;
                                }
                            case 4:
                                {
                                    VyM_mes.Semana4 = VyM_Set_Week();
                                    break;
                                }
                            case 5:
                                {
                                    VyM_mes.Semana5 = VyM_Set_Week();
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
                                    RP_mes.Semana1 = RP_Set_Week();
                                    break;
                                }
                            case 2:
                                {
                                    RP_mes.Semana2 = RP_Set_Week();
                                    break;
                                }
                            case 3:
                                {
                                    RP_mes.Semana3 = RP_Set_Week();
                                    break;
                                }
                            case 4:
                                {
                                    RP_mes.Semana4 = RP_Set_Week();
                                    break;
                                }
                            case 5:
                                {
                                    RP_mes.Semana5 = RP_Set_Week();
                                    break;
                                }
                        }
                        break;
                    }
                case 2:
                    {
                        AC_mes.Semana1.Cap_Aseo = txt_Aseo_1.Text;
                        AC_mes.Semana1.Vym_Cap = txt_Cap_L_1.Text;
                        AC_mes.Semana1.Vym_Izq = txt_AC1_L_1.Text;
                        AC_mes.Semana1.Vym_Der = txt_AC2_L_1.Text;
                        AC_mes.Semana1.Rp_Cap = txt_Cap_S_1.Text;
                        AC_mes.Semana1.Rp_Izq = txt_AC1_S_1.Text;
                        AC_mes.Semana1.Rp_Der = txt_AC2_S_1.Text;

                        AC_mes.Semana2.Cap_Aseo = txt_Aseo_2.Text;
                        AC_mes.Semana2.Vym_Cap = txt_Cap_L_2.Text;
                        AC_mes.Semana2.Vym_Izq = txt_AC1_L_2.Text;
                        AC_mes.Semana2.Vym_Der = txt_AC2_L_2.Text;
                        AC_mes.Semana2.Rp_Cap = txt_Cap_S_2.Text;
                        AC_mes.Semana2.Rp_Izq = txt_AC1_S_2.Text;
                        AC_mes.Semana2.Rp_Der = txt_AC2_S_2.Text;

                        AC_mes.Semana3.Cap_Aseo = txt_Aseo_3.Text;
                        AC_mes.Semana3.Vym_Cap = txt_Cap_L_3.Text;
                        AC_mes.Semana3.Vym_Izq = txt_AC1_L_3.Text;
                        AC_mes.Semana3.Vym_Der = txt_AC2_L_3.Text;
                        AC_mes.Semana3.Rp_Cap = txt_Cap_S_3.Text;
                        AC_mes.Semana3.Rp_Izq = txt_AC1_S_3.Text;
                        AC_mes.Semana3.Rp_Der = txt_AC2_S_3.Text;

                        AC_mes.Semana4.Cap_Aseo = txt_Aseo_4.Text;
                        AC_mes.Semana4.Vym_Cap = txt_Cap_L_4.Text;
                        AC_mes.Semana4.Vym_Izq = txt_AC1_L_4.Text;
                        AC_mes.Semana4.Vym_Der = txt_AC2_L_4.Text;
                        AC_mes.Semana4.Rp_Cap = txt_Cap_S_4.Text;
                        AC_mes.Semana4.Rp_Izq = txt_AC1_S_4.Text;
                        AC_mes.Semana4.Rp_Der = txt_AC2_S_4.Text;

                        AC_mes.Semana5.Cap_Aseo = txt_Aseo_5.Text;
                        AC_mes.Semana5.Vym_Cap = txt_Cap_L_5.Text;
                        AC_mes.Semana5.Vym_Izq = txt_AC1_L_5.Text;
                        AC_mes.Semana5.Vym_Der = txt_AC2_L_5.Text;
                        AC_mes.Semana5.Rp_Cap = txt_Cap_S_5.Text;
                        AC_mes.Semana5.Rp_Izq = txt_AC1_S_5.Text;
                        AC_mes.Semana5.Rp_Der = txt_AC2_S_5.Text;
                        break;
                    }
            }
        }

        public VyM_Sem VyM_Set_Week()
        {
            VyM_Sem sem = new VyM_Sem();
            sem.Fecha       = lbl_DateVyM.Text;
            sem.Sem_Biblia  = txt_Date.Text;
            sem.Presidente  = txt_Pres.Text;
            sem.Discurso    = txt_TdlB_1.Text;
            sem.Discurso_A  = txt_TdlB_A1.Text;
            sem.Perlas      = txt_TdlB_A2.Text;
            sem.Lectura     = txt_TdlB_A3.Text;
            sem.SMM1        = txt_SMM1.Text;
            sem.SMM1_A      = txt_SMM_A1.Text;
            sem.SMM2        = txt_SMM2.Text;
            sem.SMM2_A      = txt_SMM_A2.Text;
            sem.SMM3        = txt_SMM3.Text;
            sem.SMM3_A      = txt_SMM_A3.Text;
            sem.NVC1        = txt_NVC1.Text;
            sem.NVC1_A      = txt_NVC_A1.Text;
            sem.NVC2        = txt_NVC2.Text;
            sem.NVC2_A      = txt_NVC2.Text;
            sem.Libro_A     = txt_NVC_A3.Text;
            sem.Libro_L     = txt_NVC_A4.Text;
            sem.Oracion     = txt_Ora2VyM.Text;
            return sem;
        }

        public RP_Sem RP_Set_Week()
        {
            RP_Sem sem = new RP_Sem();
            sem.Fecha           = lbl_DateRP.Text;
            sem.Titulo          = txt_RP_Speech.Text;
            sem.Presidente      = txt_PresRP.Text;
            sem.Congregacion    = txt_RP_Disc.Text;
            sem.Discursante     = txt_RP_Cong.Text;
            sem.Titulo_Atalaya  = txt_Title_Atly.Text;
            sem.Conductor       = txt_Con_Atly.Text;
            sem.Lector          = txt_Lect_Atly.Text;
            sem.Oracion         = txt_OraRP.Text;
            sem.Discu_Sal       = txt_Sal_Disc.Text;
            sem.Ttl_Sal         = txt_Sal_Title.Text;
            sem.Cong_Sal        = txt_Sal_Cong.Text;
            return sem;
        }
    }
}
