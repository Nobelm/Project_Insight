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
            Fenix,
            Selendis,
            Oracle,
            DarkTemplar,
            HunterKiller,
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
        private Excel.Worksheet Sheet_PA;
        private Excel.Range range_1;
        private Excel.Range range_2;
        private Excel.Range range_3;
        public static bool excel_ready = false;
        private DateTime start_time = new DateTime(2018, 1, 1, 7, 00, 00);
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
        public static bool DB_form_show = false;
        public static string message_form2 = null;
        public static bool pending_refresh_DB = false;
        public static int m_dia = 1;
        public static int m_mes = 1;
        public static int m_año = DateTime.Today.Year;
        public static int m_semana = 1;
        public static DateTime[,] meetings_days = new DateTime[5, 2];
        public static string[] guard_cbx_names = new string[10];
        public static int date_checksum = 0;
        public static string[] Command_history = new string[10];
        //public static string[] Command_input = new string[] {"op_xlsx", "op_db", "sv", "clc", "rst", "mnth", "wk", "autofill", "exit"};
        //public static string[] Command_input = new string[] {"new", "open", "save", "exit", "month"};
        public static string[] month = new string[] { "ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic" };
        public static int command_iterator = 0;
        DB_Form DB_Form = new DB_Form();
        public static string Path = "";
        public static bool is_new_instance = false;
        public static VyM_Mes VyM_mes = new VyM_Mes();
        public static RP_Mes RP_mes = new RP_Mes();
        public static AC_Mes AC_mes = new AC_Mes();
        //public static string[] VyM_Names = new string[] { "ig_1", "ig_2", "tb_1", "tb_2", "tb_3", "tb_4", "sm_1", "sm_2", "sm_3", "sm_4", "sm_5", "sm_6",
        //   "nv_1", "nv_2", "nv3", "nv_4", "nv_5", "nv_6", "nv_7" };
        //public static string[] RP_Names = new string[] { "rp_101", "rp_12", "rp_13", "rp_14", "rp_15", "rp_16", "rp_17", "rp_18", "rp_19", "rp_s1", "rp_s1", };
        IDictionary<string, object> Dict_vym = new Dictionary<string, object>();
        IDictionary<string, object> Dict_rp = new Dictionary<string, object>();
        IDictionary<string, object> Dict_ac = new Dictionary<string, object>();
        public static int tab_meeting = 0;
        public static string aux_command;
        public static bool selected_txt = false;
        public static int loading_delta = 1;
        public static int loading = 0;
        public static bool is_loading = false;



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
            Dict_vym.Add("nv_11", txt_NVC1);
            Dict_vym.Add("nv_12", txt_NVC_A1);
            Dict_vym.Add("nv_21", txt_NVC2);
            Dict_vym.Add("nv_22", txt_NVC_A2);
            Dict_vym.Add("nv_30", txt_NVC_A3);
            Dict_vym.Add("nv_40", txt_NVC_A4);
            Dict_vym.Add("nv_50", txt_Ora2VyM);

            Dict_rp.Add("rp_00", txt_DateRP);
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
                    case p.Fenix:
                        {
                            picPresenter.Image = Project_Insight.Properties.Resources.Fenix;
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
                log_txtBx.SelectionColor = Color.White;
                log_txtBx.AppendText("L-" + lineNumber + ": ");
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
            if (busy_trace)
            {
                await Task.Delay(1000);
            }
            busy_trace = true;
            log_txtBx.SelectionColor = Color.White;
            log_txtBx.AppendText("\nLoading:  ");
            aux = log_txtBx.Text;
            log_txtBx.Text = "";
            log_txtBx.SelectionStart = log_txtBx.Text.Length;
            log_txtBx.ScrollToCaret();
            while (is_loading)
            {
                aux += loading.ToString() + " ....... \\";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                log_txtBx.ScrollToCaret();
                await Task.Delay(delay);
                aux = aux.Substring(0, aux.Length - 10 - loading.ToString().Length);
                aux += loading.ToString() + " ....... |";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                log_txtBx.ScrollToCaret();
                await Task.Delay(delay);
                aux = aux.Substring(0, aux.Length - 10 - loading.ToString().Length);
                aux += loading.ToString() + " ....... /";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                log_txtBx.ScrollToCaret();
                await Task.Delay(delay);
                aux = aux.Substring(0, aux.Length - 10 - loading.ToString().Length);
                aux += loading.ToString() + " ....... -";
                log_txtBx.Text = aux;
                log_txtBx.SelectionStart = log_txtBx.Text.Length;
                log_txtBx.ScrollToCaret();
                await Task.Delay(delay);
                aux = aux.Substring(0, aux.Length - 10 - loading.ToString().Length);
            }
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
                    case 3:
                        {
                            data += "3";
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
                }/*
                else
                {
                    Command(str_stack[0], int_stack[0]);
                }*/
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
                                if (int.TryParse(sup, out int sv))
                                {
                                    Thread Save_thread = new Thread(() => Process_save(sv));
                                    Save_thread.Start();
                                    Loading_Trace();
                                }
                                break;
                            }
                        case "tab":
                            {
                                if (int.TryParse(sup, out int tab))
                                {
                                    tab_Control.SelectedIndex = tab;
                                }
                                break;
                            }
                        case "week":
                            {
                                if (int.TryParse(sup, out int wk))
                                {
                                    if ((wk != m_semana) && (wk > 0))
                                    {
                                        Pre_save_info();
                                        m_semana = wk;
                                        Week_Handler();
                                    }
                                }//else if x3
                                //Loading_Trace();
                                break;
                            }
                        case "stop":
                            {
                                //is_loading = false;
                                break;
                            }
                        default:
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
                                                //Pre_save_info();
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
                                                //Pre_save_info();
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
                                                //Pre_save_info();
                                            }
                                            break;
                                        }
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
            tab_Control.Enabled = true;
            Get_Meetings();
            Week_Handler();
            var autocomplete = new AutoCompleteStringCollection();
            autocomplete.AddRange(Dict_vym.Keys.ToArray());
            txt_Command.AutoCompleteCustomSource = autocomplete;
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

                Sheet_PA = (Excel.Worksheet)objSheets.get_Item(3);
                range_3 = Sheet_PA.get_Range("A1", "H70");
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

        public string Check_null(object sender)
        {
            string Str = "";
            TextBox txt_bx = (TextBox)sender;
            if ((txt_bx.Text != null) && (txt_bx.Text != " "))
            {
                txt_bx.BackColor = Color.White;
            }
            else
            {
                txt_bx.BackColor = Color.LightCoral;
                Warn("Empty field in [" + txt_bx.Name.ToString() + "]");
            }
            Str = txt_bx.Text;
            return Str;
        }


        public void Get_index_time(Excel.Range cell)
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
            for (int d = 1; d <= days; d++)
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
                if (i == 4)
                {
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
        }

        private void Process_save(int save)
        {
            is_loading = true;
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
                RP_Handler(false);
                FileName += "_RP";
            }
            if ((save == 3) || (save == 4))
            {
                AC_Handler(false);
                FileName += "_AC";
            }
            Notify("FileName: " + FileName);
            pending_refresh_DB = true;
            //Notify(Application.StartupPath + FileName + ".xlsx");
            loading = 80;
            if (is_new_instance)
            {
                //objBooks.SaveAs(@"c:\test2.xlsx");
                //objBooks.SaveAs(Application.StartupPath + FileName + ".xlsx");
                string createfolder = "c:\\Project_Insight";
                System.IO.Directory.CreateDirectory(createfolder);
                objBooks.SaveAs(createfolder + "\\" + FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                Path = createfolder + "\\" + FileName;
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
            loading = 100;
            Notify("Saved file for JW Meetings" + ", Week [" + m_semana.ToString() + "]");
            Notify("Saved date: [" + m_dia.ToString() + "-" + m_mes.ToString() + "-" + m_año.ToString() + "]");
            //Check_time(this, null);
            is_loading = false;
        }

        /*--------------------------------------- Meeting handlers ---------------------------------------*/
        public void VyM_Handler(bool save)
        {
            //Notify((read ? "Reading": "Saving") + " VyM meeting");
            if (save)
            {
                int primary_cell = Get_cell();
                Sheet_VyM.Cells[primary_cell, A] = VyM_mes.Semana1.Fecha.ToUpper();
                Sheet_VyM.Cells[primary_cell, D] = VyM_mes.Semana1.Sem_Biblia.ToUpper();
                Sheet_VyM.Cells[primary_cell, G] = VyM_mes.Semana1.Presidente;
                Sheet_VyM.Cells[primary_cell + 6, C] = VyM_mes.Semana1.Discurso;
                Sheet_VyM.Cells[primary_cell + 6, G] = VyM_mes.Semana1.Discurso_A;
                Sheet_VyM.Cells[primary_cell + 7, G] = VyM_mes.Semana1.Perlas;
                Sheet_VyM.Cells[primary_cell + 8, G] = VyM_mes.Semana1.Lectura;
                Sheet_VyM.Cells[primary_cell + 11, C] = VyM_mes.Semana1.SMM1;
                Sheet_VyM.Cells[primary_cell + 11, G] = VyM_mes.Semana1.SMM1_A;
                Sheet_VyM.Cells[primary_cell + 12, C] = VyM_mes.Semana1.SMM2;
                Sheet_VyM.Cells[primary_cell + 12, G] = VyM_mes.Semana1.SMM2_A;
                Sheet_VyM.Cells[primary_cell + 13, C] = VyM_mes.Semana1.SMM3;
                Sheet_VyM.Cells[primary_cell + 13, G] = VyM_mes.Semana1.SMM3_A;
                Sheet_VyM.Cells[primary_cell + 17, C] = VyM_mes.Semana1.NVC1;
                Sheet_VyM.Cells[primary_cell + 17, G] = VyM_mes.Semana1.NVC1_A;
                Sheet_VyM.Cells[primary_cell + 18, C] = VyM_mes.Semana1.NVC2;
                Sheet_VyM.Cells[primary_cell + 18, G] = VyM_mes.Semana1.NVC2_A;
                Sheet_VyM.Cells[primary_cell + 19, G] = VyM_mes.Semana1.Libro_A;
                Sheet_VyM.Cells[primary_cell + 20, G] = VyM_mes.Semana1.Libro_L;
                Sheet_VyM.Cells[primary_cell + 22, G] = VyM_mes.Semana1.Oracion;

                loading += (15/loading_delta);

                //Sheet_VyM.Cells[28, A] = VyM_mes.Semana2.Fecha.ToUpper() + " | " + VyM_mes.Semana2.Sem_Biblia.ToUpper();
                //Sheet_VyM.Cells[28, G] = VyM_mes.Semana2.Presidente;
                /*txt_TdlB_1.Text = VyM_mes.Semana2.Discurso;
                txt_TdlB_A1.Text = VyM_mes.Semana2.Discurso_A;
                txt_TdlB_A2.Text = VyM_mes.Semana2.Perlas;
                txt_TdlB_A3.Text = VyM_mes.Semana2.Lectura;
                txt_SMM1.Text = VyM_mes.Semana2.SMM1;
                txt_SMM_A1.Text = VyM_mes.Semana2.SMM1_A;
                txt_SMM2.Text = VyM_mes.Semana2.SMM2;
                txt_SMM_A2.Text = VyM_mes.Semana2.SMM2_A;
                txt_SMM3.Text = VyM_mes.Semana2.SMM3;
                txt_SMM_A3.Text = VyM_mes.Semana2.SMM3_A;
                txt_NVC1.Text = VyM_mes.Semana2.NVC1;
                txt_NVC_A1.Text = VyM_mes.Semana2.NVC1_A;
                txt_NVC2.Text = VyM_mes.Semana2.NVC2;
                txt_NVC2.Text = VyM_mes.Semana2.NVC2_A;
                txt_NVC_A3.Text = VyM_mes.Semana2.Libro_A;
                txt_NVC_A4.Text = VyM_mes.Semana2.Libro_L;
                txt_Ora2VyM.Text = VyM_mes.Semana2.Oracion;*/
            
                loading += (15/loading_delta);

                //lbl_Date.Text = VyM_mes.Semana3.Fecha;
                /*txt_Date.Text = VyM_mes.Semana3.Sem_Biblia;
                txt_Pres.Text = VyM_mes.Semana3.Presidente;
                txt_TdlB_1.Text = VyM_mes.Semana3.Discurso;
                txt_TdlB_A1.Text = VyM_mes.Semana3.Discurso_A;
                txt_TdlB_A2.Text = VyM_mes.Semana3.Perlas;
                txt_TdlB_A3.Text = VyM_mes.Semana3.Lectura;
                txt_SMM1.Text = VyM_mes.Semana3.SMM1;
                txt_SMM_A1.Text = VyM_mes.Semana3.SMM1_A;
                txt_SMM2.Text = VyM_mes.Semana3.SMM2;
                txt_SMM_A2.Text = VyM_mes.Semana3.SMM2_A;
                txt_SMM3.Text = VyM_mes.Semana3.SMM3;
                txt_SMM_A3.Text = VyM_mes.Semana3.SMM3_A;
                txt_NVC1.Text = VyM_mes.Semana3.NVC1;
                txt_NVC_A1.Text = VyM_mes.Semana3.NVC1_A;
                txt_NVC2.Text = VyM_mes.Semana3.NVC2;
                txt_NVC2.Text = VyM_mes.Semana3.NVC2_A;
                txt_NVC_A3.Text = VyM_mes.Semana3.Libro_A;
                txt_NVC_A4.Text = VyM_mes.Semana3.Libro_L;
                txt_Ora2VyM.Text = VyM_mes.Semana3.Oracion;*/
                
                loading += (15 / loading_delta);

                //lbl_Date.Text = VyM_mes.Semana4.Fecha;
                /*txt_Date.Text = VyM_mes.Semana4.Sem_Biblia;
                txt_Pres.Text = VyM_mes.Semana4.Presidente;
                txt_TdlB_1.Text = VyM_mes.Semana4.Discurso;
                txt_TdlB_A1.Text = VyM_mes.Semana4.Discurso_A;
                txt_TdlB_A2.Text = VyM_mes.Semana4.Perlas;
                txt_TdlB_A3.Text = VyM_mes.Semana4.Lectura;
                txt_SMM1.Text = VyM_mes.Semana4.SMM1;
                txt_SMM_A1.Text = VyM_mes.Semana4.SMM1_A;
                txt_SMM2.Text = VyM_mes.Semana4.SMM2;
                txt_SMM_A2.Text = VyM_mes.Semana4.SMM2_A;
                txt_SMM3.Text = VyM_mes.Semana4.SMM3;
                txt_SMM_A3.Text = VyM_mes.Semana4.SMM3_A;
                txt_NVC1.Text = VyM_mes.Semana4.NVC1;
                txt_NVC_A1.Text = VyM_mes.Semana4.NVC1_A;
                txt_NVC2.Text = VyM_mes.Semana4.NVC2;
                txt_NVC2.Text = VyM_mes.Semana4.NVC2_A;
                txt_NVC_A3.Text = VyM_mes.Semana4.Libro_A;
                txt_NVC_A4.Text = VyM_mes.Semana4.Libro_L;
                txt_Ora2VyM.Text = VyM_mes.Semana4.Oracion;*/

                loading += (15 / loading_delta);

                //lbl_Date.Text = VyM_mes.Semana5.Fecha;
                /*txt_Date.Text = VyM_mes.Semana5.Sem_Biblia;
                txt_Pres.Text = VyM_mes.Semana5.Presidente;
                txt_TdlB_1.Text = VyM_mes.Semana5.Discurso;
                txt_TdlB_A1.Text = VyM_mes.Semana5.Discurso_A;
                txt_TdlB_A2.Text = VyM_mes.Semana5.Perlas;
                txt_TdlB_A3.Text = VyM_mes.Semana5.Lectura;
                txt_SMM1.Text = VyM_mes.Semana5.SMM1;
                txt_SMM_A1.Text = VyM_mes.Semana5.SMM1_A;
                txt_SMM2.Text = VyM_mes.Semana5.SMM2;
                txt_SMM_A2.Text = VyM_mes.Semana5.SMM2_A;
                txt_SMM3.Text = VyM_mes.Semana5.SMM3;
                txt_SMM_A3.Text = VyM_mes.Semana5.SMM3_A;
                txt_NVC1.Text = VyM_mes.Semana5.NVC1;
                txt_NVC_A1.Text = VyM_mes.Semana5.NVC1_A;
                txt_NVC2.Text = VyM_mes.Semana5.NVC2;
                txt_NVC2.Text = VyM_mes.Semana5.NVC2_A;
                txt_NVC_A3.Text = VyM_mes.Semana5.Libro_A;
                txt_NVC_A4.Text = VyM_mes.Semana5.Libro_L;
                txt_Ora2VyM.Text = VyM_mes.Semana5.Oracion;*/

                loading += (15 / loading_delta);

                /*cellValue_1 = (System.Object[,])range_1.get_Value();
                txt_Date.Text = Check_null_string(cellValue_1[primary_cell, A]);
                txt_Pres.Text = Check_null_string(cellValue_1[primary_cell, G]);
                txt_TdlB_1.Text = Check_null_string(cellValue_1[primary_cell + 6, C]);
                txt_TdlB_A1.Text = Check_null_string(cellValue_1[primary_cell + 6, G]);
                txt_TdlB_A2.Text = Check_null_string(cellValue_1[primary_cell + 7, G]);
                txt_TdlB_A3.Text = Check_null_string(cellValue_1[primary_cell + 8, G]);
                txt_SMM1.Text = Check_null_string(cellValue_1[primary_cell + 11, C]);
                txt_SMM_A1.Text = Check_null_string(cellValue_1[primary_cell + 11, G]);
                txt_SMM2.Text = Check_null_string(cellValue_1[primary_cell + 12, C]);
                txt_SMM_A2.Text = Check_null_string(cellValue_1[primary_cell + 12, G]);
                txt_SMM3.Text = Check_null_string(cellValue_1[primary_cell + 13, C]);
                txt_SMM_A3.Text = Check_null_string(cellValue_1[primary_cell + 13, G]);
                txt_NVC1.Text = Check_null_string(cellValue_1[primary_cell + 17, C]);
                txt_NVC_A1.Text = Check_null_string(cellValue_1[primary_cell + 17, G]);
                txt_NVC2.Text = Check_null_string(cellValue_1[primary_cell + 18, C]);
                txt_NVC_A2.Text = Check_null_string(cellValue_1[primary_cell + 18, G]);
                txt_NVC_A3.Text = Check_null_string(cellValue_1[primary_cell + 19, G]);
                //Compare_cbx_string(Check_null_string(cellValue_1[primary_cell + 20, G]), cbx_NVC_A3L);
                //Compare_cbx_string(Check_null_string(cellValue_1[primary_cell + 22, G]), cbx_Ora2VyM);*/
            }
            else
            {
                Get_month_from_Excel(cellValue_1[3, A]);
                VyM_mes.Semana1.Sem_Biblia = Check_null_string(cellValue_1[3, D]);
                VyM_mes.Semana1.Presidente = Check_null_string(cellValue_1[3, G]);

            }
        }

        public void RP_Handler(bool read)
        {
            Notify((read ? "Reading" : "Saving") + " RP meeting");
            if (read)
            {
                cellValue_2 = (System.Object[,])range_2.get_Value();
                int primary_cell = 4;
                //Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_1);
                txt_RP_Speech.Text = Check_null_string(cellValue_2[primary_cell + 2, D]);
                txt_RP_Disc.Text = Check_null_string(cellValue_2[primary_cell + 2, H]);
                txt_RP_Cong.Text = Check_null_string(cellValue_2[primary_cell + 3, E]);
                //Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 5, H]), cbx_CondAtly_1);
                txt_Title_Atly.Text = Check_null_string(cellValue_2[primary_cell + 6, D]);
                //Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 6, H]), cbx_LectRP_1);
                txt_Sal_Disc.Text = Check_null_string(cellValue_2[primary_cell + 10, C]);
                txt_Sal_Title.Text = Check_null_string(cellValue_2[primary_cell + 10, E]);
                txt_Sal_Cong.Text = Check_null_string(cellValue_2[primary_cell + 10, H]);
                //Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_1);

                primary_cell = 17;
                /*Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_2);
                txt_RP_Speech_2.Text = Check_null_string(cellValue_2[primary_cell + 2, D]);
                txt_RP_Disc_2.Text = Check_null_string(cellValue_2[primary_cell + 2, H]);
                txt_RP_Cong_2.Text = Check_null_string(cellValue_2[primary_cell + 3, E]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 5, H]), cbx_CondAtly_2);
                txt_AdlA_Title_2.Text = Check_null_string(cellValue_2[primary_cell + 6, D]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 6, H]), cbx_LectRP_2);
                txt_Sal_Disc_2.Text = Check_null_string(cellValue_2[primary_cell + 10, C]);
                txt_Sal_Title_2.Text = Check_null_string(cellValue_2[primary_cell + 10, E]);
                txt_Sal_Cong_2.Text = Check_null_string(cellValue_2[primary_cell + 10, H]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_1);

                primary_cell = 30;
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_3);
                txt_RP_Speech_3.Text = Check_null_string(cellValue_2[primary_cell + 2, D]);
                txt_RP_Disc_3.Text = Check_null_string(cellValue_2[primary_cell + 2, H]);
                txt_RP_Cong_3.Text = Check_null_string(cellValue_2[primary_cell + 3, E]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 5, H]), cbx_CondAtly_3);
                txt_AdlA_Title_3.Text = Check_null_string(cellValue_2[primary_cell + 6, D]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 6, H]), cbx_LectRP_3);
                txt_Sal_Disc_3.Text = Check_null_string(cellValue_2[primary_cell + 10, C]);
                txt_Sal_Title_3.Text = Check_null_string(cellValue_2[primary_cell + 10, E]);
                txt_Sal_Cong_3.Text = Check_null_string(cellValue_2[primary_cell + 10, H]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_3);

                primary_cell = 43;
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_4);
                txt_RP_Speech_4.Text = Check_null_string(cellValue_2[primary_cell + 2, D]);
                txt_RP_Disc_4.Text = Check_null_string(cellValue_2[primary_cell + 2, H]);
                txt_RP_Cong_4.Text = Check_null_string(cellValue_2[primary_cell + 3, E]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 5, H]), cbx_CondAtly_4);
                txt_AdlA_Title_4.Text = Check_null_string(cellValue_2[primary_cell + 6, D]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 6, H]), cbx_LectRP_4);
                txt_Sal_Disc_4.Text = Check_null_string(cellValue_2[primary_cell + 10, C]);
                txt_Sal_Title_4.Text = Check_null_string(cellValue_2[primary_cell + 10, E]);
                txt_Sal_Cong_4.Text = Check_null_string(cellValue_2[primary_cell + 10, H]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_4);

                primary_cell = 59;
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_5);
                txt_RP_Speech_5.Text = Check_null_string(cellValue_2[primary_cell + 2, D]);
                txt_RP_Disc_5.Text = Check_null_string(cellValue_2[primary_cell + 2, H]);
                txt_RP_Cong_5.Text = Check_null_string(cellValue_2[primary_cell + 3, E]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 5, H]), cbx_CondAtly_5);
                txt_AdlA_Title_5.Text = Check_null_string(cellValue_2[primary_cell + 6, D]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 6, H]), cbx_LectRP_5);
                txt_Sal_Disc_5.Text = Check_null_string(cellValue_2[primary_cell + 10, C]);
                txt_Sal_Title_5.Text = Check_null_string(cellValue_2[primary_cell + 10, E]);
                txt_Sal_Cong_5.Text = Check_null_string(cellValue_2[primary_cell + 10, H]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_5);*/
            }
            else
            {
                int primary_cell = 4;
                //Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_1);
                Sheet_RP.Cells[primary_cell + 2, D] = Check_null(txt_RP_Speech);
                Sheet_RP.Cells[primary_cell + 2, H] = Check_null(txt_RP_Disc);
                Sheet_RP.Cells[primary_cell + 3, E] = Check_null(txt_RP_Cong);
                //Sheet_RP.Cells[primary_cell + 5, H] = Check_null_cbx(cbx_CondAtly_1);
                Sheet_RP.Cells[primary_cell + 6, D] = Check_null(txt_Title_Atly);
                //Sheet_RP.Cells[primary_cell + 6, H] = Check_null_cbx(cbx_LectRP_1);
                //Sheet_RP.Cells[primary_cell + 7, H] = Check_null_cbx(cbx_OraRP_1);
                Sheet_RP.Cells[primary_cell + 10, C] = txt_Sal_Disc.Text;
                Sheet_RP.Cells[primary_cell + 10, E] = txt_Sal_Title.Text;
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong.Text;

                primary_cell = 17;
               /* Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_2);
                Sheet_RP.Cells[primary_cell + 2, D] = Check_null(txt_RP_Speech_2);
                Sheet_RP.Cells[primary_cell + 2, H] = Check_null(txt_RP_Disc_2);
                Sheet_RP.Cells[primary_cell + 3, E] = Check_null(txt_RP_Cong_2);
                Sheet_RP.Cells[primary_cell + 5, H] = Check_null_cbx(cbx_CondAtly_2);
                Sheet_RP.Cells[primary_cell + 6, D] = Check_null(txt_AdlA_Title_2);
                Sheet_RP.Cells[primary_cell + 6, H] = Check_null_cbx(cbx_LectRP_2);
                Sheet_RP.Cells[primary_cell + 7, H] = Check_null_cbx(cbx_OraRP_2);
                Sheet_RP.Cells[primary_cell + 10, C] = txt_Sal_Disc_2.Text;
                Sheet_RP.Cells[primary_cell + 10, E] = txt_Sal_Title_2.Text;
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong_2.Text;

                primary_cell = 30;
                Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_3);
                Sheet_RP.Cells[primary_cell + 2, D] = Check_null(txt_RP_Speech_3);
                Sheet_RP.Cells[primary_cell + 2, H] = Check_null(txt_RP_Disc_3);
                Sheet_RP.Cells[primary_cell + 3, E] = Check_null(txt_RP_Cong_3);
                Sheet_RP.Cells[primary_cell + 5, H] = Check_null_cbx(cbx_CondAtly_3);
                Sheet_RP.Cells[primary_cell + 6, D] = Check_null(txt_AdlA_Title_3);
                Sheet_RP.Cells[primary_cell + 6, H] = Check_null_cbx(cbx_LectRP_3);
                Sheet_RP.Cells[primary_cell + 7, H] = Check_null_cbx(cbx_OraRP_3);
                Sheet_RP.Cells[primary_cell + 10, C] = txt_Sal_Disc_3.Text;
                Sheet_RP.Cells[primary_cell + 10, E] = txt_Sal_Title_3.Text;
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong_3.Text;

                primary_cell = 43;
                Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_4);
                Sheet_RP.Cells[primary_cell + 2, D] = Check_null(txt_RP_Speech_4);
                Sheet_RP.Cells[primary_cell + 2, H] = Check_null(txt_RP_Disc_4);
                Sheet_RP.Cells[primary_cell + 3, E] = Check_null(txt_RP_Cong_4);
                Sheet_RP.Cells[primary_cell + 5, H] = Check_null_cbx(cbx_CondAtly_4);
                Sheet_RP.Cells[primary_cell + 6, D] = Check_null(txt_AdlA_Title_4);
                Sheet_RP.Cells[primary_cell + 6, H] = Check_null_cbx(cbx_LectRP_4);
                Sheet_RP.Cells[primary_cell + 7, H] = Check_null_cbx(cbx_OraRP_4);
                Sheet_RP.Cells[primary_cell + 10, C] = txt_Sal_Disc_4.Text;
                Sheet_RP.Cells[primary_cell + 10, E] = txt_Sal_Title_4.Text;
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong_4.Text;

                primary_cell = 59;
                Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_5);
                Sheet_RP.Cells[primary_cell + 2, D] = Check_null(txt_RP_Speech_5);
                Sheet_RP.Cells[primary_cell + 2, H] = Check_null(txt_RP_Disc_5);
                Sheet_RP.Cells[primary_cell + 3, E] = Check_null(txt_RP_Cong_5);
                Sheet_RP.Cells[primary_cell + 5, H] = Check_null_cbx(cbx_CondAtly_5);
                Sheet_RP.Cells[primary_cell + 6, D] = Check_null(txt_AdlA_Title_5);
                Sheet_RP.Cells[primary_cell + 6, H] = Check_null_cbx(cbx_LectRP_5);
                Sheet_RP.Cells[primary_cell + 7, H] = Check_null_cbx(cbx_OraRP_5);
                Sheet_RP.Cells[primary_cell + 10, C] = txt_Sal_Disc_5.Text;
                Sheet_RP.Cells[primary_cell + 10, E] = txt_Sal_Title_5.Text;
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong_5.Text;*/
            }
        }

        public void AC_Handler(bool read)
        {
            Notify((read ? "Reading" : "Saving") + " AC program");
            if (read)
            {
                cellValue_3 = (System.Object[,])range_3.get_Value();
                int primary_cell = 5;
                /*Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, A]), cbx_Aseo_1);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, C]), cbx_Cap_L_1);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, C]), cbx_AC1_L_1);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, C]), cbx_AC2_L_1);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, E]), cbx_Cap_S_1);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, E]), cbx_AC1_S_1);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, E]), cbx_AC2_S_1);

                primary_cell = 13;
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, A]), cbx_Aseo_2);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, C]), cbx_Cap_L_2);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, C]), cbx_AC1_L_2);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, C]), cbx_AC2_L_2);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, E]), cbx_Cap_S_2);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, E]), cbx_AC1_S_2);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, E]), cbx_AC2_S_2);

                primary_cell = 21;
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, A]), cbx_Aseo_3);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, C]), cbx_Cap_L_3);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, C]), cbx_AC1_L_3);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, C]), cbx_AC2_L_3);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, E]), cbx_Cap_S_3);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, E]), cbx_AC1_S_3);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, E]), cbx_AC2_S_3);

                primary_cell = 29;
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, A]), cbx_Aseo_4);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, C]), cbx_Cap_L_4);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, C]), cbx_AC1_L_4);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, C]), cbx_AC2_L_4);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, E]), cbx_Cap_S_4);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, E]), cbx_AC1_S_4);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, E]), cbx_AC2_S_4);

                primary_cell = 37;
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, A]), cbx_Aseo_5);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, C]), cbx_Cap_L_5);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, C]), cbx_AC1_L_5);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, C]), cbx_AC2_L_5);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, E]), cbx_Cap_S_5);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 2, E]), cbx_AC1_S_5);
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, E]), cbx_AC2_S_5);*/
            }
            else
            {
                int primary_cell = 5;
               /* Sheet_PA.Cells[primary_cell + 1, A] = Check_null_cbx(cbx_Aseo_1);
                Sheet_PA.Cells[primary_cell + 1, C] = Check_null_cbx(cbx_Cap_L_1);
                Sheet_PA.Cells[primary_cell + 2, C] = Check_null_cbx(cbx_AC1_L_1);
                Sheet_PA.Cells[primary_cell + 3, C] = Check_null_cbx(cbx_AC2_L_1);
                Sheet_PA.Cells[primary_cell + 1, E] = Check_null_cbx(cbx_Cap_S_1);
                Sheet_PA.Cells[primary_cell + 2, E] = Check_null_cbx(cbx_AC1_S_1);
                Sheet_PA.Cells[primary_cell + 3, E] = Check_null_cbx(cbx_AC2_S_1);

                primary_cell = 13;
                Sheet_PA.Cells[primary_cell + 1, A] = Check_null_cbx(cbx_Aseo_2);
                Sheet_PA.Cells[primary_cell + 1, C] = Check_null_cbx(cbx_Cap_L_2);
                Sheet_PA.Cells[primary_cell + 2, C] = Check_null_cbx(cbx_AC1_L_2);
                Sheet_PA.Cells[primary_cell + 3, C] = Check_null_cbx(cbx_AC2_L_2);
                Sheet_PA.Cells[primary_cell + 1, E] = Check_null_cbx(cbx_Cap_S_2);
                Sheet_PA.Cells[primary_cell + 2, E] = Check_null_cbx(cbx_AC1_S_2);
                Sheet_PA.Cells[primary_cell + 3, E] = Check_null_cbx(cbx_AC2_S_2);

                primary_cell = 21;
                Sheet_PA.Cells[primary_cell + 1, A] = Check_null_cbx(cbx_Aseo_3);
                Sheet_PA.Cells[primary_cell + 1, C] = Check_null_cbx(cbx_Cap_L_3);
                Sheet_PA.Cells[primary_cell + 2, C] = Check_null_cbx(cbx_AC1_L_3);
                Sheet_PA.Cells[primary_cell + 3, C] = Check_null_cbx(cbx_AC2_L_3);
                Sheet_PA.Cells[primary_cell + 1, E] = Check_null_cbx(cbx_Cap_S_3);
                Sheet_PA.Cells[primary_cell + 2, E] = Check_null_cbx(cbx_AC1_S_3);
                Sheet_PA.Cells[primary_cell + 3, E] = Check_null_cbx(cbx_AC2_S_3);

                primary_cell = 29;
                Sheet_PA.Cells[primary_cell + 1, A] = Check_null_cbx(cbx_Aseo_4);
                Sheet_PA.Cells[primary_cell + 1, C] = Check_null_cbx(cbx_Cap_L_4);
                Sheet_PA.Cells[primary_cell + 2, C] = Check_null_cbx(cbx_AC1_L_4);
                Sheet_PA.Cells[primary_cell + 3, C] = Check_null_cbx(cbx_AC2_L_4);
                Sheet_PA.Cells[primary_cell + 1, E] = Check_null_cbx(cbx_Cap_S_4);
                Sheet_PA.Cells[primary_cell + 2, E] = Check_null_cbx(cbx_AC1_S_4);
                Sheet_PA.Cells[primary_cell + 3, E] = Check_null_cbx(cbx_AC2_S_4);

                primary_cell = 37;
                Sheet_PA.Cells[primary_cell + 1, A] = Check_null_cbx(cbx_Aseo_5);
                Sheet_PA.Cells[primary_cell + 1, C] = Check_null_cbx(cbx_Cap_L_5);
                Sheet_PA.Cells[primary_cell + 2, C] = Check_null_cbx(cbx_AC1_L_5);
                Sheet_PA.Cells[primary_cell + 3, C] = Check_null_cbx(cbx_AC2_L_5);
                Sheet_PA.Cells[primary_cell + 1, E] = Check_null_cbx(cbx_Cap_S_5);
                Sheet_PA.Cells[primary_cell + 2, E] = Check_null_cbx(cbx_AC1_S_5);
                Sheet_PA.Cells[primary_cell + 3, E] = Check_null_cbx(cbx_AC2_S_5);*/
            }
        }


/*--------------------------------------- Auxiliar functions to set/read strings ---------------------------------------*/

        /*public void Compare_cbx_string(object cell_value, object sender)
        {
            if (cell_value.ToString() != null)
            {
                int value = 100;
                ComboBox cbx = (ComboBox)sender;
                cbx.SelectedIndex = -1;
                for (int i = 0; i <= cbx.Items.Count - 1; i++)
                {
                    value = cell_value.ToString().CompareTo(cbx.Items[i].ToString());
                    if (value == 0)
                    {
                        cbx.SelectedIndex = i;
                        break;
                    }
                }
            }
        }*/

        private void Get_month_from_Excel(object cellvalue)
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


        public string Check_null_string(object cellvalue)
        {
            if (cellvalue == null)
            {
                cellvalue = "";
            }
            return cellvalue.ToString();
        }

        public int Get_cell()
        {
            int cell = 0;
            switch (m_semana)
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
            Presenter(p.Fenix);
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
                tab_meeting = 0;
                autocomplete.AddRange(Dict_vym.Keys.ToArray());
                Presenter(p.Executor);
                Notify("Overview");
            }
            else if (tab_Control.SelectedIndex == 1)
            {
                tab_meeting = 1;
                autocomplete.AddRange(Dict_rp.Keys.ToArray());
                Presenter(p.Artanis);
                Notify("Section 'Reunion Publica y analisis de La Atalaya'");
                
            }
            else
            {
                tab_meeting = 2;
                autocomplete.AddRange(Dict_ac.Keys.ToArray());
                Presenter(p.Oracle);
                Notify("Section 'Acomodadores'");
            }
            Week_Handler();
            txt_Command.AutoCompleteCustomSource = autocomplete;
        }

        private void TextChanged(object sender, EventArgs e)
        {

        }

        private void Save_time_from_string(VyM_Sem vyM_Sem, int wk, bool save)
        {
            DateTime Aux_dateTime = new DateTime(2018, 1, 1, 7, 00, 00);
            if (save)
            {

            }
            else // read
            {
                time_0.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(5);
                time_1.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(3);
                time_2.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(10);
                time_3.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(8);
                time_4.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(5+1); //adjusting to real time
                time_5.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(vyM_Sem.SMM1, true));
                time_6.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(vyM_Sem.SMM2, true));
                time_7.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(vyM_Sem.SMM3, true) + 1); //adjusting to real time
                time_8.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(3);
                time_9.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(vyM_Sem.NVC1, false));
                if ((txt_NVC2.Text == null) || (txt_NVC2.Text == " ") || (txt_NVC2.Text == " -"))
                {
                    time_10.Text = " ";                
                }
                else
                {
                    time_10.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                    Aux_dateTime = Aux_dateTime.AddMinutes(Get_time_from_string(vyM_Sem.NVC2, false));
                }
                Aux_dateTime = Aux_dateTime.AddMinutes(1); //adjusting to real time
                time_11.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(30);
                time_12.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                Aux_dateTime = Aux_dateTime.AddMinutes(3);
                time_13.Text = Aux_dateTime.Hour.ToString() + ":" + Aux_dateTime.Minute.ToString();
                if (Aux_dateTime.Hour == 8 && Aux_dateTime.Minute == 40)
                {
                    time_13.ForeColor = Color.Green;
                }
                else
                {
                    time_13.ForeColor = Color.Red;
                }
            }
        }

        public int Get_time_from_string(string Str, bool SMM)
        {
            string min = "mins.";
            string video = "video";
            Str = Str.ToLower();
            string number = "";
            var array = Str.ToCharArray();
            int time = 0;
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
            if (SMM)
            {
                if (!Str.Contains(video))
                {
                    time++;
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

        private void open_DB()
        {
            if (!DB_form_show)
            {
                DB_form_show = true;
                timer_Form2.Enabled = true;
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
        }

        /*public void save_DB(object sender)
        {
            ComboBox cbx = (ComboBox)sender;
            int column = 0;
            string s = "0";
            switch (cbx.Name.ToString())
            {
                case "cbx_Ora1VyM":
                    {
                        column = 9;
                        break;
                    }
                case "cbx_NVC_A3L":
                    {
                        column = 10;
                        break;
                    }
                case "cbx_Ora2VyM":
                    {
                        column = 9;
                        break;
                    }
                case "cbx_PresRP":
                    {
                        column = 5;
                        break;
                    }
                case "cbx_LectRP":
                    {
                        column = 6;
                        break;
                    }
                case "cbx_OraRP":
                    {
                        column = 7;
                        break;
                    }
                case "cbx_CondAtly":
                    {
                        column = 8;
                        break;
                    }
                case "cbx_Cap":
                    {
                        column = 3 ;
                        break;
                    }
                case "cbx_AC1":
                    {
                        column = 4;
                        break;
                    }
                case "cbx_AC2":
                    {
                        column = 4;
                        break;
                    }
                case "cbx_Aseo":
                    {
                        column = 3;
                        break;
                    }
            }
            for (int i = 5; i <= 42; i++)
            {
                int value = 100;
                if ((cellValue_4[i, B] != null) && (cbx.SelectedItem != null))
                {
                    value = cbx.SelectedItem.ToString().CompareTo(cellValue_4[i, B].ToString());
                    if (value == 0)
                    {
                        s = (m_año-2000).ToString() + m_mes.ToString("00") + m_dia.ToString("00");
                        //Sheet_DB.Cells[i, column] = s;
                        break;
                    }
                }
            }
        }

        private void Autofill_Handler()
        {

        }

        /*Function so set local variables' info into form*/
        public void Week_Handler()
        {
            int lun = 0;
            lbl_Week.Text = "Semana: " + m_semana.ToString();
            switch (tab_meeting)
            {
                case 0:
                    {
                        lbl_Date.Text = meetings_days[m_semana - 1, 0].ToString("dddd, dd MMMM");
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    //lbl_Date.Text = meetings_days[m_semana, 0].ToString();
                                    //lbl_Date.Text = VyM_mes.Semana1.Fecha;
                                    txt_Date.Text = VyM_mes.Semana1.Sem_Biblia;
                                    txt_Pres.Text = VyM_mes.Semana1.Presidente;
                                    txt_TdlB_1.Text = VyM_mes.Semana1.Discurso;
                                    txt_TdlB_A1.Text = VyM_mes.Semana1.Discurso_A;
                                    txt_TdlB_A2.Text = VyM_mes.Semana1.Perlas;
                                    txt_TdlB_A3.Text = VyM_mes.Semana1.Lectura;
                                    txt_SMM1.Text = VyM_mes.Semana1.SMM1;
                                    txt_SMM_A1.Text = VyM_mes.Semana1.SMM1_A;
                                    txt_SMM2.Text = VyM_mes.Semana1.SMM2;
                                    txt_SMM_A2.Text = VyM_mes.Semana1.SMM2_A;
                                    txt_SMM3.Text = VyM_mes.Semana1.SMM3;
                                    txt_SMM_A3.Text = VyM_mes.Semana1.SMM3_A;
                                    txt_NVC1.Text = VyM_mes.Semana1.NVC1;
                                    txt_NVC_A1.Text = VyM_mes.Semana1.NVC1_A;
                                    txt_NVC2.Text = VyM_mes.Semana1.NVC2;
                                    txt_NVC2.Text = VyM_mes.Semana1.NVC2_A;
                                    txt_NVC_A3.Text = VyM_mes.Semana1.Libro_A;
                                    txt_NVC_A4.Text = VyM_mes.Semana1.Libro_L;
                                    txt_Ora2VyM.Text = VyM_mes.Semana1.Oracion;
                                    break;
                                }
                            case 2:
                                {
                                    //lbl_Date.Text = VyM_mes.Semana2.Fecha;
                                    txt_Date.Text = VyM_mes.Semana2.Sem_Biblia;
                                    txt_Pres.Text = VyM_mes.Semana2.Presidente;
                                    txt_TdlB_1.Text = VyM_mes.Semana2.Discurso;
                                    txt_TdlB_A1.Text = VyM_mes.Semana2.Discurso_A;
                                    txt_TdlB_A2.Text = VyM_mes.Semana2.Perlas;
                                    txt_TdlB_A3.Text = VyM_mes.Semana2.Lectura;
                                    txt_SMM1.Text = VyM_mes.Semana2.SMM1;
                                    txt_SMM_A1.Text = VyM_mes.Semana2.SMM1_A;
                                    txt_SMM2.Text = VyM_mes.Semana2.SMM2;
                                    txt_SMM_A2.Text = VyM_mes.Semana2.SMM2_A;
                                    txt_SMM3.Text = VyM_mes.Semana2.SMM3;
                                    txt_SMM_A3.Text = VyM_mes.Semana2.SMM3_A;
                                    txt_NVC1.Text = VyM_mes.Semana2.NVC1;
                                    txt_NVC_A1.Text = VyM_mes.Semana2.NVC1_A;
                                    txt_NVC2.Text = VyM_mes.Semana2.NVC2;
                                    txt_NVC2.Text = VyM_mes.Semana2.NVC2_A;
                                    txt_NVC_A3.Text = VyM_mes.Semana2.Libro_A;
                                    txt_NVC_A4.Text = VyM_mes.Semana2.Libro_L;
                                    txt_Ora2VyM.Text = VyM_mes.Semana2.Oracion;
                                    break;
                                }
                            case 3:
                                {
                                    //lbl_Date.Text = VyM_mes.Semana3.Fecha;
                                    txt_Date.Text = VyM_mes.Semana3.Sem_Biblia;
                                    txt_Pres.Text = VyM_mes.Semana3.Presidente;
                                    txt_TdlB_1.Text = VyM_mes.Semana3.Discurso;
                                    txt_TdlB_A1.Text = VyM_mes.Semana3.Discurso_A;
                                    txt_TdlB_A2.Text = VyM_mes.Semana3.Perlas;
                                    txt_TdlB_A3.Text = VyM_mes.Semana3.Lectura;
                                    txt_SMM1.Text = VyM_mes.Semana3.SMM1;
                                    txt_SMM_A1.Text = VyM_mes.Semana3.SMM1_A;
                                    txt_SMM2.Text = VyM_mes.Semana3.SMM2;
                                    txt_SMM_A2.Text = VyM_mes.Semana3.SMM2_A;
                                    txt_SMM3.Text = VyM_mes.Semana3.SMM3;
                                    txt_SMM_A3.Text = VyM_mes.Semana3.SMM3_A;
                                    txt_NVC1.Text = VyM_mes.Semana3.NVC1;
                                    txt_NVC_A1.Text = VyM_mes.Semana3.NVC1_A;
                                    txt_NVC2.Text = VyM_mes.Semana3.NVC2;
                                    txt_NVC2.Text = VyM_mes.Semana3.NVC2_A;
                                    txt_NVC_A3.Text = VyM_mes.Semana3.Libro_A;
                                    txt_NVC_A4.Text = VyM_mes.Semana3.Libro_L;
                                    txt_Ora2VyM.Text = VyM_mes.Semana3.Oracion;
                                    break;
                                }
                            case 4:
                                {
                                    //lbl_Date.Text = VyM_mes.Semana4.Fecha;
                                    txt_Date.Text = VyM_mes.Semana4.Sem_Biblia;
                                    txt_Pres.Text = VyM_mes.Semana4.Presidente;
                                    txt_TdlB_1.Text = VyM_mes.Semana4.Discurso;
                                    txt_TdlB_A1.Text = VyM_mes.Semana4.Discurso_A;
                                    txt_TdlB_A2.Text = VyM_mes.Semana4.Perlas;
                                    txt_TdlB_A3.Text = VyM_mes.Semana4.Lectura;
                                    txt_SMM1.Text = VyM_mes.Semana4.SMM1;
                                    txt_SMM_A1.Text = VyM_mes.Semana4.SMM1_A;
                                    txt_SMM2.Text = VyM_mes.Semana4.SMM2;
                                    txt_SMM_A2.Text = VyM_mes.Semana4.SMM2_A;
                                    txt_SMM3.Text = VyM_mes.Semana4.SMM3;
                                    txt_SMM_A3.Text = VyM_mes.Semana4.SMM3_A;
                                    txt_NVC1.Text = VyM_mes.Semana4.NVC1;
                                    txt_NVC_A1.Text = VyM_mes.Semana4.NVC1_A;
                                    txt_NVC2.Text = VyM_mes.Semana4.NVC2;
                                    txt_NVC2.Text = VyM_mes.Semana4.NVC2_A;
                                    txt_NVC_A3.Text = VyM_mes.Semana4.Libro_A;
                                    txt_NVC_A4.Text = VyM_mes.Semana4.Libro_L;
                                    txt_Ora2VyM.Text = VyM_mes.Semana4.Oracion;
                                    break;
                                }
                            case 5:
                                {
                                    //lbl_Date.Text = VyM_mes.Semana5.Fecha;
                                    txt_Date.Text = VyM_mes.Semana5.Sem_Biblia;
                                    txt_Pres.Text = VyM_mes.Semana5.Presidente;
                                    txt_TdlB_1.Text = VyM_mes.Semana5.Discurso;
                                    txt_TdlB_A1.Text = VyM_mes.Semana5.Discurso_A;
                                    txt_TdlB_A2.Text = VyM_mes.Semana5.Perlas;
                                    txt_TdlB_A3.Text = VyM_mes.Semana5.Lectura;
                                    txt_SMM1.Text = VyM_mes.Semana5.SMM1;
                                    txt_SMM_A1.Text = VyM_mes.Semana5.SMM1_A;
                                    txt_SMM2.Text = VyM_mes.Semana5.SMM2;
                                    txt_SMM_A2.Text = VyM_mes.Semana5.SMM2_A;
                                    txt_SMM3.Text = VyM_mes.Semana5.SMM3;
                                    txt_SMM_A3.Text = VyM_mes.Semana5.SMM3_A;
                                    txt_NVC1.Text = VyM_mes.Semana5.NVC1;
                                    txt_NVC_A1.Text = VyM_mes.Semana5.NVC1_A;
                                    txt_NVC2.Text = VyM_mes.Semana5.NVC2;
                                    txt_NVC2.Text = VyM_mes.Semana5.NVC2_A;
                                    txt_NVC_A3.Text = VyM_mes.Semana5.Libro_A;
                                    txt_NVC_A4.Text = VyM_mes.Semana5.Libro_L;
                                    txt_Ora2VyM.Text = VyM_mes.Semana5.Oracion;
                                    break;
                                }
                        }
                        lun = 0;
                        break;
                    }
                case 1:
                    {
                        switch (m_semana)
                        {
                            case 1:
                                {
                                    txt_DateRP.Text= RP_mes.Semana1.Fecha;
                                    txt_RP_Speech.Text = RP_mes.Semana1.Titulo;
                                    txt_PresRP.Text = RP_mes.Semana1.Presidente;
                                    txt_RP_Disc.Text = RP_mes.Semana1.Congregacion;
                                    txt_RP_Cong.Text = RP_mes.Semana1.Discursante;
                                    txt_Title_Atly.Text = RP_mes.Semana1.Titulo_Atalaya;
                                    txt_Con_Atly.Text = RP_mes.Semana1.Conductor;
                                    txt_Lect_Atly.Text = RP_mes.Semana1.Lector;
                                    txt_OraRP.Text = RP_mes.Semana1.Oracion;
                                    txt_Sal_Disc.Text = RP_mes.Semana1.Discu_Sal;
                                    txt_Sal_Title.Text = RP_mes.Semana1.Ttl_Sal;
                                    txt_Sal_Cong.Text = RP_mes.Semana1.Cong_Sal;
                                    break;
                                }
                            case 2:
                                {
                                    txt_DateRP.Text = RP_mes.Semana2.Fecha;
                                    txt_RP_Speech.Text = RP_mes.Semana2.Titulo;
                                    txt_PresRP.Text = RP_mes.Semana2.Presidente;
                                    txt_RP_Disc.Text = RP_mes.Semana2.Congregacion;
                                    txt_RP_Cong.Text = RP_mes.Semana2.Discursante;
                                    txt_Title_Atly.Text = RP_mes.Semana2.Titulo_Atalaya;
                                    txt_Con_Atly.Text = RP_mes.Semana2.Conductor;
                                    txt_Lect_Atly.Text = RP_mes.Semana2.Lector;
                                    txt_OraRP.Text = RP_mes.Semana2.Oracion;
                                    txt_Sal_Disc.Text = RP_mes.Semana2.Discu_Sal;
                                    txt_Sal_Title.Text = RP_mes.Semana2.Ttl_Sal;
                                    txt_Sal_Cong.Text = RP_mes.Semana2.Cong_Sal;
                                    break;
                                }
                            case 3:
                                {
                                    txt_DateRP.Text = RP_mes.Semana3.Fecha;
                                    txt_RP_Speech.Text = RP_mes.Semana3.Titulo;
                                    txt_PresRP.Text = RP_mes.Semana3.Presidente;
                                    txt_RP_Disc.Text = RP_mes.Semana3.Congregacion;
                                    txt_RP_Cong.Text = RP_mes.Semana3.Discursante;
                                    txt_Title_Atly.Text = RP_mes.Semana3.Titulo_Atalaya;
                                    txt_Con_Atly.Text = RP_mes.Semana3.Conductor;
                                    txt_Lect_Atly.Text = RP_mes.Semana3.Lector;
                                    txt_OraRP.Text = RP_mes.Semana3.Oracion;
                                    txt_Sal_Disc.Text = RP_mes.Semana3.Discu_Sal;
                                    txt_Sal_Title.Text = RP_mes.Semana3.Ttl_Sal;
                                    txt_Sal_Cong.Text = RP_mes.Semana3.Cong_Sal;
                                    break;
                                }
                            case 4:
                                {
                                    txt_DateRP.Text = RP_mes.Semana4.Fecha;
                                    txt_RP_Speech.Text = RP_mes.Semana4.Titulo;
                                    txt_PresRP.Text = RP_mes.Semana4.Presidente;
                                    txt_RP_Disc.Text = RP_mes.Semana4.Congregacion;
                                    txt_RP_Cong.Text = RP_mes.Semana4.Discursante;
                                    txt_Title_Atly.Text = RP_mes.Semana4.Titulo_Atalaya;
                                    txt_Con_Atly.Text = RP_mes.Semana4.Conductor;
                                    txt_Lect_Atly.Text = RP_mes.Semana4.Lector;
                                    txt_OraRP.Text = RP_mes.Semana4.Oracion;
                                    txt_Sal_Disc.Text = RP_mes.Semana4.Discu_Sal;
                                    txt_Sal_Title.Text = RP_mes.Semana4.Ttl_Sal;
                                    txt_Sal_Cong.Text = RP_mes.Semana4.Cong_Sal;
                                    break;
                                }
                            case 5:
                                {
                                    txt_DateRP.Text = RP_mes.Semana5.Fecha;
                                    txt_RP_Speech.Text = RP_mes.Semana5.Titulo;
                                    txt_PresRP.Text = RP_mes.Semana5.Presidente;
                                    txt_RP_Disc.Text = RP_mes.Semana5.Congregacion;
                                    txt_RP_Cong.Text = RP_mes.Semana5.Discursante;
                                    txt_Title_Atly.Text = RP_mes.Semana5.Titulo_Atalaya;
                                    txt_Con_Atly.Text = RP_mes.Semana5.Conductor;
                                    txt_Lect_Atly.Text = RP_mes.Semana5.Lector;
                                    txt_OraRP.Text = RP_mes.Semana5.Oracion;
                                    txt_Sal_Disc.Text = RP_mes.Semana5.Discu_Sal;
                                    txt_Sal_Title.Text = RP_mes.Semana5.Ttl_Sal;
                                    txt_Sal_Cong.Text = RP_mes.Semana5.Cong_Sal;
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
                                    VyM_mes.Semana1.Fecha = lbl_Date.Text;
                                    VyM_mes.Semana1.Sem_Biblia = txt_Date.Text;
                                    VyM_mes.Semana1.Presidente = txt_Pres.Text;
                                    VyM_mes.Semana1.Discurso = txt_TdlB_1.Text;
                                    VyM_mes.Semana1.Discurso_A = txt_TdlB_A1.Text;
                                    VyM_mes.Semana1.Perlas = txt_TdlB_A2.Text;
                                    VyM_mes.Semana1.Lectura = txt_TdlB_A3.Text;
                                    VyM_mes.Semana1.SMM1 = txt_SMM1.Text;
                                    VyM_mes.Semana1.SMM1_A = txt_SMM_A1.Text;
                                    VyM_mes.Semana1.SMM2 = txt_SMM2.Text;
                                    VyM_mes.Semana1.SMM2_A = txt_SMM_A2.Text;
                                    VyM_mes.Semana1.SMM3 = txt_SMM3.Text;
                                    VyM_mes.Semana1.SMM3_A = txt_SMM_A3.Text;
                                    VyM_mes.Semana1.NVC1 = txt_NVC1.Text;
                                    VyM_mes.Semana1.NVC1_A = txt_NVC_A1.Text;
                                    VyM_mes.Semana1.NVC2 = txt_NVC2.Text;
                                    VyM_mes.Semana1.NVC2_A = txt_NVC2.Text;
                                    VyM_mes.Semana1.Libro_A = txt_NVC_A3.Text;
                                    VyM_mes.Semana1.Libro_L = txt_NVC_A4.Text;
                                    VyM_mes.Semana1.Oracion = txt_Ora2VyM.Text;
                                    break;
                                }
                            case 2:
                                {
                                    VyM_mes.Semana2.Fecha = lbl_Date.Text;
                                    VyM_mes.Semana2.Sem_Biblia = txt_Date.Text;
                                    VyM_mes.Semana2.Presidente = txt_Pres.Text;
                                    VyM_mes.Semana2.Discurso = txt_TdlB_1.Text;
                                    VyM_mes.Semana2.Discurso_A = txt_TdlB_A1.Text;
                                    VyM_mes.Semana2.Perlas = txt_TdlB_A2.Text;
                                    VyM_mes.Semana2.Lectura = txt_TdlB_A3.Text;
                                    VyM_mes.Semana2.SMM1 = txt_SMM1.Text;
                                    VyM_mes.Semana2.SMM1_A = txt_SMM_A1.Text;
                                    VyM_mes.Semana2.SMM2 = txt_SMM2.Text;
                                    VyM_mes.Semana2.SMM2_A = txt_SMM_A2.Text;
                                    VyM_mes.Semana2.SMM3 = txt_SMM3.Text;
                                    VyM_mes.Semana2.SMM3_A = txt_SMM_A3.Text;
                                    VyM_mes.Semana2.NVC1 = txt_NVC1.Text;
                                    VyM_mes.Semana2.NVC1_A = txt_NVC_A1.Text;
                                    VyM_mes.Semana2.NVC2 = txt_NVC2.Text;
                                    VyM_mes.Semana2.NVC2_A = txt_NVC2.Text;
                                    VyM_mes.Semana2.Libro_A = txt_NVC_A3.Text;
                                    VyM_mes.Semana2.Libro_L = txt_NVC_A4.Text;
                                    VyM_mes.Semana2.Oracion = txt_Ora2VyM.Text;
                                    break;
                                }
                            case 3:
                                {
                                    VyM_mes.Semana3.Fecha = lbl_Date.Text;
                                    VyM_mes.Semana3.Sem_Biblia = txt_Date.Text;
                                    VyM_mes.Semana3.Presidente = txt_Pres.Text;
                                    VyM_mes.Semana3.Discurso = txt_TdlB_1.Text;
                                    VyM_mes.Semana3.Discurso_A = txt_TdlB_A1.Text;
                                    VyM_mes.Semana3.Perlas = txt_TdlB_A2.Text;
                                    VyM_mes.Semana3.Lectura = txt_TdlB_A3.Text;
                                    VyM_mes.Semana3.SMM1 = txt_SMM1.Text;
                                    VyM_mes.Semana3.SMM1_A = txt_SMM_A1.Text;
                                    VyM_mes.Semana3.SMM2 = txt_SMM2.Text;
                                    VyM_mes.Semana3.SMM2_A = txt_SMM_A2.Text;
                                    VyM_mes.Semana3.SMM3 = txt_SMM3.Text;
                                    VyM_mes.Semana3.SMM3_A = txt_SMM_A3.Text;
                                    VyM_mes.Semana3.NVC1 = txt_NVC1.Text;
                                    VyM_mes.Semana3.NVC1_A = txt_NVC_A1.Text;
                                    VyM_mes.Semana3.NVC2 = txt_NVC2.Text;
                                    VyM_mes.Semana3.NVC2_A = txt_NVC2.Text;
                                    VyM_mes.Semana3.Libro_A = txt_NVC_A3.Text;
                                    VyM_mes.Semana3.Libro_L = txt_NVC_A4.Text;
                                    VyM_mes.Semana3.Oracion = txt_Ora2VyM.Text;
                                    break;
                                }
                            case 4:
                                {
                                    VyM_mes.Semana4.Fecha = lbl_Date.Text;
                                    VyM_mes.Semana4.Sem_Biblia = txt_Date.Text;
                                    VyM_mes.Semana4.Presidente = txt_Pres.Text;
                                    VyM_mes.Semana4.Discurso = txt_TdlB_1.Text;
                                    VyM_mes.Semana4.Discurso_A = txt_TdlB_A1.Text;
                                    VyM_mes.Semana4.Perlas = txt_TdlB_A2.Text;
                                    VyM_mes.Semana4.Lectura = txt_TdlB_A3.Text;
                                    VyM_mes.Semana4.SMM1 = txt_SMM1.Text;
                                    VyM_mes.Semana4.SMM1_A = txt_SMM_A1.Text;
                                    VyM_mes.Semana4.SMM2 = txt_SMM2.Text;
                                    VyM_mes.Semana4.SMM2_A = txt_SMM_A2.Text;
                                    VyM_mes.Semana4.SMM3 = txt_SMM3.Text;
                                    VyM_mes.Semana4.SMM3_A = txt_SMM_A3.Text;
                                    VyM_mes.Semana4.NVC1 = txt_NVC1.Text;
                                    VyM_mes.Semana4.NVC1_A = txt_NVC_A1.Text;
                                    VyM_mes.Semana4.NVC2 = txt_NVC2.Text;
                                    VyM_mes.Semana4.NVC2_A = txt_NVC2.Text;
                                    VyM_mes.Semana4.Libro_A = txt_NVC_A3.Text;
                                    VyM_mes.Semana4.Libro_L = txt_NVC_A4.Text;
                                    VyM_mes.Semana4.Oracion = txt_Ora2VyM.Text;
                                    break;
                                }
                            case 5:
                                {
                                    VyM_mes.Semana5.Fecha = lbl_Date.Text;
                                    VyM_mes.Semana5.Sem_Biblia = txt_Date.Text;
                                    VyM_mes.Semana5.Presidente = txt_Pres.Text;
                                    VyM_mes.Semana5.Discurso = txt_TdlB_1.Text;
                                    VyM_mes.Semana5.Discurso_A = txt_TdlB_A1.Text;
                                    VyM_mes.Semana5.Perlas = txt_TdlB_A2.Text;
                                    VyM_mes.Semana5.Lectura = txt_TdlB_A3.Text;
                                    VyM_mes.Semana5.SMM1 = txt_SMM1.Text;
                                    VyM_mes.Semana5.SMM1_A = txt_SMM_A1.Text;
                                    VyM_mes.Semana5.SMM2 = txt_SMM2.Text;
                                    VyM_mes.Semana5.SMM2_A = txt_SMM_A2.Text;
                                    VyM_mes.Semana5.SMM3 = txt_SMM3.Text;
                                    VyM_mes.Semana5.SMM3_A = txt_SMM_A3.Text;
                                    VyM_mes.Semana5.NVC1 = txt_NVC1.Text;
                                    VyM_mes.Semana5.NVC1_A = txt_NVC_A1.Text;
                                    VyM_mes.Semana5.NVC2 = txt_NVC2.Text;
                                    VyM_mes.Semana5.NVC2_A = txt_NVC2.Text;
                                    VyM_mes.Semana5.Libro_A = txt_NVC_A3.Text;
                                    VyM_mes.Semana5.Libro_L = txt_NVC_A4.Text;
                                    VyM_mes.Semana5.Oracion = txt_Ora2VyM.Text;
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
                                    RP_mes.Semana1.Fecha = txt_DateRP.Text;
                                    RP_mes.Semana1.Titulo = txt_RP_Speech.Text;
                                    RP_mes.Semana1.Presidente = txt_PresRP.Text;
                                    RP_mes.Semana1.Congregacion = txt_RP_Disc.Text;
                                    RP_mes.Semana1.Discursante = txt_RP_Cong.Text;
                                    RP_mes.Semana1.Titulo_Atalaya = txt_Title_Atly.Text;
                                    RP_mes.Semana1.Conductor = txt_Con_Atly.Text;
                                    RP_mes.Semana1.Lector = txt_Lect_Atly.Text;
                                    RP_mes.Semana1.Oracion = txt_OraRP.Text;
                                    RP_mes.Semana1.Discu_Sal = txt_Sal_Disc.Text;
                                    RP_mes.Semana1.Ttl_Sal = txt_Sal_Title.Text;
                                    RP_mes.Semana1.Cong_Sal = txt_Sal_Cong.Text;
                                    break;
                                }
                            case 2:
                                {
                                    RP_mes.Semana2.Fecha = txt_DateRP.Text;
                                    RP_mes.Semana2.Titulo = txt_RP_Speech.Text;
                                    RP_mes.Semana2.Presidente = txt_PresRP.Text;
                                    RP_mes.Semana2.Congregacion = txt_RP_Disc.Text;
                                    RP_mes.Semana2.Discursante = txt_RP_Cong.Text;
                                    RP_mes.Semana2.Titulo_Atalaya = txt_Title_Atly.Text;
                                    RP_mes.Semana2.Conductor = txt_Con_Atly.Text;
                                    RP_mes.Semana2.Lector = txt_Lect_Atly.Text;
                                    RP_mes.Semana2.Oracion = txt_OraRP.Text;
                                    RP_mes.Semana2.Discu_Sal = txt_Sal_Disc.Text;
                                    RP_mes.Semana2.Ttl_Sal = txt_Sal_Title.Text;
                                    RP_mes.Semana2.Cong_Sal = txt_Sal_Cong.Text;
                                    break;
                                }
                            case 3:
                                {
                                    RP_mes.Semana3.Fecha = txt_DateRP.Text;
                                    RP_mes.Semana3.Titulo = txt_RP_Speech.Text;
                                    RP_mes.Semana3.Presidente = txt_PresRP.Text;
                                    RP_mes.Semana3.Congregacion = txt_RP_Disc.Text;
                                    RP_mes.Semana3.Discursante = txt_RP_Cong.Text;
                                    RP_mes.Semana3.Titulo_Atalaya = txt_Title_Atly.Text;
                                    RP_mes.Semana3.Conductor = txt_Con_Atly.Text;
                                    RP_mes.Semana3.Lector = txt_Lect_Atly.Text;
                                    RP_mes.Semana3.Oracion = txt_OraRP.Text;
                                    RP_mes.Semana3.Discu_Sal = txt_Sal_Disc.Text;
                                    RP_mes.Semana3.Ttl_Sal = txt_Sal_Title.Text;
                                    RP_mes.Semana3.Cong_Sal = txt_Sal_Cong.Text;
                                    break;
                                }
                            case 4:
                                {
                                    RP_mes.Semana4.Fecha = txt_DateRP.Text;
                                    RP_mes.Semana4.Titulo = txt_RP_Speech.Text;
                                    RP_mes.Semana4.Presidente = txt_PresRP.Text;
                                    RP_mes.Semana4.Congregacion = txt_RP_Disc.Text;
                                    RP_mes.Semana4.Discursante = txt_RP_Cong.Text;
                                    RP_mes.Semana4.Titulo_Atalaya = txt_Title_Atly.Text;
                                    RP_mes.Semana4.Conductor = txt_Con_Atly.Text;
                                    RP_mes.Semana4.Lector = txt_Lect_Atly.Text;
                                    RP_mes.Semana4.Oracion = txt_OraRP.Text;
                                    RP_mes.Semana4.Discu_Sal = txt_Sal_Disc.Text;
                                    RP_mes.Semana4.Ttl_Sal = txt_Sal_Title.Text;
                                    RP_mes.Semana4.Cong_Sal = txt_Sal_Cong.Text;
                                    break;
                                }
                            case 5:
                                {
                                    RP_mes.Semana5.Fecha = txt_DateRP.Text;
                                    RP_mes.Semana5.Titulo = txt_RP_Speech.Text;
                                    RP_mes.Semana5.Presidente = txt_PresRP.Text;
                                    RP_mes.Semana5.Congregacion = txt_RP_Disc.Text;
                                    RP_mes.Semana5.Discursante = txt_RP_Cong.Text;
                                    RP_mes.Semana5.Titulo_Atalaya = txt_Title_Atly.Text;
                                    RP_mes.Semana5.Conductor = txt_Con_Atly.Text;
                                    RP_mes.Semana5.Lector = txt_Lect_Atly.Text;
                                    RP_mes.Semana5.Oracion = txt_OraRP.Text;
                                    RP_mes.Semana5.Discu_Sal = txt_Sal_Disc.Text;
                                    RP_mes.Semana5.Ttl_Sal = txt_Sal_Title.Text;
                                    RP_mes.Semana5.Cong_Sal = txt_Sal_Cong.Text;
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
    }
}
