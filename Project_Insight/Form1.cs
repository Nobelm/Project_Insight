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
        private Excel.Application objApp = new Excel.Application();
        private Excel.Workbook objBooks = null;
        private Excel.Sheets objSheets;
        private Excel.Worksheet Sheet_VyM;
        private Excel.Worksheet Sheet_RP;
        private Excel.Worksheet Sheet_PA;
        private Excel.Worksheet Sheet_DB;
        private Excel.Range range_1;
        private Excel.Range range_2;
        private Excel.Range range_3;
        private Excel.Range range_4;
        public static bool excel_ready = false;
        private DateTime dateTime = new DateTime(2018, 1, 1, 7, 00, 00);
        private DateTime start_time = new DateTime(2018, 1, 1, 7, 00, 00);
        private DateTime date;
        private object[,] cellValue_1 = null;
        private object[,] cellValue_2 = null;
        private object[,] cellValue_3 = null;
        public static Object[,] cellValue_4 = null;
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
        public static int m_año = 2018;
        public static int m_semana = 0;
        public static DateTime[,] meetings_days = new DateTime[5,2];
        public static string[]  guard_cbx_names = new string[10];
        public static int date_checksum = 0;
        public static string[] Command_history = new string[10];
        //public static string[] Command_input = new string[] {"op_xlsx", "op_db", "sv", "clc", "rst", "mnth", "wk", "autofill", "exit"};
        //public static string[] Command_input = new string[] {"new", "open", "save", "exit", "month"};
        public static string[] month = new string[] { "ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic" };
        public static int command_iterator = 0;
        DB_Form DB_Form = new DB_Form();
        public static string Path = "";
        public static bool is_new_instance = false;


        public Main_Form()
        {
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            InitializeComponent();
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            Notify("UI up and ready \nWelcome back Hierarch!");
            Presenter(p.Executor);
            Warn("Pending changes:");
            Warn("[1] Make DB static in the code, and implement \"save as pdf\" logic");
            Warn("[2] Implementing autofill function");
            Warn("[3] Implementing handling for \"Asambleas\"");
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
                if(caller != "String_stack")
                {
                    log_txtBx.AppendText("(" + caller + ") ");
                }
                for (int i=0; i<= array.Length - 1; i++)
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

        public async void Command(string data, [CallerLineNumber] int lineNumber = 0)
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
                    String_stack("", false, 3, lineNumber);
                }
            }
            else
            {
                String_stack(data, true, 3, lineNumber);
            }
        }

        private void Process_txt_Command(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txt_Command.Text != "")
                {
                    string Str = txt_Command.Text.ToLower();
                    string cmd = Str;
                    string sup = "";
                    int index = 0;
                    if (Str.Length > 4)
                    {
                        index = cmd.IndexOf(" ");
                        cmd = cmd.Substring(0, index);
                        sup = Str.Substring(index+1);
                    }
                    Command("Executing [" + Str + "] command");
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
                                string[] file = new string[] {"vym", "rp", "ac", "all"};
                                for (int i = 0; i <= file.Length - 1; i++)
                                {
                                    if (sup.Contains(file[i]))
                                    {
                                        Process_save(i + 1);
                                        break;
                                    }
                                }
                                break;
                            }
                        default:
                            {

                                break;
                            }
                    }
                    txt_Command.Text = "";
                }
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (command_iterator < Command_history.Length-1)
                {
                    command_iterator++;
                    if (Command_history[command_iterator] != null)
                    {
                        txt_Command.Text = Command_history[command_iterator];
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
                    }
                }
                if (command_iterator == 0)
                {
                    txt_Command.Text = "";
                }
            }
        }

        public void Save_command(string cmd)
        {
            for (int i = Command_history.Length - 1; i >= 2; i--)
            {
                Command_history[i] = Command_history[i-1];
            }
            Command_history[1] = cmd;
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
                else if(notify_warn == 2)
                {
                    Warn(str_stack[0], int_stack[0]);
                }
                else
                {
                    Command(str_stack[0], int_stack[0]);
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

        public void New_Instance()
        {
            Path = Application.StartupPath + "\\\\Programs.xlsx";
            is_new_instance = true;
            /*if (excel_ready)
            {
                excel_ready = false;
                objBooks.Close(0);
                objApp.Quit();
            }
            string path = Application.StartupPath + "\\\\Programs.xlsx";
            Opening_Excel(path);*/
            tab_Control.Enabled = true;
        }

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
                }
            }
            else
            {
                Warn("File not loaded");
            }
            is_new_instance = false;
            tab_Control.Enabled = true;
        }

        public void Opening_Excel(string path)
        {
            objBooks = (Excel.Workbook)objApp.Workbooks.Open(path, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);

            objSheets = objBooks.Worksheets;
            Sheet_VyM = (Excel.Worksheet)objSheets.get_Item(1);
            range_1 = Sheet_VyM.get_Range("A1", "H137");
            cellValue_1 = (System.Object[,])range_1.get_Value();

            if ((cellValue_1[53, 1] != null) && (cellValue_1[53, 1].ToString() == "S-140 AGR-Technologies"))
            {
                Notify("File decoded correctly");

                Sheet_RP = (Excel.Worksheet)objSheets.get_Item(2);
                range_2 = Sheet_RP.get_Range("A1", "H70");
                cellValue_2 = (System.Object[,])range_2.get_Value();

                Sheet_PA = (Excel.Worksheet)objSheets.get_Item(3);
                range_3 = Sheet_PA.get_Range("A1", "H70");
                cellValue_3 = (System.Object[,])range_3.get_Value();

                Notify("Path: " + path);
                excel_ready = true;
                //open_DB();
                //Process_read();
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

        public void Process_read()
        {
            start_time = dateTime;
            if (excel_ready)
            {
                Fill_cbx();
                Get_Meetings();
                VyM_Handler(true);
                RP_Handler(true);
                AC_Handler(true);
            }
        }

        private void Process_save(int save)
        {
            string FileName = "";
            FileName += CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m_mes);
            if (true)
            {
                Opening_Excel(Path);
                if ((save == 1) || (save == 4))
                {
                    VyM_Handler(false);
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
                Notify(Application.StartupPath + FileName + ".xlsx");
                if (is_new_instance)
                {
                    objBooks.SaveAs(Application.StartupPath + FileName + ".xlsx");
                }
                else
                {
                    objBooks.Save();
                }
                if (excel_ready)
                {
                    excel_ready = false;
                    objBooks.Close(0);
                    objApp.Quit();
                }
                Notify("Saved file for JW Meetings" + ", Week [" + m_semana.ToString() + "]");
                Notify("Saved date: [" + m_dia.ToString() + "-" + m_mes.ToString() + "-" + m_año.ToString() + "]");
                Check_time(this, null);
            }
        }

        private string Check_null_cbx(object sender)
        {
            string str = "";
            ComboBox cbx = (ComboBox)sender;
            if (cbx.SelectedItem == null)
            {
                str = "";
            }
            else
            {
                str = cbx.SelectedItem.ToString();
            }
            return str;
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

        private void Cmbx_Week_SelectedIndexChanged(object sender, EventArgs e) 
        {
            if (excel_ready)
            {
                Process_read();
                //cbx_Day_SelectedIndexChanged(cbx_Day, null);
            }
        }
               
        private void Process_restore()
        {
            Notify("Restore info");
            Process_read();
        }

        public void Get_Meetings()
        {
           

        }

        public void VyM_Handler(bool read)
        {
            Notify((read ? "Reading": "Saving") + " VyM meeting");
            if (read)
            {
                int primary_cell = Get_cell();
                cellValue_1 = (System.Object[,])range_1.get_Value();
                txt_Date.Text = Check_null_string(cellValue_1[primary_cell, A]);
                txt_Pres.Text = Check_null_string(cellValue_1[primary_cell, G]);
                txt_CSA.Text = Check_null_string(cellValue_1[primary_cell + 1, G]);
                Compare_cbx_string(Check_null_string(cellValue_1[primary_cell + 2, G]), cbx_Ora1VyM);
                txt_TdlB_1.Text = Check_null_string(cellValue_1[primary_cell + 6, C]);
                txt_TdlB_A1.Text = Check_null_string(cellValue_1[primary_cell + 6, G]);
                txt_TdlB_A2.Text = Check_null_string(cellValue_1[primary_cell + 7, G]);
                txt_TdlB_A3.Text = Check_null_string(cellValue_1[primary_cell + 8, G]);
                txt_TdlB_B3.Text = Check_null_string(cellValue_1[primary_cell + 8, F]);
                txt_SMM1.Text = Check_null_string(cellValue_1[primary_cell + 11, C]);
                txt_SMM_A1.Text = Check_null_string(cellValue_1[primary_cell + 11, G]);
                txt_SMM_B1.Text = Check_null_string(cellValue_1[primary_cell + 11, F]);
                txt_SMM2.Text = Check_null_string(cellValue_1[primary_cell + 12, C]);
                txt_SMM_A2.Text = Check_null_string(cellValue_1[primary_cell + 12, G]);
                txt_SMM_B2.Text = Check_null_string(cellValue_1[primary_cell + 12, F]);
                txt_SMM3.Text = Check_null_string(cellValue_1[primary_cell + 13, C]);
                txt_SMM_A3.Text = Check_null_string(cellValue_1[primary_cell + 13, G]);
                txt_SMM_B3.Text = Check_null_string(cellValue_1[primary_cell + 13, F]);
                txt_NVC1.Text = Check_null_string(cellValue_1[primary_cell + 17, C]);
                txt_NVC_A1.Text = Check_null_string(cellValue_1[primary_cell + 17, G]);
                txt_NVC2.Text = Check_null_string(cellValue_1[primary_cell + 18, C]);
                txt_NVC_A2.Text = Check_null_string(cellValue_1[primary_cell + 18, G]);
                txt_NVC_A3.Text = Check_null_string(cellValue_1[primary_cell + 19, G]);
                Compare_cbx_string(Check_null_string(cellValue_1[primary_cell + 20, G]), cbx_NVC_A3L);
                Compare_cbx_string(Check_null_string(cellValue_1[primary_cell + 22, G]), cbx_Ora2VyM);
            }
            else
            {
                int primary_cell = Get_cell();
                Sheet_VyM.Cells[primary_cell, A] = Check_null(txt_Date).ToUpper();
                Sheet_VyM.Cells[primary_cell, G] = Check_null(txt_Pres);
                Sheet_VyM.Cells[primary_cell + 1, G] = Check_null(txt_CSA);

                Sheet_VyM.Cells[primary_cell + 2, G] = Check_null_cbx(cbx_Ora1VyM);
                Sheet_VyM.Cells[primary_cell + 6, C] = Check_null(txt_TdlB_1);
                Get_index_time(Sheet_VyM.get_Range("C" + (primary_cell + 6).ToString()));
                Sheet_VyM.Cells[primary_cell + 6, G] = Check_null(txt_TdlB_A1);
                Sheet_VyM.Cells[primary_cell + 7, G] = Check_null(txt_TdlB_A2);
                Sheet_VyM.Cells[primary_cell + 8, G] = Check_null(txt_TdlB_A3);
                Sheet_VyM.Cells[primary_cell + 8, F] = Check_null(txt_TdlB_B3);
                Sheet_VyM.Cells[primary_cell + 11, C] = Check_null(txt_SMM1);
                Get_index_time(Sheet_VyM.get_Range("C" + (primary_cell + 11).ToString()));
                Sheet_VyM.Cells[primary_cell + 11, G] = Check_null(txt_SMM_A1);
                Sheet_VyM.Cells[primary_cell + 11, F] = Check_null(txt_SMM_B1);
                Sheet_VyM.Cells[primary_cell + 12, C] = Check_null(txt_SMM2);
                Get_index_time(Sheet_VyM.get_Range("C" + (primary_cell + 12).ToString()));
                Sheet_VyM.Cells[primary_cell + 12, G] = Check_null(txt_SMM_A2);
                Sheet_VyM.Cells[primary_cell + 12, F] = Check_null(txt_SMM_B2);
                Sheet_VyM.Cells[primary_cell + 13, C] = Check_null(txt_SMM3);
                Get_index_time(Sheet_VyM.get_Range("C" + (primary_cell + 13).ToString()));
                Sheet_VyM.Cells[primary_cell + 13, G] = Check_null(txt_SMM_A3);
                Sheet_VyM.Cells[primary_cell + 13, F] = Check_null(txt_SMM_B3);
                Sheet_VyM.Cells[primary_cell + 17, C] = Check_null(txt_NVC1);
                Get_index_time(Sheet_VyM.get_Range("C" + (primary_cell + 17).ToString()));
                Sheet_VyM.Cells[primary_cell + 17, G] = Check_null(txt_NVC_A1);
                Sheet_VyM.Cells[primary_cell + 18, C] = Check_null(txt_NVC2);
                Get_index_time(Sheet_VyM.get_Range("C" + (primary_cell + 18).ToString()));
                Sheet_VyM.Cells[primary_cell + 18, G] = Check_null(txt_NVC_A2);
                Sheet_VyM.Cells[primary_cell + 19, G] = Check_null(txt_NVC_A3);
                Sheet_VyM.Cells[primary_cell + 20, G] = Check_null_cbx(cbx_NVC_A3L);
                Sheet_VyM.Cells[primary_cell + 22, G] = Check_null_cbx(cbx_Ora2VyM);
                //time
                Sheet_VyM.Cells[primary_cell + 2, A] = time_0.Text;
                Sheet_VyM.Cells[primary_cell + 3, A] = time_1.Text;
                Sheet_VyM.Cells[primary_cell + 6, A] = time_2.Text;
                Sheet_VyM.Cells[primary_cell + 7, A] = time_3.Text;
                Sheet_VyM.Cells[primary_cell + 8, A] = time_4.Text;
                Sheet_VyM.Cells[primary_cell + 11, A] = time_5.Text;
                Sheet_VyM.Cells[primary_cell + 12, A] = time_6.Text;
                Sheet_VyM.Cells[primary_cell + 13, A] = time_7.Text;
                Sheet_VyM.Cells[primary_cell + 16, A] = time_8.Text;
                Sheet_VyM.Cells[primary_cell + 17, A] = time_9.Text;
                Sheet_VyM.Cells[primary_cell + 18, A] = time_10.Text;
                Sheet_VyM.Cells[primary_cell + 19, A] = time_11.Text;
                Sheet_VyM.Cells[primary_cell + 21, A] = time_12.Text;
                Sheet_VyM.Cells[primary_cell + 22, A] = time_13.Text;
            }
        }

        public void RP_Handler(bool read)
        {
            Notify((read ? "Reading" : "Saving") + " RP meeting");
            if (read)
            {
                cellValue_2 = (System.Object[,])range_2.get_Value();
                int primary_cell = 4;
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_1);
                txt_RP_Speech_1.Text = Check_null_string(cellValue_2[primary_cell + 2, D]);
                txt_RP_Disc_1.Text = Check_null_string(cellValue_2[primary_cell + 2, H]);
                txt_RP_Cong_1.Text = Check_null_string(cellValue_2[primary_cell + 3, E]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 5, H]), cbx_CondAtly_1);
                txt_AdlA_Title_1.Text = Check_null_string(cellValue_2[primary_cell + 6, D]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 6, H]), cbx_LectRP_1);
                txt_Sal_Disc_1.Text = Check_null_string(cellValue_2[primary_cell + 10, C]);
                txt_Sal_Title_1.Text = Check_null_string(cellValue_2[primary_cell + 10, E]);
                txt_Sal_Cong_1.Text = Check_null_string(cellValue_2[primary_cell + 10, H]);
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_1);

                primary_cell = 17;
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 1, H]), cbx_PresRP_2);
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
                Compare_cbx_string(Check_null_string(cellValue_2[primary_cell + 7, H]), cbx_OraRP_5);
            }
            else
            {
                int primary_cell = 4;
                Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_1);
                Sheet_RP.Cells[primary_cell + 2, D] = Check_null(txt_RP_Speech_1);
                Sheet_RP.Cells[primary_cell + 2, H] = Check_null(txt_RP_Disc_1);
                Sheet_RP.Cells[primary_cell + 3, E] = Check_null(txt_RP_Cong_1);
                Sheet_RP.Cells[primary_cell + 5, H] = Check_null_cbx(cbx_CondAtly_1);
                Sheet_RP.Cells[primary_cell + 6, D] = Check_null(txt_AdlA_Title_1);
                Sheet_RP.Cells[primary_cell + 6, H] = Check_null_cbx(cbx_LectRP_1);
                Sheet_RP.Cells[primary_cell + 7, H] = Check_null_cbx(cbx_OraRP_1);
                Sheet_RP.Cells[primary_cell + 10, C] = txt_Sal_Disc_1.Text;
                Sheet_RP.Cells[primary_cell + 10, E] = txt_Sal_Title_1.Text;
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong_1.Text;

                primary_cell = 17;
                Sheet_RP.Cells[primary_cell + 1, H] = Check_null_cbx(cbx_PresRP_2);
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
                Sheet_RP.Cells[primary_cell + 10, H] = txt_Sal_Cong_5.Text;
            }
        }

        public void AC_Handler(bool read)
        {
            Notify((read ? "Reading" : "Saving") + " AC program");
            if (read)
            {
                cellValue_3 = (System.Object[,])range_3.get_Value();
                int primary_cell = 5;
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 1, A]), cbx_Aseo_1);
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
                Compare_cbx_string(Check_null_string(cellValue_3[primary_cell + 3, E]), cbx_AC2_S_5);
            }
            else
            {
                int primary_cell = 5;
                Sheet_PA.Cells[primary_cell + 1, A] = Check_null_cbx(cbx_Aseo_1);
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
                Sheet_PA.Cells[primary_cell + 3, E] = Check_null_cbx(cbx_AC2_S_5);
            }
        }

        public void Compare_cbx_string(object cell_value, object sender)
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
        }

        public void Fill_cbx()  
        {
            cbx_Ora1VyM.Items.Clear();
            cbx_NVC_A3L.Items.Clear();
            cbx_Ora2VyM.Items.Clear();

            for (int i = 0; i <= DB_Form.Generals.Count - 1; i++)
            {
                cbx_Ora1VyM.Items.Add(DB_Form.Generals[i].Nombre);
                cbx_NVC_A3L.Items.Add(DB_Form.Generals[i].Nombre);
                cbx_Ora2VyM.Items.Add(DB_Form.Generals[i].Nombre);
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
            if (tab_Control.SelectedIndex == 0)
            {
                Presenter(p.Executor);
                Notify("Overview");
            }
            else if (tab_Control.SelectedIndex == 1)
            {
                Presenter(p.Artanis);
                Notify("Section 'Reunion Publica y analisis de La Atalaya'");
            }
            else
            {
                Presenter(p.Oracle);
                Notify("Section 'Acomodadores'");
            }
        }

        private void Check_time(object sender, EventArgs e)
        {
            dateTime = new DateTime(2018, 1, 1, 7, 00, 00);
            time_0.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(5);
            time_1.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(3);
            time_2.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(10);
            time_3.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(8);
            time_4.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(5+1); //adjusting to real time
            time_5.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(Analyze_string(txt_SMM1.Text, true));
            time_6.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(Analyze_string(txt_SMM2.Text, true));
            time_7.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(Analyze_string(txt_SMM3.Text, true) + 1); //adjusting to real time
            time_8.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(3);
            time_9.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(Analyze_string(txt_NVC1.Text, false));
            if ((txt_NVC2.Text == null) || (txt_NVC2.Text == " ") || (txt_NVC2.Text == " -"))
            {
                time_10.Text = " ";                
            }
            else
            {
                time_10.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
                dateTime = dateTime.AddMinutes(Analyze_string(txt_NVC2.Text, false));
            }
            dateTime = dateTime.AddMinutes(1); //adjusting to real time
            time_11.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(30);
            time_12.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            dateTime = dateTime.AddMinutes(3);
            time_13.Text = dateTime.Hour.ToString() + ":" + dateTime.Minute.ToString();
            if (dateTime.Hour == 8 && dateTime.Minute == 40)
            {
                time_13.ForeColor = Color.Green;
            }
            else
            {
                time_13.ForeColor = Color.Red;
            }
        }

        public int Analyze_string(string Str, bool SMM)
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

        private void Txt_Date_TextChanged(object sender, EventArgs e)
        {
            TextBox txtbx = (TextBox)sender;
            string Str = txtbx.Text.ToLower();
            var array = txtbx.Text.ToCharArray();
            int result = 0;
            bool converted = false;
            for (int i = 0; i <= month.Length-1; i++)
            {
                if (Str.Contains(month[i]))
                {
                    m_mes = i + 1;
                    break;
                }
            }
            if (array.Length >= 3)
            {
                if (array[1] == ' ')
                {
                    converted = int.TryParse(array[0].ToString(), out result);
                    if (converted)
                    {
                        Notify("Detecting Day [" + result.ToString() + "]");
                        m_dia = result;
                        Set_date();
                    }
                }
                else if (array[2] == ' ')
                {
                    converted = int.TryParse(array[0].ToString() + array[1].ToString(), out result);
                    if (converted)
                    {
                        Notify("Detecting Day [" + result.ToString() + "]");
                        m_dia = result;
                        Set_date();
                    }
                }
            }
        }

        private void Set_date()
        {
            int checksum_aux = 0;
            checksum_aux = m_año + m_mes + m_dia;
            lbl_Month.Text = "Mes: " + month[m_mes - 1];
            if (checksum_aux != date_checksum)
            {
                date_checksum = checksum_aux;
                if ((m_año != 0) && (m_mes != 0) && (m_dia != 0))
                {
                    date = new DateTime(m_año, m_mes, m_dia);
                    Calendar.SetDate(date);
                    Notify("Date Set in: [" + m_dia.ToString() + "-" + m_mes.ToString() + "-" + m_año.ToString() + "]");
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

        public void save_DB(object sender)
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
                        Sheet_DB.Cells[i, column] = s;
                        break;
                    }
                }
            }
        }

        private void Autofill_Handler()
        {

        }
    }
}
