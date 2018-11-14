using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.IO;

namespace Project_Insight
{
    public partial class DB_Form : Form
    {
        public List<DB_Eld> Elders = new List<DB_Eld>();
        public List<DB_Mns> Ministerials = new List<DB_Mns>();
        public List<DB_Gnr> Generals = new List<DB_Gnr>();
        public List<DB_Cln> Cleaners = new List <DB_Cln>();
        private StreamReader sr;
        private StreamWriter wr;
        private int lenght = File.ReadAllLines(Application.StartupPath + "\\\\DB.csv").Length;

        public DB_Form()
        {
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            InitializeComponent();
        }

        private void DB_Form_Load(object sender, EventArgs e)
        {

            //implementing new DB logic 
            Refresh_2_0();
            //Refresh_DB();
            Send_messsage("Opening DB");
            timer_refresh.Enabled = true;
            
        }

        private void Refresh_2_0()
        {
            //open CSV file
            Elders.Clear();
            Ministerials.Clear();
            Generals.Clear();
            Cleaners.Clear();

            sr = new StreamReader(Application.StartupPath + "\\\\DB.csv", Encoding.UTF8, false);
            //wr = new StreamWriter(Application.StartupPath + "\\\\DB.csv");
            int section = 0;
            bool readable = false;
            string temp = "";
            for (int i = 0; i <= lenght-1; i++)
            {
                temp = sr.ReadLine();
                temp = temp.Replace("�", "ñ");
                if (temp.Contains("end"))
                {
                    section++;
                    readable = false;
                }
                else
                {
                    readable = true;
                }
                if (readable)
                {
                    switch (section)
                    {
                        case 0:
                            {
                                string[] data = temp.Split(',');
                                Elders.Add(new DB_Eld(data[0], data[1], data[3], data[4], data[5], data[6]));
                                break;
                            }
                        case 1:
                            {
                                string[] data = temp.Split(',');
                                Ministerials.Add(new DB_Mns(data[0], data[1], data[2], data[3], data[4], data[5]));
                                break;
                            }
                        case 2:
                            {
                                string[] data = temp.Split(',');
                                Generals.Add(new DB_Gnr(data[0], data[2], data[4], data[7], data[8]));
                                break;
                            }
                        case 3:
                            {
                                string[] data = temp.Split(',');
                                Cleaners.Add(new DB_Cln(data[0], data[1]));
                                break;
                            }
                    }
                }
            }
            Eld_Grid.DataSource = Elders;
            Min_Grid.DataSource = Ministerials;
            Gen_Grid.DataSource = Generals;
            Cln_Grid.DataSource = Cleaners;
           /* Eld_Grid.AutoSize = true;
            Min_Grid.AutoSize = true;
            Gen_Grid.AutoSize = true;
            Cln_Grid.AutoSize = true;*/
            Eld_Grid.Refresh();
            Min_Grid.Refresh();
            Gen_Grid.Refresh();
            Cln_Grid.Refresh();
            sr.Close();
        }

        
        private void DB_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            Main_Form.DB_form_show = false;
            Send_messsage("Closing DB");
        }

        public  void Refresh_DB()
        {
            Eld_Grid.Rows.Clear();
            string[] row = new string[40];
            for (int j = 0; j <= 37; j++)
            {
                for (int i = 2; i <= 10; i++)
                {
                    if (Main_Form.cellValue_4[j + 5, i] != null)
                    {
                        row[i - 2] = Main_Form.cellValue_4[j + 5, i].ToString();
                    }
                    else
                    {
                        row[i - 2] = "-";
                    }

                }
                Eld_Grid.Rows.Insert(j, row);
            }
        }

        public async void Send_messsage(string Message)
        {
            Main_Form.message_form2 = Message;
            await Task.Delay(100);
        }

        private void timer_refresh_Tick(object sender, EventArgs e)
        {
            if (Main_Form.pending_refresh_DB)
            {
                //Refresh_DB();
                Main_Form.pending_refresh_DB = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Refresh_2_0();
        }

        private void btn_Hide_Click(object sender, EventArgs e)
        {
            Main_Form.DB_form_show = false;
            Send_messsage("Hiding DB");
            this.Hide();
        }
    }
}
