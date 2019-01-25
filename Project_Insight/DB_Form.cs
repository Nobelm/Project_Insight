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

        private StreamWriter wr;
        public string Path_CSV = Application.StartupPath + "\\\\DB.csv";

        public DB_Form()
        {
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            InitializeComponent();
        }

        private void DB_Form_Load(object sender, EventArgs e)
        {
            Read_CSV();
        }


        private void DB_Form_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        public async void Read_CSV()
        {
            StreamReader reader;
            int lenght = File.ReadAllLines(Path_CSV).Length;
            string temp = "";
            await Task.Delay(10);
            reader = new StreamReader(Path_CSV);
            for(int i = 0; i < lenght; i++)
            {
                temp = reader.ReadLine();
                string[] data = temp.Split(',');

                Elders.Add(new DB_Eld(data[0], data[1], data[2], data[3], data[4], data[5]));
            }

            reader.Close();
            Eld_Grid.DataSource = Elders;
            Eld_Grid.Refresh();

        }

    }
}
