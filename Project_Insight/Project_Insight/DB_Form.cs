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
        public static List<DB_Eld> Elders = new List<DB_Eld>();
        public static List<DB_Mns> Ministerials = new List<DB_Mns>();
        public static List<DB_Gnr> Generals = new List<DB_Gnr>();
        //public List<DB_Cln> Cleaners = new List <DB_Cln>();
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
            bool read;
            int lenght = File.ReadAllLines(Path_CSV).Length;
            string temp = "";
            await Task.Delay(10);
            reader = new StreamReader(Path_CSV);
            short iterator = 0;
            for (int i = 0; i < lenght; i++)
            {
                read = true;
                temp = reader.ReadLine();
                string[] data = temp.Split(',');
                if (data[0] == "end section")
                {
                    iterator++;
                    read = false;
                }
                if (read)
                {
                    switch (iterator)
                    {
                        case 0:
                            {
                                Elders.Add(new DB_Eld(data[0], data[1], data[2], data[3], data[4], data[5], data[6]));
                                break;
                            }
                        case 1:
                            {
                                Ministerials.Add(new DB_Mns(data[0], data[1], data[2], data[3], data[4], data[5], data[6]));
                                break;
                            }
                        case 2:
                            {
                                Generals.Add(new DB_Gnr(data[0], data[1], data[2], data[3], data[4]));
                                break;
                            }
                    }
                }
            }

            reader.Close();
            Eld_Grid.DataSource = Elders;
            Min_Grid.DataSource = Ministerials;
            Gen_Grid.DataSource = Generals;
            Eld_Grid.Refresh();
            Min_Grid.Refresh();
            Gen_Grid.Refresh();
            /*Message: "Read Succesfull"*/
        }

        public async void Persistence_VyM(VyM_Sem sem, string date)
        {
            await Task.Delay(10);
            for (int i = 0; i < Generals.Count; i++)
            {
                if (Generals[i].Nombre == sem.Libro_L)
                {
                    Generals[i].Libro_L = date;
                }
                else if (Generals[i].Nombre == sem.Oracion)
                {
                    Generals[i].Oracion_VyM = date;
                }
            }
        }

        public async void Persistence_RP(RP_Sem sem, string date)
        {
            await Task.Delay(10);
            for (int i = 0; i < Elders.Count; i++)
            {
                if (Elders[i].Nombre == sem.Presidente)
                {
                    Elders[i].Pres_RP = date;
                }
                else if (Elders[i].Nombre == sem.Lector)
                {
                    Elders[i].Lec_RP = date;
                }
                else if (Elders[i].Nombre == sem.Oracion)
                {
                    Elders[i].Ora_RP = date;
                }
                else if (Elders[i].Nombre == sem.Conductor)
                {
                    Elders[i].Atalaya = date;
                }
            }
            for (int i = 0; i < Ministerials.Count; i++)
            {
                if (Ministerials[i].Nombre == sem.Presidente)
                {
                    Ministerials[i].Pres_RP = date;
                }
                else if (Ministerials[i].Nombre == sem.Lector)
                {
                    Ministerials[i].Lec_RP = date;
                }
                else if (Ministerials[i].Nombre == sem.Oracion)
                {
                    Ministerials[i].Ora_RP = date;
                }
            }
            for (int i = 0; i < Generals.Count; i++)
            {
                if (Generals[i].Nombre == sem.Lector)
                {
                    Generals[i].Lec_RP = date;
                }
            }
        }

        public async void Persistence_AC(AC_Sem sem, string date_vym, string date_rp)
        {
            await Task.Delay(10);
            for (int i = 0; i < Elders.Count; i++)
            {
                if (Elders[i].Nombre == sem.Vym_Cap)
                {
                    Elders[i].Capitan = date_vym;
                }
                else if (Elders[i].Nombre == sem.Rp_Cap)
                {
                    Elders[i].Capitan = date_rp;
                }
                else if (Elders[i].Nombre == sem.Cp_Aseo_VyM)
                {
                    Elders[i].Cpt_Aseo = date_vym;
                }
                else if (Elders[i].Nombre == sem.Cp_Aseo_RP)
                {
                    Elders[i].Capitan = date_rp;
                }
            }
            for (int i = 0; i < Ministerials.Count; i++)
            {
                if (Ministerials[i].Nombre == sem.Vym_Cap)
                {
                    Ministerials[i].Capitan = date_vym;
                }
                else if (Ministerials[i].Nombre == sem.Rp_Cap)
                {
                    Ministerials[i].Capitan = date_rp;
                }
                else if (Ministerials[i].Nombre == sem.Cp_Aseo_VyM)
                {
                    Ministerials[i].Cpt_Aseo = date_vym;
                }
                else if (Ministerials[i].Nombre == sem.Cp_Aseo_RP)
                {
                    Ministerials[i].Cpt_Aseo = date_rp;
                }
                else if (Ministerials[i].Nombre == sem.Rp_Der || Ministerials[i].Nombre == sem.Rp_Izq)
                {
                    Ministerials[i].Acom = date_rp;
                }
                else if (Ministerials[i].Nombre == sem.Vym_Der || Ministerials[i].Nombre == sem.Vym_Izq)
                {
                    Ministerials[i].Acom = date_vym;
                }
            }
            for (int i = 0; i < Generals.Count; i++)
            {
                if (Generals[i].Nombre == sem.Rp_Der || Generals[i].Nombre == sem.Rp_Izq)
                {
                    Generals[i].Acom = date_rp;
                }
                else if (Generals[i].Nombre == sem.Vym_Der || Generals[i].Nombre == sem.Vym_Izq)
                {
                    Generals[i].Acom = date_vym;
                }
            }
        }
        
        public void Write_CSV()
        {
            StreamWriter writer = new StreamWriter(Path_CSV);
            for (int i = 0; i < Elders.Count; i++)
            {
                writer.WriteLine(Elders[i].Nombre + "," + Elders[i].Capitan + "," + Elders[i].Pres_RP + "," + Elders[i].Lec_RP + "," + Elders[i].Ora_RP + "," + Elders[i].Atalaya + "," + Elders[i].Cpt_Aseo);
            }
            writer.WriteLine("end section");
            for (int i = 0; i < Ministerials.Count; i++)
            {
                writer.WriteLine(Ministerials[i].Nombre + "," + Ministerials[i].Capitan + "," + Ministerials[i].Acom + "," + Ministerials[i].Pres_RP + "," + Ministerials[i].Lec_RP + "," + Ministerials[i].Ora_RP + "," + Ministerials[i].Cpt_Aseo);
            }
            writer.WriteLine("end section");
            for (int i = 0; i < Generals.Count; i++)
            {
                writer.WriteLine(Generals[i].Nombre + "," + Generals[i].Acom + "," + Generals[i].Lec_RP + "," + Generals[i].Libro_L + "," + Generals[i].Oracion_VyM);
            }
            writer.Close();
            /*Message: "Write Succesfull"*/
        }

        public void Edit_DB()
        {
            Eld_Grid.ReadOnly = false;
            Min_Grid.ReadOnly = false;
            Gen_Grid.ReadOnly = false;
        }

        public void Save_DB()
        {
            Elders = Eld_Grid.DataSource as List<DB_Eld>;
            Ministerials = Min_Grid.DataSource as List<DB_Mns>;
            Generals = Gen_Grid.DataSource as List<DB_Gnr>;
            Eld_Grid.ReadOnly = true;
            Min_Grid.ReadOnly = true;
            Gen_Grid.ReadOnly = true;
            Write_CSV();
        }

        public static string[] Get_VyM_Assigned(string ID)
        {
            string[] Str_name = new string[3];
            DateTime[] min_num = new DateTime[3];
            if (ID == "Libro_L")
            {
                min_num[0] = Convert.ToDateTime(Generals[0].Libro_L);
                for (int i = 0; i < Generals.Count; i++)
                {
                    if (0 > DateTime.Compare(Convert.ToDateTime(Generals[i].Libro_L), min_num[0]))
                    {
                        min_num[2] = min_num[1];
                        min_num[1] = min_num[0];
                        min_num[0] = Convert.ToDateTime(Generals[i].Libro_L);
                        Str_name[2] = Str_name[1];
                        Str_name[1] = Str_name[0];
                        Str_name[0] = Generals[i].Nombre;
                    }
                    else if ((min_num[1] != null) && (0 > DateTime.Compare(Convert.ToDateTime(Generals[i].Libro_L), min_num[1])))
                    {
                        min_num[2] = min_num[1];
                        min_num[1] = Convert.ToDateTime(Generals[i].Libro_L);
                        Str_name[2] = Str_name[1];
                        Str_name[1] = Generals[i].Nombre;
                    }
                    else if ((min_num[2] != null) && (0 > DateTime.Compare(Convert.ToDateTime(Generals[i].Libro_L), min_num[2])))
                    {
                        min_num[2] = Convert.ToDateTime(Generals[i].Libro_L);
                        Str_name[2] = Generals[i].Nombre;
                    }
                }
            }
            else if(ID == "Oracion")
            {
                min_num[0] = Convert.ToDateTime(Generals[0].Oracion_VyM);
                for (int i = 0; i < Generals.Count; i++)
                {
                    if (0 > DateTime.Compare(Convert.ToDateTime(Generals[i].Oracion_VyM), min_num[0]))
                    {
                        min_num[2] = min_num[1];
                        min_num[1] = min_num[0];
                        min_num[0] = Convert.ToDateTime(Generals[i].Oracion_VyM);
                        Str_name[2] = Str_name[1];
                        Str_name[1] = Str_name[0];
                        Str_name[0] = Generals[i].Nombre;
                    }
                    else if ((min_num[1] != null) && (0 > DateTime.Compare(Convert.ToDateTime(Generals[i].Oracion_VyM), min_num[1])))
                    {
                        min_num[2] = min_num[1];
                        min_num[1] = Convert.ToDateTime(Generals[i].Oracion_VyM);
                        Str_name[2] = Str_name[1];
                        Str_name[1] = Generals[i].Nombre;
                    }
                    else if ((min_num[2] != null) && (0 > DateTime.Compare(Convert.ToDateTime(Generals[i].Oracion_VyM), min_num[2])))
                    {
                        min_num[2] = Convert.ToDateTime(Generals[i].Oracion_VyM);
                        Str_name[2] = Generals[i].Nombre;
                    }
                }
            }
            return Str_name;
        }

        public void Get_Generals_Elements(DB_Gnr gnr)
        {

        }

        public static string[] Get_RP_Assigned(string ID)
        {
            string[] Str_name = new string[3];
            DateTime[] min_num = new DateTime[3];

            return Str_name;
        }
    }
}
