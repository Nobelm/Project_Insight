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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Project_Insight
{
    public partial class DB_Form : Form
    {
        private Excel.Application DBApp;
        private Excel.Workbook DBBooks = null;
        private Excel.Sheets DBSheets;
        private Excel.Worksheet Sheet_DB;
        private Excel.Range range_1;
        private object[,] cellValue_1 = null;
        public static bool db_open = false;

        public static List<DB_Eld> Elders = new List<DB_Eld>();
        public static List<DB_Mns> Ministerials = new List<DB_Mns>();
        public static List<DB_Gnr> Generals = new List<DB_Gnr>();

        public delegate void Updater();

        public string Path_DB = Application.StartupPath + "\\\\DB.xlsx";

        public DB_Form()
        {
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            InitializeComponent();
            Event_Handler.DB_EXCEL_LOAD += DB_Control;
        }

        private void DB_Form_Load(object sender, EventArgs e)
        {
            Thread DB_open_thread = new Thread(() => Open_DB());
            DB_open_thread.Start();
        }

        private void DB_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            DBBooks.Close(0);
            DBApp.Quit();
            Marshal.ReleaseComObject(Sheet_DB);
            Marshal.ReleaseComObject(DBBooks);
            Marshal.ReleaseComObject(DBApp);
        }

        public void DB_Control(bool save)
        {
            if (save)
            {
                Write_DB();
            }
            else
            {
                Read_DB();
            }
        }

        public void Open_DB()
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "DB_Open_Thread";
                Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            }
            db_open = true;
            DBApp = new Excel.Application();
            DBBooks = DBApp.Workbooks.Open(Path_DB, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
            DBSheets = DBBooks.Worksheets;
            Sheet_DB = (Excel.Worksheet)DBSheets.get_Item(1);
            range_1 = Sheet_DB.get_Range("A1", "G137");
            cellValue_1 = range_1.get_Value();
            Event_Handler.Db_Excel_Load(false);
        }

        public void Read_DB()
        {
            bool read;
            string data = "";
            short iterator = 0;
            for (int i = 1; i < 50; i++)
            {
                if (cellValue_1[i, 1] == null)
                {
                    break;
                }
                data = cellValue_1[i, 1].ToString();
                read = true;
                if (data == "end section")
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
                                Elders.Add(new DB_Eld(cellValue_1[i, 1].ToString(), cellValue_1[i, 2].ToString(), cellValue_1[i, 3].ToString(), cellValue_1[i, 4].ToString(), cellValue_1[i, 5].ToString(), cellValue_1[i, 6].ToString(), cellValue_1[i, 7].ToString()));
                                break;
                            }
                        case 1:
                            {
                                Ministerials.Add(new DB_Mns(cellValue_1[i, 1].ToString(), cellValue_1[i, 2].ToString(), cellValue_1[i, 3].ToString(), cellValue_1[i, 4].ToString(), cellValue_1[i, 5].ToString(), cellValue_1[i, 7].ToString(), cellValue_1[i, 7].ToString()));
                                break;
                            }
                        case 2:
                            {
                                Generals.Add(new DB_Gnr(cellValue_1[i, 1].ToString(), cellValue_1[i, 2].ToString(), cellValue_1[i, 3].ToString(), cellValue_1[i, 4].ToString(), cellValue_1[i, 5].ToString()));
                                break;
                            }
                    }
                }
            }
            Refresh_Grid();
            /*Message: "Read Succesfull"*/
        }

        public void Write_DB()
        {
            range_1.NumberFormat = "dd/mm/yyyy";
            int aux = 1, i = 0;
            for (i = 0; i < Elders.Count; i++)
            {
                Sheet_DB.Cells[i + aux, 1] = '\'' + Elders[i].Nombre;
                Sheet_DB.Cells[i + aux, 2] = '\'' + Elders[i].Capitan.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 3] = '\'' + Elders[i].Pres_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 4] = '\'' + Elders[i].Lec_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 5] = '\'' + Elders[i].Ora_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 6] = '\'' + Elders[i].Atalaya .ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 7] = '\'' + Elders[i].Cpt_Aseo.ToString("dd/MM/yyyy");
            }
            aux += i++;
            Sheet_DB.Cells[aux, 1] = "end section";
            aux++;
            for (i = 0; i < Ministerials.Count; i++)
            {
                Sheet_DB.Cells[i + aux, 1] = '\'' + Ministerials[i].Nombre;
                Sheet_DB.Cells[i + aux, 2] = '\'' + Ministerials[i].Capitan.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 3] = '\'' + Ministerials[i].Acom.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 4] = '\'' + Ministerials[i].Pres_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 5] = '\'' + Ministerials[i].Lec_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 6] = '\'' + Ministerials[i].Ora_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 7] = '\'' + Ministerials[i].Cpt_Aseo.ToString("dd/MM/yyyy");
            }
            aux += i++;
            Sheet_DB.Cells[aux, 1] = "end section";
            aux++;
            for (i = 0; i < Generals.Count; i++)
            {
                Sheet_DB.Cells[i + aux, 1] = '\'' + Generals[i].Nombre;
                Sheet_DB.Cells[i + aux, 2] = '\'' + Generals[i].Acom.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 3] = '\'' + Generals[i].Lec_RP.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 4] = '\'' + Generals[i].Lec_VyM.ToString("dd/MM/yyyy");
                Sheet_DB.Cells[i + aux, 5] = '\'' + Generals[i].Ora_VyM.ToString("dd/MM/yyyy");
            }

            DBBooks.Save();
            DB_Control(false);
        }

        public void Persistence_VyM(VyM_Sem sem, DateTime date)
        {
            for (int i = 0; i < Generals.Count; i++)
            {
                if (Generals[i].Nombre == sem.Libro_L)
                {
                    Generals[i].Lec_VyM = date;
                }
                else if (Generals[i].Nombre == sem.Oracion)
                {
                    Generals[i].Ora_VyM = date;
                }
            }
            Refresh_Grid();
        }

        public void Persistence_RP(RP_Sem sem, DateTime date)
        {
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
            Refresh_Grid();
        }

        public void Persistence_AC(AC_Sem sem, DateTime date_vym, DateTime date_rp)
        {
            for (int i = 0; i < Elders.Count; i++)
            {
                if (Elders[i].Nombre == sem.Vym_Cap)
                {
                    Elders[i].Capitan = date_vym;
                }
                if (Elders[i].Nombre == sem.Rp_Cap)
                {
                    Elders[i].Capitan = date_rp;
                }
                if (Elders[i].Nombre == sem.Cp_Aseo_VyM)
                {
                    Elders[i].Cpt_Aseo = date_vym;
                }
                if (Elders[i].Nombre == sem.Cp_Aseo_RP)
                {
                    Elders[i].Cpt_Aseo = date_rp;
                }
            }
            for (int i = 0; i < Ministerials.Count; i++)
            {
                if (Ministerials[i].Nombre == sem.Vym_Cap)
                {
                    Ministerials[i].Capitan = date_vym;
                }
                if (Ministerials[i].Nombre == sem.Rp_Cap)
                {
                    Ministerials[i].Capitan = date_rp;
                }
                if (Ministerials[i].Nombre == sem.Cp_Aseo_VyM)
                {
                    Ministerials[i].Cpt_Aseo = date_vym;
                }
                if (Ministerials[i].Nombre == sem.Cp_Aseo_RP)
                {
                    Ministerials[i].Cpt_Aseo = date_rp;
                }
                if (Ministerials[i].Nombre == sem.Rp_Der || Ministerials[i].Nombre == sem.Rp_Izq)
                {
                    Ministerials[i].Acom = date_rp;
                }
                if (Ministerials[i].Nombre == sem.Vym_Der || Ministerials[i].Nombre == sem.Vym_Izq)
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
                if (Generals[i].Nombre == sem.Vym_Der || Generals[i].Nombre == sem.Vym_Izq)
                {
                    Generals[i].Acom = date_vym;
                }
            }
            Refresh_Grid();
        }

        public void Refresh_Grid()
        {
            if (Eld_Grid.InvokeRequired)
            {
                Updater updater = Refresh_Grid;
                Invoke(updater);
            }
            else
            {
                Eld_Grid.DataSource = Elders;
                Min_Grid.DataSource = Ministerials;
                Gen_Grid.DataSource = Generals;
                Eld_Grid.Refresh();
                Min_Grid.Refresh();
                Gen_Grid.Refresh();
            }
        }
    }
}
