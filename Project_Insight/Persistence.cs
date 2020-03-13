using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Threading;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using System.Diagnostics;

namespace Project_Insight
{
    public class Persistence
    {
        public static Excel.Application DBApp;
        public static Excel.Workbook DBBooks = null;
        public static Excel.Sheets DBSheets;
        public static Excel.Worksheet Sheet_Stat;
        public static Excel.Range range_stat;
       // public static Excel.Range range_2;
        //private static object[,] cellValue_1 = null;
        private static object[,] cellValue_stat = null;
        public static bool db_open = false;
        public static bool attending_persistance = false;
        public static bool attending_db_request = false;
        private static bool Initial_Check = false;
        private static bool DB_Allowed = false;
        public static bool Close_Persistence = false;
        public static List<DB_Request> DB_Requests_List = new List<DB_Request>();
        public static List<Males> Male_Status_List = new List<Males>();
        public static List<Persistence_Request> Persistence_Requests_List = new List<Persistence_Request>();

        public enum DB_Request
        {
            read,
            write
        }

        public class Persistence_Request
        {
            public Insight_Sem persistence_insight;
        }

        /*-------------------- Initialize methods -------------------- */
        public static void Start_DataBase()
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Persistence";
                Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            }
            Persistence_Thread_Handler();
        }

        private static void Persistence_Thread_Handler()
        {
            while (true)
            {
                if (DB_Requests_List.Count > 0 && !attending_db_request && DB_Allowed)
                {
                    if (DB_Allowed)
                    {
                        attending_db_request = true;
                        DB_Hub(DB_Requests_List[0]);
                    }
                    else
                    {
                        Main_Form.Warn("Persistence functions disabled");
                        DB_Requests_List.RemoveAt(0);
                    }
                }
                if (Persistence_Requests_List.Count > 0 && !attending_persistance && DB_Allowed)
                {
                    Persistence_Hub(Persistence_Requests_List[0]);
                    Persistence_Requests_List.RemoveAt(0);
                }
                if (!Initial_Check)
                {
                    Initial_Check = true;
                    Initial_Database_check();
                }
                Thread.Sleep(1000);
            }
        }

        private static void Initial_Database_check()
        {
            if (File.Exists(Main_Form.Path_DB))
            {
                Main_Form.Notify("Initial Check: Database file exist");
                DB_Allowed = true;
            }
            else
            {
                Main_Form.Warn("Initial Check: Database file missing");
                Main_Form.Warn("Disabling Persistence features");
                DB_Allowed = false;
            }
        }

        /*-------------------- Attending request -------------------- */

        public static void DB_Hub(DB_Request _Request)
        {
            CultureInfo en = new CultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = en;
            if (!db_open)
            {
                Open_DB();
            }

            switch (_Request)
            {
                case DB_Request.read:
                    {
                        Read_DB();
                        break;
                    }
                case DB_Request.write:
                    {
                        Write_DB();
                        break;
                    }
            }
            DB_Requests_List.RemoveAt(0);
            attending_db_request = false;
        }

        public static void Open_DB()
        {
            db_open = true;
            Main_Form.Notify("Opening Database file");
            DBApp = new Excel.Application();
            DBBooks = DBApp.Workbooks.Open(Main_Form.Path_DB, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
            DBSheets = DBBooks.Worksheets;
            Sheet_Stat = (Excel.Worksheet)DBSheets.get_Item(1);
            range_stat = Sheet_Stat.get_Range("A1", "I137");
            cellValue_stat = range_stat.get_Value();
        }

        private static void Read_DB()
        {
            /*Read Status males*/
            Main_Form.Notify("Retrieve values from Status database");
            Main_Form.Elders_Count = 0;
            Main_Form.Ministerials_Count = 0;
            Main_Form.Generals_Count = 0;
            Main_Form.Males_Count = 0;
            for (int i = 1; i < 100; i++)
            {
                if (cellValue_stat[i, 1] == null)
                {
                    break;
                }
                else
                {
                    Males aux_male = new Males
                    {
                        Name       = cellValue_stat[i, 1].ToString(),
                        Atalaya    = cellValue_stat[i, 2].ToString(),
                        Capitan    = cellValue_stat[i, 3].ToString(),
                        Acomodador = cellValue_stat[i, 4].ToString(),
                        Lector     = cellValue_stat[i, 5].ToString(),
                        Pres_RP    = cellValue_stat[i, 6].ToString(),
                        Oracion    = cellValue_stat[i, 7].ToString(),
                        male_type  = (Main_Form.Male_Type)Convert.ToInt16(cellValue_stat[i, 8].ToString()),
                    };
                    switch (aux_male.male_type)
                    {
                        case Main_Form.Male_Type.Anciano:
                            {
                                Main_Form.Elders_Count++;
                                break;
                            }
                        case Main_Form.Male_Type.Ministerial:
                            {
                                Main_Form.Ministerials_Count++;
                                break;
                            }
                        case Main_Form.Male_Type.Publicador:
                            {
                                Main_Form.Generals_Count++;
                                break;
                            }
                    }
                    Main_Form.Male_List.Add(aux_male);
                }
            }
            Main_Form.Male_List_filled = true;
            Males_Rules_Handler();
            Main_Form.Pending_refresh_status_grids = true;
            Main_Form.Males_Count = Main_Form.Elders_Count + Main_Form.Ministerials_Count + Main_Form.Generals_Count;
            Main_Form.Notify("Read Successfull:\nElders: " + Main_Form.Elders_Count.ToString() + "\nMinisterials: " + Main_Form.Ministerials_Count.ToString() + "\nGeneral Males: " + Main_Form.Generals_Count.ToString() + "\nMales Count: " + Main_Form.Males_Count);
        }

        public static void Males_Rules_Handler()
        {
            for (int i = 0; i < Main_Form.Male_List.Count; i++)
            {
                switch (Main_Form.Male_List[i].male_type)
                {
                    case Main_Form.Male_Type.Anciano:
                        {
                            Main_Form.Male_List[i] = Set_Status(Main_Form.Rule_Elders, Main_Form.Male_List[i]);
                            break;
                        }
                    case Main_Form.Male_Type.Ministerial:
                        {
                            Main_Form.Male_List[i] = Set_Status(Main_Form.Rule_Ministerials, Main_Form.Male_List[i]);
                            break;
                        }
                    case Main_Form.Male_Type.Publicador:
                        {
                            Main_Form.Male_List[i] = Set_Status(Main_Form.Rule_Generals, Main_Form.Male_List[i]);
                            break;
                        }
                }
            }
        }

        public static Males Set_Status(Males local_rule, Males male)
        {
            male.Atalaya    = Sub_State_Set(local_rule.Atalaya, male.Atalaya);
            male.Capitan    = Sub_State_Set(local_rule.Capitan, male.Capitan);
            male.Acomodador = Sub_State_Set(local_rule.Acomodador, male.Acomodador);
            male.Lector     = Sub_State_Set(local_rule.Lector, male.Lector);
            male.Pres_RP    = Sub_State_Set(local_rule.Pres_RP, male.Pres_RP);
            male.Oracion    = Sub_State_Set(local_rule.Oracion, male.Oracion);
            male.male_type  = local_rule.male_type;
            return male;
        }

        public static string Sub_State_Set(string m_Rule, string m_State)
        {
            if (m_Rule.Equals("Allowed"))
            {
                if (m_State == null)
                {
                    m_State = "";
                }
                if (m_State.Equals("Blocked"))
                {
                    m_State = "Blocked";
                }
                else if (!m_State.Contains('/'))
                {
                    DateTime date = new DateTime(2019, 01, 01);
                    m_State = date.ToString("dd/MM/yyyy");
                }
                else
                {
                    DateTime date = Convert.ToDateTime(m_State);
                    m_State = date.ToString("dd/MM/yyyy");
                }
            }
            else
            {
                m_State = "Non_Status";
            }
            return m_State;
        }

        private static void Write_DB()
        {
            Main_Form.Notify("Update and saving values in database");
            int j;
            for (j = 1; j <= Main_Form.Male_List.Count; j++)
            {
                Sheet_Stat.Cells[j, 1] = '\'' + Main_Form.Male_List[j - 1].Name;
                Sheet_Stat.Cells[j, 2] = '\'' + Main_Form.Male_List[j - 1].Atalaya;
                Sheet_Stat.Cells[j, 3] = '\'' + Main_Form.Male_List[j - 1].Capitan;
                Sheet_Stat.Cells[j, 4] = '\'' + Main_Form.Male_List[j - 1].Acomodador;
                Sheet_Stat.Cells[j, 5] = '\'' + Main_Form.Male_List[j - 1].Lector;
                Sheet_Stat.Cells[j, 6] = '\'' + Main_Form.Male_List[j - 1].Pres_RP;
                Sheet_Stat.Cells[j, 7] = '\'' + Main_Form.Male_List[j - 1].Oracion;
                Sheet_Stat.Cells[j, 8] = Main_Form.Male_List[j - 1].male_type;
            }
            while (cellValue_stat[j, 1] != null)
            {
                Sheet_Stat.Cells[j, 1] = "";
                Sheet_Stat.Cells[j, 2] = "";
                Sheet_Stat.Cells[j, 3] = "";
                Sheet_Stat.Cells[j, 4] = "";
                Sheet_Stat.Cells[j, 5] = "";
                Sheet_Stat.Cells[j, 6] = "";
                Sheet_Stat.Cells[j, 7] = "";
                Sheet_Stat.Cells[j, 8] = "";
                j++;
            }
            DBBooks.Save();
            Main_Form.Notify("Saved DB with new values");
        }

        public static string Convert_Datetime(string date)
        {
            string retval = date;
            if (date.Contains('/'))
            {
                DateTime time = Convert.ToDateTime(date);
                retval = time.ToString("dd/MM/yyyy");

            }
            return retval;
        }

        public static void Close_DB()
        {
            if (db_open)
            {
                DBBooks.Close(0);
                DBApp.Quit();

                Marshal.ReleaseComObject(Sheet_Stat);
                Marshal.ReleaseComObject(DBBooks);
                Marshal.ReleaseComObject(DBApp);
            }
        }

        public static void Persistence_Hub(Persistence_Request request)
        {
            attending_persistance = true;
            string req = "";
            if (request.persistence_insight != null)
            {
                req = "VyM week: " + request.persistence_insight.Num_of_Week;
                Persistence_Worker(request.persistence_insight);
            }
            Main_Form.Notify("Persistence Request for: " + req);
            attending_persistance = false;
        }

        public static void Persistence_Worker(Insight_Sem sem)
        {
            if (DB_Allowed)
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Name.Equals(sem.Libro_Lector) && Main_Form.Male_List[i].Lector.Contains('/'))
                    {
                        Main_Form.Male_List[i].Lector = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Oracion_End_VyM) && Main_Form.Male_List[i].Oracion.Contains('/'))
                    {
                        Main_Form.Male_List[i].Oracion = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Presidente_RP) && Main_Form.Male_List[i].Pres_RP.Contains('/'))
                    {
                        Main_Form.Male_List[i].Pres_RP = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Conductor_Atalaya) && Main_Form.Male_List[i].Atalaya.Contains('/'))
                    {
                        Main_Form.Male_List[i].Atalaya = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Lector_Atalaya) && Main_Form.Male_List[i].Lector.Contains('/'))
                    {
                        Main_Form.Male_List[i].Lector = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Oracion_End_RP) && Main_Form.Male_List[i].Oracion.Contains('/'))
                    {
                        Main_Form.Male_List[i].Oracion = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Vym_Cap) && Main_Form.Male_List[i].Capitan.Contains('/'))
                    {
                        Main_Form.Male_List[i].Capitan = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Rp_Cap) && Main_Form.Male_List[i].Capitan.Contains('/'))
                    {
                        Main_Form.Male_List[i].Capitan = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Vym_Der) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Vym_Izq) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Rp_Der) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Rp_Izq) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                }
            }
        }

        public static void Persistence_RP(RP_Sem sem)
        {
            if(DB_Allowed)
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Name.Equals(sem.Presidente) && Main_Form.Male_List[i].Pres_RP.Contains('/'))
                    {
                        Main_Form.Male_List[i].Pres_RP = sem.Fecha.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Conductor) && Main_Form.Male_List[i].Atalaya.Contains('/'))
                    {
                        Main_Form.Male_List[i].Atalaya = sem.Fecha.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Lector) && Main_Form.Male_List[i].Lector.Contains('/'))
                    {
                        Main_Form.Male_List[i].Lector = sem.Fecha.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Oracion) && Main_Form.Male_List[i].Oracion.Contains('/'))
                    {
                        Main_Form.Male_List[i].Oracion = sem.Fecha.ToString("dd/MM/yyyy");
                    }
                }
            }
        }

        public static void Persistence_AC(AC_Sem sem)
        {
            if (DB_Allowed)
            {
                for (int i = 0; i < Main_Form.Male_List.Count; i++)
                {
                    if (Main_Form.Male_List[i].Name.Equals(sem.Vym_Cap) && Main_Form.Male_List[i].Capitan.Contains('/'))
                    {
                        Main_Form.Male_List[i].Capitan = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Rp_Cap) && Main_Form.Male_List[i].Capitan.Contains('/'))
                    {
                        Main_Form.Male_List[i].Capitan = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Vym_Der) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Vym_Izq) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_VyM.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Rp_Der) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                    if (Main_Form.Male_List[i].Name.Equals(sem.Rp_Izq) && Main_Form.Male_List[i].Acomodador.Contains('/'))
                    {
                        Main_Form.Male_List[i].Acomodador = sem.Fecha_RP.ToString("dd/MM/yyyy");
                    }
                }
            }
        }
    }
}
