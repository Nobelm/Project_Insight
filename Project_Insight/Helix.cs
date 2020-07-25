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
using System.Windows.Forms;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.Office.Interop.Outlook;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Excel;

namespace Project_Insight
{
    class Helix
    {
        private static Excel.Application objApp;
        private static Excel.Workbook objBooks = null;
        private static Excel.Sheets objSheets;
        private static Excel.Worksheet Sheet_VyM;
        private static Excel.Worksheet Sheet_RP;
        private static Excel.Worksheet Sheet_AC;
        private static Excel.Worksheet Sheet_Week;
        private static Excel.Range range_1;
        private static Excel.Range range_2;
        private static Excel.Range range_3;
        private static Excel.Range range_main;
        private static object[,] cellValue_1 = null;
        private static object[,] cellValue_2 = null;
        private static object[,] cellValue_3 = null;
        private static object[,] cellValue_main = null;
        public static List<Helix_Request> List_Helix_Requests = new List<Helix_Request>();
        private static bool Attending_Helix_Request = false;
        public static bool excel_ready = false;
        public static bool Close_Helix = false;
        private static bool Initial_Check = false;
        public static int loading_delta = 1;
        public static int loading = 0;
        public static int A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7, H = 8, I = 9, J = 10, K = 11, L = 12, M = 13;
        private static int cell = 0;

        public enum Helix_Request
        {
            Save,
            Save_Db,
            Open_Ex,
            Open_Known_Ins
        };

        public static void Start_Helix()
        {
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Helix";
                Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            }
            Helix_Thread_Handler();
        }

        public static void Helix_Thread_Handler()
        {
            while (true)
            {
                if (List_Helix_Requests.Count > 0 && !Attending_Helix_Request)
                {
                    Attending_Helix_Request = true;
                    Helix_Hub(List_Helix_Requests[0]);
                }
                if (!Initial_Check)
                {
                    Initial_Check = true;
                    Initial_Helix_Check();
                }
                Thread.Sleep(1000);
            }
        }

        public static void Initial_Helix_Check()
        {
            if (File.Exists(Main_Form.File_Path))
            {
                Main_Form.Notify("Initial Check: Main File exist");
                Main_Form.Main_Allowed = true;
            }
            else
            {
                Main_Form.Warn("Initial Check: Main file missing");
                Main_Form.Warn("Disabling Main features");
                Main_Form.Main_Allowed = false;
            }
        }

        public static void Helix_Hub(Helix_Request hx)
        {
            Main_Form.Notify("Executing Helix Request: " + hx.ToString());
            switch (hx)
            {
                case Helix_Request.Save:
                    {
                        Gaia_Protocol();
                        break;
                    }
                case Helix_Request.Save_Db:
                    {
                        Persistence.DB_Requests_List.Add(Persistence.DB_Request.write);
                        break;
                    }
                case Helix_Request.Open_Ex:
                    {
                        Opening_Excel(Main_Form.File_Path);
                        break;
                    }
                case Helix_Request.Open_Known_Ins:
                    {
                        Read_Handler();
                        Close_Ex();
                        Main_Form.UI_running = true;
                        Main_Form.Main_Allowed = true;
                        Main_Form.Pending_Week_Handler_Refresh = true;
                        Overwatch.OW_Request = true;
                        Persistence.Persistence_Request request = new Persistence.Persistence_Request();
                        request.persistence_insight = Main_Form.Insight_month.Semana1;
                        Persistence.Persistence_Requests_List.Add(request);
                        request.persistence_insight = Main_Form.Insight_month.Semana2;
                        Persistence.Persistence_Requests_List.Add(request);
                        request.persistence_insight = Main_Form.Insight_month.Semana3;
                        Persistence.Persistence_Requests_List.Add(request);
                        request.persistence_insight = Main_Form.Insight_month.Semana4;
                        Persistence.Persistence_Requests_List.Add(request);
                        request.persistence_insight = Main_Form.Insight_month.Semana5;
                        Persistence.Persistence_Requests_List.Add(request);
                        break;
                    }
            }
            List_Helix_Requests.RemoveAt(0);
            Attending_Helix_Request = false;
        }

        public static bool Opening_Excel(string path)
        {
            bool open_excel = false;
            try
            {
                objApp = new Excel.Application();
                objBooks = objApp.Workbooks.Open(path, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);

                /*Begin Experimental*/
                objSheets = objBooks.Worksheets;
                Sheet_Week = (Excel.Worksheet)objSheets.get_Item(1);
                range_main = Sheet_Week.get_Range("A1", "M250");
                cellValue_main = (object[,])range_main.get_Value();
                /*End Experimental
                objSheets = objBooks.Worksheets;
                Sheet_VyM = (Excel.Worksheet)objSheets.get_Item(1);
                range_1 = Sheet_VyM.get_Range("A1", "H137");
                cellValue_1 = (object[,])range_1.get_Value();
                excel_ready = false;

                if ((cellValue_1[53, 1] != null) && (cellValue_1[53, 1].ToString() == "S-140 AGR-Technologies"))
                {
                    Main_Form.Notify("File decoded correctly");

                    Sheet_RP = (Excel.Worksheet)objSheets.get_Item(2);
                    range_2 = Sheet_RP.get_Range("A1", "H70");
                    cellValue_2 = (object[,])range_2.get_Value();

                    Sheet_AC = (Excel.Worksheet)objSheets.get_Item(3);
                    range_3 = Sheet_AC.get_Range("A1", "H70");
                    cellValue_3 = (object[,])range_3.get_Value();

                    Main_Form.Notify("Opening excel Main file");
                    excel_ready = true;
                }
                else
                {
                    Main_Form.Warn("Invalid file");
                }*/
                open_excel = true;
            }
            catch
            {
                open_excel = false;
                Main_Form.Warn("Something wrong happend opening excel file. Canceling process");
            }
            return open_excel;
        }

        /*-----------------------------------------Gaia Protocol-------------------------------------------*/

        public static void Gaia_Protocol()
        {
            Main_Form.Helix_Saving = true;
            loading = 1;
            Main_Form.Notify("Executing Gaia Protocol. Helix Saving");
            string FileName = Main_Form.meetings_days[0, 0].ToString("MMMM");
            string path = System.Windows.Forms.Application.StartupPath + "\\\\Experimental.xlsx";
            if (Opening_Excel(path))
            {
                if (Main_Form.Week_Format)
                {
                    loading = 10;
                    Persistence.Persistence_Request request = new Persistence.Persistence_Request();
                    Main_Form.Notify("Running Full Week Format");
                    Sheet_Week.PageSetup.LeftHeader = "&16&B" + Main_Form.Cong_Name;
                    Sheet_Week.PageSetup.RightHeader = "&16&B" + "Programa de Reuniones";
                    /*----------Semana 1-----------*/
                    Main_Form.Notify("Saving week 1");
                    cell = 1;
                    VyM_Writer(Main_Form.Insight_month.Semana1);
                    cell++;
                    RP_Writer(Main_Form.Insight_month.Semana1);
                    cell++;
                    AC_Writer(Main_Form.Insight_month.Semana1);
                    request.persistence_insight = Main_Form.Insight_month.Semana1;
                    Persistence.Persistence_Requests_List.Add(request);
                    /*----------Semana 2-----------*/
                    Main_Form.Notify("Saving week 2");
                    cell = 51;
                    VyM_Writer(Main_Form.Insight_month.Semana2); //52
                    cell++;
                    RP_Writer(Main_Form.Insight_month.Semana2);
                    cell++;
                    AC_Writer(Main_Form.Insight_month.Semana2);
                    request.persistence_insight = Main_Form.Insight_month.Semana2;
                    Persistence.Persistence_Requests_List.Add(request);
                    /*----------Semana 3-----------*/
                    Main_Form.Notify("Saving week 3");
                    cell = 101;
                    VyM_Writer(Main_Form.Insight_month.Semana3); //102
                    cell++;
                    RP_Writer(Main_Form.Insight_month.Semana3);
                    cell++;
                    AC_Writer(Main_Form.Insight_month.Semana3);
                    request.persistence_insight = Main_Form.Insight_month.Semana3;
                    Persistence.Persistence_Requests_List.Add(request);
                    /*----------Semana 4-----------*/
                    Main_Form.Notify("Saving week 4");
                    cell = 151;
                    VyM_Writer(Main_Form.Insight_month.Semana4); //152
                    cell++;
                    RP_Writer(Main_Form.Insight_month.Semana4);
                    cell++;
                    AC_Writer(Main_Form.Insight_month.Semana4);
                    request.persistence_insight = Main_Form.Insight_month.Semana4;
                    Persistence.Persistence_Requests_List.Add(request);
                    /*----------Semana 5-----------*/
                    if (Main_Form.week_five_exist)
                    {
                        Main_Form.Notify("Saving week 5");
                        cell = 201;
                        VyM_Writer(Main_Form.Insight_month.Semana5); //202
                        cell++;
                        RP_Writer(Main_Form.Insight_month.Semana5);
                        cell++;
                        AC_Writer(Main_Form.Insight_month.Semana5);
                        request.persistence_insight = Main_Form.Insight_month.Semana5;
                        Persistence.Persistence_Requests_List.Add(request);
                    }
                    loading = 85;
                    cell++;
                    Excel.Range range = Sheet_Week.get_Range("A" + cell.ToString(), "B" + cell.ToString());
                    range.Merge();
                    range.Cells.Font.Color = Color.White;
                    Sheet_Week.Cells[cell, A] = "End of File";
                }
                else
                {
                    Main_Form.Notify("Running Individual Week Format");
                }
                Persistence.DB_Requests_List.Add(Persistence.DB_Request.write);
                Main_Form.Notify("FileName: " + FileName);
                string createfolder = "c:\\Project_Insight";
                System.IO.Directory.CreateDirectory(createfolder);
                objApp.DisplayAlerts = false;
                objBooks.SaveAs(createfolder + "\\" + FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                loading = 90;
                //objBooks.SaveAs(createfolder + "\\" + FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                Main_Form.File_Path = createfolder + "\\" + FileName;
                if (Main_Form.Save_as_pdf)
                {
                    objBooks.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Main_Form.File_Path);
                    Main_Form.Notify("Saving as Pdf");
                    Main_Form.Save_as_pdf = false;
                }
                loading = 100;
                Main_Form.Notify("Saved path: " + Main_Form.File_Path);
                objBooks.Close(0);
                objApp.Quit();
                Main_Form.Notify("Save complete!");
                Main_Form.Helix_Saving = false;
            }
        }

        public static void VyM_Writer(Insight_Sem sem)
        {
            CultureInfo spanish = new CultureInfo("es-MX");
            Excel.Range range;
            //Code
            range = Sheet_Week.get_Range("A" + cell.ToString(), "B" + cell.ToString());
            range.Merge();
            range.Cells.Font.Color = Color.White;
            Sheet_Week.Cells[cell, A] = "vym_" + sem.Num_of_Week.ToString();
            cell++;
            //Informacion de semana
            range = Sheet_Week.get_Range("A" + cell.ToString(), "M" + cell.ToString());
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;
            range = Sheet_Week.get_Range("A" + cell.ToString(), "I" + cell.ToString());
            range.Cells.Merge();
            if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
            {
                range.Cells.Font.Bold = true;
                range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(255, 204, 153);
                Sheet_Week.Cells[cell, A] = sem.Special_VyM_Meeting_Info;
            }
            //Horario
            range = Sheet_Week.get_Range("J" + cell.ToString(), "M" + cell.ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            Sheet_Week.Cells[cell, J] = "Horario: " + sem.Fecha_VyM.ToString("dddd", spanish) + " " + Main_Form.VyM_horary.ToString("hh:mm tt");
            cell += 1;
            //Fecha
            range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + (cell + 1).ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous; 
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
            Sheet_Week.Cells[cell, A] = sem.Fecha_VyM.ToString("dddd, dd MMMM");

            if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Conv_type)
            {
                //Handler for Convention type
                cell += 4;
                range = Sheet_Week.get_Range("A" + cell.ToString(), "M" + (cell + 4).ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 14;
                range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(255, 204, 153);
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Sheet_Week.Cells[cell, A] = sem.Special_VyM_Meeting_Info;
                cell += 12;
            }
            else
            {
                //Lectura de la biblia semanal
                range = Sheet_Week.get_Range("D" + cell.ToString(), "G" + (cell + 1).ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 11;
                Sheet_Week.Cells[cell, D] = sem.Sem_Biblia;
                //Presidente y Consejero auxiliar
                range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Size = 8;
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                Sheet_Week.Cells[cell, H] = "Presidente";
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 9;
                Sheet_Week.Cells[cell, K] = sem.Presidente_VyM;
                if (Main_Form.Room_B_enabled)
                {
                    range = Sheet_Week.get_Range("H" + (cell + 1).ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell + 1, H] = "Consejero Auxiliar";
                    range = Sheet_Week.get_Range("K" + (cell + 1).ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 9;
                    Sheet_Week.Cells[cell + 1, K] = sem.Consejero_Aux;
                }
                cell += 2;
                //Cancion y Palabras de Introduccion
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_VyM_1;
                range = Sheet_Week.get_Range("A" + (cell + 1).ToString(), "E" + (cell + 1).ToString());
                range.Cells.Merge();
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                Sheet_Week.Cells[cell + 1, A] = "• Palabras de introduccion (1 min.)";
                Set_Font(Sheet_Week.get_Range("A" + (cell + 1)));
                cell += 2;
                //Seccion "Tesoros de la biblia"
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 11;
                range.Cells.Font.Color = Color.White;
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(87, 90, 93);
                Sheet_Week.Cells[cell, A] = "TESOROS DE LA BIBLIA";
                if (Main_Form.Room_B_enabled)
                {
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Sheet_Week.Cells[cell, H] = "Sala Auxiliar";
                }
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Size = 8;
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Sheet_Week.Cells[cell, K] = "Sala Principal";
                cell++;
                //Discurso
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Discurso_VyM;
                Set_Font(Sheet_Week.get_Range("A" + cell));
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, K] = sem.Discurso_VyM_A;
                cell++;
                //Perlas                    
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• Busquemos Perlas Escondidas (10 mins.)";
                Set_Font(Sheet_Week.get_Range("A" + cell));
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, K] = sem.Perlas;
                cell++;
                //Lectura
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Lectura_Biblia;
                Set_Font(Sheet_Week.get_Range("A" + cell));
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, K] = sem.Lectura_Biblia_A;
                range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, H] = sem.Lectura_Biblia_B;
                cell++;
                //Seccion "Seamos mejores maestros"
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 11;
                range.Cells.Font.Color = Color.White;
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(190, 137, 0);
                Sheet_Week.Cells[cell, A] = "SEAMOS MEJORES MAESTROS";
                cell++;
                //Asignacion 1
                string aux_column = "I";
                if (Main_Form.Room_B_enabled)
                {
                    aux_column = "G";
                }
                range = Sheet_Week.get_Range("A" + cell.ToString(), aux_column + (cell + 1).ToString());
                range.Cells.Merge();
                range.Cells.WrapText = true;
                Sheet_Week.Cells[cell, A] = "• " + sem.SMM1;
                Set_Font(Sheet_Week.get_Range("A" + cell));
                if (Main_Form.Room_B_enabled)
                {
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.WrapText = true;
                    Sheet_Week.Cells[cell, H] = sem.SMM1_B;
                }
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                range.Cells.Merge();
                range.Cells.WrapText = true;
                Sheet_Week.Cells[cell, K] = sem.SMM1_A;
                cell += 2;
                //Asignacion 2
                range = Sheet_Week.get_Range("A" + cell.ToString(), aux_column + (cell + 1).ToString());
                range.Cells.Merge();
                range.Cells.WrapText = true;
                Sheet_Week.Cells[cell, A] = "• " + sem.SMM2;
                Set_Font(Sheet_Week.get_Range("A" + cell));
                if (Main_Form.Room_B_enabled)
                {
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.WrapText = true;
                    Sheet_Week.Cells[cell, H] = sem.SMM2_B;
                }
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                range.Cells.Merge();
                range.Cells.WrapText = true;
                Sheet_Week.Cells[cell, K] = sem.SMM2_A;
                cell += 2;
                //Asignacion 3
                if (sem.SMM3 != null && sem.SMM3.Length > 5)
                {
                    range = Sheet_Week.get_Range("A" + cell.ToString(), aux_column + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.WrapText = true;
                    Sheet_Week.Cells[cell, A] = "• " + sem.SMM3;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
                    if (Main_Form.Room_B_enabled)
                    {
                        range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                        range.Cells.Merge();
                        range.Cells.WrapText = true;
                        Sheet_Week.Cells[cell, H] = sem.SMM3_B;
                    }
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.WrapText = true;
                    Sheet_Week.Cells[cell, K] = sem.SMM3_A;
                    cell += 2;
                }
                //Asignacion 4
                if (sem.SMM4 != null && sem.SMM4.Length > 5)
                {
                    range = Sheet_Week.get_Range("A" + cell.ToString(), aux_column + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.WrapText = true;
                    Sheet_Week.Cells[cell, A] = "• " + sem.SMM4;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
                    if (Main_Form.Room_B_enabled)
                    {
                        range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                        range.Cells.Merge();
                        range.Cells.WrapText = true;
                        Sheet_Week.Cells[cell, H] = sem.SMM4_B;
                    }
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.WrapText = true;
                    Sheet_Week.Cells[cell, K] = sem.SMM4_A;
                    cell += 2;
                }
                //Seccion "Nuestra vida cristiana"
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 11;
                range.Cells.Font.Color = Color.White;
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(126, 0, 36);
                Sheet_Week.Cells[cell, A] = "NUESTRA VIDA CRISTIANA";
                cell++;
                //Cancion
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_VyM_2;
                cell++;
                //Parte 1
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.NVC1;
                Set_Font(Sheet_Week.get_Range("A" + cell)); 
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, K] = sem.NVC1_A;
                cell++;
                //Parte 2
                if (sem.NVC2 != null && sem.NVC2.Length > 5)
                {
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.NVC2;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.NVC2_A;
                    cell++;
                }
                //Estudio Biblico de Congregacion
                if (sem.Libro_Titulo == null)
                {
                    sem.Libro_Titulo = "Estudio biblico de congregacion (30 min.)";
                }
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + (cell + 1).ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Libro_Titulo;
                Set_Font(Sheet_Week.get_Range("A" + cell));
                if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
                {
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    Sheet_Week.Cells[cell, K] = sem.Libro_Conductor;
                }
                else
                {
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, H] = "Conductor";
                    range = Sheet_Week.get_Range("H" + (cell + 1).ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell + 1, H] = "Lector";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Libro_Conductor;
                    range = Sheet_Week.get_Range("K" + (cell + 1).ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, K] = sem.Libro_Lector;
                }
                cell += 2;
                //Repaso
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• Palabras de conclusión (3 mins.o menos)";
                Set_Font(Sheet_Week.get_Range("A" + cell));
                cell++;
                //Cancion y Oracion Final
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_VyM_3; 
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, K] = sem.Oracion_End_VyM;
            }
            loading += 5;
        }
        public static void RP_Writer(Insight_Sem sem)
        {
            CultureInfo spanish = new CultureInfo("es-MX");
            Excel.Range range;
            //Code
            range = Sheet_Week.get_Range("A" + cell.ToString(), "B" + cell.ToString());
            range.Merge();
            range.Cells.Font.Color = Color.White;
            Sheet_Week.Cells[cell, A] = "rp_" + sem.Num_of_Week.ToString();
            cell++;
            //Informacion de semana
            range = Sheet_Week.get_Range("A" + cell.ToString(), "M" + cell.ToString());
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;
            range = Sheet_Week.get_Range("A" + cell.ToString(), "I" + cell.ToString());
            range.Cells.Merge();
            range = Sheet_Week.get_Range("J" + cell.ToString(), "M" + cell.ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            range = Sheet_Week.get_Range("A" + cell.ToString(), "I" + cell.ToString());
            range.Cells.Merge();
            if (sem.Special_RP_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
            {
                range.Cells.Font.Bold = true;
                range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(255, 204, 153);
                Sheet_Week.Cells[cell, A] = sem.Special_RP_Meeting_Info;
            }
            //Horario
            Sheet_Week.Cells[cell, J] = "Horario: " + sem.Fecha_RP.ToString("dddd", spanish) + " " + Main_Form.RP_horary.ToString("hh:mm tt");
            cell += 1;
            //Fecha
            range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + (cell + 1).ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
            Sheet_Week.Cells[cell, A] = sem.Fecha_RP.ToString("dddd, dd MMMM");
            cell++;
            if (sem.Special_RP_Meeting == Main_Form.Special_Meeting_Type.Conv_type)
            {                
                //Handler for Convention type
                cell += 1;
                range = Sheet_Week.get_Range("A" + cell.ToString(), "M" + (cell + 4).ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 14;
                range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(255, 204, 153);
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Sheet_Week.Cells[cell, A] = sem.Special_RP_Meeting_Info;
                cell += 7;
            }
            else
            {
                //Presidente
                range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Size = 8;
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                Sheet_Week.Cells[cell, H] = "Presidente";
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 9;
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                Sheet_Week.Cells[cell, K] = sem.Presidente_RP; 
                cell++;
                //Seccion "Reunion Publica"
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 11;
                range.Cells.Font.Color = Color.White;
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(68, 84, 106);
                Sheet_Week.Cells[cell, A] = "REUNION PUBLICA";
                cell++;
                //Cancion de inicio
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_RP_1;
                cell++;
                //Informacion del discurso
                Sheet_Week.Cells[cell, A] = "• Tema";
                range = Sheet_Week.get_Range("B" + cell.ToString(), "H" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, B] = sem.Titulo_Discurso_RP;
                range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Size = 8;
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                Sheet_Week.Cells[cell, I] = "Discursante";
                range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                Sheet_Week.Cells[cell, K] = sem.Discursante_RP;
                cell++;
                range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "Congregacion";
                range = Sheet_Week.get_Range("D" + cell.ToString(), "I" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, D] = sem.Congregacion_RP;
                cell++;
                //Seccion "Analisis de La Atalaya"
                range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Size = 11;
                range.Cells.Font.Color = Color.White;
                Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(84, 130, 53);
                Sheet_Week.Cells[cell, A] = "ANALISIS DE LA ATALAYA";
                cell++;
                //Cancion intermedia
                range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                range.Cells.Merge();
                Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_RP_2;
                cell++;
                //Informacion de La Atalaya
                if (sem.Special_RP_Meeting_Info != null && sem.Special_RP_Meeting_Info.Contains("Visita"))
                {
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.Titulo_Atalaya;
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Conductor";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Conductor_Atalaya;
                    cell++;
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• Discurso de Servicio";
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Superintendente";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Discursante_RP; 
                    cell++;
                    //Cancion final y oracion
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_RP_3;
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Oracion Final";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Oracion_End_RP;
                }
                else //Normal Meeting
                {
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "H" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.Titulo_Atalaya;
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Conductor";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Conductor_Atalaya;
                    cell++;
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Lector";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Lector_Atalaya;
                    cell++;
                    //Cancion final y oracion
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.Cancion_RP_3;
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Oracion Final";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Oracion_End_RP;
                }
                cell++;
                //Salidas a Discursar
                if (sem.Discu_Sal != null && sem.Discu_Sal.Length > 2)
                {
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "D" + cell.ToString());
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 2d;
                    range.Cells.Font.Bold = true;
                    Sheet_Week.Cells[cell, A] = "Salidas a Discursar:";
                    cell++;
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = sem.Discu_Sal;
                    range = Sheet_Week.get_Range("D" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, D] = sem.Ttl_Sal;
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Cong_Sal;
                }
            }
            loading += 5;
        }
        public static void AC_Writer(Insight_Sem sem)
        {
            Excel.Range range;
            //Code
            range = Sheet_Week.get_Range("A" + cell.ToString(), "B" + cell.ToString());
            range.Merge();
            range.Cells.Font.Color = Color.White;
            Sheet_Week.Cells[cell, A] = "ac_" + sem.Num_of_Week.ToString();
            cell++;
            //Informacion de semana
            range = Sheet_Week.get_Range("A" + cell.ToString(), "M" + cell.ToString());
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;
            if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Visit_type && sem.Special_RP_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
            {
                range = Sheet_Week.get_Range("D" + cell.ToString(), "M" + cell.ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                Sheet_Week.Cells[cell, D].Interior.Color = Color.FromArgb(255, 204, 153);
                Sheet_Week.Cells[cell, D] = sem.Special_VyM_Meeting_Info;
            }
            else
            {
                if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
                {
                    range = Sheet_Week.get_Range("D" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                    Sheet_Week.Cells[cell, D].Interior.Color = Color.FromArgb(255, 204, 153);
                    Sheet_Week.Cells[cell, D] = Get_First_Word(sem.Special_VyM_Meeting_Info);
                }
                if (sem.Special_RP_Meeting == Main_Form.Special_Meeting_Type.Visit_type)
                {
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                    Sheet_Week.Cells[cell, I].Interior.Color = Color.FromArgb(255, 204, 153);
                    Sheet_Week.Cells[cell, I] = Get_First_Word(sem.Special_RP_Meeting_Info);
                }
            }
            cell += 1;
            //Fecha
            range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + (cell + 1).ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
            Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(180, 198, 231);
            Sheet_Week.Cells[cell, A] = "FECHA";
            range = Sheet_Week.get_Range("D" + cell.ToString(), "H" + (cell + 1).ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
            Sheet_Week.Cells[cell, D].Interior.Color = Color.FromArgb(180, 198, 231);
            Sheet_Week.Cells[cell, D] = sem.Fecha_VyM.ToString("dddd, dd MMMM"); 
            range = Sheet_Week.get_Range("I" + cell.ToString(), "M" + (cell + 1).ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            range.Cells.Font.Size = 11;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
            Sheet_Week.Cells[cell, I].Interior.Color = Color.FromArgb(180, 198, 231);
            Sheet_Week.Cells[cell, I] = sem.Fecha_RP.ToString("dddd, dd MMMM");
            cell += 3;
            //Aseo
            range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + cell.ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            Sheet_Week.Cells[cell, A] = "Aseo";
            range = Sheet_Week.get_Range("A" + (cell + 1).ToString(), "C" + (cell + 2).ToString());
            range.Cells.Merge();
            range.Cells.Font.Bold = true;
            Sheet_Week.Cells[cell + 1, A] = sem.Aseo;
            //Asamblea Natural
            if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Conv_type && sem.Special_RP_Meeting == Main_Form.Special_Meeting_Type.Conv_type)
            {
                range = Sheet_Week.get_Range("D" + cell.ToString(), "M" + (cell + 2).ToString());
                range.Cells.Merge();
                range.Cells.Font.Bold = true;
                range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                Sheet_Week.Cells[cell, D].Interior.Color = Color.FromArgb(255, 204, 153);
                Sheet_Week.Cells[cell, D] = sem.Special_VyM_Meeting_Info;
            }
            else
            {
                //Acomodadores VyM
                if (sem.Special_VyM_Meeting == Main_Form.Special_Meeting_Type.Conv_type)
                {
                    range = Sheet_Week.get_Range("D" + cell.ToString(), "H" + (cell + 2).ToString());
                    range.Cells.Merge();
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                    Sheet_Week.Cells[cell, D].Interior.Color = Color.FromArgb(255, 204, 153);
                    Sheet_Week.Cells[cell, D] = Get_First_Word(sem.Special_VyM_Meeting_Info);
                }
                else
                {
                    range = Sheet_Week.get_Range("D" + cell.ToString(), "E" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, D] = "Capitan";
                    range = Sheet_Week.get_Range("F" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, F] = sem.Vym_Cap;
                    range = Sheet_Week.get_Range("D" + (cell + 1).ToString(), "E" + (cell + 2).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, D] = "Acomodadores";
                    range = Sheet_Week.get_Range("F" + (cell + 1).ToString(), "H" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, F] = sem.Vym_Der;
                    range = Sheet_Week.get_Range("F" + (cell + 2).ToString(), "H" + (cell + 2).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 2, F] = sem.Vym_Izq;
                }
                //Acomodadores RP
                if (sem.Special_RP_Meeting == Main_Form.Special_Meeting_Type.Conv_type)
                {
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "M" + (cell + 2).ToString());
                    range.Cells.Merge();
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.Cells.Font.Color = Color.FromArgb(63, 63, 118);
                    Sheet_Week.Cells[cell, I].Interior.Color = Color.FromArgb(255, 204, 153);
                    Sheet_Week.Cells[cell, I] = Get_First_Word(sem.Special_RP_Meeting_Info);
                }
                else
                {
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, I] = "Capitan";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Rp_Cap;
                    range = Sheet_Week.get_Range("I" + (cell + 1).ToString(), "J" + (cell + 2).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, I] = "Acomodadores";
                    range = Sheet_Week.get_Range("K" + (cell + 1).ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, K] = sem.Rp_Der;
                    range = Sheet_Week.get_Range("K" + (cell + 2).ToString(), "M" + (cell + 2).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 2, K] = sem.Rp_Izq;
                }
            }
            loading += 5;
        }

        /*----------------------------------------- Open Handlers -------------------------------------------*/

        public static void Read_Handler()
        {
            if (Opening_Excel(Main_Form.File_Path))
            {
                Get_month_from_Excel(cellValue_main[3, 1]);
                Main_Form.Notify("Reading Week 1");
                cell = 1;
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana1 = VyM_Reader(Main_Form.Insight_month.Semana1);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana1 = RP_Reader(Main_Form.Insight_month.Semana1);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana1 = AC_Reader(Main_Form.Insight_month.Semana1);

                Main_Form.Notify("Reading Week 2");
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana2 = VyM_Reader(Main_Form.Insight_month.Semana2);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana2 = RP_Reader(Main_Form.Insight_month.Semana2);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana2 = AC_Reader(Main_Form.Insight_month.Semana2);

                Main_Form.Notify("Reading Week 3");
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana3 = VyM_Reader(Main_Form.Insight_month.Semana3);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana3 = RP_Reader(Main_Form.Insight_month.Semana3);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana3 = AC_Reader(Main_Form.Insight_month.Semana3);

                Main_Form.Notify("Reading Week 4");
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana4 = VyM_Reader(Main_Form.Insight_month.Semana4);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana4 = RP_Reader(Main_Form.Insight_month.Semana4);
                cell = Find_Meeting(cell);
                Main_Form.Insight_month.Semana4 = AC_Reader(Main_Form.Insight_month.Semana4);

                if (Main_Form.week_five_exist)
                {
                    Main_Form.Notify("Reading Week 5");
                    cell = Find_Meeting(cell);
                    Main_Form.Insight_month.Semana5 = VyM_Reader(Main_Form.Insight_month.Semana5);
                    cell = Find_Meeting(cell);
                    Main_Form.Insight_month.Semana5 = RP_Reader(Main_Form.Insight_month.Semana5);
                    cell = Find_Meeting(cell);
                    Main_Form.Insight_month.Semana5 = AC_Reader(Main_Form.Insight_month.Semana5);
                }
                objBooks.Close(0);
                objApp.Quit();
            }
        }

        public static int Find_Meeting(int start_value)
        {
            int i = start_value;
            while (!Check_null_string(cellValue_main[i, A]).Equals("End of File"))
            {
                if (Check_null_string(cellValue_main[i, A]).Contains("vym_"))
                {
                    i++;
                    break;
                }
                else if (Check_null_string(cellValue_main[i, A]).Contains("rp_"))
                {
                    i++;
                    break;
                }
                else if (Check_null_string(cellValue_main[i, A]).Contains("ac_"))
                {
                    i++;
                    break;
                }
                else
                {
                    i++;
                }
            }
            return i;
        }

        public static Insight_Sem VyM_Reader(Insight_Sem sem)
        {
            /*Check for Conv Type*/
            if (cellValue_main[cell + 5, A] != null && (cellValue_main[cell + 5, A].ToString().Contains("Conmemoracion") || (cellValue_main[cell + 5, A].ToString().Contains("Asamblea"))))
            {
                sem.Special_VyM_Meeting = Main_Form.Special_Meeting_Type.Conv_type;
                sem.Special_VyM_Meeting_Info = cellValue_main[cell + 5, 1].ToString();
                Main_Form.Notify("Week [" + sem.Num_of_Week.ToString() + "] found as Convention Type Week");
                cell += 17;
            }
            else
            {
                if (cellValue_main[cell, A] != null)
                {
                    sem.Special_VyM_Meeting = Main_Form.Special_Meeting_Type.Visit_type;
                    sem.Special_VyM_Meeting_Info = cellValue_main[cell, A].ToString();
                    Main_Form.Notify("Week [" + sem.Num_of_Week.ToString() + "] found as Visit Type Week");
                }
                cell++;
                sem.Sem_Biblia = Check_null_string(cellValue_main[cell, D]);
                sem.Presidente_VyM = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Consejero_Aux = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Cancion_VyM_1 = Check_null_string(cellValue_main[cell, A]);
                cell += 3;
                sem.Discurso_VyM = Check_null_string(cellValue_main[cell, A]);
                sem.Discurso_VyM_A = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Perlas = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Lectura_Biblia = Check_null_string(cellValue_main[cell, A]);
                sem.Lectura_Biblia_A = Check_null_string(cellValue_main[cell, K]);
                sem.Lectura_Biblia_B = Check_null_string(cellValue_main[cell, H]);
                cell += 2;
                sem.SMM1 = Check_null_string(cellValue_main[cell, A]);
                sem.SMM1_A = Check_null_string(cellValue_main[cell, K]);
                sem.SMM1_B = Check_null_string(cellValue_main[cell, H]);
                cell += 2;
                sem.SMM2 = Check_null_string(cellValue_main[cell, A]);
                sem.SMM2_A = Check_null_string(cellValue_main[cell, K]);
                sem.SMM2_B = Check_null_string(cellValue_main[cell, H]);
                cell += 2;
                if (cellValue_main[cell, A] != null && !cellValue_main[cell, A].ToString().Equals("NUESTRA VIDA CRISTIANA"))
                {
                    sem.SMM3 = Check_null_string(cellValue_main[cell, A]);
                    sem.SMM3_A = Check_null_string(cellValue_main[cell, K]);
                    sem.SMM3_B = Check_null_string(cellValue_main[cell, H]);
                    cell += 2;
                    if (cellValue_main[cell, A] != null && !cellValue_main[cell, A].ToString().Equals("NUESTRA VIDA CRISTIANA"))
                    {
                        sem.SMM4 = Check_null_string(cellValue_main[cell, A]);
                        sem.SMM4_A = Check_null_string(cellValue_main[cell, K]);
                        sem.SMM4_B = Check_null_string(cellValue_main[cell, H]);
                        cell += 3;
                    }
                    else
                    {
                        cell++;
                    }
                }
                else
                {
                    cell++;
                }
                sem.Cancion_VyM_2 = Check_null_string(cellValue_main[cell, A]);
                cell++;
                sem.NVC1 = Check_null_string(cellValue_main[cell, A]);
                sem.NVC1_A = Check_null_string(cellValue_main[cell, K]);
                cell++;
                string aux = Check_null_string(cellValue_main[cell, A]);
                if (aux != "" && !aux.Contains("Estudio bíblico") && !aux.Contains("Discurso de servicio"))
                {
                    sem.NVC2 = Check_null_string(cellValue_main[cell, A]);
                    sem.NVC2_A = Check_null_string(cellValue_main[cell, K]);
                    cell++;
                }
                sem.Libro_Titulo = Check_null_string(cellValue_main[cell, A]);
                sem.Libro_Conductor = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Libro_Lector = Check_null_string(cellValue_main[cell, K]);
                cell += 2;
                sem.Cancion_VyM_3 = Check_null_string(cellValue_main[cell, A]);
                sem.Oracion_End_VyM = Check_null_string(cellValue_main[cell, K]);
            }
                return sem;
        }

        public static Insight_Sem RP_Reader(Insight_Sem sem)
        {
            /*Check for Conv Type*/
            if (cellValue_main[cell + 5, A] != null && (cellValue_main[cell + 5, A].ToString().Contains("Conmemoracion") || (cellValue_main[cell + 5, A].ToString().Contains("Asamblea"))))
            {
                sem.Special_RP_Meeting = Main_Form.Special_Meeting_Type.Conv_type;
                sem.Special_RP_Meeting_Info = cellValue_main[cell + 5, 1].ToString();
                Main_Form.Notify("Week [" + sem.Num_of_Week.ToString() + "] found as Convention Type Week");
            }
            else
            {
                if (cellValue_main[cell, A] != null)
                {
                    sem.Special_RP_Meeting = Main_Form.Special_Meeting_Type.Visit_type;
                    sem.Special_RP_Meeting_Info = cellValue_main[cell, A].ToString();
                    Main_Form.Notify("Week [" + sem.Num_of_Week.ToString() + "] found as Visit Type Week");
                }
                cell += 2;
                sem.Presidente_RP = Check_null_string(cellValue_main[cell, K]);
                cell += 2;
                sem.Cancion_RP_1 = Check_null_string(cellValue_main[cell, A]);
                cell++;
                sem.Titulo_Discurso_RP = Check_null_string(cellValue_main[cell, B]);
                sem.Discursante_RP = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Congregacion_RP = Check_null_string(cellValue_main[cell, D]);
                cell += 2;
                sem.Cancion_RP_2 = Check_null_string(cellValue_main[cell, A]);
                cell++;
                sem.Titulo_Atalaya = Check_null_string(cellValue_main[cell, A]);
                sem.Conductor_Atalaya = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Lector_Atalaya = Check_null_string(cellValue_main[cell, K]);
                cell++;
                sem.Cancion_RP_3 = Check_null_string(cellValue_main[cell, A]);
                sem.Oracion_End_RP = Check_null_string(cellValue_main[cell, K]);
                cell += 2;
                if(!Check_null_string(cellValue_main[cell, A]).Contains("ac_"))
                {
                    sem.Discu_Sal = Check_null_string(cellValue_main[cell, A]);
                    sem.Ttl_Sal = Check_null_string(cellValue_main[cell, D]);
                    sem.Cong_Sal = Check_null_string(cellValue_main[cell, K]);
                }
            }
            return sem;
        }

        public static Insight_Sem AC_Reader(Insight_Sem sem)
        {
            cell += 4;
            sem.Vym_Cap = Check_null_string(cellValue_main[cell, F]);
            sem.Rp_Cap = Check_null_string(cellValue_main[cell, K]);
            cell++;
            sem.Aseo = Check_null_string(cellValue_main[cell, A]);
            sem.Vym_Der = Check_null_string(cellValue_main[cell, F]);
            sem.Rp_Der = Check_null_string(cellValue_main[cell, K]);
            cell++;
            sem.Vym_Izq = Check_null_string(cellValue_main[cell, F]);
            sem.Rp_Izq = Check_null_string(cellValue_main[cell, K]);


            return sem;
        }

        /*Set font size to (x min.)*/
        public static void Set_Font(Excel.Range cell)
        {
            string Str = cell.Text;
            int index = Str.IndexOf("min");
            if (index > 4)
            {
                cell.Characters[index - 3, Str.Length].Font.Size = 8;
            }
        }

        public static string Get_First_Word(string str)
        {
            string retval = "";
            if (str.Contains(" "))
            {
                int index = str.IndexOf(" ");
                retval = str.Substring(0, index + 1);
            }
            else
            {
                retval = str;
            }
            return retval;
        }


        private static void Get_month_from_Excel(object cellvalue)
        {
            if (cellvalue != null && !Main_Form.month_found)
            {
                //bool month_found = false;
                for (int i = 0; i <= Main_Form.Months.Length - 1; i++)
                {
                    if (cellvalue.ToString().ToLower().Contains(Main_Form.Months[i]))
                    {
                        Main_Form.m_mes = i + 1;
                        Main_Form.month_found = true;
                        break;
                    }
                }
                if (Main_Form.month_found)
                {
                    Main_Form.Notify("Month set in [" + Main_Form.m_mes.ToString() + "]");
                    Main_Form.Get_Meetings();
                }
                else
                {
                    Main_Form.m_mes = DateTime.Today.Month;
                    Main_Form.Warn("Month not found in first week, seeting today's month [" + Main_Form.m_mes.ToString() + "]");
                }
            }
        }

        public static string Check_null_string(object cellvalue)
        {
            string retval;
            if (cellvalue == null)
            {
                retval = "";
            }
            else
            {
                retval = cellvalue.ToString();
                if (retval.Contains("• "))
                {
                    retval = retval.Substring(2);
                }
            }
            return retval;
        }
       
        public static void Close_Ex()
        {
            if (excel_ready)
            {
                excel_ready = false;
                objBooks.Close(0);
                objApp.Quit();
                Marshal.ReleaseComObject(Sheet_VyM);
                Marshal.ReleaseComObject(Sheet_RP);
                Marshal.ReleaseComObject(Sheet_AC);
                Marshal.ReleaseComObject(Sheet_Week);
                Marshal.ReleaseComObject(objBooks);
                Marshal.ReleaseComObject(objApp);
            }
        }
    }
}
