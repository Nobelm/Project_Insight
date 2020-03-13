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
                        Opening_Excel(Main_Form.File_Path);
                        VyM_Handler(false);
                        RP_Handler(false);
                        AC_Handler(false);
                        Close_Ex();
                        Main_Form.Get_Meetings();
                        Main_Form.UI_running = true;
                        Main_Form.Main_Allowed = true;
                        Main_Form.Pending_Week_Handler_Refresh = true;
                        break;
                    }
            }
            List_Helix_Requests.RemoveAt(0);
            Attending_Helix_Request = false;
        }

        public static void Process_save(int save)
        {
            Main_Form.Helix_Saving = true;
            string FileName = Main_Form.meetings_days[0, 0].ToString("MMMM");
            loading = 1;
            Opening_Excel(Main_Form.File_Path);
            loading += 4;
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
            Main_Form.Notify("FileName: " + FileName);
            loading = 80;
            if (Main_Form.is_new_instance)
            {
                string createfolder = "c:\\Project_Insight";
                System.IO.Directory.CreateDirectory(createfolder);
                objApp.DisplayAlerts = false;
                objBooks.SaveAs(createfolder + "\\" + FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

                //objBooks.SaveAs(createfolder + "\\" + FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                Main_Form.File_Path = createfolder + "\\" + FileName;
                if (Main_Form.Save_as_pdf)
                {
                    objBooks.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Main_Form.File_Path);
                    Main_Form.Notify("Saving as Pdf");
                    Main_Form.Save_as_pdf = false;
                }
                Main_Form.Notify("Saved path: " + Main_Form.File_Path);
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
            Persistence.DB_Requests_List.Add(Persistence.DB_Request.write);
            Main_Form.Notify("Save successful!");
            //Heavensward request
            loading = 100;
        }

        public static void Opening_Excel(string path)
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
        }

        public static void VyM_Handler(bool save)
        {
            if (save)
            {
                VyM_Save_Week(Main_Form.VyM_mes.Semana1);
                loading += (15 / loading_delta);
                VyM_Save_Week(Main_Form.VyM_mes.Semana2);
                loading += (15 / loading_delta);
                VyM_Save_Week(Main_Form.VyM_mes.Semana3);
                loading += (15 / loading_delta);
                VyM_Save_Week(Main_Form.VyM_mes.Semana4);
                loading += (15 / loading_delta);
                if (Main_Form.week_five_exist)
                {
                    VyM_Save_Week(Main_Form.VyM_mes.Semana5);
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


        /*-----------------------------------------BEGIN EXPERIMENTAL-------------------------------------------*/
        /*public class Coordinates
        {
            public string Column;
            public int Row;
        }*/

        public static void Gaia_Protocol()
        {
            Main_Form.Notify("Executing Gaia Protocol");
            string FileName = "Experimental Gaia Protocol";//Main_Form.meetings_days[0, 0].ToString("MMMM");
            string path = Application.StartupPath + "\\\\Experimental.xlsx";
            Opening_Excel(path);  /**/
            if (Main_Form.Week_Format)
            {
                Main_Form.Notify("Running Full Week Format");
                Sheet_Week.PageSetup.LeftHeader = "&16&B" + Main_Form.Cong_Name;
                Sheet_Week.PageSetup.RightHeader = "&16&B" + "Programa de Reuniones";
                Main_Form.Notify("Saving week 1");
                cell = 2;
                VyM_Builder(Main_Form.Insight_month.Semana1);
                cell += 2;
                RP_Builder(Main_Form.Insight_month.Semana1);
                cell += 2;
                AC_Builder(Main_Form.Insight_month.Semana1);
                Main_Form.Notify("Saving week 2");
                cell = 52;
                VyM_Builder(Main_Form.Insight_month.Semana2); //52
                cell += 2;
                RP_Builder(Main_Form.Insight_month.Semana2);
                cell += 2;
                AC_Builder(Main_Form.Insight_month.Semana2);
                Main_Form.Notify("Saving week 3");
                cell = 102;
                VyM_Builder(Main_Form.Insight_month.Semana3); //102
                cell += 2;
                RP_Builder(Main_Form.Insight_month.Semana3);
                cell += 2;
                AC_Builder(Main_Form.Insight_month.Semana3);
                Main_Form.Notify("Saving week 4");
                cell = 152;
                VyM_Builder(Main_Form.Insight_month.Semana4); //152
                cell += 2;
                RP_Builder(Main_Form.Insight_month.Semana4);
                cell += 2;
                AC_Builder(Main_Form.Insight_month.Semana4);
                if (Main_Form.week_five_exist)
                {
                    Main_Form.Notify("Saving week 5");
                    cell = 202;
                    VyM_Builder(Main_Form.Insight_month.Semana5); //202
                    cell += 2;
                    RP_Builder(Main_Form.Insight_month.Semana5);
                    cell += 2;
                    AC_Builder(Main_Form.Insight_month.Semana5);
                }

            }
            else
            {
                Main_Form.Notify("Running Individual Week Format");




            }

            Main_Form.Notify("FileName: " + FileName);
            if (Main_Form.is_new_instance)
            {
                string createfolder = "c:\\Project_Insight";
                System.IO.Directory.CreateDirectory(createfolder);
                objApp.DisplayAlerts = false;
                objBooks.SaveAs(createfolder + "\\" + FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

                //objBooks.SaveAs(createfolder + "\\" + FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                Main_Form.File_Path = createfolder + "\\" + FileName;
                if (Main_Form.Save_as_pdf)
                {
                    objBooks.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Main_Form.File_Path);
                    Main_Form.Notify("Saving as Pdf");
                    Main_Form.Save_as_pdf = false;
                }
                Main_Form.Notify("Saved path: " + Main_Form.File_Path);
                objBooks.Close(0);
                objApp.Quit();
            }

            Main_Form.Notify("Adding Persistence Request");
            Persistence.Persistence_Request request = new Persistence.Persistence_Request();
            request.persistence_insight = Main_Form.Insight_month.Semana1;
            Persistence.Persistence_Requests_List.Add(request);
            request.persistence_insight = Main_Form.Insight_month.Semana2;
            Persistence.Persistence_Requests_List.Add(request);
            request.persistence_insight = Main_Form.Insight_month.Semana3;
            Persistence.Persistence_Requests_List.Add(request);
            request.persistence_insight = Main_Form.Insight_month.Semana4;
            Persistence.Persistence_Requests_List.Add(request);
            if (Main_Form.week_five_exist)
            {
                request.persistence_insight = Main_Form.Insight_month.Semana5;
                Persistence.Persistence_Requests_List.Add(request);
            }
            Main_Form.Notify("Save complete!");
        }

        public static void VyM_Builder(Insight_Sem sem)
        {
            //if(sem.Overwatch_Aprobal)
            {
                CultureInfo spanish = new CultureInfo("es-MX");
                Excel.Range range;
                if(sem.Special_Week == Main_Form.Special_Week_Type.Conv_type)
                {
                    //Handler for Convention type
                }
                else
                {
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
                    if (sem.Special_Week == Main_Form.Special_Week_Type.Visit_type)
                    {
                        //Handler for Visit type
                    }
                    //Horario
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
                    range = Sheet_Week.get_Range("H" + (cell + 1).ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell + 1, H] = "Consejero Auxiliar";
                    range = Sheet_Week.get_Range("K" + (cell + 1).ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 9;
                    Sheet_Week.Cells[cell + 1, K] = sem.Consejero_Aux;
                    cell += 2;
                    //Cancion y Palabras de Introduccion
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "E" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    Sheet_Week.Cells[cell, A] = "• Cancion y Oracion Inicial";
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
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Sheet_Week.Cells[cell, H] = "Sala Auxiliar";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Sheet_Week.Cells[cell, K] = "Sala Principal";
                    cell++;
                    //Discurso
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
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
                    Sheet_Week.Cells[cell, A] = "• " + sem.Lectura;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Lectura_A;
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, H] = sem.Lectura_B;
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
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.SMM1;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
                    range.Cells.WrapText = true;
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, H] = sem.SMM1_B;
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.SMM1_A;
                    cell += 2;
                    //Asignacion 2                    
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.SMM2;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
                    range.Cells.WrapText = true;
                    range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, H] = sem.SMM2_B;
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.SMM2_A;
                    cell += 2;
                    //Asignacion 3
                    if(sem.SMM3 != null && sem.SMM3.Length > 5)
                    {
                        range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + (cell + 1).ToString());
                        range.Cells.Merge();
                        Sheet_Week.Cells[cell, A] = "• " + sem.SMM3;
                        Set_Font(Sheet_Week.get_Range("A" + cell));
                        range.Cells.WrapText = true;
                        range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                        range.Cells.Merge();
                        Sheet_Week.Cells[cell, H] = sem.SMM3_B;
                        range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                        range.Cells.Merge();
                        Sheet_Week.Cells[cell, K] = sem.SMM3_A;
                        cell += 2;
                    }
                    //Asignacion 4
                    if (sem.SMM4 != null && sem.SMM4.Length > 5)
                    {
                        range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + (cell + 1).ToString());
                        range.Cells.Merge();
                        Sheet_Week.Cells[cell, A] = "• " + sem.SMM4;
                        Set_Font(Sheet_Week.get_Range("A" + cell));
                        range.Cells.WrapText = true;
                        range = Sheet_Week.get_Range("H" + cell.ToString(), "J" + (cell + 1).ToString());
                        range.Cells.Merge();
                        Sheet_Week.Cells[cell, H] = sem.SMM4_B;
                        range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + (cell + 1).ToString());
                        range.Cells.Merge();
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
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• Cancion";
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
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.Libro_Titulo;
                    Set_Font(Sheet_Week.get_Range("A" + cell));
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
                    cell += 2;
                    //Repaso
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• Repaso de esta reunion, adelanto e la proxima y anuncios (3 mins.)";
                    Set_Font(Sheet_Week.get_Range("B" + cell));
                    cell++;
                    //Cancion y Oracion Final
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "E" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• Cancion y Oracion Final"; 
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Oracion_End_VyM;

                }
            }
        }
        public static void RP_Builder(Insight_Sem sem)
        {
            //if (sem.Overwatch_Aprobal)
            {
                CultureInfo spanish = new CultureInfo("es-MX");
                Excel.Range range;
                if (sem.Special_Week == Main_Form.Special_Week_Type.Conv_type)
                {
                    //Handler for Convention type
                }
                else
                {
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
                    if (sem.Special_Week == Main_Form.Special_Week_Type.Visit_type)
                    {
                        //Handler for Visit type
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
                    Sheet_Week.Cells[cell, K] = sem.Presidente_RP; 
                    cell += 2;
                    //Seccion "Reunion Publica"
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Bold = true;
                    range.Cells.Font.Size = 11;
                    range.Cells.Font.Color = Color.White;
                    Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(68, 84, 106);
                    Sheet_Week.Cells[cell, A] = "REUNION PUBLICA";
                    cell++;
                    //Informacion del discurso
                    Sheet_Week.Cells[cell, A] = "• Tema";
                    range = Sheet_Week.get_Range("B" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, B] = sem.Titulo_RP;
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Sheet_Week.Cells[cell, K] = sem.Discursante;
                    cell++;
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "C" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "Congregacion";
                    range = Sheet_Week.get_Range("D" + cell.ToString(), "I" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, D] = sem.Congregacion_RP;
                    cell += 2;
                    //Seccion "Analisis de La Atalaya"
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "G" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Bold = true;
                    range.Cells.Font.Size = 11;
                    range.Cells.Font.Color = Color.White;
                    Sheet_Week.Cells[cell, A].Interior.Color = Color.FromArgb(84, 130, 53);
                    Sheet_Week.Cells[cell, A] = "ANALISIS DE LA ATALAYA";
                    cell++;
                    //Informacion de La Atalaya
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "H" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, A] = "• " + sem.Titulo_Atalaya;
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + (cell + 1).ToString());
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
                    range = Sheet_Week.get_Range("I" + cell.ToString(), "J" + cell.ToString());
                    range.Cells.Merge();
                    range.Cells.Font.Size = 8;
                    range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    Sheet_Week.Cells[cell, I] = "Oracion Final";
                    range = Sheet_Week.get_Range("K" + cell.ToString(), "M" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, K] = sem.Oracion_End_RP;
                    cell++;
                    //Salidas a Discursar
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
        }
        public static void AC_Builder(Insight_Sem sem)
        {
            //if (sem.Overwatch_Aprobal)
            {
                CultureInfo spanish = new CultureInfo("es-MX");
                Excel.Range range;
                if (sem.Special_Week == Main_Form.Special_Week_Type.Conv_type)
                {
                    //Handler for Convention type
                }
                else 
                {
                    //Informacion de semana
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "M" + cell.ToString());
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;
                    range = Sheet_Week.get_Range("A" + cell.ToString(), "I" + cell.ToString());
                    range.Cells.Merge();
                    if (sem.Special_Week == Main_Form.Special_Week_Type.Visit_type)
                    {
                        //Handler for Visit type
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
                    //Acomodadores VyM
                    range = Sheet_Week.get_Range("D" + cell.ToString(), "E" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, D] = "Capitan";
                    range = Sheet_Week.get_Range("F" + cell.ToString(), "H" + cell.ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell, F] = sem.Vym_Cap; 
                    range = Sheet_Week.get_Range("D" + (cell+1).ToString(), "E" + (cell+2).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, D] = "Acomodadores";
                    range = Sheet_Week.get_Range("F" + (cell + 1).ToString(), "H" + (cell + 1).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 1, F] = sem.Vym_Der;
                    range = Sheet_Week.get_Range("F" + (cell + 2).ToString(), "H" + (cell + 2).ToString());
                    range.Cells.Merge();
                    Sheet_Week.Cells[cell + 2, F] = sem.Vym_Izq;
                    //Acomodadores RP
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
        }

        /*-----------------------------------------END EXPERIMENTAL-------------------------------------------*/


        public static void VyM_Save_Week(VyM_Sem sem)
        {
            short num_sem = sem.Num_of_Week;
            short primary_cell = Get_vym_cell(num_sem);
            short increment_smm = 0;
            bool increment_nvc = false;
            Sheet_VyM.PageSetup.LeftHeader = "&16&B" + Main_Form.Cong_Name;
            if (num_sem == Main_Form.Vst_Wk)
            {
                Excel.Range range;
                Sheet_VyM.Cells[primary_cell, A] = "Visita del Superintendente de Circuito";
                range = Sheet_VyM.get_Range("A" + primary_cell.ToString());
                range.Interior.Color = Color.Orange;
            }
            if (((num_sem + 10) % 2) != 0)
            {
                CultureInfo spanish = new CultureInfo("es-MX");
                Sheet_VyM.Cells[primary_cell, G] = Main_Form.meetings_days[(num_sem - 1), 0].ToString("dddd", spanish) + " " + Main_Form.VyM_horary.ToString("hh:mm tt");
            }
            Excel.Range aux_range;
            primary_cell++;
            Sheet_VyM.Cells[primary_cell, A] = sem.Fecha.ToString("dddd, dd MMMM");
            if (num_sem != Main_Form.Conv_Wk)
            {
                if ((sem.Sem_Biblia != null) && (sem.Sem_Biblia != ""))
                {
                    Main_Form.Notify("Saving VyM Week: " + sem.Num_of_Week.ToString());
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
                        if (Main_Form.Room_B_enabled)
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
                        if (Main_Form.Room_B_enabled)
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
                    if (Main_Form.Room_B_enabled)
                    {
                        Sheet_VyM.Cells[primary_cell + 8, F] = sem.Lectura_B;
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
                    if (sem.Num_of_Week == Main_Form.Vst_Wk)
                    {
                        Excel.Range range;
                        range = Sheet_VyM.get_Range("F" + (primary_cell + 19).ToString(), g + (primary_cell + 20).ToString());
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

                    /*string[] Time_data = Get_time_from_week(sem);
                    Sheet_VyM.Cells[primary_cell + 2, A] = Time_data[0];
                    Sheet_VyM.Cells[primary_cell + 3, A] = Time_data[1];
                    Sheet_VyM.Cells[primary_cell + 6, A] = Time_data[2];
                    Sheet_VyM.Cells[primary_cell + 7, A] = Time_data[3];
                    Sheet_VyM.Cells[primary_cell + 8, A] = Time_data[4];
                    Sheet_VyM.Cells[primary_cell + 11, A] = Time_data[5];
                    Sheet_VyM.Cells[primary_cell + 12, A] = Time_data[6];
                    Sheet_VyM.Cells[primary_cell + 13, A] = Time_data[7];
                    Sheet_VyM.Cells[primary_cell + 14, A] = Time_data[8];
                    Sheet_VyM.Cells[primary_cell + 16, A] = Time_data[9];
                    Sheet_VyM.Cells[primary_cell + 17, A] = Time_data[10];
                    Sheet_VyM.Cells[primary_cell + 18, A] = Time_data[11];
                    Sheet_VyM.Cells[primary_cell + 19, A] = Time_data[12];
                    Sheet_VyM.Cells[primary_cell + 21, A] = Time_data[13];
                    Sheet_VyM.Cells[primary_cell + 22, A] = Time_data[14];*/


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

                    /*Persistence Request*/
                    Persistence.Persistence_Request request_pers = new Persistence.Persistence_Request
                    {
                        //persistence_insight = sem,
                    };
                    Persistence.Persistence_Requests_List.Add(request_pers);

                    /*Heavensward Request*/
                    Heavensward.HW_Oracle_Request request_hw = new Heavensward.HW_Oracle_Request
                    {
                        //hw_oracle_vym = sem
                    };
                    Heavensward.HW_Oracle_Requests_List.Add(request_hw);
                }
            }
            else
            {
                Convention_Handler(1);
            }
        }

        public static void VyM_Read_Week(short num_sem)
        {
            VyM_Sem sem = new VyM_Sem();
            short primary_cell = Get_vym_cell(num_sem);
            Main_Form.Notify("Read VyM week: " + num_sem.ToString()) ;
            Get_month_from_Excel(cellValue_1[primary_cell + 1, A]);
            if (Check_null_string(cellValue_1[primary_cell, A]).Contains("Visita"))
            {
                Main_Form.Vst_Wk = num_sem;
                Main_Form.Notify("Marked current week [" + num_sem + "] as Visit");
                /*Alert_Label_VyM.Text = "Semana de la Visita del Superintendente de Circuito";
                Alert_Label_VyM.Visible = true;*/
            }
            else if (Check_null_string(cellValue_1[primary_cell + 10, A]).Contains("Asamblea"))
            {
                Main_Form.Conv_Wk = num_sem;
                Main_Form.Notify("Current week [" + num_sem.ToString() + "] setting as Convention [" + (Main_Form.Conv_Wk > 0 ? "True" : "False") + "]");
                Main_Form.Conv_Name = Check_null_string(cellValue_1[primary_cell + 10, A]);
                /*Alert_Label_VyM.Text = "Semana de Asamblea!";
                Alert_Label_VyM.Visible = true;*/
            }
            sem.Sem_Biblia = Check_null_string(cellValue_1[primary_cell + 1, D]);
            sem.Presidente = Check_null_string(cellValue_1[primary_cell + 1, G]);
            sem.Consejero_Aux = Check_null_string(cellValue_1[primary_cell + 2, G]);
            sem.Discurso = Check_null_string(cellValue_1[primary_cell + 7, C]);
            sem.Discurso_A = Check_null_string(cellValue_1[primary_cell + 7, G]);
            sem.Perlas = Check_null_string(cellValue_1[primary_cell + 8, G]);
            sem.Lectura = Check_null_string(cellValue_1[primary_cell + 9, C]);
            sem.Lectura_A = Check_null_string(cellValue_1[primary_cell + 9, G]);
            sem.Lectura_B = Check_null_string(cellValue_1[primary_cell + 9, F]);
            sem.SMM1 = Check_null_string(cellValue_1[primary_cell + 12, C]);
            sem.SMM1_A = Check_null_string(cellValue_1[primary_cell + 12, G]);
            sem.SMM1_B = Check_null_string(cellValue_1[primary_cell + 12, F]);
            sem.SMM2 = Check_null_string(cellValue_1[primary_cell + 13, C]);
            sem.SMM2_A = Check_null_string(cellValue_1[primary_cell + 13, G]);
            sem.SMM2_B = Check_null_string(cellValue_1[primary_cell + 13, F]);
            sem.SMM3 = Check_null_string(cellValue_1[primary_cell + 14, C]);
            sem.SMM3_A = Check_null_string(cellValue_1[primary_cell + 14, G]);
            sem.SMM3_B = Check_null_string(cellValue_1[primary_cell + 14, F]);
            sem.SMM4 = Check_null_string(cellValue_1[primary_cell + 15, C]);
            sem.SMM4_A = Check_null_string(cellValue_1[primary_cell + 15, G]);
            sem.SMM4_B = Check_null_string(cellValue_1[primary_cell + 15, F]);
            sem.NVC1 = Check_null_string(cellValue_1[primary_cell + 18, C]);
            sem.NVC1_A = Check_null_string(cellValue_1[primary_cell + 18, G]);
            sem.NVC2 = Check_null_string(cellValue_1[primary_cell + 19, C]);
            sem.NVC2_A = Check_null_string(cellValue_1[primary_cell + 19, G]);
            sem.Libro_Titulo = Check_null_string(cellValue_1[primary_cell + 20, C]);
            sem.Libro_A = Check_null_string(cellValue_1[primary_cell + 20, G]);
            sem.Libro_L = Check_null_string(cellValue_1[primary_cell + 21, G]);
            sem.Oracion = Check_null_string(cellValue_1[primary_cell + 23, G]);
            sem.Num_of_Week = num_sem;

            switch (num_sem)
            {
                case 1:
                    {
                        Main_Form.VyM_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        Main_Form.VyM_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        Main_Form.VyM_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        Main_Form.VyM_mes.Semana4 = sem;
                        break;
                    }
                case 5:
                    {
                        Main_Form.VyM_mes.Semana5 = sem;
                        break;
                    }
            }
        }

        public static void RP_Handler(bool save)
        {
            if (save)
            {
                RP_Save_Week(Main_Form.RP_mes.Semana1);
                loading += (15 / loading_delta);
                RP_Save_Week(Main_Form.RP_mes.Semana2);
                loading += (15 / loading_delta);
                RP_Save_Week(Main_Form.RP_mes.Semana3);
                loading += (15 / loading_delta);
                RP_Save_Week(Main_Form.RP_mes.Semana4);
                loading += (15 / loading_delta);
                if (Main_Form.week_five_exist)
                {
                    RP_Save_Week(Main_Form.RP_mes.Semana5);
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

        public static void RP_Save_Week(RP_Sem sem)
        {
            short num_sem = sem.Num_of_Week;
            short primary_cell = Get_rp_cell(num_sem);
            Sheet_RP.PageSetup.LeftHeader = "&16&B" + Main_Form.Cong_Name;
            if (num_sem == Main_Form.Vst_Wk)
            {
                Excel.Range range;
                Sheet_RP.Cells[primary_cell, A] = "Visita del Superintendente de Circuito";
                range = Sheet_RP.get_Range("A" + primary_cell.ToString());
                range.Interior.Color = Color.Orange;
            }
            primary_cell++;
            Sheet_RP.Cells[primary_cell, C] = sem.Fecha.ToString("dddd, dd MMMM");
            if (num_sem != Main_Form.Conv_Wk)
            {
                if (sem.Presidente != null)
                {
                    Main_Form.Notify("Saving RP Week: " + sem.Num_of_Week.ToString());
                    Sheet_RP.Cells[primary_cell + 1, H] = sem.Presidente;
                    Sheet_RP.Cells[primary_cell + 2, D] = sem.Titulo;
                    if (sem.Titulo.ToLower().Contains("pendiente"))
                    {
                        Sheet_RP.Cells[primary_cell + 2, D].Font.Italic = true; // D = 4
                    }
                    Sheet_RP.Cells[primary_cell + 2, H] = sem.Discursante;
                    Sheet_RP.Cells[primary_cell + 3, E] = sem.Congregacion;
                    Sheet_RP.Cells[primary_cell + 6, D] = sem.Titulo_Atalaya;
                    Sheet_RP.Cells[primary_cell + 5, H] = sem.Conductor;
                    Sheet_RP.Cells[primary_cell + 6, H] = sem.Lector;
                    Sheet_RP.Cells[primary_cell + 7, H] = sem.Oracion;
                    Sheet_RP.Cells[primary_cell + 10, C] = sem.Discu_Sal;
                    Sheet_RP.Cells[primary_cell + 10, E] = sem.Ttl_Sal;
                    Sheet_RP.Cells[primary_cell + 10, H] = sem.Cong_Sal;

                    /*Persistence Request*/
                    Persistence.Persistence_Request request_pers = new Persistence.Persistence_Request
                    {
                        //persistence_rp = sem,
                    };
                    Persistence.Persistence_Requests_List.Add(request_pers);

                    //Heavensward Oracle set request
                    Heavensward.HW_Oracle_Request request_hw = new Heavensward.HW_Oracle_Request
                    {
                        hw_oracle_rp = sem
                    };
                    Heavensward.HW_Oracle_Requests_List.Add(request_hw);
                }
            }
            else
            {
                Convention_Handler(2);
            }
        }

        public static void RP_Read_Week(short num_sem)
        {
            RP_Sem sem = new RP_Sem();
            short primary_cell = Get_rp_cell(num_sem);
            Main_Form.Notify("Read RP week: " + num_sem.ToString());
            primary_cell++;
            sem.Presidente = Check_null_string(cellValue_2[primary_cell + 1, H]);
            sem.Titulo = Check_null_string(cellValue_2[primary_cell + 2, D]);
            sem.Discursante = Check_null_string(cellValue_2[primary_cell + 2, H]);
            sem.Congregacion = Check_null_string(cellValue_2[primary_cell + 3, E]);
            sem.Titulo_Atalaya = Check_null_string(cellValue_2[primary_cell + 6, D]);
            sem.Conductor = Check_null_string(cellValue_2[primary_cell + 5, H]);
            sem.Lector = Check_null_string(cellValue_2[primary_cell + 6, H]);
            sem.Oracion = Check_null_string(cellValue_2[primary_cell + 7, H]);
            sem.Discu_Sal = Check_null_string(cellValue_2[primary_cell + 10, C]);
            sem.Ttl_Sal = Check_null_string(cellValue_2[primary_cell + 10, E]);
            sem.Cong_Sal = Check_null_string(cellValue_2[primary_cell + 10, H]);
            sem.Num_of_Week = num_sem;
            switch (num_sem)
            {
                case 1:
                    {
                        Main_Form.RP_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        Main_Form.RP_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        Main_Form.RP_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        Main_Form.RP_mes.Semana4 = sem;
                        break;
                    }
                case 5:
                    {
                        Main_Form.RP_mes.Semana5 = sem;
                        break;
                    }
            }
        }

        public static void AC_Handler(bool save)
        {
            if (save)
            {
                AC_Save_Week(Main_Form.AC_mes.Semana1);
                loading += (15 / loading_delta);
                AC_Save_Week(Main_Form.AC_mes.Semana2);
                loading += (15 / loading_delta);
                AC_Save_Week(Main_Form.AC_mes.Semana3);
                loading += (15 / loading_delta);
                AC_Save_Week(Main_Form.AC_mes.Semana4);
                loading += (15 / loading_delta);
                if (Main_Form.week_five_exist)
                {
                    AC_Save_Week(Main_Form.AC_mes.Semana5);
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

        public static void AC_Save_Week(AC_Sem sem)
        {
            short num_sem = sem.Num_of_Week;
            short primary_cell = Get_ac_cell(num_sem);
            Sheet_AC.PageSetup.LeftHeader = "&16&B" + Main_Form.Cong_Name;
            if (num_sem == Main_Form.Vst_Wk)
            {
                Excel.Range range;
                Sheet_AC.Cells[primary_cell, A] = "Visita del Superintendente de Circuito";
                range = Sheet_AC.get_Range("A" + primary_cell.ToString());
                range.Interior.Color = Color.Orange;
            }
            primary_cell++;
            Main_Form.Notify("Saving AC Week: " + sem.Num_of_Week.ToString());
            if (Main_Form.Ac_same_all_week)
            {
                string Date = "Semana del " + sem.Fecha_VyM.ToString("dddd dd") + " de " + sem.Fecha_VyM.ToString("MMMM")
                    + " al " + sem.Fecha_RP.ToString("dddd dd") + " de " + sem.Fecha_RP.ToString("MMMM");
                string b = "B", c = "C", d = "D", e = "E";
                Excel.Range aux_range = Sheet_AC.get_Range(b + primary_cell.ToString(), d + primary_cell.ToString());
                aux_range.Cells.Merge();
                Sheet_AC.Cells[primary_cell, B] = Date;
                if (num_sem != Main_Form.Conv_Wk)
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

                        /*Persistence Request*/
                        Persistence.Persistence_Request request_pers = new Persistence.Persistence_Request
                        {
                            //persistence_ac = sem,
                        };
                        Persistence.Persistence_Requests_List.Add(request_pers);

                        //Heavensward Oracle set request
                        Heavensward.HW_Oracle_Request request_hw = new Heavensward.HW_Oracle_Request
                        {
                            hw_oracle_ac = sem
                        };
                        Heavensward.HW_Oracle_Requests_List.Add(request_hw);
                    }
                }
                else
                {
                    Convention_Handler(3);
                }
            }
            else
            {
                Sheet_AC.Cells[primary_cell, B] = sem.Fecha_VyM.ToString("dddd, dd MMMM");
                Sheet_AC.Cells[primary_cell, D] = sem.Fecha_RP.ToString("dddd, dd MMMM");
                if (num_sem != Main_Form.Conv_Wk)
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

                        /*Persistence Request*/
                        Persistence.Persistence_Request request_pers = new Persistence.Persistence_Request
                        {
                            //persistence_ac = sem,
                        };
                        Persistence.Persistence_Requests_List.Add(request_pers);

                        //Heavensward Oracle set request                        
                        Heavensward.HW_Oracle_Request request_hw = new Heavensward.HW_Oracle_Request
                        {
                            hw_oracle_ac = sem
                        };
                        Heavensward.HW_Oracle_Requests_List.Add(request_hw);
                    }
                }
                else
                {
                    Convention_Handler(3);
                }
            }
        }

        public static void AC_Read_Week(short num_sem)
        {
            AC_Sem sem = new AC_Sem();
            short primary_cell = Get_ac_cell(num_sem);
            Main_Form.Notify("Read AC week: " + num_sem.ToString());
            primary_cell++;
            sem.Aseo = Check_null_string(cellValue_3[primary_cell + 2, A]);
            sem.Vym_Cap = Check_null_string(cellValue_3[primary_cell + 2, C]);
            sem.Vym_Izq = Check_null_string(cellValue_3[primary_cell + 3, C]);
            sem.Vym_Der = Check_null_string(cellValue_3[primary_cell + 4, C]);
            sem.Rp_Cap = Check_null_string(cellValue_3[primary_cell + 2, E]);
            sem.Rp_Izq = Check_null_string(cellValue_3[primary_cell + 3, E]);
            sem.Rp_Der = Check_null_string(cellValue_3[primary_cell + 4, E]);
            sem.Num_of_Week = num_sem;
            switch (num_sem)
            {
                case 1:
                    {
                        Main_Form.AC_mes.Semana1 = sem;
                        break;
                    }
                case 2:
                    {
                        Main_Form.AC_mes.Semana2 = sem;
                        break;
                    }
                case 3:
                    {
                        Main_Form.AC_mes.Semana3 = sem;
                        break;
                    }
                case 4:
                    {
                        Main_Form.AC_mes.Semana4 = sem;
                        break;
                    }
                case 5:
                    {
                        Main_Form.AC_mes.Semana5 = sem;
                        break;
                    }
            }
        }

        /*Function to set proper format to program when a Convention ocurrs*/
        public static void Convention_Handler(short program)
        {
            Excel.Range range;
            string a = "A", g = "G", e = "E", h = "H", f = "F";
            switch (program)
            {
                case 1: //VyM
                    {
                        short cell = Get_vym_cell(Main_Form.Conv_Wk);
                        range = Sheet_VyM.get_Range(f + (cell + 1).ToString(), g + (cell + 2).ToString());
                        range.Cells.UnMerge();
                        range.Cells.Clear();
                        range = Sheet_VyM.get_Range(a + (cell + 3).ToString(), g + (cell + 23).ToString());
                        range.Cells.UnMerge();
                        range.Cells.Clear();
                        range = Sheet_VyM.get_Range(a + (cell + 10).ToString(), g + (cell + 15).ToString());
                        range.Cells.Merge();
                        range.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Characters.Font.Size = 16;
                        range.Interior.Color = Color.Orange;
                        Sheet_VyM.Cells[cell + 10, A] = Main_Form.Conv_Name;
                        Sheet_VyM.Cells[cell, F] = "";
                        Sheet_VyM.Cells[cell + 1, F] = "";
                        break;
                    }
                case 2: //RP
                    {
                        short cell = Get_rp_cell(Main_Form.Conv_Wk);
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
                        Sheet_RP.Cells[cell + 2, A] = Main_Form.Conv_Name;
                        break;
                    }
                case 3: //AC
                    {
                        short cell = Get_ac_cell(Main_Form.Conv_Wk);
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
                        Sheet_AC.Cells[cell + 1, A] = Main_Form.Conv_Name;
                        break;
                    }
            }
        }

        /*Set font size to (x min.)*/
        public static void Set_Font(Excel.Range cell)
        {
            string Str = cell.Text;
            int index = Str.IndexOf("min");
            if (index > 4)
            {
                cell.Characters[index - 3, Str.Length].Font.Size = 8;
            }/*
            if (Str.Length > 135)
            {
                //find next space
                for (int i = 135; i < Str.Length; i++)
                {
                    if (Str.Substring(i, 1).Equals(" "))
                    {
                        Str = Str.Substring(0, i);
                        cell.Characters[0, i].Text = Str;
                    }

                }
            }*/
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
            if (cellvalue == null)
            {
                cellvalue = "";
            }
            return cellvalue.ToString();
        }

        public static short Get_vym_cell(short num_sem)
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

        public static short Get_rp_cell(short num_sem)
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

        public static short Get_ac_cell(short num_sem)
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

        public static string[] Get_time_from_week(Insight_Sem sem)
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

        public static int Get_time_from_string(string Str)
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
                        Main_Form.Warn("Must be numbers");
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
                Marshal.ReleaseComObject(objBooks);
                Marshal.ReleaseComObject(objApp);
            }
        }
    }
}
