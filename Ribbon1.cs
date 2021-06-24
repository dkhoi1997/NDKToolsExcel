using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using System.Linq;
using Autodesk.AutoCAD.Interop;

namespace NDKToolsExcel
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void LoadDataButton_Click(object sender, RibbonControlEventArgs e) // Lấy dữ liệu cột, vách
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            int last_row = xlSheet.UsedRange.Rows.Count;

            // Clean data cũ trước
            if (last_row > 10)
            {
                xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[last_row, 28]].EntireRow.Delete();
                _ = xlSheet.UsedRange;
            }

            // Lấy dữ liệu theo TypeMember
            var instance = new NDKFunction();
            string TypeMember = xlSheet.Range["TypeMember"].Value2;
            object[,] out_import = instance.MultiThreadingImport(TypeMember);
            if (out_import != null)
            {
                xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[out_import.GetUpperBound(0) + 11, 16]].Value2 = out_import;
                xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[out_import.GetUpperBound(0) + 11, 29]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void ModifyButton_Click(object sender, RibbonControlEventArgs e) // Điều chỉnh thông số cột, vách
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            int last_row = xlSheet.UsedRange.Rows.Count;
            var instance = new NDKFunction();
            object[,] in_modify = xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[last_row, 16]].Value2;
            object[,] out_modify = instance.ModifyMultiThreading(in_modify);
            xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[last_row, 16]].Value2 = out_modify;
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void AssignCADButton_Click(object sender, RibbonControlEventArgs e) // Gán tên CAD cột, vách
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            int last_row = xlSheet.UsedRange.Rows.Count;
            object[,] in_assign = xlSheet.Range[xlSheet.Cells[11, 3], xlSheet.Cells[last_row, 5]].Value2;
            var instance = new NDKFunction();
            object[,] out_assign = instance.AssignCADName(in_assign);
            xlSheet.Range[xlSheet.Cells[11, 3], xlSheet.Cells[last_row, 5]].Value2 = out_assign;
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void CalculateRebarPercentButton_Click(object sender, RibbonControlEventArgs e) // Tính toán hàm lượng sơ bộ
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            int last_row = xlSheet.UsedRange.Rows.Count;
            int abv = Convert.ToInt32(xlSheet.Range["abv"].Value2);
            double Rb = xlSheet.Range["Rb"].Value2 * xlSheet.Range["gb_1"].Value2 * xlSheet.Range["gb_2"].Value2;
            double Eb = xlSheet.Range["Eb"].Value2;
            double Rs = xlSheet.Range["Rs"].Value2 * xlSheet.Range["gs"].Value2;
            double Rsc = xlSheet.Range["Rsc"].Value2 * xlSheet.Range["gs"].Value2;
            double Es = xlSheet.Range["Es"].Value2;
            double umin = xlSheet.Range["umin"].Value2;
            double umax = xlSheet.Range["umax"].Value2;
            object[,] in_rebarpercent = xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[last_row, 16]].Value2;
            var instance = new NDKFunction();
            object[,] out_rebarpercent = instance.MultiThreadingCalculate(in_rebarpercent, abv, Rb, Eb, Rs, Rsc, Es, umin, umax);
            xlSheet.Range[xlSheet.Cells[11, 17], xlSheet.Cells[last_row, 19]].Value2 = out_rebarpercent;
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void DesignCheck_Click(object sender, RibbonControlEventArgs e) // Thiết kế và kiểm tra
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            int last_row = xlSheet.UsedRange.Rows.Count;
            int abv = Convert.ToInt32(xlSheet.Range["abv"].Value2);
            double Rb = xlSheet.Range["Rb"].Value2 * xlSheet.Range["gb_1"].Value2 * xlSheet.Range["gb_2"].Value2; ;
            double Rbt = xlSheet.Range["Rbt"].Value2 * xlSheet.Range["gb_1"].Value2;
            double Eb = xlSheet.Range["Eb"].Value2;
            double Rs = xlSheet.Range["Rs"].Value2 * xlSheet.Range["gs"].Value2;
            double Rsc = xlSheet.Range["Rsc"].Value2 * xlSheet.Range["gs"].Value2;
            double Rsw = xlSheet.Range["Rsw"].Value2 * xlSheet.Range["gsw"].Value2;
            double Es = xlSheet.Range["Es"].Value2;
            int nsec = Convert.ToInt32(xlSheet.Range["nsec"].Value2);
            var instance = new NDKFunction();
            string[] loadcomb = xlSheet.Range["EQCase"].Cells.Cast<Excel.Range>().Select(instance.Selector).ToArray();
            object[,] in_designcheck = xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[last_row + 1, 26]].Value2;
            object[,] out_designcheck = instance.DesignResult(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, abv, nsec, in_designcheck, loadcomb);
            xlSheet.Range[xlSheet.Cells[11, 20], xlSheet.Cells[last_row, 29]].Value2 = out_designcheck;
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void Filter_Click(object sender, RibbonControlEventArgs e) // Lọc dữ liệu
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            Excel.Worksheet xlFilter = window.Application.Sheets["Filter"];
            int last_row1 = xlSheet.UsedRange.Rows.Count;
            int last_row2 = xlFilter.UsedRange.Rows.Count;
            var instance = new NDKFunction();
            object[,] in_filter = xlSheet.Range[xlSheet.Cells[11, 2], xlSheet.Cells[last_row1, 29]].Value2;
            object[,] out_filter = instance.Filter(in_filter);
            

            // Clean data cũ trước
            if (last_row2 > 4)
            {
                xlFilter.Range[xlFilter.Cells[5, 2], xlFilter.Cells[last_row2, 32]].EntireRow.Delete();
                _ = xlFilter.UsedRange;
            }
            xlFilter.Range[xlFilter.Cells[5, 2], xlFilter.Cells[out_filter.GetUpperBound(0) + 5, 32]] = out_filter;
            xlFilter.Range[xlFilter.Cells[5, 2], xlFilter.Cells[out_filter.GetUpperBound(0) + 5, 32]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            xlFilter.Activate();
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void ShowMember_Click(object sender, RibbonControlEventArgs e) // Hiện thị chi tiết kết quả
        {
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            Excel.Worksheet xlInstance = window.Application.ActiveSheet;
            if (xlInstance.Name != "Filter") MessageBox.Show("Chuyển sang Sheet Filter");
            else
            {
                int abv = Convert.ToInt32(xlSheet.Range["abv"].Value2);
                int nsec = Convert.ToInt32(xlSheet.Range["nsec"].Value2);
                double Rb = xlSheet.Range["Rb"].Value2 * xlSheet.Range["gb_1"].Value2 * xlSheet.Range["gb_2"].Value2; ;
                double Rbt = xlSheet.Range["Rbt"].Value2 * xlSheet.Range["gb_1"].Value2;
                double Eb = xlSheet.Range["Eb"].Value2;
                double Rs = xlSheet.Range["Rs"].Value2 * xlSheet.Range["gs"].Value2;
                double Rsc = xlSheet.Range["Rsc"].Value2 * xlSheet.Range["gs"].Value2;
                double Rsw = xlSheet.Range["Rsw"].Value2 * xlSheet.Range["gsw"].Value2;
                double Es = xlSheet.Range["Es"].Value2;
                Excel.Range rng = window.Application.ActiveCell;
                int row = rng.Row;
                // Các giá trị cần thiết để pass sang form show
                string shape = Convert.ToString(xlInstance.Cells[row, 6].Value2);
                double L = Convert.ToDouble(xlInstance.Cells[row, 7].Value2);
                double Cx = Convert.ToDouble(xlInstance.Cells[row, 8].Value2);
                double Cy = Convert.ToDouble(xlInstance.Cells[row, 9].Value2);
                int P = -1 * Convert.ToInt32(xlInstance.Cells[row, 11].Value2);
                int Mx = Math.Abs(Convert.ToInt32(xlInstance.Cells[row, 12].Value2));
                int My = Math.Abs(Convert.ToInt32(xlInstance.Cells[row, 13].Value2));
                int nx = Convert.ToInt32(xlInstance.Cells[row, 17].Value2);
                int ny = Convert.ToInt32(xlInstance.Cells[row, 18].Value2);
                string dmain = Convert.ToString(xlInstance.Cells[row, 19].Value2);
                int Pq = -1 * Convert.ToInt32(xlInstance.Cells[row, 24].Value2);
                int Qx = Math.Abs(Convert.ToInt32(xlInstance.Cells[row, 25].Value2));
                int Qy = Math.Abs(Convert.ToInt32(xlInstance.Cells[row, 26].Value2));
                string dstir = Convert.ToString(xlInstance.Cells[row, 27].Value2);
                int nstir = Convert.ToInt32(xlInstance.Cells[row, 28].Value2);
                int Ped = -1 * Convert.ToInt32(xlInstance.Cells[row, 31].Value2);
                string etabs_label = Convert.ToString(xlInstance.Cells[row, 3].Value2);
                string cad_label = Convert.ToString(xlInstance.Cells[row, 5].Value2);
                string story = Convert.ToString(xlInstance.Cells[row, 2].Value2);
                double umin = xlSheet.Range["umin"].Value2;
                double umax = xlSheet.Range["umax"].Value2;
                if ((shape == "Rec") || (shape == "Wall"))
                {
                    RecColShow_Form myform = new RecColShow_Form(Cx, Cy, Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, L, abv, nsec,
                    dmain, nx, ny, dstir, nstir, P, Mx, My, Pq, Qx, Qy, Ped, shape, cad_label, etabs_label, story, umin, umax);
                    myform.ShowDialog();
                    object[] recieve1 = myform.SendBack().Item1;
                    object[] recieve2 = myform.SendBack().Item2;
                    object[] recieve3 = myform.SendBack().Item3;
                    object[] recieve4 = myform.SendBack().Item4;

                    xlInstance.Range[xlInstance.Cells[row, 7], xlInstance.Cells[row, 9]] = recieve1;
                    xlInstance.Range[xlInstance.Cells[row, 11], xlInstance.Cells[row, 22]] = recieve2;
                    xlInstance.Range[xlInstance.Cells[row, 24], xlInstance.Cells[row, 29]] = recieve3;
                    xlInstance.Range[xlInstance.Cells[row, 31], xlInstance.Cells[row, 32]] = recieve4;
                }
                else if (shape == "Cir")
                {
                    CirColShow_Form myform = new CirColShow_Form(Cx, Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, L, abv, nsec,
                    dmain, nx, dstir, nstir, P, Mx, My, Pq, Qx, Qy, Ped, shape, cad_label, etabs_label, story, umin, umax);
                    myform.ShowDialog();

                    object[] recieve1 = myform.SendBack().Item1;
                    object[] recieve2 = myform.SendBack().Item2;
                    object[] recieve3 = myform.SendBack().Item3;
                    object[] recieve4 = myform.SendBack().Item4;

                    xlInstance.Range[xlInstance.Cells[row, 7], xlInstance.Cells[row, 9]] = recieve1;
                    xlInstance.Range[xlInstance.Cells[row, 11], xlInstance.Cells[row, 22]] = recieve2;
                    xlInstance.Range[xlInstance.Cells[row, 24], xlInstance.Cells[row, 29]] = recieve3;
                    xlInstance.Range[xlInstance.Cells[row, 31], xlInstance.Cells[row, 32]] = recieve4;
                }
                else
                {
                    MessageBox.Show("Loại cấu kiện không phù hợp");
                }
            }




        }

        private void DrawSection_Click(object sender, RibbonControlEventArgs e) // Vẽ cấu kiện
        {
            Excel.Window window = e.Control.Context;
            Excel.Worksheet xlSheet = window.Application.Sheets["Main Sheet"];
            Excel.Worksheet xlInstance = window.Application.ActiveSheet;
            if (xlInstance.Name != "Filter") MessageBox.Show("Chuyển sang Sheet Filter");
            else
            {
                var instance = new NDKFunction();
                int abv = Convert.ToInt32(xlSheet.Range["abv"].Value2);
                Excel.Range rng = window.Application.ActiveCell;
                int row = rng.Row;
                // Các giá trị cần thiết để pass sang form show
                string shape = Convert.ToString(xlInstance.Cells[row, 6].Value2);
                double Cx = Convert.ToDouble(xlInstance.Cells[row, 8].Value2);
                double Cy = Convert.ToDouble(xlInstance.Cells[row, 9].Value2);
                int nx = Convert.ToInt32(xlInstance.Cells[row, 17].Value2);
                int ny = Convert.ToInt32(xlInstance.Cells[row, 18].Value2);
                int dmain = instance.DiameterExtract(Convert.ToString(xlInstance.Cells[row, 19].Value2));
                int dstir = instance.StirrupExtract(Convert.ToString(xlInstance.Cells[row, 27].Value2)).Item1;
                int sw = instance.StirrupExtract(Convert.ToString(xlInstance.Cells[row, 27].Value2)).Item2;
                //int nstir = Convert.ToInt32(xlInstance.Cells[row, 28].Value2);
                if (shape == "Rec")
                {
                    try
                    {
                        dynamic acApp = Marshal.GetActiveObject("AutoCAD.Application");
                        acApp.ActiveDocument.SendCommand("VCHAuto" + "\n" + Cx * 1000 + "-" + Cy * 1000 + "-" + abv + "-" + 
                            nx + "-" + ny + "-" + dmain + "-" + dstir + "-" + sw + "\n");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Không tìm thấy AutoCAD hoặc chưa load module vẽ");
                    }
                }
                else if (shape == "Wall")
                {
                    try
                    {
                        dynamic acApp = Marshal.GetActiveObject("AutoCAD.Application");
                        acApp.ActiveDocument.SendCommand("VVAuto" + "\n" + Cx * 1000 + "-" + Cy * 1000 + "-" + abv + "-" +
                            nx + "-" + ny + "\n");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Không tìm thấy AutoCAD hoặc chưa load module vẽ");
                    }
                }
                else if (shape == "Cir")
                {
                    try
                    {
                        dynamic acApp = Marshal.GetActiveObject("AutoCAD.Application");
                        acApp.ActiveDocument.SendCommand("VCTAuto" + "\n" + Cx * 1000 + "-" + abv + "-" + nx + "\n");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Không tìm thấy AutoCAD hoặc chưa load module vẽ");
                    }
                }
            }
        }
    }





    public class NDKFunction
    {
        //CÁC HÀM LOAD DỮ LIỆU
        public void DataWallTask(object[,] import_wall, object[,] arr_force,
            object[,] arr_assign, object[,] arr_section, int start, int end) // Hàm nạp chia khối lượng task nạp dữ liệu vách
        {
            int i; int j;
            for (i = start; i <= end; i++)
            {
                bool rotate = false;
                import_wall[i - 4, 0] = arr_force[i, 1]; //Story
                import_wall[i - 4, 1] = arr_force[i, 2]; //Pier Name
                for (j = 4; j <= arr_assign.GetUpperBound(0); j++)
                {
                    if ((Convert.ToString(import_wall[i - 4, 0]) == Convert.ToString(arr_assign[j, 1])) && (Convert.ToString(import_wall[i - 4, 1]) == Convert.ToString(arr_assign[j, 6])))
                    {
                        import_wall[i - 4, 2] = arr_assign[j, 4]; //Section Name
                        break;
                    }
                }
                for (j = 4; j <= arr_section.GetUpperBound(0); j++)
                {
                    if ((Convert.ToString(import_wall[i - 4, 0]) == Convert.ToString(arr_section[j, 1])) && (Convert.ToString(import_wall[i - 4, 1]) == Convert.ToString(arr_section[j, 2])))
                    {
                        import_wall[i - 4, 4] = "Wall"; //Shape
                        if (Convert.ToString(arr_section[j, 1]) == "90") //Vách đang bị xoay 90 độ
                        {
                            rotate = true;
                        }

                        import_wall[i - 4, 7] = Convert.ToDouble(arr_section[j, 16]) - Convert.ToDouble(arr_section[j, 13]); //Length
                        import_wall[i - 4, 8] = Math.Round(Convert.ToDouble(arr_section[j, 6]), 2); //L
                        import_wall[i - 4, 9] = Math.Round(Convert.ToDouble(arr_section[j, 7]), 2); //tw
                        break;
                    }
                }
                import_wall[i - 4, 3] = null;
                import_wall[i - 4, 5] = arr_force[i, 3]; //Load
                import_wall[i - 4, 6] = arr_force[i, 4]; //Station
                import_wall[i - 4, 10] = Convert.ToInt32(arr_force[i, 5]); //P
                import_wall[i - 4, 11] = Math.Abs(Convert.ToInt32(arr_force[i, 6])); //Vx
                import_wall[i - 4, 12] = Math.Abs(Convert.ToInt32(arr_force[i, 7])); //Vy
                if (rotate == true)
                {
                    import_wall[i - 4, 13] = Math.Abs(Convert.ToInt32(arr_force[i, 9])); //Mx
                    import_wall[i - 4, 14] = Math.Abs(Convert.ToInt32(arr_force[i, 10])); //My
                }
                else
                {
                    import_wall[i - 4, 13] = Math.Abs(Convert.ToInt32(arr_force[i, 10])); //Mx
                    import_wall[i - 4, 14] = Math.Abs(Convert.ToInt32(arr_force[i, 9])); //My
                }
            }
        }

        public void DataColumnTask(object[,] import_col, object[,] arr_force,
            object[,] arr_assign, object[,] arr_section, int start, int end) // Hàm nạp chia khối lượng task nạp dữ liệu cột
        {
            int i; int j;
            for (i = start; i <= end; i++)
            {
                import_col[i - 4, 0] = arr_force[i, 1]; //Story
                import_col[i - 4, 1] = arr_force[i, 2]; //Column Name
                for (j = 4; j <= arr_assign.GetUpperBound(0); j++)
                {
                    if ((Convert.ToString(import_col[i - 4, 0]) == Convert.ToString(arr_assign[j, 1])) && (Convert.ToString(import_col[i - 4, 1]) == Convert.ToString(arr_assign[j, 2])))
                    {
                        import_col[i - 4, 2] = arr_assign[j, 6]; //Section Name
                        import_col[i - 4, 7] = arr_assign[j, 5]; //Length
                        break;
                    }
                }
                for (j = 4; j <= arr_section.GetUpperBound(0); j++)
                {
                    if (Convert.ToString(import_col[i - 4, 2]) == Convert.ToString(arr_section[j, 1]))
                    {
                        import_col[i - 4, 4] = Convert.ToString(arr_section[j, 3]).Substring(9, 3); //Shape
                        import_col[i - 4, 8] = Math.Round(Convert.ToDouble(arr_section[j, 4]), 2); //Cx
                        import_col[i - 4, 9] = Math.Round(Convert.ToDouble(arr_section[j, 5]), 2); //Cy
                        break;
                    }
                }
                import_col[i - 4, 3] = null;
                import_col[i - 4, 5] = arr_force[i, 3]; //Load
                import_col[i - 4, 6] = arr_force[i, 4]; //Station
                import_col[i - 4, 10] = Convert.ToInt32(arr_force[i, 5]); //P
                import_col[i - 4, 11] = Math.Abs(Convert.ToInt32(arr_force[i, 6])); //Vx
                import_col[i - 4, 12] = Math.Abs(Convert.ToInt32(arr_force[i, 7])); //Vy
                import_col[i - 4, 13] = Math.Abs(Convert.ToInt32(arr_force[i, 10])); //Mx
                import_col[i - 4, 14] = Math.Abs(Convert.ToInt32(arr_force[i, 9])); //My
            }
        }

        public object[,] MultiThreadingImport(string mode) // Hàm nạp dữ liệu cột, vách bằng multi threading
        {
            string txtFilename; string[] name_sheet;
            object[,] import_data = null; bool col_mode = false; bool wall_mode = false;
            if (mode == "Cột")
            {
                name_sheet = new string[3] { "Column Design Forces", "Frame Assignments - Summary", "Frame Sections" };
                col_mode = true;
            }
            else if (mode == "Vách")
            {
                name_sheet = new string[3] { "Pier Design Forces", "Shell Assignments - Summary", "Pier Section Properties" };
                wall_mode = true;
            }
            else return import_data;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename = openFileDialog.FileName;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtFilename);
                    try
                    {
                        Excel.Worksheet xlForce = xlWorkbook.Worksheets[name_sheet[0]];
                        Excel.Worksheet xlAssign = xlWorkbook.Worksheets[name_sheet[1]];
                        Excel.Worksheet xlSection = xlWorkbook.Worksheets[name_sheet[2]];
                        xlApp.Visible = false;
                        int last_row; int last_col;
                        last_row = xlForce.UsedRange.Rows.Count;
                        last_col = xlForce.UsedRange.Columns.Count;
                        object[,] arr_force = xlForce.Range[xlForce.Cells[1, 1], xlForce.Cells[last_row, last_col]].Value2;
                        last_row = xlAssign.UsedRange.Rows.Count;
                        last_col = xlAssign.UsedRange.Columns.Count;
                        object[,] arr_assign = xlAssign.Range[xlAssign.Cells[1, 1], xlAssign.Cells[last_row, last_col]].Value2;
                        last_row = xlSection.UsedRange.Rows.Count;
                        last_col = xlSection.UsedRange.Columns.Count;
                        object[,] arr_section = xlSection.Range[xlSection.Cells[1, 1], xlSection.Cells[last_row, last_col]].Value2;
                        xlWorkbook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkbook);
                        Marshal.ReleaseComObject(xlApp);
                        int num; int i;
                        num = arr_force.GetLength(0) - 3;
                        import_data = new object[num, 15];
                        // Chia luồng để load dữ liệu vào
                        int[,] range = DetermineTask(4, num + 3).Item1;
                        int cpu_thread = DetermineTask(4, num + 3).Item2;
                        Thread[] thread = new Thread[cpu_thread];
                        for (i = 0; i <= cpu_thread - 1; i++)
                        {
                            int temp = i;
                            if (col_mode == true)
                            {
                                thread[temp] = new Thread(() => DataColumnTask(import_data, arr_force, arr_assign, arr_section, range[temp, 0], range[temp, 1]));
                            }
                            else if (wall_mode == true)
                            {
                                thread[temp] = new Thread(() => DataWallTask(import_data, arr_force, arr_assign, arr_section, range[temp, 0], range[temp, 1]));
                            }
                            thread[temp].Start();
                        }
                        for (i = 0; i <= cpu_thread - 1; i++)
                        {
                            int temp = i;
                            thread[temp].Join();
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Sai định dạng dữ liệu !!", "Thông báo", MessageBoxButtons.OK);
                        return import_data;
                    }
                }
                else
                {
                    MessageBox.Show("Không có dữ liệu được chọn !!", "Thông báo", MessageBoxButtons.OK);
                    return import_data;
                }
            }
            return import_data;
        }





        //CÁC HÀM ĐIỀU CHỈNH THÔNG SỐ ĐẦU VÀO
        public List<string[]> GetDataModify(object[,] import, int name) // Hàm lấy dữ liệu điều chỉnh thông số 
        {
            bool new_stage = false;
            string add_label;
            Dictionary<string, string[]> dic_modify = new Dictionary<string, string[]>();
            List<string[]> list_modify = new List<string[]>();
            int num = import.GetUpperBound(0);
            for (int i = 1; i < num; i++)
            {
                string[] store = new string[5];
                if (Convert.ToString(import[i, 1]) != Convert.ToString(import[i + 1, 1]) || (i == num - 1))
                {
                    new_stage = true;
                }
                add_label = Convert.ToString(import[i, name]);
                store[0] = Convert.ToString(import[i, 1]);
                store[1] = Convert.ToString(import[i, 5]);
                store[2] = Convert.ToString(import[i, 8]);
                store[3] = Convert.ToString(import[i, 9]);
                store[4] = Convert.ToString(import[i, 10]);
                if (dic_modify.ContainsKey(add_label) == false)
                {
                    dic_modify.Add(add_label, store);
                }
                if (new_stage == true)
                {
                    foreach (KeyValuePair<string, string[]> entry in dic_modify)
                    {
                        string[] add_value = new string[6];
                        add_value[0] = Convert.ToString(entry.Value[0]);
                        add_value[1] = entry.Key;
                        add_value[2] = Convert.ToString(entry.Value[1]);
                        add_value[3] = Convert.ToString(entry.Value[2]);
                        add_value[4] = Convert.ToString(entry.Value[3]);
                        add_value[5] = Convert.ToString(entry.Value[4]);
                        list_modify.Add(add_value);
                    }
                    new_stage = false;
                    dic_modify.Clear();
                }
            }
            return list_modify;
        }

        public void ModifyTask(object[,] import, List<string[]> modify_output, byte pos, int start, int end) // Hàm chia khối lượng điều chỉnh
        {
            double ex; double ey;
            int P; int Mx; int My; double delta_Mx; double delta_My;
            for (int i = start; i <= end; i++)
            {
                P = Convert.ToInt32(import[i, 11]);
                Mx = Convert.ToInt32(import[i, 14]);
                My = Convert.ToInt32(import[i, 15]);
                ex = 0;
                ey = 0;
                for (int j = 0; j < modify_output.Count; j++)
                {
                    if ((Convert.ToString(import[i, 1]) == modify_output[j][0]) && (Convert.ToString(import[i, pos]) == modify_output[j][1]))
                    {
                        if (string.IsNullOrEmpty(modify_output[j][6]) != true) ex = Convert.ToDouble(modify_output[j][6]);
                        if (string.IsNullOrEmpty(modify_output[j][7]) != true) ey = Convert.ToDouble(modify_output[j][7]);
                        delta_Mx = Math.Abs(P) * ex;
                        delta_My = Math.Abs(P) * ey;
                        if (Mx <= 0) Mx = Convert.ToInt32(Mx - delta_Mx);
                        else Mx = Convert.ToInt32(Mx + delta_Mx);
                        if (My <= 0) My = Convert.ToInt32(My - delta_My);
                        else My = Convert.ToInt32(My + delta_My);
                        import[i, 8] = modify_output[j][3];
                        import[i, 9] = modify_output[j][4];
                        import[i, 10] = modify_output[j][5];
                        import[i, 14] = Mx;
                        import[i, 15] = My;
                        break;
                    }
                }
            }
        }

        public object[,] ModifyMultiThreading(object[,] import) // Hàm điều chỉnh lại thông số bằng multi threading
        {
            List<string[]> modify_label = GetDataModify(import, 2);
            List<string[]> modify_cad = GetDataModify(import, 4);
            Modify_Form myform = new Modify_Form(modify_label, modify_cad);
            myform.ShowDialog();
            string mode = myform.SendBack().Item1;
            List<string[]> modify_output = myform.SendBack().Item2;
            byte pos;
            if (mode == "Tên ETABS") pos = 2;
            else pos = 4;
            int num = import.GetUpperBound(0); int i;
            // Chia luồng để load dữ liệu vào
            int[,] range = DetermineTask(1, num).Item1;
            int cpu_thread = DetermineTask(1, num).Item2;
            Thread[] thread = new Thread[cpu_thread];
            for (i = 0; i <= cpu_thread - 1; i++)
            {
                int temp = i;
                thread[temp] = new Thread(() => ModifyTask(import, modify_output, pos, range[temp, 0], range[temp, 1]));
                thread[temp].Start();
            }
            for (i = 0; i <= cpu_thread - 1; i++)
            {
                int temp = i;
                thread[temp].Join();
            }
            return import;

        }





        //CÁC HÀM GÁN TÊN CHO CỘT, VÁCH
        public object[,] AssignCADName(object[,] import)
        {
            string add_label; string add_cad;
            int num = import.GetUpperBound(0);
            Dictionary<string, string> dic_label = new Dictionary<string, string>();
            for (int i = 0; i < num; i++)
            {
                add_label = Convert.ToString(import[i + 1, 1]);
                add_cad = Convert.ToString(import[i + 1, 3]);
                if (dic_label.ContainsKey(add_label) == false)
                {
                    dic_label.Add(add_label, add_cad);
                }
            }
            AssignCAD_Form myform = new AssignCAD_Form(dic_label);
            myform.ShowDialog();
            Dictionary<string, string> receive = myform.SendBack();
            object[,] assign_name = new object[num, 3];
            for (int i = 0; i < num; i++)
            {
                string label = Convert.ToString(import[i + 1, 1]);
                assign_name[i, 0] = label;
                assign_name[i, 1] = Convert.ToString(import[i + 1, 2]);
                assign_name[i, 2] = receive[label];
            }
            return assign_name;
        }





        //CÁC HÀM TÍNH TOÁN HÀM LƯỢNG THÉP
        public (int, int) RectangleUpperMoment(int P, int Mx, int My, double Cx,
            double Cy, double L, double Eb, double k) //Hàm tính moment gia tăng tiết diện chữ nhật
        {
            int Mx_up; int My_up;
            Eb = Eb * 1000;

            //Tính toán moment gia tăng
            if (P <= 0)
            {
                Mx_up = Math.Max(Math.Abs(Mx), 1);
                My_up = Math.Max(Math.Abs(My), 1);
                return (Mx_up, My_up);
            }
            else
            {
                //Độ lệch tâm
                double e1_x = (double)Mx / P;
                double e1_y = (double)My / P;
                double ea1_x = Math.Max(Math.Max(L / 600, Cx / 30), 0.01);
                double ea1_y = Math.Max(Math.Max(L / 600, Cy / 30), 0.01);
                double e0_x = Math.Max(e1_x, ea1_x);
                double e0_y = Math.Max(e1_y, ea1_y);
                double Ib_x = Cy * Math.Pow(Cx, 3) / 12;
                double Ib_y = Cx * Math.Pow(Cy, 3) / 12;
                double delta_ex = Math.Min(Math.Max(e0_x / Cx, 0.15), 1.5);
                double delta_ey = Math.Min(Math.Max(e0_y / Cy, 0.15), 1.5);
                double kb_x = 0.15 / (1.5 * (0.3 + delta_ex));
                double kb_y = 0.15 / (1.5 * (0.3 + delta_ey));
                double Ncr_x = Math.Pow(Math.PI, 2) * kb_x * Eb * Ib_x / Math.Pow(L * k, 2);
                double Ncr_y = Math.Pow(Math.PI, 2) * kb_y * Eb * Ib_y / Math.Pow(L * k, 2);
                if (Ncr_x < P) Mx_up = ushort.MaxValue;
                else
                {
                    double eta_x = 1 / (1 - P / Ncr_x);
                    Mx_up = Math.Max((int)(P * e0_x * eta_x), 1);
                }
                if (Ncr_y < P) My_up = ushort.MaxValue;
                else
                {
                    double eta_y = 1 / (1 - P / Ncr_y);
                    My_up = Math.Max((int)(P * e0_y * eta_y), 1);
                }
                return (Mx_up, My_up);
            }
        }

        public (int, int) CircleUpperMoment(int P, int Mx, int My, double D,
            double L, double Eb, double k) //Hàm tính moment gia tăng tiết diện tròn
        {
            int Mx_up; int My_up;
            Eb = Eb * 1000;

            //Tính toán moment gia tăng
            if (P <= 0)
            {
                Mx_up = Math.Max(Math.Abs(Mx), 1);
                My_up = Math.Max(Math.Abs(My), 1);
                return (Mx_up, My_up);
            }
            else
            {
                //Độ lệch tâm
                double e1_x = (double)Mx / P;
                double e1_y = (double)My / P;
                double ea1 = Math.Max(Math.Max(L / 600, D / 30), 0.01);
                double e0_x = Math.Max(e1_x, ea1);
                double e0_y = Math.Max(e1_y, ea1);
                double Ib = Math.PI / 4 * Math.Pow(D / 2, 4);
                double delta_ex = Math.Min(Math.Max(e0_x / D, 0.15), 1.5);
                double delta_ey = Math.Min(Math.Max(e0_y / D, 0.15), 1.5);
                double kb_x = 0.15 / (1.5 * (0.3 + delta_ex));
                double kb_y = 0.15 / (1.5 * (0.3 + delta_ey));
                double Ncr_x = Math.Pow(Math.PI, 2) * kb_x * Eb * Ib / Math.Pow(L * k, 2);
                double Ncr_y = Math.Pow(Math.PI, 2) * kb_y * Eb * Ib / Math.Pow(L * k, 2);
                if (Ncr_x < P) Mx_up = ushort.MaxValue;
                else
                {
                    double eta_x = 1 / (1 - P / Ncr_x);
                    Mx_up = Math.Max((int)(P * e0_x * eta_x), 1);
                }
                if (Ncr_y < P) My_up = ushort.MaxValue;
                else
                {
                    double eta_y = 1 / (1 - P / Ncr_y);
                    My_up = Math.Max((int)(P * e0_y * eta_y), 1);
                }
                return (Mx_up, My_up);
            }
        }

        public object SimplifiedRectangleID(int P, int Mx_up, int My_up, double Cx, double Cy, double abv,
            double Rb, double Rs, double Rsc, double Es, double umin, double umax) //Hàm tính hàm lượng thép tiết diện chữ nhật 
        {
            object ucal = 0;

            //Đổi đơn vị
            abv /= 1000;
            Rb *= 1000;
            Rs *= 1000;
            Rsc *= 1000;
            Es *= 1000;
            umin /= 100;
            umax /= 100;

            double ciR = 0.8 / (1 + Rs / Es / 0.0035);
            double ks = Rsc / Rs;
            double delta_x = abv / (Cx - abv);
            double delta_y = abv / (Cy - abv);
            double u; double alpha; double i; double n; double m_x; double m_y;
            double ci; int count;

            //Trường hợp cột kéo thì trả về ucal = 0
            if (P < 0) return ucal;

            //Lặp hàm lượng từ giá trị min tới giá trị max
            for (i = umin; i <= umax; i += 0.001)
            {
                u = i / 2;
                alpha = Rs * u / Rb;
                count = 0;
                double[,] ver_value = new double[12, 4];
                for (ci = 0; ci <= 1; ci += 0.1)
                {
                    if (ci <= ciR)
                    {
                        n = ci - alpha + ks * alpha;
                    }
                    else
                    {
                        n = ci + alpha * (2 * ci - 1 - ciR + ks - ks * ciR) / (1 - ciR);
                    }
                    m_x = ci * (1 - 0.5 * ci) + (1 - delta_x) * (ks * alpha - 0.5 * n);
                    m_y = ci * (1 - 0.5 * ci) + (1 - delta_y) * (ks * alpha - 0.5 * n);
                    ver_value[count, 0] = ci;
                    ver_value[count, 1] = n;
                    ver_value[count, 2] = m_x;
                    ver_value[count, 3] = m_y;
                    count = count + 1;
                    if (Math.Round(ci, 2) == 1)
                    {
                        ver_value[count, 0] = ci;
                        ver_value[count, 1] = n;
                        ver_value[count, 2] = 0;
                        ver_value[count, 3] = 0;
                        break;
                    }
                }
                double ns_x; double ns_y; double ms_x; double ms_y; double k_i;
                double alpha_nx = 0; double alpha_ny = 0; double alpha_mx = 0; double alpha_my = 0;
                int j;
                ns_x = P / (Rb * Cy * (Cx - abv));
                ms_x = Mx_up / (Rb * Cy * Math.Pow(Cx - abv, 2));
                ns_y = P / (Rb * Cx * (Cy - abv));
                ms_y = My_up / (Rb * Cx * Math.Pow(Cy - abv, 2));
                for (j = 0; j <= 10; j++)
                {
                    k_i = Math.Round((ms_x * ver_value[j, 1] - ns_x * ver_value[j, 2]) / (ns_x * (ver_value[j + 1, 2] - ver_value[j, 2]) - ms_x * (ver_value[j + 1, 1] - ver_value[j, 1])), 3);
                    if ((0 <= k_i) && (k_i <= 1))
                    {
                        alpha_nx = ver_value[j, 1] + k_i * (ver_value[j + 1, 1] - ver_value[j, 1]);
                        alpha_mx = ver_value[j, 2] + k_i * (ver_value[j + 1, 2] - ver_value[j, 2]);
                        break;
                    }
                }
                for (j = 0; j <= 10; j++)
                {
                    k_i = Math.Round((ms_y * ver_value[j, 1] - ns_y * ver_value[j, 3]) / (ns_y * (ver_value[j + 1, 3] - ver_value[j, 3]) - ms_y * (ver_value[j + 1, 1] - ver_value[j, 1])), 3);
                    if ((0 <= k_i) && (k_i <= 1))
                    {
                        alpha_ny = ver_value[j, 1] + k_i * (ver_value[j + 1, 1] - ver_value[j, 1]);
                        alpha_my = ver_value[j, 3] + k_i * (ver_value[j + 1, 3] - ver_value[j, 3]);
                        break;
                    }
                }
                //Kiểm tra hàm lượng thép đang lặp đã đủ yêu cầu chưa
                double Pnx; double Mnx; double Pny; double Mny; double Pu; double Pnxy; double DC;
                double alpha_s; double k0; double k1; double k2;
                Pnx = alpha_nx * Rb * Cy * (Cx - abv);
                Mnx = alpha_mx * Rb * Cy * Math.Pow(Cx - abv, 2);
                Pny = alpha_ny * Rb * Cx * (Cy - abv);
                Mny = alpha_my * Rb * Cx * Math.Pow(Cy - abv, 2);
                Pu = (Rb * Cx * Cy + Rs * i * Cx * Cy);
                Pnxy = 1 / (1 / Pnx + 1 / Pny - 1 / Pu);
                if (Pnxy > 0.1 * Pu)
                {
                    DC = P / Pnxy;
                }
                else
                {
                    alpha_s = Rs * i / Rb;
                    k0 = (0.275 + alpha_s) / (0.16 + alpha_s);
                    if (alpha_nx <= 0.4)
                    {
                        k1 = (3.44 - 0.023 * alpha_s) * Math.Pow(0.4 - alpha_nx, 2) / (0.254 + alpha_s) + k0;
                    }
                    else
                    {
                        k1 = (Math.Pow(1.7 - alpha_s, 2) / 4 + 0.1775) * (Math.Pow(alpha_s, 2) - 0.16) + k0;
                    }
                    if (alpha_ny <= 0.4)
                    {
                        k2 = (3.44 - 0.023 * alpha_s) * Math.Pow(0.4 - alpha_ny, 2) / (0.254 + alpha_s) + k0;
                    }
                    else
                    {
                        k2 = (Math.Pow(1.7 - alpha_s, 2) / 4 + 0.1775) * (Math.Pow(alpha_s, 2) - 0.16) + k0;
                    }
                    k1 = Math.Min(k1, 1.6);
                    k2 = Math.Min(k2, 1.6);
                    DC = 1.1 * Math.Pow(Mx_up / Mnx, k1) + Math.Pow(My_up / Mny, k2);
                }
                if (DC < 1)
                {
                    ucal = Convert.ToString(Math.Round(i * 100, 2));
                    break;
                }
            }
            //Nếu hàm lượng lặp qua mức umax thì báo Over Stress (O.S)
            if (i > umax) ucal = "O.S";
            return ucal;
        }

        public object SimplifiedCircleID(int P, int Mx_up, int My_up, double D, double abv,
            double Rb, double Rs, double umin, double umax) //Hàm tính hàm lượng thép tiết diện tròn
        {
            object ucal = 0;

            //Đổi đơn vị
            abv /= 1000;
            Rb *= 1000;
            Rs *= 1000;
            umin /= 100;
            umax /= 100;

            double alpha; double i; double n; double m; double phi;
            double ci; int count;
            double r = D / 2;
            double rs = r - abv;
            double gamma = rs / r;

            //Trường hợp cột kéo thì trả về ucal = 0
            if (P < 0) return ucal;

            //Lặp hàm lượng từ giá trị min tới giá trị max
            for (i = umin; i <= umax; i += 0.001)
            {
                alpha = Rs * i / Rb;
                count = 0;
                double[,] ver_value = new double[11, 4];
                for (ci = 0; ci <= 1; ci += 0.1)
                {
                    n = ci + 2.55 * ci * alpha - alpha - Math.Sin(2 * Math.PI * ci) / (2 * Math.PI);
                    phi = Math.Min(1.6 * (1 - 1.55 * ci) * ci, 1);
                    if (n > 0.77 + 0.645 * alpha)
                    {
                        n = ci + ci * alpha - Math.Sin(2 * Math.PI * ci) / (2 * Math.PI);
                        phi = 0;
                    }
                    m = 2 * (Math.Pow(Math.Sin(Math.PI * ci), 3) / Math.PI) / 3 + alpha * (Math.Sin(Math.PI * ci) / Math.PI + phi) * gamma;
                    ver_value[count, 0] = ci;
                    ver_value[count, 1] = phi;
                    ver_value[count, 2] = n;
                    ver_value[count, 3] = m;
                    count += 1;
                }
                double ns; double ms; double k_i;
                double alpha_n = 0; double alpha_m = 0;
                int Mxy = Convert.ToInt32(Math.Sqrt(Math.Pow(Mx_up, 2) + Math.Pow(My_up, 2)));
                double Ar = Math.PI * Math.Pow(D, 2) / 4;
                ns = P / (Rb * Ar);
                ms = Mxy / (Rb * Ar * rs);
                for (int j = 0; j <= 9; j++)
                {
                    k_i = Math.Round((ms * ver_value[j, 2] - ns * ver_value[j, 3]) / (ns * (ver_value[j + 1, 3] - ver_value[j, 3]) - ms * (ver_value[j + 1, 2] - ver_value[j, 2])), 3);
                    if ((0 <= k_i) && (k_i <= 1))
                    {
                        alpha_n = ver_value[j, 2] + k_i * (ver_value[j + 1, 2] - ver_value[j, 2]);
                        alpha_m = ver_value[j, 3] + k_i * (ver_value[j + 1, 3] - ver_value[j, 3]);
                        break;
                    }
                }
                double Pu; double Mu; double DC;
                Pu = alpha_n * Rb * Ar;
                Mu = alpha_m * Rb * Ar * rs;
                DC = 1.1 * Math.Sqrt(Math.Pow(Mxy, 2) + Math.Pow(P, 2)) / Math.Sqrt(Math.Pow(Mu, 2) + Math.Pow(Pu, 2));
                if (DC < 1)
                {
                    ucal = Convert.ToString(Math.Round(i * 100, 2));
                    break;
                }
            }
            if (i > umax) ucal = "O.S";
            return ucal;
        }

        public void CalculateTask(object[,] import, object[,] rebar_percent, int abv, double Rb, double Eb,
            double Rs, double Rsc, double Es, double umin, double umax, int start, int end) //Hàm chia khối lượng tính toán cột, vách
        {
            int Mx_up = 0; int My_up = 0; object ucal = 0;
            int P; int Mx; int My;
            for (int i = start; i <= end; i++)
            {
                P = -1 * Convert.ToInt32(import[i, 11]);
                Mx = Math.Abs(Convert.ToInt32(import[i, 14]));
                My = Math.Abs(Convert.ToInt32(import[i, 15]));
                if ((Convert.ToString(import[i, 5]) == "Rec") || Convert.ToString(import[i, 5]) == "Wall")
                {
                    double L = Convert.ToDouble(import[i, 8]);
                    double Cx = Convert.ToDouble(import[i, 9]);
                    double Cy = Convert.ToDouble(import[i, 10]);
                    Mx_up = RectangleUpperMoment(P, Mx, My, Cx, Cy, L, Eb, 1).Item1;
                    My_up = RectangleUpperMoment(P, Mx, My, Cx, Cy, L, Eb, 1).Item2;
                    ucal = SimplifiedRectangleID(P, Mx_up, My_up, Cx, Cy, abv, Rb, Rs, Rsc, Es, umin, umax);
                }
                if (Convert.ToString(import[i, 5]) == "Cir")
                {
                    double L = Convert.ToDouble(import[i, 8]);
                    double D = Convert.ToDouble(import[i, 9]);
                    Mx_up = CircleUpperMoment(P, Mx, My, D, L, Eb, 1).Item1;
                    My_up = CircleUpperMoment(P, Mx, My, D, L, Eb, 1).Item2;
                    ucal = SimplifiedCircleID(P, Mx_up, My_up, D, abv, Rb, Rs, umin, umax);
                }
                rebar_percent[i - 1, 0] = Mx_up;
                rebar_percent[i - 1, 1] = My_up;
                rebar_percent[i - 1, 2] = ucal;
            }
        }

        public object[,] MultiThreadingCalculate(object[,] import, int abv, double Rb, double Eb,
            double Rs, double Rsc, double Es, double umin, double umax) //Hàm tính toán cột, vách bằng multi threading
        {
            int num = import.GetUpperBound(0); int i;
            object[,] rebar_percent = new object[num, 3];
            // Chia luồng để load dữ liệu vào
            int[,] range = DetermineTask(1, num).Item1;
            int cpu_thread = DetermineTask(1, num).Item2;
            Thread[] thread = new Thread[cpu_thread];
            for (i = 0; i <= cpu_thread - 1; i++)
            {
                int temp = i;
                thread[temp] = new Thread(() => CalculateTask(import, rebar_percent, abv, Rb, Eb, Rs, Rsc, Es, umin, umax, range[temp, 0], range[temp, 1]));
                thread[temp].Start();
            }
            for (i = 0; i <= cpu_thread - 1; i++)
            {
                int temp = i;
                thread[temp].Join();
            }
            return rebar_percent;
        }





        //CÁC HÀM THIẾT KẾ VÀ KIỂM TRA CỐT THÉP
        public string[,] GetDataDesign(object[,] data, int name) //Hàm lấy các dữ liệu cần thiết để pass form thiết kế
        {
            int num = data.GetUpperBound(0); int i; int count;
            string[,] design_name = new string[num, 13];
            string add_label; string[] current_value;
            bool new_stage = false;
            Dictionary<string, string[]> dic_name = new Dictionary<string, string[]>();
            count = 0;
            for (i = 1; i < num; i++)
            {
                string[] store = new string[11];
                if (Convert.ToString(data[i, 1]) != Convert.ToString(data[i + 1, 1]) || (i == num))
                {
                    new_stage = true;
                }
                add_label = Convert.ToString(data[i, name]);
                store[0] = Convert.ToString(data[i, 5]);
                store[1] = Convert.ToString(data[i, 9]);
                store[2] = Convert.ToString(data[i, 10]);
                store[3] = Convert.ToString(data[i, 18]);
                store[4] = Convert.ToString(data[i, 19]);
                store[5] = Convert.ToString(data[i, 20]);
                store[6] = Convert.ToString(data[i, 21]);
                store[7] = Convert.ToString(data[i, 22]);
                store[8] = Convert.ToString(data[i, 23]);
                store[9] = Convert.ToString(data[i, 24]);
                store[10] = Convert.ToString(data[i, 25]);
                if (dic_name.ContainsKey(add_label) == false)
                {
                    dic_name.Add(add_label, store);
                }
                else
                {
                    dic_name.TryGetValue(add_label, out current_value);
                    if (store[3] == "O.S")
                    {
                        dic_name[add_label] = store;
                    }
                    else
                    {
                        if (current_value[3] != "O.S")
                        {
                            if (Convert.ToDouble(current_value[3]) < Convert.ToDouble(store[3]))
                            {
                                dic_name[add_label] = store;
                            }
                        }
                    }

                }
                if (new_stage == true)
                {
                    foreach (KeyValuePair<string, string[]> entry in dic_name)
                    {
                        design_name[count, 0] = Convert.ToString(data[i, 1]);
                        design_name[count, 1] = entry.Key;
                        design_name[count, 2] = Convert.ToString(entry.Value[0]);
                        design_name[count, 3] = Convert.ToString(entry.Value[1]);
                        design_name[count, 4] = Convert.ToString(entry.Value[2]);
                        design_name[count, 5] = Convert.ToString(entry.Value[3]);
                        design_name[count, 6] = Convert.ToString(entry.Value[4]);
                        design_name[count, 7] = Convert.ToString(entry.Value[5]);
                        design_name[count, 8] = Convert.ToString(entry.Value[6]);
                        design_name[count, 9] = Convert.ToString(entry.Value[7]);
                        design_name[count, 10] = Convert.ToString(entry.Value[8]);
                        design_name[count, 11] = Convert.ToString(entry.Value[9]);
                        design_name[count, 12] = Convert.ToString(entry.Value[10]);
                        count = count + 1;
                    }
                    new_stage = false;
                    dic_name.Clear();
                }
            }
            return design_name;
        }

        public (double[,], double[,]) IDSurfaceRectangle(double Cx, double Cy, double Eb, double Es, double abv, double Rb,
            double Rs, double Rsc, int nx, int ny, int dmain, int dstir, int nsec) //Hàm xây dựng mặt cong tương tác tiết diện chữ nhật
        {
            int di;
            if (Math.Min(Cx, Cy) <= 300) di = 5;
            else di = 10;
            int du = di * di;
            int sump = Convert.ToInt32(Cx * Cy / du);
            double[,] concrete_element = new double[sump, 2];
            int k; int i; int j; int u;
            k = 0;
            for (i = 0; i < Convert.ToInt32(Cy / di); i++)
            {
                for (j = 0; j < Convert.ToInt32(Cx / di); j++)
                {
                    concrete_element[k, 0] = Cx / 2 - di / 2 - di * j;
                    concrete_element[k, 1] = Cy / 2 - di / 2 - di * i;
                    k = k + 1;
                }
            }
            double dre = Math.PI * Math.Pow(dmain, 2) / 4;
            double Cx_cen = Cx - 2 * abv - 2 * dstir - dmain;
            double Cy_cen = Cy - 2 * abv - 2 * dstir - dmain;
            int sum_rebar = 2 * (nx + ny) - 4;
            double[,] rebar_element = new double[sum_rebar, 2];
            j = 0;
            for (i = 0; i < nx; i++)
            {
                rebar_element[j, 0] = -Cx_cen / 2 + Cx_cen / (nx - 1) * i;
                rebar_element[j, 1] = Cy_cen / 2;
                rebar_element[j + 1, 0] = rebar_element[j, 0];
                rebar_element[j + 1, 1] = -Cy_cen / 2;
                j = j + 2;
            }
            for (i = 1; i < ny - 1; i++)
            {
                rebar_element[j, 0] = Cx_cen / 2;
                rebar_element[j, 1] = -Cy_cen / 2 + Cy_cen / (ny - 1) * i;
                rebar_element[j + 1, 0] = -Cx_cen / 2;
                rebar_element[j + 1, 1] = rebar_element[j, 1];
                j = j + 2;
            }
            double xM = Cx / 2 + 100;
            double yM = Cy / 2 + 100;
            double epc_b2 = 0.0035;
            int a_init = 0; int b_init = 1;
            int col; int row;
            double rotate_angle; double a; double b; double c_min; double c_max; double c;
            double na_depth; double k_j; double k_jj;
            double dc_i; double epc_i; double sigc_i; double Pc_i; double Mcx_i; double Mcy_i;
            double ds_i; double eps_i; double sigs_i; double Ps_i; double Msx_i; double Msy_i;
            double sum_Pc; double sum_Mcx; double sum_Mcy; double sum_Ps; double sum_Msx; double sum_Msy;
            double[,] ver_value = new double[nsec, 4 * nsec];
            //Xoay đường nén trong vùng I/4 mặt cắt tiết diện
            col = 0;
            for (u = 0; u < nsec; u++)
            {
                row = 0;
                rotate_angle = 0.5 * Math.PI / (nsec - 1) * u;
                a = a_init * Math.Cos(rotate_angle) + b_init * Math.Sin(rotate_angle);
                b = b_init * Math.Cos(rotate_angle) - a_init * Math.Sin(rotate_angle);
                c_min = -a * Cx / 2 - b * Cy / 2 + a_init * Cx / 2 + b_init * Cy / 2 - Cy / 2;
                c_max = -c_min;
                for (i = 0; i < nsec; i++)
                {
                    sum_Pc = 0;
                    sum_Mcx = 0;
                    sum_Mcy = 0;
                    sum_Ps = 0;
                    sum_Msx = 0;
                    sum_Msy = 0;
                    c = c_min + 2 * Math.Abs(c_min) / (nsec - 1) * i;
                    na_depth = Math.Abs(a * Cx / 2 + b * Cy / 2 + c) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2));
                    if ((Math.Round(na_depth, 2) == 0) || (Math.Round(c - c_max, 2) == 0))
                    {
                        na_depth = 1 / int.MaxValue;
                    }
                    for (j = 0; j < sump; j++) //Tính toán cho bê tông
                    {
                        k_j = a * xM + b * yM + c;
                        k_jj = a * concrete_element[j, 0] + b * concrete_element[j, 1] + c;
                        if (k_j * k_jj >= 0)
                        {
                            dc_i = Math.Abs(a * concrete_element[j, 0] + b * concrete_element[j, 1] + c) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2));
                            epc_i = dc_i * epc_b2 / na_depth;
                            sigc_i = Concrete(epc_i, Rb, Eb);
                            Pc_i = sigc_i * du / 1000;
                            Mcx_i = Pc_i * concrete_element[j, 0] / 1000;
                            Mcy_i = Pc_i * concrete_element[j, 1] / 1000;
                        }
                        else
                        {
                            Pc_i = 0;
                            Mcx_i = 0;
                            Mcy_i = 0;
                        }
                        sum_Pc = sum_Pc + Pc_i;
                        sum_Mcx = sum_Mcx + Mcx_i;
                        sum_Mcy = sum_Mcy + Mcy_i;
                    }
                    for (j = 0; j < sum_rebar; j++) //Tính toán cho cốt thép
                    {
                        ds_i = Math.Abs(a * rebar_element[j, 0] + b * rebar_element[j, 1] + c) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2));
                        eps_i = ds_i * epc_b2 / na_depth;
                        k_j = a * xM + b * yM + c;
                        k_jj = a * rebar_element[j, 0] + b * rebar_element[j, 1] + c;
                        if (k_j * k_jj <= 0)
                        {
                            eps_i = -eps_i;
                        }
                        sigs_i = Rebar(eps_i, Rs, Rsc, Es);
                        Ps_i = sigs_i * dre / 1000;
                        Msx_i = Ps_i * rebar_element[j, 0] / 1000;
                        Msy_i = Ps_i * rebar_element[j, 1] / 1000;
                        sum_Ps = sum_Ps + Ps_i;
                        sum_Msx = sum_Msx + Msx_i;
                        sum_Msy = sum_Msy + Msy_i;
                    }
                    ver_value[row, col] = Math.Round(sum_Mcx + sum_Msx, 2);
                    ver_value[row, col + 1] = Math.Round(sum_Mcy + sum_Msy, 2);
                    ver_value[row, col + 2] = Math.Round(sum_Pc + sum_Ps, 2);
                    ver_value[row, col + 3] = Math.Round(rotate_angle * (180 / Math.PI), 2);
                    row = row + 1;
                }
                col = col + 4;
            }
            //Tính toán Pz(i)
            double Pz_min = ver_value[0, 2];
            double Pz_max = ver_value[nsec - 1, 2];
            double[] Pz_delta = new double[nsec];
            double k_i;
            for (i = 0; i < nsec; i++)
            {
                Pz_delta[i] = Pz_min + (Pz_max - Pz_min) / (nsec - 1) * i;
            }
            //Tìm giao điểm Pz(i) với các đường cong tương tác
            double[,] hoz_value = new double[nsec, 4 * nsec];
            col = 0;
            for (i = 0; i < nsec; i++)
            {
                row = 0;
                for (k = 0; k < 4 * nsec; k += 4)
                {
                    for (j = 0; j < nsec - 1; j++)
                    {
                        k_i = Math.Round((Pz_delta[i] - ver_value[j, k + 2]) / (ver_value[j + 1, k + 2] - ver_value[j, k + 2]), 2);
                        if ((0 <= k_i) && (k_i <= 1))
                        {
                            hoz_value[row, col] = Math.Round(k_i * (ver_value[j + 1, k] - ver_value[j, k]) + ver_value[j, k], 2);
                            hoz_value[row, col + 1] = Math.Round(k_i * (ver_value[j + 1, k + 1] - ver_value[j, k + 1]) + ver_value[j, k + 1], 2);
                            hoz_value[row, col + 2] = Math.Round(Pz_delta[i], 2);
                            hoz_value[row, col + 3] = k_i;
                            row = row + 1;
                            break;
                        }
                    }
                }
                col = col + 4;
            }
            return (ver_value, hoz_value);
        }

        public (double[,], double[,]) IDSurfaceCircle(double D, double Eb, double Es, double abv, double Rb, double Rs,
            double Rsc, int nsum, int dmain, int dstir, int nsec) //Hàm xây dựng mặt cong tương tác tiết diện tròn
        {
            int di;
            if (D <= 300) di = 5;
            else di = 10;
            int du = di * di;
            List<double[]> concrete_element = new List<double[]>();
            int k; int i; int j; int u;
            for (i = 0; i < Convert.ToInt32(D / di); i++)
            {
                for (j = 0; j < Convert.ToInt32(D / di); j++)
                {
                    double[] point = new double[2];
                    point[0] = D / 2 - di / 2 - di * j;
                    point[1] = D / 2 - di / 2 - di * i;
                    if (Math.Sqrt(Math.Pow(point[0], 2) + Math.Pow(point[1], 2)) <= D / 2) concrete_element.Add(point);
                }
            }
            double dre = Math.PI * Math.Pow(dmain, 2) / 4;
            double D_cen = D - 2 * abv - 2 * dstir - dmain;
            double[,] rebar_element = new double[nsum, 2];
            for (i = 0; i < nsum; i++)
            {
                rebar_element[i, 0] = Math.Sin(2 * Math.PI / nsum * i) * (D_cen / 2);
                rebar_element[i, 1] = Math.Cos(2 * Math.PI / nsum * i) * (D_cen / 2);
            }
            double xM = D / 2 + 100;
            double yM = D / 2 + 100;
            double xH; double yH;
            double epc_b2 = 0.0035;
            int a_init = 0; int b_init = 1;
            int col; int row;
            double rotate_angle; double a; double b; double c_min; double c_max; double c;
            double na_depth; double k_j; double k_jj;
            double dc_i; double epc_i; double sigc_i; double Pc_i; double Mcx_i; double Mcy_i;
            double ds_i; double eps_i; double sigs_i; double Ps_i; double Msx_i; double Msy_i;
            double sum_Pc; double sum_Mcx; double sum_Mcy; double sum_Ps; double sum_Msx; double sum_Msy;
            double[,] ver_value = new double[nsec, 4 * nsec];
            //Xoay đường nén trong vùng I/4 mặt cắt tiết diện
            col = 0;
            for (u = 0; u < nsec; u++)
            {
                row = 0;
                rotate_angle = 0.5 * Math.PI / (nsec - 1) * u;
                a = a_init * Math.Cos(rotate_angle) + b_init * Math.Sin(rotate_angle);
                b = b_init * Math.Cos(rotate_angle) - a_init * Math.Sin(rotate_angle);
                yH = Math.Sqrt(Math.Pow(D / 2, 2) * Math.Pow(b, 2) / (Math.Pow(a, 2) + Math.Pow(b, 2)));
                xH = Math.Sqrt(Math.Pow(D / 2, 2) - Math.Pow(yH, 2));
                c_min = -a * xH - b * yH + a_init * xH + b_init * yH - yH;
                c_max = -c_min;
                for (i = 0; i < nsec; i++)
                {
                    sum_Pc = 0;
                    sum_Mcx = 0;
                    sum_Mcy = 0;
                    sum_Ps = 0;
                    sum_Msx = 0;
                    sum_Msy = 0;
                    c = c_min + 2 * Math.Abs(c_min) / (nsec - 1) * i;
                    na_depth = Math.Abs(a * xH + b * yH + c) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2));
                    if ((Math.Round(na_depth, 2) == 0) || (Math.Round(c - c_max, 2) == 0))
                    {
                        na_depth = 1 / int.MaxValue;
                    }
                    for (j = 0; j < concrete_element.Count; j++) //Tính toán cho bê tông
                    {
                        k_j = a * xM + b * yM + c;
                        k_jj = a * concrete_element[j][0] + b * concrete_element[j][1] + c;
                        if (k_j * k_jj >= 0)
                        {
                            dc_i = Math.Abs(a * concrete_element[j][0] + b * concrete_element[j][1] + c) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2));
                            epc_i = dc_i * epc_b2 / na_depth;
                            sigc_i = Concrete(epc_i, Rb, Eb);
                            Pc_i = sigc_i * du / 1000;
                            Mcx_i = Pc_i * concrete_element[j][0] / 1000;
                            Mcy_i = Pc_i * concrete_element[j][1] / 1000;
                        }
                        else
                        {
                            Pc_i = 0;
                            Mcx_i = 0;
                            Mcy_i = 0;
                        }
                        sum_Pc = sum_Pc + Pc_i;
                        sum_Mcx = sum_Mcx + Mcx_i;
                        sum_Mcy = sum_Mcy + Mcy_i;
                    }
                    for (j = 0; j < nsum; j++) //Tính toán cho cốt thép
                    {
                        ds_i = Math.Abs(a * rebar_element[j, 0] + b * rebar_element[j, 1] + c) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2));
                        eps_i = ds_i * epc_b2 / na_depth;
                        k_j = a * xM + b * yM + c;
                        k_jj = a * rebar_element[j, 0] + b * rebar_element[j, 1] + c;
                        if (k_j * k_jj <= 0)
                        {
                            eps_i = -eps_i;
                        }
                        sigs_i = Rebar(eps_i, Rs, Rsc, Es);
                        Ps_i = sigs_i * dre / 1000;
                        Msx_i = Ps_i * rebar_element[j, 0] / 1000;
                        Msy_i = Ps_i * rebar_element[j, 1] / 1000;
                        sum_Ps = sum_Ps + Ps_i;
                        sum_Msx = sum_Msx + Msx_i;
                        sum_Msy = sum_Msy + Msy_i;
                    }
                    ver_value[row, col] = Math.Round(sum_Mcx + sum_Msx, 2);
                    ver_value[row, col + 1] = Math.Round(sum_Mcy + sum_Msy, 2);
                    ver_value[row, col + 2] = Math.Round(sum_Pc + sum_Ps, 2);
                    ver_value[row, col + 3] = Math.Round(rotate_angle * (180 / Math.PI), 2);
                    row = row + 1;
                }
                col = col + 4;
            }

            //Tính toán Pz(i)
            double Pz_min = ver_value[0, 2];
            double Pz_max = ver_value[nsec - 1, 2];
            double[] Pz_delta = new double[nsec];
            double k_i;
            for (i = 0; i < nsec; i++)
            {
                Pz_delta[i] = Pz_min + (Pz_max - Pz_min) / (nsec - 1) * i;
            }
            //Tìm giao điểm Pz(i) với các đường cong tương tác
            double[,] hoz_value = new double[nsec, 4 * nsec];
            col = 0;
            for (i = 0; i < nsec; i++)
            {
                row = 0;
                for (k = 0; k < 4 * nsec; k += 4)
                {
                    for (j = 0; j < nsec - 1; j++)
                    {
                        k_i = Math.Round((Pz_delta[i] - ver_value[j, k + 2]) / (ver_value[j + 1, k + 2] - ver_value[j, k + 2]), 2);
                        if ((0 <= k_i) && (k_i <= 1))
                        {
                            hoz_value[row, col] = Math.Round(k_i * (ver_value[j + 1, k] - ver_value[j, k]) + ver_value[j, k], 2);
                            hoz_value[row, col + 1] = Math.Round(k_i * (ver_value[j + 1, k + 1] - ver_value[j, k + 1]) + ver_value[j, k + 1], 2);
                            hoz_value[row, col + 2] = Math.Round(Pz_delta[i], 2);
                            hoz_value[row, col + 3] = k_i;
                            row = row + 1;
                            break;
                        }
                    }
                }
                col = col + 4;
            }
            return (ver_value, hoz_value);
        }

        public (int, int, double, double[,], double [,]) InteractionDiagramCheck(int P, int Mx_up, int My_up, int nsec,
            double[,] ver_value, double[,] hoz_value) //Hàm kiểm tra KNCL bằng BĐTT
        {
            int k; int i; int j; double out_Mx; double out_My;
            double deno; double k_i;
            //Kết quả P-Mxy
            double[,] output_PMxy = new double[nsec, 2];
            k = 0;
            for (i = 0; i < 4 * nsec; i += 4)
            {
                for (j = 0; j < nsec - 1; j++)
                {
                    deno = My_up * (hoz_value[j + 1, i] - hoz_value[j, i]) - Mx_up * (hoz_value[j + 1, i + 1] - hoz_value[j, i + 1]);
                    if (deno == 0)
                    {
                        k_i = 0;
                        goto nextstep;
                    }
                    k_i = Math.Round((Mx_up * hoz_value[j, i + 1] - My_up * hoz_value[j, i]) / deno, 2);
                nextstep:
                    if ((0 <= k_i) && (k_i <= 1))
                    {
                        output_PMxy[k, 0] = hoz_value[j, i + 2];
                        out_Mx = hoz_value[j, i] + (hoz_value[j + 1, i] - hoz_value[j, i]) * k_i;
                        out_My = hoz_value[j, i + 1] + (hoz_value[j + 1, i + 1] - hoz_value[j, i + 1]) * k_i;
                        output_PMxy[k, 1] = Math.Sqrt(Math.Pow(out_Mx, 2) + Math.Pow(out_My, 2));
                        k = k + 1;
                        break;
                    }
                }
            }
            //Kết quả Mx-My
            double[,] output_MxMy = new double[nsec, 2];
            k = 0;
            for (i = 0; i < 4 * nsec; i += 4)
            {
                for (j = 0; j < nsec - 1; j++)
                {
                    k_i = Math.Round((P - ver_value[j, i + 2]) / (ver_value[j + 1, i + 2] - ver_value[j, i + 2]), 5);
                    if ((0 <= k_i) && (k_i <= 1))
                    {
                        output_MxMy[k, 0] = Math.Round(k_i * (ver_value[j + 1, i] - ver_value[j, i]) + ver_value[j, i], 2);
                        output_MxMy[k, 1] = Math.Round(k_i * (ver_value[j + 1, i + 1] - ver_value[j, i + 1]) + ver_value[j, i + 1], 2);
                        k = k + 1;
                        break;
                    }
                }
            }
            //Kết quả khả năng chịu lực
            double Mxy = Math.Sqrt(Math.Pow(Mx_up, 2) + Math.Pow(My_up, 2));
            int Mnxy = 0; int Pnxy = 0; double RC; double RD; double DC = 0;
            for (i = 0; i < nsec - 1; i++)
            {
                k_i = (Mxy * output_PMxy[i, 0] - P * output_PMxy[i, 1]) / (P * (output_PMxy[i + 1, 1] - output_PMxy[i, 1]) - Mxy * (output_PMxy[i + 1, 0] - output_PMxy[i, 0]));
                if ((0 <= k_i) && (k_i <= 1))
                {
                    Mnxy = Convert.ToInt32(output_PMxy[i, 1] + k_i * (output_PMxy[i + 1, 1] - output_PMxy[i, 1]));
                    Pnxy = Convert.ToInt32(output_PMxy[i, 0] + k_i * (output_PMxy[i + 1, 0] - output_PMxy[i, 0]));
                    RC = Math.Sqrt(Math.Pow(Mnxy, 2) + Math.Pow(Pnxy, 2));
                    RD = Math.Sqrt(Math.Pow(Mxy, 2) + Math.Pow(P, 2));
                    DC = Math.Round(RD / RC, 2);
                    break;
                }
            }
            return (Pnxy, Mnxy, DC, output_PMxy, output_MxMy);
        }

        public (int, int, object) StirrupRectangleCheck(double Cx, double Cy, int P, int Qx, int Qy, int dmain, int dstir, int nstir, int sw,
            double abv, double Rb, double Rbt, double Rsw) //Hàm kiểm tra chịu lực cắt
        {
            double att; object DC; double sigma; double phi_n;
            int Qnx; int Qny;

            //Đơn vị truyền vào
            //Dấu P ngược hướng
            //Cx, Cy - (m)
            //Qx, Qy - (kN)
            //dmain, dstir, nstir, sw, abv - (mm)
            //Rb, Rbt, Rsw - (MPa)

            Rb = Rb * 1000;
            Rbt = Rbt * 1000;
            Rsw = Rsw * 1000;

            att = (abv + dstir + dmain / 2) / 1000;
            if ((0.3 * Rb * Cy * (Cx - att) < Qx) || (0.3 * Rb * Cx * (Cy - att) < Qy))
            {
                goto endstep;
            }
            sigma = Math.Abs(P) / (Cx * Cy);
            if (P > 0)
            {
                if (sigma <= 0.25 * Rb)
                {
                    phi_n = 1 + sigma / Rb;
                }
                else if ((0.25 * Rb < sigma) && (sigma <= 0.75 * Rb))
                {
                    phi_n = 1.25;
                }
                else if ((0.75 * Rb < sigma) && (sigma <= Rb))
                {
                    phi_n = 5 * (1 - sigma / Rb);
                }
                else
                {
                    goto endstep;
                }
            }
            else
            {
                if (sigma <= Rbt)
                {
                    phi_n = 1 - sigma / (2 * Rbt);
                }
                else
                {
                    goto endstep;
                }
            }
            goto nextstep;
        endstep:
            {
                DC = "O.S";
                Qnx = 0;
                Qny = 0;
                return (Qnx, Qny, DC);
            }
        nextstep:
            {
                double Qsw_x; double Qsw_y; double c_x; double c_y;
                Qsw_x = 0.001 * Rsw * nstir * Math.PI * Math.Pow(dstir, 2) / 4 / sw;
                Qsw_y = 0.001 * Rsw * nstir * Math.PI * Math.Pow(dstir, 2) / 4 / sw;
                c_x = Math.Sqrt(phi_n * 1.5 * Rbt * Cy * Math.Pow(Cx - att, 2) / (0.75 * Qsw_x));
                c_y = Math.Sqrt(phi_n * 1.5 * Rbt * Cx * Math.Pow(Cy - att, 2) / (0.75 * Qsw_y));
                if (c_x < Cx - att)
                {
                    c_x = Cx - att;
                }
                if (c_x > 2 * (Cx - att))
                {
                    c_x = 2 * (Cx - att);
                }
                if (c_y < Cy - att)
                {
                    c_y = Cy - att;
                }
                if (c_y > 2 * (Cy - att))
                {
                    c_y = 2 * (Cy - att);
                }
                double Qb_x; double Qb_y;
                Qb_x = phi_n * 1.5 * Rbt * Cy * Math.Pow(Cx - att, 2) / c_x;
                Qb_y = phi_n * 1.5 * Rbt * Cx * Math.Pow(Cy - att, 2) / c_y;
                if (Qb_x < 0.5 * Rbt * Cy * (Cx - att))
                {
                    Qb_x = 0.5 * Rbt * Cy * (Cx - att);
                }
                if (Qb_x > 2.5 * Rbt * Cy * (Cx - att))
                {
                    Qb_x = 2.5 * Rbt * Cy * (Cx - att);
                }
                if (Qb_y < 0.5 * Rbt * Cx * (Cy - att))
                {
                    Qb_y = 0.5 * Rbt * Cx * (Cy - att);
                }
                if (Qb_y > 2.5 * Rbt * Cx * (Cy - att))
                {
                    Qb_y = 2.5 * Rbt * Cx * (Cy - att);
                }
                double Qs_x; double Qs_y;
                Qs_x = 0.75 * Qsw_x * c_x;
                Qs_y = 0.75 * Qsw_y * c_y;
                Qnx = Convert.ToInt32(Qb_x + Qs_x);
                Qny = Convert.ToInt32(Qb_y + Qs_y);
                DC = Math.Round(Math.Max((double)Math.Abs(Qx) / Qnx, (double)Math.Abs(Qy) / Qny),2);
                return (Qnx, Qny, DC);
            }
        }

        public double AxialCompressionRectangleCheck(double Cx, double Cy, int P,
            string[] loadcomb, string loadcheck, double Rb) //Hàm tính lực dọc quy đổi
        {
            double ved = 0;
            if (P > 0)
            {
                double fcd;
                fcd = (1868.6 * Rb - 1465.7) * 0.8 / 1.2;
                for (int i = 0; i < loadcomb.GetUpperBound(0); i++)
                {
                    if (loadcheck.Contains(loadcomb[i]) == true && string.IsNullOrEmpty(loadcomb[i]) == false)
                    {
                        ved = Math.Round(P / (Cx * Cy * fcd), 2);
                        break;
                    }
                }
            }
            return ved;
        }

        public void DesignResultTask(object[,] design_check, double Rb, double Rbt, double Rsw, double abv, int nsec, object[,] data, List<string[]> receive,
             Dictionary<string, (double[,], double[,])> dic_idsurface, string[] loadcomb, int pos, int start, int end) //Hàm chia công việc kiểm tra
        {
            int P; int Mx_up; int My_up; int Qx; int Qy; double DC; object DCs;
            int i; int j;
            double equal_area; int dmain; int dstir; int nstir; int sw;
            string loadcheck; double ved;
            for (i = start; i <= end; i++)
            {
                bool skip = true;
                for (j = 0; j < receive.Count; j++)
                {
                    if ((Convert.ToString(data[i, 1]) == receive[j][0]) && (Convert.ToString(data[i, pos]) == receive[j][1]))
                    {
                        skip = false;
                        string[] add_value = new string[7];
                        add_value[0] = receive[j][2];
                        add_value[1] = receive[j][3];
                        add_value[2] = receive[j][4];
                        add_value[3] = receive[j][6];
                        add_value[4] = receive[j][7];
                        add_value[5] = Convert.ToString(DiameterExtract(receive[j][8]));
                        add_value[6] = Convert.ToString(StirrupExtract(receive[j][11]).Item1);
                        string combine_string = add_value[0] + add_value[1] + add_value[2] + add_value[3] + add_value[4] + add_value[5] + add_value[6];
                        foreach (KeyValuePair<string, (double[,], double[,])> entry in dic_idsurface)
                        {
                            if (combine_string == entry.Key)
                            {
                                design_check[i - 1, 0] = receive[j][6];
                                design_check[i - 1, 1] = receive[j][7];
                                design_check[i - 1, 2] = receive[j][8];
                                design_check[i - 1, 3] = receive[j][9];
                                design_check[i - 1, 4] = receive[j][10];
                                design_check[i - 1, 5] = receive[j][11];
                                design_check[i - 1, 6] = receive[j][12];
                                P = -1 * Convert.ToInt32(data[i, 11]);
                                Mx_up = Convert.ToInt32(data[i, 16]);
                                My_up = Convert.ToInt32(data[i, 17]);
                                Qx = Convert.ToInt32(data[i, 12]);
                                Qy = Convert.ToInt32(data[i, 13]);
                                double Cx; double Cy;
                                if (combine_string.Contains("Cir") == true)
                                {
                                    equal_area = Math.PI * Math.Pow(Convert.ToDouble(receive[j][3]), 2) / 4;
                                    Cx = Math.Sqrt(equal_area);
                                    Cy = Math.Sqrt(equal_area);
                                }
                                else
                                {
                                    Cx = Convert.ToDouble(receive[j][3]);
                                    Cy = Convert.ToDouble(receive[j][4]);
                                }
                                loadcheck = Convert.ToString(data[i, 6]);
                                dmain = DiameterExtract(receive[j][8]);
                                dstir = StirrupExtract(receive[j][11]).Item1;
                                sw = StirrupExtract(receive[j][11]).Item2;
                                nstir = Convert.ToInt32(receive[j][12]);
                                double[,] ver_check = entry.Value.Item1;
                                double[,] hoz_check = entry.Value.Item2;
                                DC = InteractionDiagramCheck(P, Mx_up, My_up, nsec, ver_check, hoz_check).Item3;
                                DCs = StirrupRectangleCheck(Cx, Cy, P, Qx, Qy, dmain, dstir, nstir, sw, abv, Rb, Rbt, Rsw).Item3;
                                ved = AxialCompressionRectangleCheck(Cx, Cy, P, loadcomb, loadcheck, Rb);
                                design_check[i - 1, 7] = DC;
                                design_check[i - 1, 8] = DCs;
                                design_check[i - 1, 9] = ved;
                                break;
                            }
                        }
                        break;
                    }
                }
                if (skip == true)
                {
                    for (j = 0; j <= 9; j++)
                    {
                        design_check[i - 1, j] = null;
                    }
                }
            }
        }

        public object[,] DesignResult(double Rb, double Rbt, double Eb, double Rs, double Rsc, double Rsw, double Es,
            double abv, int nsec, object[,] data, string[] loadcomb) //Hàm thiết kế và kiểm tra
        {
            string[,] design_label = GetDataDesign(data, 2);
            string[,] design_cad = GetDataDesign(data, 4);
            Design_Form myform = new Design_Form(design_label, design_cad, Rb, Eb, Rs, Rsc, Es, abv, nsec);
            myform.ShowDialog();
            string mode = myform.SendBack2().Item1;
            List<string[]> receive = myform.SendBack2().Item2;
            Dictionary<string, (double[,], double[,])> dic_idsurface = myform.SendBack2().Item3;
            byte pos;
            if (mode == "Tên ETABS") pos = 2;
            else pos = 4;
            int num = data.GetUpperBound(0); int i;
            object[,] design_check = new object[num, 10];
            // Chia luồng để load dữ liệu vào
            int[,] range = DetermineTask(1, num).Item1;
            int cpu_thread = DetermineTask(1, num).Item2;
            Thread[] thread = new Thread[cpu_thread];
            for (i = 0; i <= cpu_thread - 1; i++)
            {
                int temp = i;
                thread[temp] = new Thread(() => DesignResultTask(design_check, Rb, Rbt, Rsw, abv, nsec,
                    data, receive, dic_idsurface, loadcomb, pos, range[temp, 0], range[temp, 1]));
                thread[temp].Start();
            }
            for (i = 0; i <= cpu_thread - 1; i++)
            {
                int temp = i;
                thread[temp].Join();
            }
            return design_check;
        }

        public double Concrete(double strain, double Rb, double Eb) //Ứng suất - biến dạng bê tông
        {
            double ep_b1 = 0.0015;
            double Eb_red = Rb / ep_b1;
            double sigma;
            if ((0 <= strain) && (strain <= ep_b1))
            {
                sigma = strain * Eb_red;
            }
            else
            {
                sigma = Rb;
            }
            return sigma;
        }

        public double Rebar(double strain, double Rs, double Rsc, double Es) //Ứng suất - biến dạng cốt thép
        {
            double ep_s0 = Rs / Es;
            double ep_sc0 = Rsc / Es;
            double sigma;
            if (strain <= 0)
            {
                strain = Math.Abs(strain);
                if ((0 <= strain) && (strain <= ep_s0))
                {
                    sigma = -strain * Es;
                }
                else
                {
                    sigma = -Rs;
                }
            }
            else
            {
                if ((0 <= strain) && (strain <= ep_sc0))
                {
                    sigma = strain * Es;
                }
                else
                {
                    sigma = Rsc;
                }
            }
            return sigma;
        }

        public int DiameterExtract(string input) //Hàm tách đường kính thép
        {
            int output = 0;
            for (int i = 0; i <= input.Length - 1; i++)
            {
                if (input.Substring(i, 1) == "Ø")
                {
                    int num = input.Length - i - 1;
                    output = Convert.ToInt32(input.Substring(input.Length - num, num));
                    break;
                }
            }
            return output;
        }

        public (int, int) StirrupExtract(string input) //Hàm tách thép đai
        {
            int pos_a = 0; int pos_d = 0;
            for (int i = 0; i <= input.Length - 1; i++)
            {
                if (input.Substring(i, 1) == "Ø")
                {
                    pos_d = i;
                }
                else if (input.Substring(i, 1) == "a")
                {
                    pos_a = i;
                }
            }
            int output1 = Convert.ToInt32(input.Substring(pos_d + 1, pos_a - pos_d - 1));
            int output2 = Convert.ToInt32(input.Substring(pos_a + 1, input.Length - pos_a - 1));
            return (output1, output2);
        }





        //CÁC HÀM ĐỂ LỌC DỮ LIỆU
        public List<string[]> DetermineFilter(object[,] input) //Hàm xác định dữ liệu cần lọc
        {
            int i; int j;
            int num = input.GetUpperBound(0);
            List<string[]> data_output = new List<string[]>();
            for (i = 1; i <= num; i++)
            {
                bool skip = false;
                string[] add_value = new string[28];
                for (j = 1; j <= 28; j++)
                {
                    string check_value = Convert.ToString(input[i, j]);
                    if ((j != 4) && (j != 10) && (j != 20) && (string.IsNullOrEmpty(check_value) == true))
                    {
                        skip = true;
                        break;
                    }
                    else
                    {
                        add_value[j - 1] = check_value;
                    }
                }
                if (skip == false)
                {
                    data_output.Add(add_value);
                }
            }
            return data_output;
        }

        public object[,] Filter(object[,] input) //Hàm lọc kết quả
        {
            int i; int j;
            bool new_stage = false;
            List<string[]> list_input = DetermineFilter(input);
            List<string[]> list_output = new List<string[]>();
            Dictionary<string, (string[], string[], string[], string[])> dic_filter = new Dictionary<string, (string[], string[], string[], string[])>();
            for (i = 0; i < list_input.Count - 1; i++)
            {
                if (Convert.ToString(list_input[i][1]) != Convert.ToString(list_input[i + 1][1]) || (i == list_input.Count - 2))
                {
                    new_stage = true;
                }
                string[] store_sec = new string[8];
                string[] store_PMxy = new string[13];
                string[] store_PQ = new string[7];
                string[] store_Pved = new string[3];

                store_sec[0] = Convert.ToString(list_input[i][0]);
                store_sec[1] = Convert.ToString(list_input[i][1]);
                store_sec[2] = Convert.ToString(list_input[i][2]);
                store_sec[3] = Convert.ToString(list_input[i][3]);
                store_sec[4] = Convert.ToString(list_input[i][4]);
                store_sec[5] = Convert.ToString(list_input[i][7]);
                store_sec[6] = Convert.ToString(list_input[i][8]);
                store_sec[7] = Convert.ToString(list_input[i][9]);

                store_PMxy[0] = Convert.ToString(list_input[i][5]);
                store_PMxy[1] = Convert.ToString(list_input[i][10]);
                store_PMxy[2] = Convert.ToString(list_input[i][13]);
                store_PMxy[3] = Convert.ToString(list_input[i][14]);
                store_PMxy[4] = Convert.ToString(list_input[i][15]);
                store_PMxy[5] = Convert.ToString(list_input[i][16]);
                store_PMxy[6] = Convert.ToString(list_input[i][17]);
                store_PMxy[7] = Convert.ToString(list_input[i][18]);
                store_PMxy[8] = Convert.ToString(list_input[i][19]);
                store_PMxy[9] = Convert.ToString(list_input[i][20]);
                store_PMxy[10] = Convert.ToString(list_input[i][21]);
                store_PMxy[11] = Convert.ToString(list_input[i][22]);
                store_PMxy[12] = Convert.ToString(list_input[i][25]);

                store_PQ[0] = Convert.ToString(list_input[i][5]);
                store_PQ[1] = Convert.ToString(list_input[i][10]);
                store_PQ[2] = Convert.ToString(list_input[i][11]);
                store_PQ[3] = Convert.ToString(list_input[i][12]);
                store_PQ[4] = Convert.ToString(list_input[i][23]);
                store_PQ[5] = Convert.ToString(list_input[i][24]);
                store_PQ[6] = Convert.ToString(list_input[i][26]);

                if (Convert.ToDouble(list_input[i][27]) != 0)
                {
                    store_Pved[0] = Convert.ToString(list_input[i][5]);
                    store_Pved[1] = Convert.ToString(list_input[i][10]);
                }
                else
                {
                    store_Pved[0] = null;
                    store_Pved[1] = Convert.ToString(0);
                }
                store_Pved[2] = Convert.ToString(list_input[i][27]);

                string combine = store_sec[0] + store_sec[1] + store_sec[2] + store_sec[3] + store_sec[4] + store_sec[5] + store_sec[6] + store_sec[7];
                if (dic_filter.ContainsKey(combine) == false)
                {
                    dic_filter.Add(combine, (store_sec, store_PMxy, store_PQ, store_Pved));
                }
                else
                {
                    (string[], string[], string[], string[]) current_data;
                    dic_filter.TryGetValue(combine, out current_data);
                    (string[], string[], string[], string[]) add_data;
                    string[] current_sec = current_data.Item1;
                    string[] current_PMxy = current_data.Item2;
                    string[] current_PQ = current_data.Item3;
                    string[] current_Pved = current_data.Item4;
                    if (Convert.ToDouble(current_data.Item2[12]) < Convert.ToDouble(store_PMxy[12]))
                    {
                        current_PMxy = store_PMxy;
                    }
                    if (Convert.ToDouble(current_data.Item3[6]) < Convert.ToDouble(store_PQ[6]))
                    {
                        current_PQ = store_PQ;
                    }
                    if (Convert.ToDouble(current_data.Item4[2]) < Convert.ToDouble(store_Pved[2]))
                    {
                        current_Pved = store_Pved;
                    }
                    add_data = (current_sec, current_PMxy, current_PQ, current_Pved);
                    dic_filter[combine] = add_data;
                }
                if (new_stage == true)
                {
                    foreach (KeyValuePair<string, (string[], string[], string[], string[])> entry in dic_filter)
                    {
                        string[] add_value = new string[31];
                        for (j = 0; j < 8; j++)
                        {
                            add_value[j] = entry.Value.Item1[j];
                        }
                        for (j = 0; j < 13; j++)
                        {
                            add_value[j + 8] = entry.Value.Item2[j];
                        }
                        for (j = 0; j < 7; j++)
                        {
                            add_value[j + 21] = entry.Value.Item3[j];
                        }
                        for (j = 0; j < 3; j++)
                        {
                            add_value[j + 28] = entry.Value.Item4[j];
                        }
                        list_output.Add(add_value);
                    }
                    new_stage = false;
                    dic_filter.Clear();
                }
            }
            //Chuyển list sang chuỗi để send vào excel
            int num = list_output.Count;
            object[,] result_output = new object[num, 31];
            for (i = 0; i < num; i++)
            {
                for (j = 0; j < 31; j++)
                {
                    result_output[i, j] = list_output[i][j];
                }
            }
            return result_output;
        }





        //CÁC HÀM KHÁC
        public (int[,], int) DetermineTask(int begin, int num) //Hàm tính toán điểm đầu và cuối để chia task cho CPU
        {

            int cpu_thread;
            int total = num - begin + 1;
            if (total < 50) //Chỉ dùng 1 thread
            {
                cpu_thread = 1;
                int[,] range = new int[1, 2];
                range[0, 0] = begin;
                range[0, 1] = num;
                return (range, cpu_thread);
            } //Dùng hết thread của cpu
            else
            {
                cpu_thread = Environment.ProcessorCount;
                int[,] range = new int[cpu_thread, 2];
                int delta = total / cpu_thread;
                int assign_start; int assign_end;
                for (int i = 1; i <= cpu_thread; i++)
                {
                    assign_start = begin + delta * (i - 1); ;
                    assign_end = begin + delta * i;
                    if (i > 1) assign_start += 1;
                    if (i == cpu_thread) assign_end = num;
                    range[i - 1, 0] = assign_start;
                    range[i - 1, 1] = assign_end;
                }
                return (range, cpu_thread);
            }


        }

        public string Selector(Excel.Range cell) //Hàm để đổi range Excel thành array string
        {
            if (cell.Value2 == null)
                return "";
            if (cell.Value2.GetType().ToString() == "System.Double")
                return ((double)cell.Value2).ToString();
            else if (cell.Value2.GetType().ToString() == "System.String")
                return ((string)cell.Value2);
            else if (cell.Value2.GetType().ToString() == "System.Boolean")
                return ((bool)cell.Value2).ToString();
            else
                return "unknown";
        }

    }
    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting) //Hàm tối ưu DataGridView
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}
