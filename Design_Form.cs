using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace NDKToolsExcel
{
    public partial class Design_Form : Form
    {
        string[,] design_label; string[,] design_cad;
        double Rb; double Eb; double Rs; double Rsc; double Es; double abv; int nsec;
        public Design_Form(string[,] recieve_label, string[,] recieve_cad, double Rb_rev, double Eb_rev, double Rs_rev, double Rsc_rev,
             double Es_rev, double abv_rev, int nsec_rev) //Nhận data từ Main_Class
        {
            InitializeComponent();
            dataGridView1.DoubleBuffered(true);
            design_label = recieve_label;
            design_cad = recieve_cad;
            Rb = Rb_rev;
            Eb = Eb_rev;
            Rs = Rs_rev;
            Rsc = Rsc_rev;
            Es = Es_rev;
            abv = abv_rev;
            nsec = nsec_rev;
            TypeDesignCombo.SelectedIndex = 0;
            dataGridView1.SelectionChanged += dataGridView1_SelectionChanged;
        }

        private void TypeDesignCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable data = new DataTable();
            data.Columns.Add("Story");
            data.Columns.Add("Name");
            data.Columns.Add("Shape");
            data.Columns.Add("Cx");
            data.Columns.Add("Cy");
            data.Columns.Add("μtt");
            data.Columns.Add("N.o X");
            data.Columns.Add("N.o Y");
            data.Columns.Add("Total");
            data.Columns.Add("Asc");
            data.Columns.Add("μsc");
            data.Columns.Add("Stirrup");
            data.Columns.Add("Legged");
            string[,] export_label;
            if ((string)TypeDesignCombo.SelectedItem == "Tên ETABS")
            {
                export_label = design_label;
            }
            else
            {
                export_label = design_cad;
            }
            for (int i = 0; i <= export_label.GetUpperBound(0); i++)
            {
                if (string.IsNullOrEmpty(export_label[i, 1]) == false)
                {
                    string[] temp = new string[13];
                    for (int j = 0; j <= 12; j++)
                    {
                        temp[j] = Convert.ToString(export_label[i, j]);
                    }
                    data.Rows.Add(temp);
                }
            }
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = data;
        }
        private void NoXTextBox_TextChanged(object sender, EventArgs e)
        {
            DynamicShow();
        }
        private void NoYTextBox_TextChanged(object sender, EventArgs e)
        {
            DynamicShow();
        }
        private void MainDiameterTextBox_TextChanged(object sender, EventArgs e)
        {
            DynamicShow();
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DynamicShow();
        }
        private void DynamicShow()
        {
            double Cx = 0; double Cy = 0; double D = 0;
            int NoX; int NoY; int MainDia;
            int RebarArea; double RebarPercent;
            bool switch_show = false;
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                if ((string)row.Cells["Shape"].Value == "Cir")
                {
                    D = Convert.ToDouble(row.Cells["Cx"].Value);
                    switch_show = false;
                    NoXTextBox.ReadOnly = false;
                    NoYTextBox.ReadOnly = true;
                }
                else
                {
                    Cx = Convert.ToDouble(row.Cells["Cx"].Value);
                    Cy = Convert.ToDouble(row.Cells["Cy"].Value);
                    switch_show = true;
                    NoXTextBox.ReadOnly = false;
                    NoYTextBox.ReadOnly = false;
                }
            }
            if (string.IsNullOrEmpty(NoXTextBox.Text) == true) NoX = 0;
            else NoX = Convert.ToInt32(NoXTextBox.Text);
            if (string.IsNullOrEmpty(NoYTextBox.Text) == true) NoY = 0;
            else NoY = Convert.ToInt32(NoYTextBox.Text);
            if (string.IsNullOrEmpty(MainDiameterTextBox.Text) == true) MainDia = 0;
            else MainDia = Convert.ToInt32(MainDiameterTextBox.Text);
            if (switch_show == true)
            {
                RebarArea = (int)((2 * (NoX + NoY) - 4) * Math.PI * Math.Pow(MainDia, 2) / 4);
                RebarPercent = Math.Round(RebarArea * 0.0001 / (Cx * Cy), 2);
                RebarAreaTextBox.Text = Convert.ToString(RebarArea);
                RebarPercentTextBox.Text = Convert.ToString(RebarPercent);
            }
            else
            {
                RebarArea = (int)(NoX * Math.PI * Math.Pow(MainDia, 2) / 4);
                RebarPercent = Math.Round(RebarArea * 0.0001 / (Math.PI * Math.Pow(D, 2) / 4), 2);
                RebarAreaTextBox.Text = Convert.ToString(RebarArea);
                RebarPercentTextBox.Text = Convert.ToString(RebarPercent);
            }

        }

        private void button1_Click(object sender, EventArgs e) //Gán thông số nhập
        {
            TypeDesignCombo.Enabled = false;
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                if ((string)row.Cells["Shape"].Value == "Cir")
                {
                    row.Cells["N.o X"].Value = NoXTextBox.Text;
                    row.Cells["Total"].Value = NoXTextBox.Text + "Ø" + MainDiameterTextBox.Text;

                }
                else
                {
                    row.Cells["N.o X"].Value = NoXTextBox.Text;
                    row.Cells["N.o Y"].Value = NoYTextBox.Text;
                    int sum_rebar = 0;
                    if (string.IsNullOrEmpty(NoXTextBox.Text) == false && string.IsNullOrEmpty(NoYTextBox.Text) == false)
                    {
                        sum_rebar = 2 * (Convert.ToInt32(NoXTextBox.Text) + Convert.ToInt32(NoYTextBox.Text)) - 4;
                    }
                    row.Cells["Total"].Value = Convert.ToString(sum_rebar) + "Ø" + MainDiameterTextBox.Text;
                }
                row.Cells["Asc"].Value = RebarAreaTextBox.Text;
                row.Cells["μsc"].Value = RebarPercentTextBox.Text;
                row.Cells["Stirrup"].Value = "Ø" + StirDiameterTextBox.Text + "a" + SpacingTextBox.Text;
                row.Cells["Legged"].Value = NoLeggedTextBox.Text;
            }
        }

        public (string, List<string[]>, Dictionary<string, (double[,], double[,])>) SendBack2() //Trả kết quả về Main_Form
        {
            DataTable data = (DataTable)dataGridView1.DataSource;
            string mode = TypeDesignCombo.Text;
            List<string[]> data_output = new List<string[]>();
            int i; int j;
            for (i = 0; i < data.Rows.Count; i++)
            {
                bool skip = false;
                string[] add_value = new string[13];
                for (j = 0; j < data.Columns.Count; j++)
                {
                    string check_value = Convert.ToString(data.Rows[i].ItemArray[j]);
                    if ((j != 4) && (j != 7) && (string.IsNullOrEmpty(check_value) == true))
                    {
                        skip = true;
                        break;
                    }
                    else
                    {
                        add_value[j] = check_value;
                    }
                }
                if (skip == false)
                {
                    data_output.Add(add_value);
                }
            }
            Dictionary<string, (double[,], double[,])> dic_idsurface;
            dic_idsurface = DetermineSection(data_output);
            return (mode, data_output, dic_idsurface);
        }
        public Dictionary<string, (double[,], double[,])> DetermineSection(List<string[]> data_input)
        {
            Dictionary<string, (double[,], double[,])> dic_idsurface = new Dictionary<string, (double[,], double[,])>();
            double Cx; double Cy; int nx; int ny; int dmain; int dstir;
            double D; int nsum;
            int i; int j;
            var instance = new NDKFunction();
            for (i = 0; i < data_input.Count; ++i)
            {
                string[] add_value = new string[13];
                string[] temp_value = new string[2];
                for (j = 0; j < 13; j++)
                {
                    add_value[j] = Convert.ToString(data_input[i][j]);
                }
                temp_value[0] = Convert.ToString(instance.DiameterExtract(add_value[8]));
                temp_value[1] = Convert.ToString(instance.StirrupExtract(add_value[11]).Item1);
                string combine_string = add_value[2] + add_value[3] + add_value[4] + add_value[6] + add_value[7] + temp_value[0] + temp_value[1];
                if (dic_idsurface.ContainsKey(combine_string) == false)
                {
                    if (combine_string.Contains("Cir") == true)
                    {
                        D = Convert.ToDouble(add_value[3]) * 1000;
                        nsum = Convert.ToInt32(add_value[6]);
                        dmain = Convert.ToInt32(temp_value[0]);
                        dstir = Convert.ToInt32(temp_value[1]);
                        double[,] ver_value = instance.IDSurfaceCircle(D, Eb, Es, abv, Rb, Rs, Rsc, nsum, dmain, dstir, nsec).Item1;
                        double[,] hoz_value = instance.IDSurfaceCircle(D, Eb, Es, abv, Rb, Rs, Rsc, nsum, dmain, dstir, nsec).Item2;
                        dic_idsurface.Add(combine_string, (ver_value, hoz_value));
                    }
                    else
                    {
                        Cx = Convert.ToDouble(add_value[3]) * 1000;
                        Cy = Convert.ToDouble(add_value[4]) * 1000;
                        nx = Convert.ToInt32(add_value[6]);
                        ny = Convert.ToInt32(add_value[7]);
                        dmain = Convert.ToInt32(temp_value[0]);
                        dstir = Convert.ToInt32(temp_value[1]);
                        double[,] ver_value = instance.IDSurfaceRectangle(Cx, Cy, Eb, Es, abv, Rb, Rs, Rsc, nx, ny, dmain, dstir, nsec).Item1;
                        double[,] hoz_value = instance.IDSurfaceRectangle(Cx, Cy, Eb, Es, abv, Rb, Rs, Rsc, nx, ny, dmain, dstir, nsec).Item2;
                        dic_idsurface.Add(combine_string, (ver_value, hoz_value));
                    }
                }
            }
            return dic_idsurface;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
                    Excel.Worksheet xlWorkSheet = xlWorkbook.Worksheets.get_Item(1);
                    DataTable data = (DataTable)dataGridView1.DataSource;
                    for (int i = 0; i <= data.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j <= data.Columns.Count - 1; j++)
                        {
                            xlWorkSheet.Cells[i + 1, j + 1] = Convert.ToString(data.Rows[i].ItemArray[j]);
                        }
                    }
                    xlWorkSheet.Cells[1, 14] = Convert.ToString(TypeDesignCombo.SelectedItem);
                    xlWorkbook.SaveAs(saveFileDialog.FileName, Excel.XlFileFormat.xlWorkbookDefault);
                    xlWorkbook.Close(0);
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
                else
                {
                    return;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(openFileDialog.FileName);
                    Excel.Worksheet xlData = xlWorkbook.Worksheets[1];
                    xlApp.Visible = false;
                    int last_row = xlData.UsedRange.Rows.Count;
                    object[,] arr_data = xlData.Range["A1", "M" + last_row].Value2;
                    string mode = Convert.ToString(xlData.Cells[1, 14].Value2);
                    xlWorkbook.Close(0);
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    DataTable data = new DataTable();
                    data.Columns.Add("Story");
                    data.Columns.Add("Name");
                    data.Columns.Add("Shape");
                    data.Columns.Add("Cx");
                    data.Columns.Add("Cy");
                    data.Columns.Add("μtt");
                    data.Columns.Add("N.o X");
                    data.Columns.Add("N.o Y");
                    data.Columns.Add("Total");
                    data.Columns.Add("Asc");
                    data.Columns.Add("μsc");
                    data.Columns.Add("Stirrup");
                    data.Columns.Add("Legged");
                    for (int i = 1; i <= arr_data.GetUpperBound(0); i++)
                    {
                        string[] temp = new string[13];
                        for (int j = 0; j < 13; j++)
                        {
                            temp[j] = Convert.ToString(arr_data[i, j + 1]);
                        }
                        data.Rows.Add(temp);
                    }
                    TypeDesignCombo.SelectedItem = mode;
                    TypeDesignCombo.Enabled = false;
                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = data;
                }
                else
                {
                    return;
                }
            }
            
        }
    }
}
