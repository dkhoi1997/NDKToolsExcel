using System;
using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace NDKToolsExcel
{
    public partial class AssignCAD_Form : Form
    {
        Dictionary<string, string> dic_label = new Dictionary<string, string>();
        public AssignCAD_Form(Dictionary<string, string> dic_label) //Nhận data Label và CAD
        {
            InitializeComponent();
            dataGridView1.DoubleBuffered(true);
            DataTable data = new DataTable();
            data.Columns.Add("Tên ETABS");
            data.Columns.Add("Tên CAD");
            foreach (KeyValuePair<string, string> entry in dic_label)
            {
                string[] temp = new string[2];
                temp[0] = entry.Key;
                temp[1] = entry.Value;
                data.Rows.Add(temp);
            }
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = data;
        }

        private void button1_Click(object sender, EventArgs e) //Gán tên
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                row.Cells["Tên CAD"].Value = AssignTextBox.Text;
            }
        }
        public Dictionary<string, string> SendBack() //Trả kết quả về Main_Class
        {
            DataTable data = (DataTable)dataGridView1.DataSource;
            foreach (DataRow row in data.Rows)
            {
                string label = Convert.ToString(row["Tên ETABS"]);
                string cad = Convert.ToString(row["Tên CAD"]);
                dic_label[label] = cad;
            }
            return dic_label;
        }

        private void button2_Click(object sender, EventArgs e) //Xuất kết quả ra file Excel
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
        private void button3_Click(object sender, EventArgs e) //Lấy dữ liệu gán CAD sẵn từ Excel
        {
            int last_row;
            object[,] arr_label = null;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(openFileDialog.FileName);
                    Excel.Worksheet xlLabel = xlWorkbook.Worksheets[1];
                    xlApp.Visible = false;
                    last_row = xlLabel.UsedRange.Rows.Count;
                    arr_label = xlLabel.Range["A1", "B" + last_row].Value2;
                    xlWorkbook.Close(0);
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
                else
                {
                    return;
                }
            }
            DataTable data = new DataTable();
            data.Columns.Add("Tên ETABS");
            data.Columns.Add("Tên CAD");
            for (int i = 1; i <= arr_label.GetUpperBound(0); i++)
            {
                string[] temp = new string[2];
                temp[0] = Convert.ToString(arr_label[i, 1]);
                temp[1] = Convert.ToString(arr_label[i, 2]);
                data.Rows.Add(temp);
            }
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = data;
        }
    }
}
