using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace NDKToolsExcel
{
    public partial class Modify_Form : Form
    {
        List<string[]> modify_label;
        List<string[]> modify_cad;
        public Modify_Form(List<string[]> receive_label, List<string[]> receive_cad)
        {
            InitializeComponent();
            dataGridView1.DoubleBuffered(true);
            modify_label = receive_label;
            modify_cad = receive_cad;
            TypeModifyCombo.SelectedIndex = 0;
        }
        private void TypeModifyCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable data = new DataTable();
            data.Columns.Add("Story");
            data.Columns.Add("Name");
            data.Columns.Add("Shape");
            data.Columns.Add("L");
            data.Columns.Add("Cx");
            data.Columns.Add("Cy");
            data.Columns.Add("ex");
            data.Columns.Add("ey");
            List<string[]> export;
            if ((string)TypeModifyCombo.SelectedItem == "Tên ETABS")
            {
                export = modify_label;
            }
            else
            {
                export = modify_cad;
            }

            for (int i = 0; i < export.Count; i++)
            {
                if (string.IsNullOrEmpty(export[i][1]) == false)
                {
                    string[] temp = new string[6];
                    for (int j = 0; j <= 5; j++)
                    {
                        temp[j] = Convert.ToString(export[i][j]);
                    }
                    data.Rows.Add(temp);
                }
            }
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = data;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            TypeModifyCombo.Enabled = false;
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                if (string.IsNullOrEmpty(LengthTextBox.Text) != true) row.Cells["L"].Value = LengthTextBox.Text;
                if (string.IsNullOrEmpty(CxTextBox.Text) != true) row.Cells["Cx"].Value = CxTextBox.Text;
                if (string.IsNullOrEmpty(CyTextBox.Text) != true) row.Cells["Cy"].Value = CyTextBox.Text;
                if (string.IsNullOrEmpty(EccentricXTextBox.Text) != true) row.Cells["ex"].Value = EccentricXTextBox.Text;
                if (string.IsNullOrEmpty(EccentricYTextBox.Text) != true) row.Cells["ey"].Value = EccentricYTextBox.Text;
            }
        }

        public (string, List<string[]>) SendBack() //Trả kết quả về Main_Class
        {
            string mode = TypeModifyCombo.Text;
            DataTable data = (DataTable)dataGridView1.DataSource;
            List<string[]> modify_output = new List<string[]>();
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string[] add_value = new string[8];
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    add_value[j] = Convert.ToString(data.Rows[i].ItemArray[j]);
                }
                modify_output.Add(add_value);
            }
            return (mode, modify_output);
        }
    }
}
