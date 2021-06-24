using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace NDKToolsExcel
{
    public partial class RecColShow_Form : Form
    {
        double Rb; double Rbt; double Eb; double Rs; double Rsc; double Rsw; double Es; int nsec; string shape; double umin; double umax;

        public RecColShow_Form(double re_Cx, double re_Cy, double re_Rb, double re_Rbt, double re_Eb, double re_Rs, double re_Rsc, 
            double re_Rsw, double re_Es, double re_L, int re_abv, int re_nsec, string re_dmain, int re_nx, int re_ny, string re_dstir, 
            int re_nstir, int re_P, int re_Mx, int re_My, int re_Pq, int re_Qx, int re_Qy, int re_Ped, string re_shape, string re_cad_label,
            string re_etabs_label, string re_story, double re_umin, double re_umax)
        {

            Rb = re_Rb;
            Rbt = re_Rbt;
            Eb = re_Eb;
            Rs = re_Rs;
            Rsc = re_Rsc;
            Rsw = re_Rsw;
            Es = re_Es;
            nsec = re_nsec;
            shape = re_shape;
            umin = re_umin;
            umax = re_umax;
            InitializeComponent();
            if (shape == "Rec")
            {
                label1.Text = "Phương X";
                label3.Text = "Phương Y";
                label8.Text = "Số thanh thép phương X";
                label7.Text = "Số thanh thép phương Y";
            }
            else if (shape == "Wall")
            {
                label1.Text = "Chiều dài L";
                label3.Text = "Bề dày t";
                label8.Text = "Số thanh thép cạnh dài L";
                label7.Text = "Số thanh thép cạnh bề dày t";
            }
            var instance = new NDKFunction();
            int dmain = instance.DiameterExtract(re_dmain);
            int dstir = instance.StirrupExtract(re_dstir).Item1;
            int sw = instance.StirrupExtract(re_dstir).Item2;
            FillTextBox1(re_Cx * 1000, re_Cy * 1000, re_nx, re_ny, dmain, dstir, 
                sw, re_nstir, re_abv, re_L, re_P, re_Mx, re_My, re_Pq, re_Qx, re_Qy, re_Ped, re_shape, re_cad_label, re_etabs_label, re_story);
        }

        private void RecColShow_Form_Load(object sender, EventArgs e)
        {

            ShowMemberRectangle();

        }

        public void FormatChart(Chart chart, ChartArea objectchart, double[,] arr_in, int point1, int point2, string title, 
            string Xlabel, string Ylabel, bool swap, int[] maxvalue) // Hàm format chart
        {
            int posX = 1; int posY = 0;
            if (swap == true)
            {
                posX = 0;
                posY = 1;
            }
            chart.Series.Clear();
            chart.Titles.Add(title);
            chart.Titles[0].Font = new Font("Segoe UI", 10);
            chart.Series.Add("srs");
            chart.Series[0].IsVisibleInLegend = false;
            chart.Series[0].ChartArea = "ChartArea1";
            chart.Series[0].ChartType = SeriesChartType.Spline;
            chart.Series[0].BorderWidth = 2;
            for (int i = 0; i <= arr_in.GetUpperBound(0); i++)
            {
                chart.Series[0].Points.AddXY(Convert.ToInt32(arr_in[i, posX]), Convert.ToInt32(arr_in[i, posY]));
            }
            chart.Series.Add("srs2");
            chart.Series[1].IsVisibleInLegend = false;
            chart.Series[1].ChartArea = "ChartArea1";
            chart.Series[1].ChartType = SeriesChartType.Point;
            chart.Series[1].MarkerSize = 7;
            chart.Series[1].Points.AddXY(point1, point2);
            objectchart.AxisX.Title = Xlabel;
            objectchart.AxisX.TitleAlignment = StringAlignment.Center;
            objectchart.AxisX.LabelStyle.Format = "F0";
            objectchart.AxisX.TitleFont = new Font("Segoe UI", 10);
            objectchart.AxisX.LabelStyle.Font = new Font("Segoe UI", 10);
            objectchart.AxisY.Title = Ylabel;
            objectchart.AxisY.TitleAlignment = StringAlignment.Center;
            objectchart.AxisY.LabelStyle.Format = "F0";
            objectchart.AxisY.TitleFont = new Font("Segoe UI", 10);
            objectchart.AxisY.LabelStyle.Font = new Font("Segoe UI", 10);
            objectchart.AxisX.Crossing = 0;
            objectchart.AxisX.MajorGrid.LineWidth = 0;
            objectchart.AxisX.LineWidth = 2;
            objectchart.AxisX.ArrowStyle = AxisArrowStyle.Triangle;
            objectchart.AxisY.Crossing = 0;
            objectchart.AxisY.MajorGrid.LineWidth = 0;
            objectchart.AxisY.LineWidth = 2;
            objectchart.AxisY.ArrowStyle = AxisArrowStyle.Triangle;
            objectchart.AxisX.Maximum = Convert.ToInt32(1.1 * maxvalue[0]);
            objectchart.AxisX.Minimum = Convert.ToInt32(1.1 * maxvalue[1]);
            objectchart.AxisY.Maximum = Convert.ToInt32(1.1 * maxvalue[2]);
            objectchart.AxisY.Minimum = Convert.ToInt32(1.3 * maxvalue[3]);
        }

        public void SectionChart(Chart chart, ChartArea objectchart) // Hàm chart hiện thị tiết diện thép
        {
            int Cx = Convert.ToInt32(CxTextBox.Text);
            int Cy = Convert.ToInt32(CyTextBox.Text);
            int nx = Convert.ToInt32(nxTextBox.Text);
            int ny = Convert.ToInt32(nyTextBox.Text);
            int abv = Convert.ToInt32(abvTextBox.Text);
            int dmain = Convert.ToInt32(dmainTextBox.Text);
            int dstir = Convert.ToInt32(dstirTextBox.Text);
            chart.Series.Clear();
            chart.Series.Add("srs");
            chart.Series[0].IsVisibleInLegend = false;
            chart.Series[0].ChartArea = "ChartArea1";
            chart.Series[0].ChartType = SeriesChartType.Line;
            chart.Series[0].BorderWidth = 2;
            chart.Series[0].Points.AddXY(Cx/2, Cy/2);
            chart.Series[0].Points.AddXY(Cx/2, -Cy/2);
            chart.Series[0].Points.AddXY(-Cx/2, -Cy/2);
            chart.Series[0].Points.AddXY(-Cx/2, Cy/2);
            chart.Series[0].Points.AddXY(Cx/2, Cy/2);

            chart.Series[0].Color = Color.Black;
            double Cx_cen = Cx - 2 * abv - 2 * dstir - dmain;
            double Cy_cen = Cy - 2 * abv - 2 * dstir - dmain;
            int sum_rebar = 2 * (nx + ny) - 4;
            double[,] rebar_element = new double[sum_rebar, 2];
            int j = 0;
            for (int i = 0; i < nx; i++)
            {
                rebar_element[j, 0] = -Cx_cen / 2 + Cx_cen / (nx - 1) * i;
                rebar_element[j, 1] = Cy_cen / 2;
                rebar_element[j + 1, 0] = rebar_element[j, 0];
                rebar_element[j + 1, 1] = -Cy_cen / 2;
                j = j + 2;
            }
            for (int i = 1; i < ny - 1; i++)
            {
                rebar_element[j, 0] = Cx_cen / 2;
                rebar_element[j, 1] = -Cy_cen / 2 + Cy_cen / (ny - 1) * i;
                rebar_element[j + 1, 0] = -Cx_cen / 2;
                rebar_element[j + 1, 1] = rebar_element[j, 1];
                j = j + 2;
            }
            chart.Series.Add("srs2");
            chart.Series[1].IsVisibleInLegend = false;
            chart.Series[1].ChartArea = "ChartArea1";
            chart.Series[1].ChartType = SeriesChartType.Point;
            chart.Series[1].BorderWidth = 2;
            chart.Series[1].MarkerStyle = MarkerStyle.Circle;
            chart.Series[1].Color = Color.Red;
            chart.Series[1].MarkerColor = Color.Red;
            for (j = 0; j <= rebar_element.GetUpperBound(0); j++)
            {
                chart.Series[1].Points.AddXY(rebar_element[j, 0], rebar_element[j, 1]);
            }
            float h = (float)((float)100 * 420 / 180 * Cy / Cx);
            if (h > 100) h = 100;
            float w = (float)((float)100 * 180 / 420 * Cx / Cy); 
            if (w > 100) w = 100;
            float pos_X = (100 - w) / 2;
            float pos_Y = (100 - h) / 2;
            objectchart.Position.Height = h;
            objectchart.Position.Width = w;
            objectchart.Position.X = pos_X;
            objectchart.Position.Y = pos_Y;
            objectchart.AxisX.Maximum = Cx / 2;
            objectchart.AxisX.Minimum = -Cx / 2;
            objectchart.AxisY.Maximum = Cy / 2;
            objectchart.AxisY.Minimum = -Cy / 2;
            objectchart.AxisX.Enabled = AxisEnabled.False;
            objectchart.AxisY.Enabled = AxisEnabled.False;
      
        }


        public (double [,], double [,], int, int, int, int, double, int, int, object, double) 
            ReCalculate(double Rb, double Rbt, double Eb, double Rs, double Rsc, double Rsw, double Es, int nsec) // Hàm tính toán lại để show kết quả
        {
            // Lấy giá trị tính từ TextBox
            double k = Convert.ToDouble(kTextBox.Text);
            double Cx = Convert.ToDouble(CxTextBox.Text) / 1000;
            double Cy = Convert.ToDouble(CyTextBox.Text) / 1000;
            int nx = Convert.ToInt32(nxTextBox.Text);
            int ny = Convert.ToInt32(nyTextBox.Text);
            int dmain = Convert.ToInt32(dmainTextBox.Text);
            int dstir = Convert.ToInt32(dstirTextBox.Text);
            int sw = Convert.ToInt32(swTextBox.Text);
            int nstir = Convert.ToInt32(nstirTextBox.Text);
            int abv = Convert.ToInt32(abvTextBox.Text);
            double L = Convert.ToDouble(LTextBox.Text);
            int P = Convert.ToInt32(PTextBox.Text);
            int Mx = Math.Abs(Convert.ToInt32(MxTextBox.Text));
            int My = Math.Abs(Convert.ToInt32(MyTextBox.Text));
            int Pq = Convert.ToInt32(PqTextBox.Text);
            int Qx = Math.Abs(Convert.ToInt32(QxTextBox.Text));
            int Qy = Math.Abs(Convert.ToInt32(QyTextBox.Text));
            int Ped = Convert.ToInt32(PedTextBox.Text);

            var instance = new NDKFunction();
            // PMxy
            int Mx_up = instance.RectangleUpperMoment(P, Mx, My, Cx, Cy, L, Eb, k).Item1;
            int My_up = instance.RectangleUpperMoment(P, Mx, My, Cx, Cy, L, Eb, k).Item2;
            double[,] ver_check = instance.IDSurfaceRectangle(Cx * 1000, Cy * 1000, Eb, Es, abv, Rb, Rs, Rsc, nx, ny, dmain, dstir, nsec).Item1;
            double[,] hoz_check = instance.IDSurfaceRectangle(Cx * 1000, Cy * 1000, Eb, Es, abv, Rb, Rs, Rsc, nx, ny, dmain, dstir, nsec).Item2;
            int Pnxy = instance.InteractionDiagramCheck(P, Mx_up, My_up, nsec, ver_check, hoz_check).Item1;
            int Mnxy = instance.InteractionDiagramCheck(P, Mx_up, My_up, nsec, ver_check, hoz_check).Item2;
            double DC = instance.InteractionDiagramCheck(P, Mx_up, My_up, nsec, ver_check, hoz_check).Item3;
            double[,] out_PMxy = instance.InteractionDiagramCheck(P, Mx_up, My_up, nsec, ver_check, hoz_check).Item4;
            double[,] out_MxMy = instance.InteractionDiagramCheck(P, Mx_up, My_up, nsec, ver_check, hoz_check).Item5;
            // PQ
            int Qnx = instance.StirrupRectangleCheck(Cx, Cy, Pq, Qx, Qy, dmain, dstir, nstir, sw, abv, Rb, Rbt, Rsw).Item1;
            int Qny = instance.StirrupRectangleCheck(Cx, Cy, Pq, Qx, Qy, dmain, dstir, nstir, sw, abv, Rb, Rbt, Rsw).Item2;
            object DCs = instance.StirrupRectangleCheck(Cx, Cy, Pq, Qx, Qy, dmain, dstir, nstir, sw, abv, Rb, Rbt, Rsw).Item3;
            // ved
            double ved = 0;
            if (Ped > 0)
            {
                double fcd;
                fcd = (1868.6 * Rb - 1465.7) * 0.8 / 1.2;
                ved = Math.Round(Ped / (Cx * Cy * fcd), 2);
            }

            return (out_PMxy, out_MxMy, Mx_up, My_up, Pnxy, Mnxy, DC, Qnx, Qny, DCs, ved);
        }

        public void FillTextBox1(double Cx, double Cy, int nx, int ny, int dmain, int dstir, int sw, 
            int nstir, int abv, double L, int P, int Mx, int My, int Pq, int Qx, int Qy, int Ped, 
            string shape, string cad_label, string etabs_label, string story) // Hàm show TextBox sẵn có
        {
            int sum_rebar = 2 * (nx + ny) - 4;
            int Asc = Convert.ToInt32(Math.PI * Math.Pow(dmain, 2) / 4 * sum_rebar);
            double muy = Math.Round(Asc / (Cx * Cy) * 100, 2);

            nxTextBox.Text = Convert.ToString(nx);
            nyTextBox.Text = Convert.ToString(ny);
            sumTextBox.Text = Convert.ToString(sum_rebar);
            dmainTextBox.Text = Convert.ToString(dmain);
            AscTextBox.Text = Convert.ToString(Asc);
            rebarperTextBox.Text = Convert.ToString(muy);
            dstirTextBox.Text = Convert.ToString(dstir);
            swTextBox.Text = Convert.ToString(sw);
            nstirTextBox.Text = Convert.ToString(nstir);
            abvTextBox.Text = Convert.ToString(abv);
            LTextBox.Text = Convert.ToString(L);
            PTextBox.Text = Convert.ToString(P);
            MxTextBox.Text = Convert.ToString(Mx);
            MyTextBox.Text = Convert.ToString(My);
            PqTextBox.Text = Convert.ToString(Pq);
            QxTextBox.Text = Convert.ToString(Qx);
            QyTextBox.Text = Convert.ToString(Qy);
            PedTextBox.Text = Convert.ToString(Ped);
            CxTextBox.Text = Convert.ToString(Cx);
            CyTextBox.Text = Convert.ToString(Cy);
            kTextBox.Text = Convert.ToString(1);
            TypeTextBox.Text = shape;
            CADTextBox.Text = cad_label;
            ETABSTextBox.Text = etabs_label;
            PositionTextBox.Text = story;
        }

        public void FillTextBox2 (int Mx_up, int My_up, int Pnxy, int Mnxy, 
            double DC, int Qnx, int Qny, object DCs, double ved) // Hàm show TextBox sau khi tính
        {
            int Mxy = Convert.ToInt32(Math.Sqrt(Math.Pow(Mx_up, 2) + Math.Pow(My_up, 2)));
            Mx_upTextBox.Text = Convert.ToString(Mx_up);
            My_upTextBox.Text = Convert.ToString(My_up);
            MxyTextBox.Text = Convert.ToString(Mxy);
            PnxyTextBox.Text = Convert.ToString(Pnxy);
            MnxyTextBox.Text = Convert.ToString(Mnxy);
            DCTextBox.Text = Convert.ToString(DC);
            QnxTextBox.Text = Convert.ToString(Qnx);
            QnyTextBox.Text = Convert.ToString(Qny);
            DCsTextBox.Text = Convert.ToString(DCs);
            vedTextBox.Text = Convert.ToString(ved);
        }

        public (int, int, int, int) MaxMin (double[,] arr_in, bool swap) // Hàm tìm max min cho chart
        {
            int posX = 1; int posY = 0;
            if (swap == true)
            {
                posX = 0;
                posY = 1;
            }
            double Xmax = arr_in[0, posX];
            double Xmin = arr_in[0, posX];
            double Ymax = arr_in[0, posY];
            double Ymin = arr_in[0, posY];
            for (int i = 1; i <= arr_in.GetUpperBound(0); i++)
            {
                if (Xmax < arr_in[i, posX]) Xmax = arr_in[i, posX];
                if (Xmin > arr_in[i, posX]) Xmin = arr_in[i, posX];
                if (Ymax < arr_in[i, posY]) Ymax = arr_in[i, posY];
                if (Ymin > arr_in[i, posY]) Ymin = arr_in[i, posY];
            }
            return (Convert.ToInt32(Xmax), Convert.ToInt32(Xmin), Convert.ToInt32(Ymax), Convert.ToInt32(Ymin));
        }

        public void ShowMemberRectangle () // Gom hàm vào 1 function
        {
            chart1.ChartAreas.Clear();
            chart2.ChartAreas.Clear();
            chart3.ChartAreas.Clear();
            chart1.Series.Clear();
            chart2.Series.Clear();
            chart3.Series.Clear();
            chart1.Titles.Clear();
            chart2.Titles.Clear();
            chart3.Titles.Clear();

            double[,] out_PMxy = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item1;
            double[,] out_MxMy = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item2;
            int Mx_up = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item3;
            int My_up = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item4;
            int Pnxy = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item5;
            int Mnxy = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item6;
            double DC = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item7;
            int Qnx = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item8;
            int Qny = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item9;
            object DCs = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item10;
            double ved = ReCalculate(Rb, Rbt, Eb, Rs, Rsc, Rsw, Es, nsec).Item11;
            FillTextBox2(Mx_up, My_up, Pnxy, Mnxy, DC, Qnx, Qny, DCs, ved);
            int P = Convert.ToInt32(PTextBox.Text);
            int Mxy = Convert.ToInt32(MxyTextBox.Text);

            ChartArea ChartArea1 = new ChartArea("ChartArea1");
            ChartArea ChartArea2 = new ChartArea("ChartArea1");
            ChartArea ChartArea3 = new ChartArea("ChartArea1");
            chart1.ChartAreas.Add(ChartArea1);
            chart2.ChartAreas.Add(ChartArea2);
            chart3.ChartAreas.Add(ChartArea3);
            var objectchart1 = chart1.ChartAreas[0];
            var objectchart2 = chart2.ChartAreas[0];
            var objectchart3 = chart3.ChartAreas[0];
            string title1 = "Mặt cắt đứng BĐTT"; string labelX1 = "Mxy (kNm)"; string labelY1 = "P (kN)";
            string title2 = "Mặt cắt ngang BĐTT"; string labelX2 = "Mx (kNm)"; string labelY2 = "My (kNm)";

            // Xác định max min để vẽ chart
            int[] MaxValue = new int[4] { Math.Max(MaxMin(out_PMxy, false).Item1, Mxy), 0, Math.Max(MaxMin(out_PMxy, false).Item3, P), Math.Min(MaxMin(out_PMxy, false).Item4, P) };
            int[] MaxValue2 = new int[4] { Math.Max(MaxMin(out_MxMy, true).Item1, Mx_up), 0, Math.Max(MaxMin(out_MxMy, true).Item3, My_up), 0 };

            FormatChart(chart1, objectchart1, out_PMxy, Mxy, P, title1, labelX1, labelY1, false, MaxValue);
            FormatChart(chart2, objectchart2, out_MxMy, Mx_up, My_up, title2, labelX2, labelY2, true, MaxValue2);
            objectchart2.AlignWithChartArea = objectchart1.Name;
            objectchart2.AlignmentOrientation = AreaAlignmentOrientations.Vertical;
            objectchart2.AlignmentStyle = AreaAlignmentStyles.All;
            SectionChart(chart3, objectchart3);
        } 

        private void button1_Click(object sender, EventArgs e) // Tính toán lại
        {
            int nx = Convert.ToInt32(nxTextBox.Text);
            int ny = Convert.ToInt32(nyTextBox.Text);
            int dmain = Convert.ToInt32(dmainTextBox.Text);
            double Cx = Convert.ToDouble(CxTextBox.Text);
            double Cy = Convert.ToDouble(CyTextBox.Text);
            int sum_rebar = 2 * (nx + ny) - 4;
            int Asc = Convert.ToInt32(Math.PI * Math.Pow(dmain, 2) / 4 * sum_rebar);
            double muy = Math.Round(Asc / (Cx * Cy) * 100, 2);

            sumTextBox.Text = Convert.ToString(sum_rebar);
            AscTextBox.Text = Convert.ToString(Asc);
            rebarperTextBox.Text = Convert.ToString(muy);

            ShowMemberRectangle();
        }

        public (object[], object[], object[], object[]) SendBack() // Trả kết quả về Main Form
        {
            var instance = new NDKFunction();
            object[] send_data1 = new object[3];
            object[] send_data2 = new object[12];
            object[] send_data3 = new object[6];
            object[] send_data4 = new object[2];
            int P = Convert.ToInt32(PTextBox.Text);
            int Mx_up = Convert.ToInt32(Mx_upTextBox.Text);
            int My_up = Convert.ToInt32(My_upTextBox.Text);
            double Cx = (double)Convert.ToInt32(CxTextBox.Text) / 1000;
            double Cy = (double)Convert.ToInt32(CyTextBox.Text) / 1000;
            int abv = Convert.ToInt32(abvTextBox.Text);
            send_data1[0] = LTextBox.Text;
            send_data1[1] = Cx;
            send_data1[2] = Cy;

            send_data2[0] = Convert.ToInt32(PTextBox.Text) * -1;
            send_data2[1] = MxTextBox.Text;
            send_data2[2] = MyTextBox.Text;
            send_data2[3] = Mx_upTextBox.Text;
            send_data2[4] = My_upTextBox.Text;

            send_data2[5] = Convert.ToString(instance.SimplifiedRectangleID(P, Mx_up, My_up, Cx, Cy, abv, Rb, Rs, Rsc, Es, umin, umax));
            send_data2[6] = nxTextBox.Text;
            send_data2[7] = nyTextBox.Text;
            int sumn = 2 * (Convert.ToInt32(nxTextBox.Text) + Convert.ToInt32(nyTextBox.Text)) - 4;
            send_data2[8] = Convert.ToString(sumn) + "Ø" + dmainTextBox.Text;
            send_data2[9] = AscTextBox.Text;
            send_data2[10] = rebarperTextBox.Text;
            send_data2[11] = DCTextBox.Text;

            send_data3[0] = Convert.ToInt32(PqTextBox.Text) * -1;
            send_data3[1] = QxTextBox.Text;
            send_data3[2] = QyTextBox.Text;
            send_data3[3] = "Ø" + dstirTextBox.Text + "a" + swTextBox.Text;
            send_data3[4] = nstirTextBox.Text;
            send_data3[5] = DCsTextBox.Text;

            send_data4[0] = Convert.ToInt32(PedTextBox.Text) * -1;
            send_data4[1] = vedTextBox.Text;
            return (send_data1, send_data2, send_data3, send_data4);
        }


    }
}
