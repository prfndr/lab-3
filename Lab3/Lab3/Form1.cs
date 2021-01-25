using System;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lab3
{
    public partial class Form1 : Form
    {
        dynamic xlApp;
        dynamic xlWorksheet;
        dynamic xlRange;
        Type typeExcel;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadExcel();

            for (int i = 1; i < 5; i++)
            {
                string funcName = xlRange.Cells[i, 1].Value2;
                comboBox1.Items.Add(funcName);
            }
            xlApp.Quit();
        }

        private void LoadExcel()
        {
            typeExcel = Type.GetTypeFromProgID("Excel.Application");
            xlApp = Activator.CreateInstance(typeExcel);
            dynamic xlWorkbook = xlApp.Workbooks.Open(Application.StartupPath + "\\Lab3.1.xlsm");
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void btnBuildChart_Click(object sender, EventArgs e)
        {
            LoadExcel();

            int typeFunc = comboBox1.SelectedIndex + 1;

            xlRange.Cells[2, 6].Value2 = typeFunc;
            

            int x1 = (int)numericUpDown1.Value;
            int x2 = (int)numericUpDown2.Value;

            int i = 10;

            for (int x = x1; x <= x2; x++)
            {
                xlRange.Cells[3, 6].Value2 = x;
                double y = xlRange.Cells[6, 6].Value2;

                xlRange.Cells[i, 2].Value2 = x;
                xlRange.Cells[i, 3].Value2 = y;
                i++;
            }

            dynamic shape = xlWorksheet.Shapes.AddChart2(240, Excel.XlChartType.xlXYScatterSmooth);
            dynamic series = shape.Chart.SeriesCollection(1);
            series.XValues = xlWorksheet.Range("B10", "B" + (i - 1));
            series.Values = xlWorksheet.Range("C10", "C" + (i - 1));

            //xlApp.Visible = true;
            typeExcel.InvokeMember("Visible", BindingFlags.SetProperty, null, xlApp, new object[1] { true });
        }
    }
}
