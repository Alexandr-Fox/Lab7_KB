using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Lab7
{
	public partial class Diagramma : Form
	{
		public bool spirit = true;
		private string diapazone = "1 1 1 1";
		private List<DataGridViewCell> Cells;
		public string Diapazone => diapazone;
		public Diagramma(List<DataGridViewCell> Items, string diapazon)
		{
			InitializeComponent();
            FormClosing += Diagramma_FormClosing;
			foreach (DataGridViewCell Cell in Items)
			{
				chart.Series.ToString();
				chart.Series["Цифры"].Points.AddXY(Cell.RowIndex.ToString() + ' ' + Cell.ColumnIndex.ToString(), Cell.Value.ToString());
			}
			diapazone = diapazon;
			Cells = Items;
		}

        private void Diagramma_FormClosing(object sender, FormClosingEventArgs e) => spirit = false;

        private void buttonExit_Click(object sender, EventArgs e)
		{
			Close();
		}
		public void InitChart()
        {
			ChartArea chartArea1 = new ChartArea();
			Legend legend1 = new Legend();
			Series series1 = new Series();
			chart = new Chart();
			((ISupportInitialize)(chart)).BeginInit();
			SuspendLayout();
			chartArea1.Name = "ChartArea1";
			chart.ChartAreas.Add(chartArea1);
			chart.Dock = DockStyle.Fill;
			legend1.Name = "Legend1";
			chart.Legends.Add(legend1);
			chart.Location = new Point(0, 0);
			chart.Name = "chart";
			chart.Palette = ChartColorPalette.Excel;
			series1.ChartArea = "ChartArea1";
			series1.Legend = "Legend1";
			series1.Name = "Цифры";
			chart.Series.Add(series1);
			chart.Size = new Size(573, 454);
			chart.TabIndex = 1;
			chart.Text = "chart1";
		}
		public void Reload(DataGridViewCell Item)
		{
			InitChart();
			Controls.Clear();
			Controls.Add(chart);
            if (Cells.Find(p => p.RowIndex == Item.RowIndex && p.ColumnIndex == Item.ColumnIndex) == null)
                Cells.Add(Item);

            foreach (DataGridViewCell Cell in Cells)
			{
				chart.Series.ToString();
				chart.Series["Цифры"].Points.AddXY(Cell.RowIndex.ToString() + ' ' + Cell.ColumnIndex.ToString(), Cell.Value!=null && Cell.Value != ""? Cell.Value.ToString():"0");
			}
		}
    }
}
