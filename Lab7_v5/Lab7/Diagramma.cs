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
		private string diapazone = "1 1 1 1";
		private List<DataGridViewCell> Cells;
		public string Diapazone => diapazone;
		public Diagramma(List<DataGridViewCell> Items, string diapazon)
		{
			InitializeComponent();
			foreach (DataGridViewCell Cell in Items)
			{
				chart.Series.ToString();
				chart.Series["Цифры"].Points.AddXY(Cell.RowIndex.ToString() + ' ' + Cell.ColumnIndex.ToString(), Cell.Value.ToString());
			}
			diapazone = diapazon;
			Cells = Items;
		}

		private void buttonExit_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		public void Reload(DataGridViewCell Item)
		{
			if (Cells.Find(p => p.RowIndex == Item.RowIndex && p.ColumnIndex == Item.ColumnIndex)!=null)
				Cells[Cells.IndexOf(Cells.Find(p => p.RowIndex == Item.RowIndex && p.ColumnIndex == Item.ColumnIndex))] = Item;
			else
				Cells.Add(Item);
			ChartArea chartArea1 = new ChartArea();
			Legend legend1 = new Legend();
			Series series1 = new Series();
			chart = new Chart();
			((ISupportInitialize)(chart)).BeginInit();
			this.SuspendLayout();
			// 
			// chart
			// 
			chartArea1.Name = "ChartArea1";
			chart.ChartAreas.Add(chartArea1);
			chart.Dock = DockStyle.Fill;
			legend1.Name = "Legend1";
			chart.Legends.Add(legend1);
			chart.Location = new Point(0, 0);
			chart.Name = "chart";
			chart.Palette = ChartColorPalette.None;
			series1.ChartArea = "ChartArea1";
			series1.Legend = "Legend1";
			series1.Name = "Цифры";
			chart.Series.Add(series1);
			chart.Size = new Size(573, 454);
			chart.TabIndex = 1;
			chart.Text = "chart1";
			Controls.Clear();
			Controls.Add(chart);
			foreach (DataGridViewCell Cell in Cells)
			{
				chart.Series.ToString();
				chart.Series["Цифры"].Points.AddXY(Cell.RowIndex.ToString() + ' ' + Cell.ColumnIndex.ToString(), Cell.Value.ToString());
			}
		}
    }
}
