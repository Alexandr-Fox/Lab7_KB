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
		private SeriesChartType style = SeriesChartType.SplineArea;
		private ChartColorPalette palette = ChartColorPalette.SemiTransparent;
		private string diapazone = "1 1 1 1";
		private List<DataGridViewCell> Cells;
		public string Diapazone => diapazone;
		private bool spirit = true;
		public bool Spirit => spirit;
		private int colC = 2;
		public Diagramma(List<DataGridViewCell> Items, string diapazon, int c = 2)
		{
			InitializeComponent();
			colC = c;
			InitChart(Items, c);
			diapazone = diapazon;
			var items = Diapazone.Split(' ');
			Text = $"Диаграмма: C{items[3]}R{items[2]}:C{items[1]}R{items[0]}";
			Cells = Items;
			comboBox1.Items.Add(SeriesChartType.SplineArea);
			comboBox1.Items.Add(SeriesChartType.Line);
			comboBox1.Items.Add(SeriesChartType.Bar);
			comboBox1.Items.Add(SeriesChartType.Range);
			comboBox1.Items.Add(SeriesChartType.Area);
			comboBox1.Items.Add(SeriesChartType.Pie);
			comboBox1.Items.Add(SeriesChartType.Column);
			comboBox1.Items.Add(SeriesChartType.Pyramid);
			comboBox1.SelectedIndex = 0;
			comboBox2.Items.Add(ChartColorPalette.SemiTransparent);
			comboBox2.Items.Add(ChartColorPalette.SeaGreen);
			comboBox2.Items.Add(ChartColorPalette.Fire);
			comboBox2.Items.Add(ChartColorPalette.None);
			comboBox2.Items.Add(ChartColorPalette.Pastel);
			comboBox2.Items.Add(ChartColorPalette.Excel);
			comboBox2.Items.Add(ChartColorPalette.Light);
			comboBox2.Items.Add(ChartColorPalette.Grayscale);
			comboBox2.Items.Add(ChartColorPalette.Bright);
			comboBox2.Items.Add(ChartColorPalette.Berry);
			comboBox2.SelectedIndex = 0;
			Resize += Diagramma_Resize;
            FormClosing += Diagramma_FormClosing;
		}

        private void Diagramma_FormClosing(object sender, FormClosingEventArgs e)
        {
			spirit = false;
        }

        private void Diagramma_Resize(object sender, EventArgs e)
        {
			comboBox1.Width = Width / 2;
			comboBox2.Width = Width / 2;
			Reload();
        }

        private void InitChart(List<DataGridViewCell> Items, int c = 2)
        {
			chart = new Chart();
			ChartArea chartArea1 = new ChartArea
			{
				Name = "ChartArea1"
			};
			chart.Dock = DockStyle.Bottom;
			chart.Location = new Point(0, 22);
			chart.Size = new Size(Width, Height-60);
			chart.ChartAreas.Add(chartArea1);
			chart.Name = "chart";
			chart.Palette = palette;
			string name = "";
			Legend legend1 = new Legend();
			legend1.LegendStyle = LegendStyle.Column;
			legend1.Name = $"Legend0";
			chart.Legends.Add(legend1);
			for (int i = 0; i < Items.Count; i++)
			{
				DataGridViewCell Cell = Items[i];
				if (Cell.RowIndex == 0 && Cell.ColumnIndex == 0)
					continue;
				if (Cell.RowIndex == 0)
				{
					Series series1 = new Series();
					series1.ChartArea = "ChartArea1";
					series1.Legend = $"Legend0";
					series1.Name = Cell.Value != null?Cell.Value.ToString():$"Undefined {Cell.ColumnIndex}";
					series1.ChartType = style;
					chart.Series.Add(series1);
				}
                else
				{
					if (Cell.ColumnIndex % c == 0)
					{
						name = Cell.Value != null ? Cell.Value.ToString() : $"Undefined {Cell.RowIndex}";
					}
					else
					{
						string ser = Items[Cell.ColumnIndex % c].Value != null ? Items[Cell.ColumnIndex % c].Value.ToString() : $"Undefined {Items[Cell.ColumnIndex % c].ColumnIndex}";
						chart.Series[ser].Points.AddXY(name, Cell.Value == null ? "0" : Cell.Value.ToString());
					}
				}
			}
			Controls.Clear();
			Controls.Add(chart);
			Controls.Add(panel1);
		}
		public void Reload()
		{
			InitChart(Cells,colC);
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			style = (SeriesChartType)comboBox1.SelectedItem;
			Reload();
		}

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
			palette = (ChartColorPalette)comboBox2.SelectedItem;
			Reload();
		}
    }
}
