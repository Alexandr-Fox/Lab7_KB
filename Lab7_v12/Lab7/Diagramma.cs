using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Lab7
{
	public sealed partial class Diagramma : Form
	{
		private SeriesChartType _style = SeriesChartType.SplineArea;
		private ChartColorPalette _palette = ChartColorPalette.SemiTransparent;
		private readonly List<DataGridViewCell> _cells;
		public string Diapazone { get; }

		public bool Spirit { get; private set; } = true;

		private readonly int _colC;
		public Diagramma(List<DataGridViewCell> items, string diapazon, int c = 2)
		{
			InitializeComponent();
			_colC = c;
			InitChart(items, c);
			Diapazone = diapazon;
			var dia = Diapazone.Split(' ');
			Text = $@"Диаграмма: C{dia[3]}R{dia[2]}:C{dia[1]}R{dia[0]}";
			_cells = items;
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
			Spirit = false;
        }

        private void Diagramma_Resize(object sender, EventArgs e)
        {
			comboBox1.Width = Width / 2;
			comboBox2.Width = Width / 2;
			Reload();
        }

        private void InitChart(List<DataGridViewCell> items, int c = 2)
        {
			chart = new Chart();
			var chartArea1 = new ChartArea
			{
				Name = "ChartArea1"
			};
			chart.Dock = DockStyle.Bottom;
			chart.Location = new Point(0, 22);
			chart.Size = new Size(Width, Height-60);
			chart.ChartAreas.Add(chartArea1);
			chart.Name = "chart";
			chart.Palette = _palette;
			var name = "";
			var legend1 = new Legend();
			legend1.LegendStyle = LegendStyle.Column;
			legend1.Name = "Legend0";
			chart.Legends.Add(legend1);
			foreach (var cell in items)
			{
				switch (cell.RowIndex)
				{
					case 0 when cell.ColumnIndex == 0:
						continue;
					case 0:
					{
						var series1 = new Series();
						series1.ChartArea = "ChartArea1";
						series1.Legend = "Legend0";
						series1.Name = cell.Value != null?cell.Value.ToString():$"Undefined {cell.ColumnIndex}";
						series1.ChartType = _style;
						chart.Series.Add(series1);
						break;
					}
					default:
					{
						if (cell.ColumnIndex % c == 0)
						{
							name = cell.Value != null ? cell.Value.ToString() : $"Undefined {cell.RowIndex}";
						}
						else
						{
							var ser = items[cell.ColumnIndex % c].Value != null ? items[cell.ColumnIndex % c].Value.ToString() : $"Undefined {items[cell.ColumnIndex % c].ColumnIndex}";
							chart.Series[ser].Points.AddXY(name, cell.Value == null ? "0" : cell.Value.ToString());
						}

						break;
					}
				}
			}
			Controls.Clear();
			Controls.Add(chart);
			Controls.Add(panel1);
		}
		public void Reload()
		{
			InitChart(_cells,_colC);
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			_style = (SeriesChartType)comboBox1.SelectedItem;
			Reload();
		}

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
			_palette = (ChartColorPalette)comboBox2.SelectedItem;
			Reload();
		}
    }
}
