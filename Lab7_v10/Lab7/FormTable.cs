using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;

namespace Lab7
{
    public partial class FormTable : Form
	{
		public bool onEdit = false;
		public List<Diagramma> diagrammas = new List<Diagramma>();
		public string BackImName = "";
		private object oldV;
		public string PATH { get; private set; }
		private int oldW;
		private int oldH;

		public FormTable(string Path, string Image)
		{
			InitializeComponent();
			Init(Path, Image);
		}
		private void Init(string Path, string Image, int col = 5, int row = 5)
		{
            dataGridView.CellBeginEdit += DataGridView_CellBeginEdit;
			ResizeRedraw = true;
			SetStyle(
				ControlStyles.AllPaintingInWmPaint |
				ControlStyles.DoubleBuffer,
				false);
			ResizeEnd += FormTable_ResizeEnd;
            ResizeBegin += FormTable_ResizeBegin;
			dataGridView.ColumnCount = col;
			dataGridView.RowCount = row;
			for (int i = 0; i < dataGridView.Columns.Count; i++)
			{
				dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
				dataGridView.Columns[i].HeaderText = $"C{i}";
			}
			dataGridView.ColumnHeadersVisible = true;
			for (int i = 0; i < dataGridView.Rows.Count; i++)
			{
				dataGridView.Rows[i].HeaderCell.Value = $"R{i}";
			}
			dataGridView.RowHeadersVisible = true;
			dataGridView.CellEndEdit += DataGridView_CellEndEdit;
			dataGridView.ForeColor = Color.White;
			dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
			BackImName = Image;
			BackgroundImage = (Image)Properties.Resources.ResourceManager.GetObject(Image);
			dataGridView.SetCellsTransparent();
			FormClosing += FormTable_FormClosing;
			PATH = Path;
			Text = System.IO.Path.GetFileName(Path);
		}

        private void DataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
			oldV = dataGridView[e.ColumnIndex, e.RowIndex].Value;
			dataGridView[e.ColumnIndex, e.RowIndex].Style.ForeColor = Color.Black;

		}

        private void FormTable_ResizeBegin(object sender, EventArgs e)
		{
			oldH = Height;
			oldW = Width;
		}

        private void FormTable_ResizeEnd(object sender, EventArgs e)
		{
			if (Math.Abs(Width - oldW) > 40 || Math.Abs(Height - oldH) > 40)
				BackgroundImage = (Image)Properties.Resources.ResourceManager.GetObject(BackImName);

		}

		public FormTable(string Path, SerializableMatrix Content, string Image)
		{
			InitializeComponent();
			Init(Path, Image, Content.Column, Content.Row);
			for (int i = 0; i < dataGridView.RowCount; i++)
			{
				for (int j = 0; j < dataGridView.ColumnCount; j++)
				{
					if (Content[i, j] != "")
						dataGridView[j, i].Value = Content[i, j];
				}
			}

		}

        private void FormTable_FormClosing(object sender, FormClosingEventArgs e)
        {
			if (onEdit)
			{
				DialogResult result = MessageBox.Show("Файл был изменен","Сохранить изменения?",
					MessageBoxButtons.YesNoCancel,
					MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					SaveInFile();
				}
				else if (result == DialogResult.Cancel)
				{
					e.Cancel = true;
					return;
				}
			}
			foreach(Diagramma diagramma in diagrammas)
				diagramma.Close();
        }

		private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
		{
			dataGridView[e.ColumnIndex, e.RowIndex].Style.ForeColor = Color.White;
			var item = ((DataGridView)sender)[e.ColumnIndex, e.RowIndex];
			if (oldV != item.Value)
			{
				onEdit = true;
				Text = $"* {Text}";
				for (int i = 0; i < diagrammas.Count; i++)
				{
					var diagramma = diagrammas[i];
					var items = diagramma.Diapazone.Split(' ');
					if (item.RowIndex >= Convert.ToInt32(items[2])
						&& item.RowIndex <= Convert.ToInt32(items[0])
						&& item.ColumnIndex >= Convert.ToInt32(items[3])
						&& item.ColumnIndex <= Convert.ToInt32(items[1])
						&& item.Value != "")
					{
						diagramma.Reload();
					}
					if (diagramma.Spirit==false)
						diagrammas.Remove(diagramma);
				}
				if (e.ColumnIndex == dataGridView.ColumnCount - 1)
				{
					dataGridView.Columns.Add($"C{dataGridView.ColumnCount}", $"C{dataGridView.ColumnCount}");
				}
				
			}
        }

        public void SaveInFile()
		{
			SerializableMatrix serializableMatrix = new SerializableMatrix(dataGridView.ColumnCount, dataGridView.RowCount);
			for (int i = 0; i < dataGridView.RowCount; i++)
			{
				for (int j = 0; j < dataGridView.ColumnCount; j++)
				{
					if (dataGridView[j, i].Value == null) continue;
					serializableMatrix[i, j] = dataGridView[j, i].Value.ToString();
				}
			}
			BinaryFormatter formatter = new BinaryFormatter();
			using (FileStream fs = new FileStream(PATH, FileMode.OpenOrCreate))
			{
				formatter.Serialize(fs, serializableMatrix);
			}
			Text = Path.GetFileName(PATH);
			onEdit = false;
		}

		private void dataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
		{
			if (e.RowIndex >= 0 &&
				e.ColumnIndex >= 0 &&
				e.ColumnIndex < dataGridView.ColumnCount &&
				e.RowIndex < dataGridView.RowCount &&
				!dataGridView.SelectedCells.Contains(
					dataGridView[e.ColumnIndex, e.RowIndex]))
			{
				dataGridView.CurrentCell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
			}
		}

		private void диаграммаToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (dataGridView.SelectedCells.Count <= 0)
			{
				MessageBox.Show("Выберите ячейки!");
				return;
			}
			List<DataGridViewCell> Numbers = new List<DataGridViewCell>();
			int startC = dataGridView.ColumnCount;
			int startR = dataGridView.RowCount;
			int endC = 0;
			int endR = 0;
			foreach (DataGridViewCell cell in dataGridView.SelectedCells)
            {
				if(startC > cell.ColumnIndex) startC = cell.ColumnIndex;
				if(startR > cell.RowIndex) startR = cell.RowIndex;
				if(endC < cell.ColumnIndex) endC = cell.ColumnIndex;
				if(endR < cell.RowIndex) endR = cell.RowIndex;
            }

			if (endC - startC  >= 1 && endR - startR >= 1)
				for (int i = startR; i < endR + 1;i++)
					for(int j = startC; j < endC + 1; j++)
					{
						var Cell = dataGridView[j, i];
						if (!double.TryParse(Cell.Value==null?"0":Cell.Value.ToString(), out double number) && 
							Cell.ColumnIndex != startC &&
							Cell.RowIndex!= startR)
						{
							MessageBox.Show("Для построения диаграммы в выделенных ячейках везде кроме первой строки и первого столбца должны быть числа!");
							return;
						}
						Numbers.Add(Cell);
					}
            else
            {
				MessageBox.Show("Выделенный диапазо не удовлетворяет условиям для постароения диаграммы!\n" +
								"Легенда: первая выделенная строка\n" +
								"Подписи: первый веделенный столбец\n" +
								"Данные: все остальные выделенные ячейки");
				return;
			}
			var select = $"{dataGridView.SelectedCells[0].RowIndex} {dataGridView.SelectedCells[0].ColumnIndex} " +
							   $"{dataGridView.SelectedCells[dataGridView.SelectedCells.Count - 1].RowIndex} " +
                               $"{dataGridView.SelectedCells[dataGridView.SelectedCells.Count - 1].ColumnIndex}";
            Diagramma diagramma = new Diagramma(Numbers, select, endC - startC + 1);
            diagramma.Show();
			diagrammas.Add(diagramma);
		}

		private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (dataGridView.SelectedCells.Count <= 0) return;

			foreach (DataGridViewCell Cell in dataGridView.SelectedCells)
			{
				Cell.Value = null;
			}
			onEdit = true;
		}

		private void распознатьToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (dataGridView.SelectedCells.Count <= 0) return;

			double rabbit;

			if (dataGridView.SelectedCells.Count == 1)
			{
				DataGridViewCell cell = dataGridView.SelectedCells[0];
				if (cell.Value == null)
				{
					MessageBox.Show("Выбранная ячейка пустая!");
				}
				else if (double.TryParse(cell.Value.ToString(), out rabbit))
				{
					MessageBox.Show("Выбранная ячейка содержит числовую информацию!");
				}
				else MessageBox.Show("Выбранная ячейка содержит текстовую информацию!");
				return;
			}

			bool num = false;
			bool alph = false;
			bool empty = false;

			foreach (DataGridViewCell Cell in dataGridView.SelectedCells)
			{

				if (Cell.Value == null) empty = true;
				else if (double.TryParse(Cell.Value.ToString(), out rabbit)) num = true;
				else alph = true;
			}
			string result = "";
			if (num) result += "Среди выделенных ячейках есть числовые!\n";
			if (alph) result += "Среди выделенных ячейках есть буквенные!\n";
			if (empty) result += "Среди выделенных ячеек есть пустые!";

			MessageBox.Show(result);
		}

    }
}
