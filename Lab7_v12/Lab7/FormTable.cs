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
		private bool _onEdit;
		private readonly List<Diagramma> _diagrammas = new List<Diagramma>();
		private object _oldV;
		private string Path { get; set; }
		public FormTable(string path)
		{
			InitializeComponent();
			Init(path);
		}
		private void Init(string path, int col = 5, int row = 5)
		{
			dataGridView.ColumnCount = col;
			dataGridView.RowCount = row;
			for (var i = 0; i < dataGridView.Columns.Count; i++)
			{
				dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
				dataGridView.Columns[i].HeaderText = $@"C{i}";
			}
			dataGridView.ColumnHeadersVisible = true;
			for (var i = 0; i < dataGridView.Rows.Count; i++)
			{
				dataGridView.Rows[i].HeaderCell.Value = $"R{i}";
			}
			dataGridView.RowHeadersVisible = true;
			dataGridView.CellEndEdit += DataGridView_CellEndEdit;
			dataGridView.BackgroundColor = Color.Coral;
			dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
			FormClosing += FormTable_FormClosing;
			this.Path = path;
			Text = System.IO.Path.GetFileName(path);
		}

		public FormTable(string path, SerializableMatrix content)
		{
			InitializeComponent();
			Init(path, content.Column, content.Row);
			for (var i = 0; i < dataGridView.RowCount; i++)
			{
				for (var j = 0; j < dataGridView.ColumnCount; j++)
				{
					if (content[i, j] != "")
						dataGridView[j, i].Value = content[i, j];
				}
			}

		}

        private void FormTable_FormClosing(object sender, FormClosingEventArgs e)
        {
			if (_onEdit)
			{
				var result = MessageBox.Show(@"Файл был изменен",@"Сохранить изменения?",
					MessageBoxButtons.YesNoCancel,
					MessageBoxIcon.Question);
				switch (result)
				{
					case DialogResult.Yes:
						SaveInFile();
						break;
					case DialogResult.Cancel:
						e.Cancel = true;
						return;
				}
			}
			foreach(var diagramma in _diagrammas)
				diagramma.Close();
        }

		private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
		{
			var item = ((DataGridView)sender)[e.ColumnIndex, e.RowIndex];
			if (_oldV == item.Value) return;
			_onEdit = true;
			Text = $@"* {Text}";
			for (var i = 0; i < _diagrammas.Count; i++)
			{
				var diagramma = _diagrammas[i];
				var items = diagramma.Diapazone.Split(' ');
				if (item.RowIndex >= Convert.ToInt32(items[2])
				    && item.RowIndex <= Convert.ToInt32(items[0])
				    && item.ColumnIndex >= Convert.ToInt32(items[3])
				    && item.ColumnIndex <= Convert.ToInt32(items[1])
				    && (string)item.Value != "")
					diagramma.Reload();
				if (diagramma.Spirit==false)
					_diagrammas.Remove(diagramma);
			}
			if (e.ColumnIndex == dataGridView.ColumnCount - 1)
				dataGridView.Columns.Add($"C{dataGridView.ColumnCount}", $"C{dataGridView.ColumnCount}");
		}

        public void SaveInFile()
		{
			var serializableMatrix = new SerializableMatrix(dataGridView.ColumnCount, dataGridView.RowCount);
			for (var i = 0; i < dataGridView.RowCount; i++)
				for (var j = 0; j < dataGridView.ColumnCount; j++)
				{
					if (dataGridView[j, i].Value == null) continue;
					serializableMatrix[i, j] = dataGridView[j, i].Value.ToString();
				}
			var formatter = new BinaryFormatter();
			var fs = new FileStream(Path, FileMode.OpenOrCreate);
			formatter.Serialize(fs, serializableMatrix);
			Text = System.IO.Path.GetFileName(Path);
			_onEdit = false;
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
				MessageBox.Show(@"Выберите ячейки!");
				return;
			}
			var numbers = new List<DataGridViewCell>();
			var startC = dataGridView.ColumnCount;
			var startR = dataGridView.RowCount;
			var endC = 0;
			var endR = 0;
			foreach (DataGridViewCell cell in dataGridView.SelectedCells)
            {
				if(startC > cell.ColumnIndex) startC = cell.ColumnIndex;
				if(startR > cell.RowIndex) startR = cell.RowIndex;
				if(endC < cell.ColumnIndex) endC = cell.ColumnIndex;
				if(endR < cell.RowIndex) endR = cell.RowIndex;
            }

			if (endC - startC  >= 1 && endR - startR >= 1)
				for (var i = startR; i < endR + 1;i++)
					for(var j = startC; j < endC + 1; j++)
					{
						var cell = dataGridView[j, i];
						if (!double.TryParse(cell.Value==null?"0":cell.Value.ToString(), out _) && 
							cell.ColumnIndex != startC &&
							cell.RowIndex!= startR)
						{
							MessageBox.Show(@"Для построения диаграммы в выделенных ячейках везде кроме первой строки и первого столбца должны быть числа!");
							return;
						}
						numbers.Add(cell);
					}
            else
            {
				MessageBox.Show(@"Выделенный диапазо не удовлетворяет условиям для постароения диаграммы!\n" +
								@"Легенда: первая выделенная строка\n" +
								@"Подписи: первый веделенный столбец\n" +
								@"Данные: все остальные выделенные ячейки");
				return;
			}
			var select = $"{dataGridView.SelectedCells[0].RowIndex} {dataGridView.SelectedCells[0].ColumnIndex} " +
							   $"{dataGridView.SelectedCells[dataGridView.SelectedCells.Count - 1].RowIndex} " +
                               $"{dataGridView.SelectedCells[dataGridView.SelectedCells.Count - 1].ColumnIndex}";
            var diagramma = new Diagramma(numbers, select, endC - startC + 1);
            diagramma.Show();
			_diagrammas.Add(diagramma);
		}

		private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (dataGridView.SelectedCells.Count <= 0) return;

			foreach (DataGridViewCell cell in dataGridView.SelectedCells)
			{
				cell.Value = null;
			}
			_onEdit = true;
		}

		private void распознатьToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (dataGridView.SelectedCells.Count <= 0) return;

			if (dataGridView.SelectedCells.Count == 1)
			{
				var cell = dataGridView.SelectedCells[0];
				if (cell.Value == null)
				{
					MessageBox.Show(@"Выбранная ячейка пустая!");
				}
				else if (double.TryParse(cell.Value.ToString(), out _))
				{
					MessageBox.Show(@"Выбранная ячейка содержит числовую информацию!");
				}
				else MessageBox.Show(@"Выбранная ячейка содержит текстовую информацию!");
				return;
			}

			var num = false;
			var alph = false;
			var empty = false;

			foreach (DataGridViewCell cell in dataGridView.SelectedCells)
			{

				if (cell.Value == null) empty = true;
				else if (double.TryParse(cell.Value.ToString(), out _)) num = true;
				else alph = true;
			}
			var result = "";
			if (num) result += "Среди выделенных ячейках есть числовые!\n";
			if (alph) result += "Среди выделенных ячейках есть буквенные!\n";
			if (empty) result += "Среди выделенных ячеек есть пустые!";

			MessageBox.Show(result);
		}

    }
}
