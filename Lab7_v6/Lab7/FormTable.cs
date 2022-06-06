using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lab7
{
    public partial class FormTable : Form
	{
		public bool onEdit = false;
		public int LastC => lastC;
		public int LastR => lastR;
		private int lastC = 0;
		private int lastR = 0;
		public List<Diagramma> diagrammas = new List<Diagramma>();
		public string PATH { get; private set; }
		public void NewSave(string Path)
		{
			PATH = Path;
			Text = System.IO.Path.GetFileName(Path);
		}
		public FormTable(string Path)
		{
			InitializeComponent();
			Init(Path);
		}
		private void Init(string Path)
        {
			dataGridView.ColumnCount = 50;
			dataGridView.RowCount = 50;
			dataGridView.CellEndEdit += DataGridView_CellEndEdit;
			FormClosing += FormTable_FormClosing;
			PATH = Path;
			Text = System.IO.Path.GetFileName(Path);
		}
        private void FormTable_FormClosing(object sender, FormClosingEventArgs e)
        {
			if (onEdit)
			{
				DialogResult result = MessageBox.Show("Файл был изменен","Сохранить изменения?",
					MessageBoxButtons.YesNoCancel,
					MessageBoxIcon.Stop);
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

        public FormTable(string Path, SerializableMatrix Content)
		{
			InitializeComponent();
			Init(Path);
			for (int i = 0; i < dataGridView.RowCount; i++)
			{
				for (int j = 0; j < dataGridView.ColumnCount; j++)
				{
					dataGridView[j, i].Value = Content[i, j];
				}
			}
		}

        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
			onEdit = true;
			var item = ((DataGridView)sender)[e.ColumnIndex, e.RowIndex];
			for (int i = 0; i< diagrammas.Count;i++)
            {
				var diagramma = diagrammas[i];
				if (diagramma.spirit)
				{
					var items = diagramma.Diapazone.Split(' ');
					if (item.RowIndex >= Convert.ToInt32(items[2])
						&& item.RowIndex <= Convert.ToInt32(items[0])
						&& item.ColumnIndex >= Convert.ToInt32(items[3])
						&& item.ColumnIndex <= Convert.ToInt32(items[1]))
					{
						diagramma.Reload(item);
					}
				}
				else 
				{ 
					diagramma.Dispose(); 
				}
			}
			if (e.ColumnIndex > LastC && (item.Value != null || item.Value != ""))
				lastC = e.ColumnIndex;
			if (e.RowIndex > LastR && (item.Value != null || item.Value != ""))
				lastR = e.RowIndex;
		}

        public void SaveInFile()
		{
			SerializableMatrix serializableMatrix = new SerializableMatrix();
			for (int i = 0; i < dataGridView.RowCount; i++)
			{
				for (int j = 0; j < dataGridView.ColumnCount; j++)
				{
					if (dataGridView[j, i].Value == null) continue;
					serializableMatrix[i, j] = dataGridView[j, i].Value.ToString();
				}
			}
			if (Path.GetFileName(PATH).Split('.')[0] == "*")
			{
				SaveFileDialog saveFileDialog = new SaveFileDialog
				{
					InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal),
					Filter = "Текстовые файлы (*.table)|*.table"
				};
				if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
				{
					NewSave(saveFileDialog.FileName);
				}
				else
				{
					MessageBox.Show("Операция отменена");
					return;
				}
			}
			FileStream fs = new FileStream(PATH, FileMode.OpenOrCreate);
			BinaryFormatter formatter = new BinaryFormatter();
			formatter.Serialize(fs, serializableMatrix);
			fs.Close();
			onEdit = false;
		}

		private void dataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
		{
			if (e.ColumnIndex < dataGridView.ColumnCount &&
				e.RowIndex < dataGridView.RowCount && !dataGridView.SelectedCells.Contains(dataGridView[e.ColumnIndex, e.RowIndex]))
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
			bool check = false;
			List<DataGridViewCell> Numbers = new List<DataGridViewCell>();
			foreach (DataGridViewCell Cell in dataGridView.SelectedCells)
			{
				if (Cell.Value == null) continue;
				double number;
				if (!double.TryParse(Cell.Value.ToString(), out number))
				{
					MessageBox.Show("Для построения диаграммы в выбранных ячейках должны быть только численные значения!");
					return;
				}
				check = true;
				Numbers.Add(Cell);
			}
			if (!check)
			{
				MessageBox.Show("В выбранных строках нет цифр!");
				return;
			}
			var select = $"{dataGridView.SelectedCells[0].RowIndex} {dataGridView.SelectedCells[0].ColumnIndex} " +
							   $"{dataGridView.SelectedCells[dataGridView.SelectedCells.Count - 1].RowIndex} " +
                               $"{dataGridView.SelectedCells[dataGridView.SelectedCells.Count - 1].ColumnIndex}";
            Diagramma diagramma = new Diagramma(Numbers, select)
            {
                FormBorderStyle = FormBorderStyle.Sizable
            };
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
