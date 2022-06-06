using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Lab7
{
    public partial class MainForm : Form
    {
        public string BackImName = "";
        public string themes = "";
        private int childFormNumber = 0;
        private int oldW;
        private int oldH;


        public MainForm()
        {
            InitializeComponent();
            FormClosing += MainForm_FormClosing;
            ResizeBegin += MainForm_ResizeBegin;
            Resize += MainForm_Resize;
            BackImName = "wolf1";
            themes = "wolf";
            ResizeRedraw = true;
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.DoubleBuffer,
                false);
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (Math.Abs(Width - oldW) > 40 || Math.Abs(Height - oldH) > 40)
                BackgroundImage = (Image)Properties.Resources.ResourceManager.GetObject(BackImName);
        }
        private void MainForm_ResizeBegin(object sender, EventArgs e)
        {
            oldH = Height;
            oldW = Width;
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Form childForm in MdiChildren)
                childForm.Close();
        }
        private void ShowNewForm(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;


            Dictionary<string,Image[]> images = new Dictionary<string,Image[]>
            {{"wolf", new Image[] {
                Properties.Resources.wolf1,
                Properties.Resources.wolf,
                Properties.Resources.wolf2,
                Properties.Resources.wolf3}},
                {"hell", new Image[] {
                Properties.Resources.hell,
                Properties.Resources.hell1,
                Properties.Resources.hell2,
                Properties.Resources.hell3}}
            };
            Random random = new Random();
            BackgroundImage = images[themes][random.Next(images[themes].Length)];
            FormTable childForm = new FormTable(saveFileDialog1.FileName, BackImName);

            childForm.MdiParent = this;
            childForm.Text = "Окно " + ++childFormNumber;
            childForm.Show();

        }
        private void OpenFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog.Filter = "Текстовые файлы (*.HELL)|*.hell";
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = openFileDialog.FileName;

                Form[] forms = MdiChildren;
                foreach (Form form in forms)
                {
                    if (form.Text == Path.GetFileName(openFileDialog.FileName))
                    {
                        MessageBox.Show("Этот файл уже открыт!");
                        return;
                    }
                }

                try
                {
                    BinaryFormatter formatter = new BinaryFormatter();
                    SerializableMatrix content;

                    using (FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate))
                    {
                        content = (SerializableMatrix)formatter.Deserialize(fs);
                    }

                    FormTable formTable = new FormTable(FileName, content, BackImName)
                    {
                        MdiParent = this
                    };
                    formTable.Show();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            saveFileDialog.Filter = "Файлы Exel (*.xlsx)|*.xlsx|Файлы Word (*.docx)|*.docx";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
                if (FileName.Split('.')[FileName.Split('.').Length-1]=="docx")
                    saveWord(FileName);
                else
                    saveExel(FileName);
            }
        }
        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStrip.Visible = toolBarToolStripMenuItem.Checked;
        }
        private void saveWord(string FileName)
        {
            Form[] forms = MdiChildren;
            if (forms.Length == 0)
            {
                MessageBox.Show("Нечего сохранять!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (ActiveMdiChild == null)
            {
                MessageBox.Show("Нет выбранной таблицы!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Word.Application applicationWord = new Word.Application();
            Word.Document documentWORD = applicationWord.Documents.Add();
            Word.Range wordRange = documentWORD.Range();
            FormTable formTable = (FormTable)ActiveMdiChild;
            var wordTable = documentWORD.Tables.Add(wordRange, formTable.dataGridView.RowCount, formTable.dataGridView.ColumnCount);
            wordTable.set_Style("Сетка таблицы");
            for (int i = 0; i < formTable.dataGridView.ColumnCount; i++)
                for (int j = 0; j < formTable.dataGridView.RowCount; j++)
                    if (formTable.dataGridView[i, j].Value != null)
                        wordTable.Cell(j + 1, i + 1).Range.Text = formTable.dataGridView[i, j].Value.ToString();
            documentWORD.SaveAs2(FileName);
            documentWORD.Close();
            applicationWord.Quit();
            MessageBox.Show("Готово!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void saveExel(string FileName)
        {
            if (ActiveMdiChild == null)
            {
                MessageBox.Show("Нечего сохранять!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Excel.Application applicationExcel = new Excel.Application();
            Excel.Workbook workbook = applicationExcel.Workbooks.Add();
            FormTable form = (FormTable)ActiveMdiChild;
            string Name = form.Text.Split('.')[0];
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            worksheet.Name = Name;


            for (int i = 0; i < form.dataGridView.ColumnCount; i++)
            {
                for (int j = 0; j < form.dataGridView.RowCount; j++)
                {
                    if (form.dataGridView[i, j].Value != null)
                        worksheet.Cells[j + 1, i + 1] = form.dataGridView[i, j].Value.ToString();
                }
            }
            workbook.SaveAs(FileName);
            workbook.Close();
            applicationExcel.Quit();
            MessageBox.Show("Готово!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void exelToolStripButton_Click(object sender, EventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            saveFileDialog.Filter = "Файлы Exel (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
                saveExel(FileName);
            }
        }
        private void wordToolStripButton_Click(object sender, EventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            saveFileDialog.Filter = "Файлы Exel (*.docx)|*.docx";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
                saveWord(FileName);
            }
        }
        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if (ActiveMdiChild == null)
            {
                MessageBox.Show("Нечего сохранять!","",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            ((FormTable)ActiveMdiChild).SaveInFile();
            MessageBox.Show("Сохранено!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.Cascade);
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileVertical);
        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {

            this.LayoutMdi(MdiLayout.TileHorizontal);
        }
        private void словаToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (ActiveMdiChild == null)
            {
                MessageBox.Show("Нет выбранной формы!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            FormTable ActiveForm = (FormTable)ActiveMdiChild;
            if (ActiveForm.dataGridView.SelectedCells == null)
            {
                MessageBox.Show("Нет выбранных ячеек!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var form = new Form1();
                if (form.ShowDialog() == DialogResult.OK)
                {
                    string[] RowAndColumn = form.res.Split(' ');
                    if (RowAndColumn[0][0]=='R' && RowAndColumn[1][0] == 'C')
                    {
                        RowAndColumn[0] = RowAndColumn[0].TrimStart('R');
                        RowAndColumn[1] = RowAndColumn[1].TrimStart('C');
                    }
                    else
                    {
                        string s = RowAndColumn[1];
                        RowAndColumn[1] = RowAndColumn[0].TrimStart('C');
                        RowAndColumn[0] = s.TrimStart('R');
                    }
                    int Row;
                    int Column;
                    if (int.TryParse(RowAndColumn[0], out Row) && int.TryParse(RowAndColumn[1], out Column))
                    {
                        bool check = false;
                        List<string> result = new List<string>();
                        foreach (DataGridViewCell Cell in ActiveForm.dataGridView.SelectedCells)
                        {
                            if (Cell.Value != null)
                                foreach (string word in Cell.Value.ToString().Split(' ', ',', '!', '?', '.'))
                                {
                                    if (word.Length < 4) { result.Add(word); check = true; }
                                }
                        }
                        if (check)
                        {
                            ActiveForm.dataGridView[Column - 1, Row - 1].Value = result.ToString();
                            ActiveForm.dataGridView.UpdateCellValue(Column - 1, Row - 1);
                        }
                        else MessageBox.Show("В выбранных ячейках нет слов короче 4 символов!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Неверно указано значение ячейки в textBox!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void среднееЗначениеToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (ActiveMdiChild == null)
            {
                MessageBox.Show("Нет выбранной формы!");
                return;
            }
            FormTable ActiveForm = (FormTable)ActiveMdiChild;

            if (ActiveForm.dataGridView.SelectedCells == null)
            {
                MessageBox.Show("Нет выбранных ячеек!");
                return;
            }
            double sum = 0;
            int c = 0;
            foreach (DataGridViewCell Cell in ActiveForm.dataGridView.SelectedCells)
            {
                if (Cell.Value == null) continue;

                if (double.TryParse(Cell.Value.ToString(), out double number))
                {
                    sum+=Convert.ToDouble(Cell.Value.ToString());
                    c++;
                }
            }
            if (c<0) MessageBox.Show("В выбранных ячейках нет чисел!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else MessageBox.Show($"Среднее значение чисел в выделенных ячейках: {Math.Round(sum/c,2)}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void about_Click(object sender, EventArgs e)
        {
            var form = new ScreenSaver();
            form.Click += Form_Click;
            form.ShowDialog();
        }

        private void Form_Click(object sender, EventArgs e)
        {
            ScreenSaver screenSaver = (ScreenSaver)sender;
            screenSaver.Close();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            Image[] images = new Image[]
            {
                Properties.Resources.wolf1,
                Properties.Resources.wolf,
                Properties.Resources.wolf2,
                Properties.Resources.wolf3
            };
            themes = "wolf";
            Random random = new Random();
            BackgroundImage = images[random.Next(images.Length)];
            foreach (FormTable form in MdiChildren)
                form.BackgroundImage = images[random.Next(images.Length)];
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Image[] images = new Image[]
            {
                Properties.Resources.hell,
                Properties.Resources.hell1,
                Properties.Resources.hell2,
                Properties.Resources.hell3
            };
            themes = "hell";
            Random random = new Random();
            BackgroundImage = images[random.Next(images.Length)];
            foreach (FormTable form in MdiChildren)
                form.BackgroundImage = images[random.Next(images.Length)];
        }
    }
}
