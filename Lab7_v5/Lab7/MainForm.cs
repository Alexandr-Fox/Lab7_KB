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
        private int childFormNumber = 0;

        public MainForm()
        {
            InitializeComponent();
            FormClosing += MainForm_FormClosing;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void ShowNewForm(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;

            FormTable childForm = new FormTable(saveFileDialog1.FileName);

            childForm.MdiParent = this;
            childForm.Text = "Окно " + ++childFormNumber;
            childForm.Show();

        }

        private void OpenFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog.Filter = "Текстовые файлы (*.maks)|*.maks";
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

                    FormTable formTable = new FormTable(FileName, content)
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
                MessageBox.Show("Нечего сохранять!");
                return;
            }
            Word.Application applicationWord = new Word.Application();
            Word.Document documentWORD = applicationWord.Documents.Add();
            Word.Range wordRange = documentWORD.Range();
            for (int i = 0; i < forms.Length; i++)
            {
                FormTable formTable = (FormTable)forms[i];
                documentWORD.Tables.Add(wordRange, formTable.dataGridView.RowCount, formTable.dataGridView.ColumnCount);
                wordRange.InsertBreak();
            }
            for (int k = 0; k < forms.Length; k++)
            {
                Form form = forms[k];
                FormTable formTable = (FormTable)form;
                string Name = form.Text.Split('.')[0];
                var wordTable = documentWORD.Tables[k];

                wordTable.set_Style("Сетка таблицы");

                for (int i = 0; i < formTable.dataGridView.ColumnCount; i++)
                {
                    for (int j = 0; j < formTable.dataGridView.RowCount; j++)
                    {
                        if (formTable.dataGridView[i, j].Value != null)
                            wordTable.Cell(j + 1, i + 1).Range.Text = formTable.dataGridView[i, j].Value.ToString();
                    }
                }
            }
            documentWORD.SaveAs2(FileName);
            documentWORD.Close();
            applicationWord.Quit();
            MessageBox.Show("Готово!");
        }
        private void saveExel(string FileName)
        {
            Form[] forms = MdiChildren;
            if (forms.Length == 0)
            {
                MessageBox.Show("Нечего сохранять!");
                return;
            }
            Excel.Application applicationExcel = new Excel.Application();
            Excel.Workbook workbook = applicationExcel.Workbooks.Add();
            for (int k = 0; k < forms.Length - 1; k++)
                workbook.Worksheets.Add();
            for (int k = 0; k < forms.Length; k++)
            {
                Form form = forms[k];
                FormTable formTable = (FormTable)form;
                string Name = form.Text.Split('.')[0];
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[k + 1];
                worksheet.Name = Name;


                for (int i = 0; i < formTable.dataGridView.ColumnCount; i++)
                {
                    for (int j = 0; j < formTable.dataGridView.RowCount; j++)
                    {
                        if (formTable.dataGridView[i, j].Value != null)
                            worksheet.Cells[j + 1, i + 1] = formTable.dataGridView[i, j].Value.ToString();
                    }
                }
            }
            workbook.SaveAs(FileName);
            workbook.Close();
            applicationExcel.Quit();
            MessageBox.Show("Готово!");
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

            Form[] forms = MdiChildren;
            if (forms.Length == 0)
            {
                MessageBox.Show("Нечего сохранять!");
                return;
            }
            foreach (FormTable form in forms)
            {
                SerializableMatrix serializableMatrix = new SerializableMatrix();
                for (int i = 0; i < form.dataGridView.RowCount; i++)
                {
                    for (int j = 0; j < form.dataGridView.ColumnCount; j++)
                    {
                        if (form.dataGridView[j, i].Value == null) continue;
                        serializableMatrix[i, j] = form.dataGridView[j, i].Value.ToString();
                    }
                }
                BinaryFormatter formatter = new BinaryFormatter();
                using (FileStream fs = new FileStream(form.PATH, FileMode.OpenOrCreate))
                {
                    formatter.Serialize(fs, serializableMatrix);
                }
            }
            MessageBox.Show("Сохранено!");
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

        private void записатьВЯчейкуВсесловаToolStripMenuItem_Click(object sender, EventArgs e)
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

            try
            {
                var form = new Form1();
                if (form.ShowDialog() == DialogResult.OK)
                {
                    string[] RowAndColumn = form.res.Split(' ');
                    int Row;
                    int Column;
                    if (int.TryParse(RowAndColumn[0], out Row) && int.TryParse(RowAndColumn[1], out Column))
                    {
                        bool check = false;
                        string result = "";
                        foreach (DataGridViewCell Cell in ActiveForm.dataGridView.SelectedCells)
                        {
                            if (Cell.Value != null)
                                foreach (string word in Cell.Value.ToString().Split(' ', ',', '!', '?', '.'))
                                {
                                    if (Regex.IsMatch(word, @".*Del.*")) { result += word + ' '; check = true; }
                                }
                        }
                        if (check)
                        {
                            ActiveForm.dataGridView[Column - 1, Row - 1].Value = result.ToString();
                            ActiveForm.dataGridView.UpdateCellValue(Column - 1, Row - 1);
                        }
                        else MessageBox.Show("В выбранных ячейках нет слов c Del!");
                    }
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Неверно указано значение ячейки в textBox!");
            }

        }

        private void количествоЧетныхЧиселToolStripMenuItem_Click(object sender, EventArgs e)
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
            float result = 0;
            bool check = false;
            foreach (DataGridViewCell Cell in ActiveForm.dataGridView.SelectedCells)
            {
                float number;
                if (Cell.Value == null) continue;

                if (float.TryParse(Cell.Value.ToString(), out number))
                {
                    if (Convert.ToInt32(Cell.Value.ToString()) % 2 == 0)
                    {
                        result++;
                        check = true;
                    }
                }
            }
            if (!check) MessageBox.Show("В выбранных ячейках нет цифр!");
            else MessageBox.Show($"Количество четных чисел: {result}" );
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
    }
}
