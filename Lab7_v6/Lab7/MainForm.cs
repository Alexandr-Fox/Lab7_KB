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
            FormTable childForm = new FormTable("*.table");

            childForm.MdiParent = this;
            childForm.Text = "Окно " + ++childFormNumber;
            childForm.Show();

        }

        private void OpenFile(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog1.Filter = "Текстовые файлы (*.table)|*.table";
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = openFileDialog1.FileName;
                Form[] forms = MdiChildren;
                foreach (Form form in forms)
                {
                    if (form.Text == Path.GetFileName(openFileDialog1.FileName))
                    {
                        MessageBox.Show("Этот файл уже открыт!");
                        return;
                    }
                }

                try
                {
                    BinaryFormatter formatter = new BinaryFormatter();
                    FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate);

                    FormTable formTable = new FormTable(FileName, (SerializableMatrix)formatter.Deserialize(fs))
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
                if (FileName.Split('.')[FileName.Split('.').Length - 1] == "docx")
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
            FormTable formTable = (FormTable)ActiveMdiChild;
            string Name = formTable.Text.Split('.')[0];
            var wordTable = documentWORD.Tables.Add(wordRange, formTable.dataGridView.ColumnCount, formTable.dataGridView.RowCount);
            wordTable.set_Style("Сетка таблицы");
            for (int i = 0; i < formTable.dataGridView.ColumnCount; i++)
            {
                for (int j = 0; j < formTable.dataGridView.RowCount; j++)
                {
                    if (formTable.dataGridView[i, j].Value != null)
                        wordTable.Cell(j + 1, i + 1).Range.Text = formTable.dataGridView[i, j].Value.ToString();
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
            FormTable formTable = (FormTable)ActiveMdiChild;
            string Name = formTable.Text.Split('.')[0];
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            worksheet.Name = Name;
            for (int i = 0; i < formTable.dataGridView.ColumnCount; i++)
            {
                for (int j = 0; j < formTable.dataGridView.RowCount; j++)
                {
                    if (formTable.dataGridView[i, j].Value != null)
                        worksheet.Cells[j + 1, i + 1] = formTable.dataGridView[i, j].Value.ToString();
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
            if (MdiChildren.Length == 0)
            {
                MessageBox.Show("Нечего сохранять!");
                return;
            }
            BinaryFormatter formatter = new BinaryFormatter();
            FormTable form = (FormTable)ActiveMdiChild;
            SerializableMatrix serializableMatrix = new SerializableMatrix();
            for (int i = 0; i < form.dataGridView.RowCount; i++)
            {
                for (int j = 0; j < form.dataGridView.ColumnCount; j++)
                {
                    if (form.dataGridView[j, i].Value == null) continue;
                    serializableMatrix[i, j] = form.dataGridView[j, i].Value.ToString();
                }
            }
            if (Path.GetFileName(form.PATH).Split('.')[0] == "*")
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                    Filter = "Текстовые файлы (*.table)|*.table"
                };
                if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
                {
                    form.NewSave(saveFileDialog.FileName);
                }
                else
                {
                    MessageBox.Show("Операция отменена");
                    return;
                }
            }
            FileStream fs = new FileStream(form.PATH, FileMode.OpenOrCreate);
            formatter.Serialize(fs, serializableMatrix);
            fs.Close();
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
                Form1 form = new Form1();
                if (form.ShowDialog() == DialogResult.OK)
                {
                    int Row;
                    int Column;
                    if (int.TryParse(form.res.Split(' ')[0], out Row) && int.TryParse(form.res.Split(' ')[1], out Column))
                    {
                        bool check = false;
                        string result = "";
                        foreach (DataGridViewCell Cell in ActiveForm.dataGridView.SelectedCells)
                        {
                            if (Cell.Value != null)
                                foreach (string word in Cell.Value.ToString().Split(' ', ',', '!', '?', '.'))
                                {
                                    if (word[0] == 'а' || word[0] == 'А') { result += word + ' '; check = true; }
                                }
                        }
                        if (check)
                        {
                            ActiveForm.dataGridView[Column - 1, Row - 1].Value = result.ToString();
                            ActiveForm.dataGridView.UpdateCellValue(Column - 1, Row - 1);
                        }
                        else MessageBox.Show("В выбранных ячейках нет слов на букву А!");
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Неверно указано значение ячейки!");
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
                    result += number;
                    check = true;
                }
            }
            if (!check) MessageBox.Show("В выбранных ячейках нет цифр!");
            else MessageBox.Show("Сумма цифр в выбранных ячейках равна: " + result.ToString());
        }
    }
}
