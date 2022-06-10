using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Lab7
{
    public partial class MainForm : Form
    {
        private int _childFormNumber;
        public MainForm()
        {
            InitializeComponent();
            FormClosing += MainForm_FormClosing;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var childForm in MdiChildren)
                childForm.Close();
        }
        private void ShowNewForm(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            var childForm = new FormTable(saveFileDialog1.FileName);
            childForm.MdiParent = this;
            childForm.Text = $@"Окно {++_childFormNumber}";
            childForm.Show();

        }
        private void OpenFile(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog.Filter = @"Текстовые файлы (*.HELL)|*.hell";
            if (openFileDialog.ShowDialog(this) != DialogResult.OK) return;
            var fileName = openFileDialog.FileName;

            var forms = MdiChildren;
            if (forms.Any(form => form.Text == Path.GetFileName(openFileDialog.FileName)))
            {
                MessageBox.Show(@"Этот файл уже открыт!");
                return;
            }

            try
            {
                var formatter = new BinaryFormatter();
                SerializableMatrix content;

                using (var fs = new FileStream(fileName, FileMode.OpenOrCreate))
                {
                    content = (SerializableMatrix)formatter.Deserialize(fs);
                }

                var formTable = new FormTable(fileName, content)
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
        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            saveFileDialog.Filter = @"Файлы Exel (*.xlsx)|*.xlsx|Файлы Word (*.docx)|*.docx";
            if (saveFileDialog.ShowDialog(this) != DialogResult.OK) return;
            var fileName = saveFileDialog.FileName;
            if (fileName.Split('.')[fileName.Split('.').Length-1]=="docx")
                SaveWord(fileName);
            else
                SaveExel(fileName);
        }
        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e) => toolStrip.Visible = toolBarToolStripMenuItem.Checked;
        private void SaveWord(string fileName)
        {
            var forms = MdiChildren;
            if (forms.Length == 0)
            {
                MessageBox.Show(@"Нечего сохранять!", @"", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (ActiveMdiChild == null)
            {
                MessageBox.Show(@"Нет выбранной таблицы!", @"", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var applicationWord = new Application();
            var documentWord = applicationWord.Documents.Add();
            var wordRange = documentWord.Range();
            var formTable = (FormTable)ActiveMdiChild;
            var wordTable = documentWord.Tables.Add(wordRange, formTable.dataGridView.RowCount, formTable.dataGridView.ColumnCount);
            wordTable.set_Style("Сетка таблицы");
            for (var i = 0; i < formTable.dataGridView.ColumnCount; i++)
                for (var j = 0; j < formTable.dataGridView.RowCount; j++)
                    if (formTable.dataGridView[i, j].Value != null)
                        wordTable.Cell(j + 1, i + 1).Range.Text = formTable.dataGridView[i, j].Value.ToString();
            documentWord.SaveAs2(fileName);
            documentWord.Close();
            applicationWord.Quit();
            MessageBox.Show(@"Готово!", @"", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SaveExel(string fileName)
        {
            if (ActiveMdiChild == null)
            {
                MessageBox.Show(@"Нечего сохранять!", @"", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var applicationExcel = new Microsoft.Office.Interop.Excel.Application();
            var workbook = applicationExcel.Workbooks.Add();
            var form = (FormTable)ActiveMdiChild;
            var name = form.Text.Split('.')[0];
            var worksheet = (Worksheet)workbook.Worksheets[1];
            worksheet.Name = name;


            for (var i = 0; i < form.dataGridView.ColumnCount; i++)
            {
                for (var j = 0; j < form.dataGridView.RowCount; j++)
                {
                    if (form.dataGridView[i, j].Value != null)
                        worksheet.Cells[j + 1, i + 1] = form.dataGridView[i, j].Value.ToString();
                }
            }
            workbook.SaveAs(fileName);
            workbook.Close();
            applicationExcel.Quit();
            MessageBox.Show(@"Готово!", @"", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void exelToolStripButton_Click(object sender, EventArgs e)
        {

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            saveFileDialog.Filter = @"Файлы Exel (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog(this) != DialogResult.OK) return;
            var fileName = saveFileDialog.FileName;
            SaveExel(fileName);
        }
        private void wordToolStripButton_Click(object sender, EventArgs e)
        {

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            saveFileDialog.Filter = @"Файлы Exel (*.docx)|*.docx";
            if (saveFileDialog.ShowDialog(this) != DialogResult.OK) return;
            SaveWord(saveFileDialog.FileName);
        }
        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if (ActiveMdiChild == null)
            {
                MessageBox.Show(@"Нечего сохранять!",@"",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            ((FormTable)ActiveMdiChild).SaveInFile();
            MessageBox.Show(@"Сохранено!", @"", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {

            LayoutMdi(MdiLayout.TileHorizontal);
        }
        private void словаToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (ActiveMdiChild == null)
            {
                MessageBox.Show(@"Нет выбранной формы!", @"", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var activeForm = (FormTable)ActiveMdiChild;

            try
            {
                var form = new Form1();
                if (form.ShowDialog() == DialogResult.OK)
                {
                    var rowAndColumn = form.Res.Split(' ');
                    if (rowAndColumn[0][0]=='R' && rowAndColumn[1][0] == 'C')
                    {
                        rowAndColumn[0] = rowAndColumn[0].TrimStart('R');
                        rowAndColumn[1] = rowAndColumn[1].TrimStart('C');
                    }
                    else
                    {
                        var s = rowAndColumn[1];
                        rowAndColumn[1] = rowAndColumn[0].TrimStart('C');
                        rowAndColumn[0] = s.TrimStart('R');
                    }
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
                        bool check = false;
                        string result = "";
                        foreach (DataGridViewCell Cell in ActiveForm.dataGridView.SelectedCells)
                            if (Cell.Value != null)
                                foreach (string word in Cell.Value.ToString().Split(' ', ',', '!', '?', '.'))
                                    if (word[0] == 'а' || word[0] == 'А') { result += word + ' '; check = true; }
                        if (check)
                        {
                            ActiveForm.dataGridView[Convert.ToInt32(rowAndColumn[1]) - 1, Convert.ToInt32(rowAndColumn[0]) - 1].Value = result.ToString();
                            ActiveForm.dataGridView.UpdateCellValue(Convert.ToInt32(rowAndColumn[1]) - 1, Convert.ToInt32(rowAndColumn[0]) - 1);
                        }
                        else MessageBox.Show("В выбранных ячейках нет слов на букву А!");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Неверно указано значение ячейки!");
                    }
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                MessageBox.Show(@"Неверно указано значение ячейки в textBox!", @"", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void среднееЗначениеToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (ActiveMdiChild == null)
            {
                MessageBox.Show(@"Нет выбранной формы!");
                return;
            }
            var activeForm = (FormTable)ActiveMdiChild;
            float result = 0;
            bool check = false;
            foreach (DataGridViewCell Cell in activeForm.dataGridView.SelectedCells)
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

        private void about_Click(object sender, EventArgs e)
        {
            var form = new ScreenSaver();
            form.Click += Form_Click;
            form.ShowDialog();
        }

        private static void Form_Click(object sender, EventArgs e)
        {
            var screenSaver = (ScreenSaver)sender;
            screenSaver.Close();
        }

    }
}
