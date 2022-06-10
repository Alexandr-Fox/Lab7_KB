using System;
using System.Windows.Forms;

namespace Lab7
{
    public partial class Form1 : Form
    {
        public string Res { get => textBox1.Text; }
        public  Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
