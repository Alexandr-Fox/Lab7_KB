using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lab7
{
    public partial class ScreenSaver : Form
    {

        public ScreenSaver()
        {
            InitializeComponent();
        }
        private void TimerStart_Tick(object sender, EventArgs e)
        {
            timerStart.Enabled = false;
            this.Visible = false;
            MainForm formMain = new MainForm();
            formMain.ShowDialog();
            Application.Exit();
        }

    }
}
