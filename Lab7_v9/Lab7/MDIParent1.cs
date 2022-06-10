using System;
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
            Visible = false;
            var formMain = new MainForm();
            formMain.ShowDialog();
            Application.Exit();
        }

    }
}
