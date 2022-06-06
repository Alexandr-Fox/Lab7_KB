using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lab7
{
    [Serializable]
    public class MyGrid : DataGridView

    {

        private Image _backgroundPic;



        //[Browsable(true)]

        public override Image BackgroundImage

        {

            get { return _backgroundPic; }

            set { _backgroundPic = value; }

        }

        
        protected override CreateParams CreateParams
        {
            get
            {
                if (Environment.OSVersion.Version.Major >= 6)
                {
                    CreateParams cp = base.CreateParams;
                    cp.ExStyle |= 0x02000000;
                    return cp;
                }
                else
                {
                    return base.CreateParams;
                }
            }
        }
        protected override void PaintBackground(Graphics graphics, Rectangle clipBounds, Rectangle gridBounds)

        {

            base.PaintBackground(graphics, clipBounds, gridBounds);
            Rectangle rectSource = new Rectangle(this.Location.X, this.Location.Y, this.Width, this.Height);
            Rectangle rectDest = new Rectangle(0, 0, rectSource.Width, rectSource.Height);

            Bitmap b = new Bitmap(Parent.ClientRectangle.Width, Parent.ClientRectangle.Height);
            Graphics.FromImage(b).DrawImage(this.Parent.BackgroundImage, Parent.ClientRectangle);


            graphics.DrawImage(b, rectDest, rectSource, GraphicsUnit.Pixel);
            SetCellsTransparent();
            /*
            base.PaintBackground(graphics, clipBounds, gridBounds);



            if (((BackgroundImage != null)))

            {

                graphics.FillRectangle(Brushes.White, gridBounds);

                graphics.DrawImage(BackgroundImage, gridBounds);

            }
            */

        }

        

        //Make BackgroundImage can be seen in all cells

        public void SetCellsTransparent()

        {

            EnableHeadersVisualStyles = false;

            Font  = new Font("Segoe UI", 9.75F, FontStyle.Bold);

            ColumnHeadersDefaultCellStyle.BackColor = Color.Transparent;
            ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            RowHeadersDefaultCellStyle.BackColor = Color.Transparent;
            RowHeadersDefaultCellStyle.ForeColor = Color.White;
            RowHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            ForeColor = Color.White;
            

            foreach (DataGridViewColumn col in Columns)

            {

                col.DefaultCellStyle.BackColor = Color.Transparent;
                
            }
            for(int i = 0; i < ColumnCount; i++)
                for (int j = 0; j < RowCount; j++)
                {
                    this[i, j].Style.SelectionBackColor = Color.BlueViolet;
                    this[i, j].Style.SelectionForeColor = Color.White;
                }
        }

    }
}
