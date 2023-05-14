using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Prot_Sistemas
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();            
        }        
        private void Form2_Load(object sender, EventArgs e)
        {
            draw_design();
        }
        void draw_design()
        {

            float x0 = (pictureBox1.Size.Width / 2);
            float y0 = (pictureBox1.Height / 2);
            float x1 = (pictureBox1.Size.Width / 2);
            float y2 = (pictureBox1.Size.Height / 2);

            Bitmap bmp = new Bitmap(pictureBox1.Size.Width, pictureBox1.Size.Height, PixelFormat.Format32bppArgb);
            Graphics graph1 = Graphics.FromImage(bmp);
            try
            {
                

            }
            finally
            {
            }
        }
    }
}
