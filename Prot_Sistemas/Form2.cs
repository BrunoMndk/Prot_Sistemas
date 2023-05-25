using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
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
            for (int x = 0; x < 10; x++)
            {
                if (Analogico_A_forms2[x] != null)
                {
                    comboBox4.Items.Add(Analogico_A_forms2[x]);
                    comboBox5.Items.Add(Analogico_A_forms2[x]);
                    comboBox6.Items.Add(Analogico_A_forms2[x]);
                    
                }
                if (Analogico_V_forms2[x] != null)
                {
                    comboBox1.Items.Add(Analogico_V_forms2[x]);
                    comboBox2.Items.Add(Analogico_V_forms2[x]);
                    comboBox3.Items.Add(Analogico_V_forms2[x]);
                    
                }
            }
            comboBox4.SelectedItem = comboBox4.Items[0];
            comboBox5.SelectedItem = comboBox5.Items[1];
            comboBox6.SelectedItem = comboBox6.Items[2];
            comboBox1.SelectedItem = comboBox1.Items[0];
            comboBox2.SelectedItem = comboBox2.Items[1];
            comboBox3.SelectedItem = comboBox3.Items[2];
            textBox1.Text = "" + RTP_P;
            textBox2.Text = "" + RTP_S;
            textBox3.Text = "" + RTC_P;
            textBox4.Text = "" + RTC_S;
            RT();
        }
        public double F2_RTP = 0;
        public double F2_RTC = 0;
        void RT()
        {
            F2_RTP = Convert.ToDouble(textBox1.Text.ToString()) / Convert.ToDouble(textBox2.Text.ToString());
            F2_RTC = Convert.ToDouble(textBox3.Text.ToString()) / Convert.ToDouble(textBox4.Text.ToString());
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
        public string[] Analogico_A_forms2;
        public string[] Analogico_V_forms2;
        public string Tensao_VA;
        public string Tensao_VB;
        public string Tensao_VC;
        public string Corrente_IA;
        public string Corrente_IB;
        public string Corrente_IC;
        public string RTP_P;
        public string RTP_S;
        public string RTC_P;
        public string RTC_S;
        public string ZL1_re;
        public string ZL1_im;
        public string ZL0_re;
        public string ZL0_im;
        public string Lengh_Line;

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Tensao_VA = comboBox1.Text.ToString();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Tensao_VB = comboBox2.Text.ToString();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Tensao_VC = comboBox3.Text.ToString();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Corrente_IA = comboBox4.Text.ToString();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            Corrente_IB = comboBox5.Text.ToString();
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            Corrente_IC = comboBox6.Text.ToString();
        }
    }
}
