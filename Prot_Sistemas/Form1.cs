using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Numerics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;
using System.Xml.Linq;

namespace Prot_Sistemas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            polarization_calc();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != ','))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        Complex LT_Z1 = new Complex(0, 0);
        Complex LT_Z0 = new Complex(0, 0);
        Complex k0 = 0;
        Complex Va = 0;
        Complex Vb = 0;
        Complex Vc = 0;
        Complex Ia = 0;
        Complex Ib = 0;
        Complex Ic = 0;
        Complex In = 0;
        void polarization_calc()
        {
            double ang_LT1 = Convert.ToDouble(Im_seq_pos_TB.Text);
            double ang_LT0 = Convert.ToDouble(Im_seq_neg_TB.Text);
            ang_LT1 = ang_LT1 * Math.PI / 180;
            ang_LT0 = ang_LT0 * Math.PI / 180;
            LT_Z1 = new Complex(Convert.ToDouble(RE_seq_pos_TB.Text) * Math.Cos(ang_LT1), Convert.ToDouble(RE_seq_pos_TB.Text) * Math.Sin(ang_LT1));
            LT_Z0 = new Complex(Convert.ToDouble(Re_seq_neg_TB.Text) * Math.Cos(ang_LT0), Convert.ToDouble(Re_seq_neg_TB.Text) * Math.Sin(ang_LT0));
            k0 = (LT_Z0 - LT_Z1) / (3 * LT_Z1);
            Complex Zr = Convert.ToDouble(textBox1.Text)*LT_Z1;
            Va = new Complex(Convert.ToDouble(textBox2.Text) * Math.Cos(Convert.ToDouble(textBox3.Text) * Math.PI / 180), Convert.ToDouble(textBox2.Text) * Math.Sin(Convert.ToDouble(textBox3.Text) * Math.PI / 180))*1000;
            Vb = new Complex(Convert.ToDouble(textBox4.Text) * Math.Cos(Convert.ToDouble(textBox5.Text) * Math.PI / 180), Convert.ToDouble(textBox4.Text) * Math.Sin(Convert.ToDouble(textBox5.Text) * Math.PI / 180))*1000;
            Vc = new Complex(Convert.ToDouble(textBox6.Text) * Math.Cos(Convert.ToDouble(textBox7.Text) * Math.PI / 180), Convert.ToDouble(textBox6.Text) * Math.Sin(Convert.ToDouble(textBox7.Text) * Math.PI / 180)) * 1000;
            Ia = new Complex(Convert.ToDouble(textBox8.Text) * Math.Cos(Convert.ToDouble(textBox9.Text) * Math.PI / 180), Convert.ToDouble(textBox8.Text) * Math.Sin(Convert.ToDouble(textBox9.Text) * Math.PI / 180));
            Ib = new Complex(Convert.ToDouble(textBox10.Text) * Math.Cos(Convert.ToDouble(textBox11.Text) * Math.PI / 180), Convert.ToDouble(textBox10.Text) * Math.Sin(Convert.ToDouble(textBox11.Text) * Math.PI / 180));
            Ic = new Complex(Convert.ToDouble(textBox12.Text) * Math.Cos(Convert.ToDouble(textBox13.Text) * Math.PI / 180), Convert.ToDouble(textBox12.Text) * Math.Sin(Convert.ToDouble(textBox13.Text) * Math.PI / 180));
            In = Ia + Ib + Ic;
            label19.Text = "" + Math.Round(k0.Real,3)+"+i"+Math.Round(k0.Imaginary,3);
            

            Complex IaG = Ia + (k0 * In);
            Complex Sop_AG = (Zr * IaG) - Va;
            Complex Spol_AG = Va;
            Complex[] Zm =  { 0,Va / IaG};
            chart1.Series[1].Points.AddXY(0, 0);
            chart1.Series[1].Points.AddXY(Zm[1].Real, Zm[1].Imaginary);
            Complex[] Sop = { Zm[1], (Zm[1] + (Sop_AG / IaG)) };
            chart1.Series[2].Points.AddXY(Sop[0].Real, Sop[0].Imaginary);
            chart1.Series[2].Points.AddXY(Sop[1].Real, Sop[1].Imaginary);
            chart1.Series[3].Points.AddXY(0, 0);
            chart1.Series[3].Points.AddXY(LT_Z1.Real, LT_Z1.Imaginary);
            double radius = Complex.Abs(Zr) / 2;
            for (int k = 0; k <= 1000; k++)
            {
                double x = (radius * Math.Cos(Zr.Phase)) + (radius) * Math.Cos(k * 2 * Math.PI / 1000);
                double y = (radius * Math.Sin(Zr.Phase)) + radius * Math.Sin(k * 2 * Math.PI / 1000);
                chart1.Series[0].Points.AddXY(x, y);
            }

        }
    }
}
