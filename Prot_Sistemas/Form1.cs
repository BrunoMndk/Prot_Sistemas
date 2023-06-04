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
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Runtime.CompilerServices;
using System.Drawing.Imaging;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Prot_Sistemas
{
    public partial class Form1 : Form
    {
        public string[] analogic_A = new string[100];
        public string[] analogic_V = new string[100];
        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true;
            polarization_calc();
            chart1.MouseWheel += chart1_MouseWheel;
            chart2.MouseWheel += chart2_MouseWheel;
            chart1.MouseMove += chart1_MouseMove;
            chart1.MouseDown += chart1_MouseDown;
            chart2.MouseMove += chart2_MouseMove;
            chart2.MouseDown += chart2_MouseDown;

            chart1.ChartAreas[0].AxisX.Maximum = 50;
            chart1.ChartAreas[0].AxisX.Minimum = -30;
            chart1.ChartAreas[0].AxisY.Maximum = 50;
            chart1.ChartAreas[0].AxisY.Minimum = -30;
            chart2.ChartAreas[0].AxisX.Maximum = 50;
            chart2.ChartAreas[0].AxisX.Minimum = -30;
            chart2.ChartAreas[0].AxisY.Maximum = 50;
            chart2.ChartAreas[0].AxisY.Minimum = -30;
            comboBox1.SelectedItem = comboBox1.Items[0];

        }
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                polarization_calc();                
            }
            
        }
        double mDown_x = double.NaN;
        double mDown_y = double.NaN;
        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            Axis ax = chart1.ChartAreas[0].AxisX;
            mDown_x = Math.Round(ax.PixelPositionToValue(e.Location.X),0);
            Axis ay = chart1.ChartAreas[0].AxisY;
            mDown_y = Math.Round(ay.PixelPositionToValue(e.Location.Y), 0);
            if(e.Button.HasFlag(MouseButtons.Right))
            {
                chart1.ChartAreas[0].AxisX.Maximum = 50;
                chart1.ChartAreas[0].AxisX.Minimum = -30;
                chart1.ChartAreas[0].AxisY.Maximum = 50;
                chart1.ChartAreas[0].AxisY.Minimum = -30;
            }
        }
        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (!e.Button.HasFlag(MouseButtons.Left)) return;
                Axis ax = chart1.ChartAreas[0].AxisX;
                Axis ay = chart1.ChartAreas[0].AxisY;
                double range_x = ax.Maximum - ax.Minimum;
                double range_y = ay.Maximum - ay.Minimum;
                double xv = ax.PixelPositionToValue(e.Location.X);
                double yv = ay.PixelPositionToValue(e.Location.Y);
                ax.Minimum -= Math.Round(xv - mDown_x, 0);
                ax.Maximum = Math.Round(ax.Minimum + range_x, 0);

                ay.Minimum -= Math.Round(yv - mDown_y, 0);
                ay.Maximum = Math.Round(ay.Minimum + range_y, 0);
            }
            catch { }
        }
        private void chart2_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                Axis ax = chart2.ChartAreas[0].AxisX;
                mDown_x = Math.Round(ax.PixelPositionToValue(e.Location.X), 0);
                Axis ay = chart2.ChartAreas[0].AxisY;
                mDown_y = Math.Round(ay.PixelPositionToValue(e.Location.Y), 0);
                if (e.Button.HasFlag(MouseButtons.Right))
                {
                    chart2.ChartAreas[0].AxisX.Maximum = 50;
                    chart2.ChartAreas[0].AxisX.Minimum = -30;
                    chart2.ChartAreas[0].AxisY.Maximum = 50;
                    chart2.ChartAreas[0].AxisY.Minimum = -30;
                }
            }
            catch
            { }
        }
        private void chart2_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (!e.Button.HasFlag(MouseButtons.Left)) return;
                Axis ax = chart2.ChartAreas[0].AxisX;
                Axis ay = chart2.ChartAreas[0].AxisY;
                double range_x = ax.Maximum - ax.Minimum;
                double range_y = ay.Maximum - ay.Minimum;
                double xv = ax.PixelPositionToValue(e.Location.X);
                double yv = ay.PixelPositionToValue(e.Location.Y);
                ax.Minimum -= Math.Round(xv - mDown_x, 0);
                ax.Maximum = Math.Round(ax.Minimum + range_x, 0);

                ay.Minimum -= Math.Round(yv - mDown_y, 0);
                ay.Maximum = Math.Round(ay.Minimum + range_y, 0);
            }
            catch
            { }
        }
        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            var chart = (Chart)sender;
            var xAxis = chart.ChartAreas[0].AxisX;
            var yAxis = chart.ChartAreas[0].AxisY;
            double cursor = (e.X);
            
                try
                {                    
                    if (chart1.ChartAreas[0].AxisX.Maximum > 20)
                    {
                        if (e.Delta < 0) // Scrolled down.
                        {
                            chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Maximum + 10;
                            chart1.ChartAreas[0].AxisX.Minimum = chart1.ChartAreas[0].AxisX.Minimum - 10;
                            chart1.ChartAreas[0].AxisY.Maximum = chart1.ChartAreas[0].AxisY.Maximum + 10;
                            chart1.ChartAreas[0].AxisY.Minimum = chart1.ChartAreas[0].AxisY.Minimum - 10;

                        }
                    }
                    if (chart1.ChartAreas[0].AxisX.Maximum >= 20)
                    {
                        if (e.Delta > 0) // Scrolled up.
                        {
                            chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Maximum - 10;
                            chart1.ChartAreas[0].AxisX.Minimum = chart1.ChartAreas[0].AxisX.Minimum + 10;
                            chart1.ChartAreas[0].AxisY.Maximum = chart1.ChartAreas[0].AxisY.Maximum - 10;
                            chart1.ChartAreas[0].AxisY.Minimum = chart1.ChartAreas[0].AxisY.Minimum + 10;
                        }
                    }
                }
                catch { }
            
        }
        private void chart2_MouseWheel(object sender, MouseEventArgs e)
        {
            var chart = (Chart)sender;
            var xAxis = chart.ChartAreas[0].AxisX;
            var yAxis = chart.ChartAreas[0].AxisY;
            double cursor = (e.X);
            try
            {
                
                if (chart2.ChartAreas[0].AxisX.Maximum > 20)
                {
                    if (e.Delta < 0) // Scrolled down.
                    {
                        chart2.ChartAreas[0].AxisX.Maximum = chart2.ChartAreas[0].AxisX.Maximum + 10;
                        chart2.ChartAreas[0].AxisX.Minimum = chart2.ChartAreas[0].AxisX.Minimum - 10;
                        chart2.ChartAreas[0].AxisY.Maximum = chart2.ChartAreas[0].AxisY.Maximum + 10;
                        chart2.ChartAreas[0].AxisY.Minimum = chart2.ChartAreas[0].AxisY.Minimum - 10;

                    }
                }
                if (chart2.ChartAreas[0].AxisX.Maximum >= 10)
                {
                    if (e.Delta > 0) // Scrolled up.
                    {
                        chart2.ChartAreas[0].AxisX.Maximum = chart2.ChartAreas[0].AxisX.Maximum - 10;
                        chart2.ChartAreas[0].AxisX.Minimum = chart2.ChartAreas[0].AxisX.Minimum + 10;
                        chart2.ChartAreas[0].AxisY.Maximum = chart2.ChartAreas[0].AxisY.Maximum - 10;
                        chart2.ChartAreas[0].AxisY.Minimum = chart2.ChartAreas[0].AxisY.Minimum + 10;
                    }
                }
            }
            catch { }
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
        Complex Va_mem = 0;
        Complex Vb_mem = 0;
        Complex Vc_mem = 0;
        void polarization_calc()
        {
            for (int x = 0; x < 50; x++)
            {
                if(x<6)
                    chart4.Series[x].Points.Clear();
                if(x<14)
                    chart1.Series[x].Points.Clear();
                if(x<25)
                    chart2.Series[x].Points.Clear();
            }

            Complex k_pol = 0;
            double kpol_abs = Convert.ToDouble(textBox20.Text);            
            Complex a = new Complex(1 * Math.Cos(120 * Math.PI / 180), 1 * Math.Sin(120 * Math.PI / 180));
            double alpha = Convert.ToDouble(textBox21.Text);
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
            Va_mem = new Complex(Convert.ToDouble(textBox14.Text) * Math.Cos(Convert.ToDouble(textBox15.Text) * Math.PI / 180), Convert.ToDouble(textBox14.Text) * Math.Sin(Convert.ToDouble(textBox15.Text) * Math.PI / 180)) * 1000;
            Vb_mem = new Complex(Convert.ToDouble(textBox16.Text) * Math.Cos(Convert.ToDouble(textBox17.Text) * Math.PI / 180), Convert.ToDouble(textBox16.Text) * Math.Sin(Convert.ToDouble(textBox17.Text) * Math.PI / 180)) * 1000;
            Vc_mem = new Complex(Convert.ToDouble(textBox18.Text) * Math.Cos(Convert.ToDouble(textBox19.Text) * Math.PI / 180), Convert.ToDouble(textBox18.Text) * Math.Sin(Convert.ToDouble(textBox19.Text) * Math.PI / 180)) * 1000;
            double Sbase = Convert.ToDouble(textBox22.Text);
            double Vbase = Convert.ToDouble(textBox23.Text);
            double Ibase = Sbase / (Math.Sqrt(3) * Vbase);

            chart4.Series[0].Points.AddXY(0, 0);
            chart4.Series[1].Points.AddXY(0, 0);
            chart4.Series[2].Points.AddXY(0, 0);

            chart4.Series[3].Points.AddXY(0, 0);
            chart4.Series[4].Points.AddXY(0, 0);
            chart4.Series[5].Points.AddXY(0, 0);

            chart4.Series[0].Points.AddXY(-1 * (Va.Phase) * 180 / Math.PI, (Va.Magnitude / Vbase));
            chart4.Series[1].Points.AddXY(-1 * (Vb.Phase) * 180 / Math.PI, (Vb.Magnitude / Vbase));
            chart4.Series[2].Points.AddXY(-1 * (Vc.Phase) * 180 / Math.PI, (Vc.Magnitude / Vbase));

            chart4.Series[3].Points.AddXY(-1 * (Ia.Phase) * 180 / Math.PI, (Ia.Magnitude / Ibase));
            chart4.Series[4].Points.AddXY(-1 * (Ib.Phase) * 180 / Math.PI, (Ib.Magnitude / Ibase));
            chart4.Series[5].Points.AddXY(-1 * (Ic.Phase) * 180 / Math.PI, (Ic.Magnitude / Ibase));


            Complex a1 = Complex.FromPolarCoordinates(1, 120 * (Math.PI / 180));
            Complex a2 = Complex.FromPolarCoordinates(1, 240 * (Math.PI / 180));
            Complex Va_0 = (Va + Vb + Vc) / 3;
            Complex Va_1 = ((a1 * Vb) + (a2 * Vc) + Va)/3;
            Complex Va_2 = ((a2 * Vb) + (a1 * Vc) + Va) / 3;
            Complex Vb_0 = Va_0;
            Complex Vb_1 = a2 * Va_1;
            Complex Vb_2 = a1 * Va_2;
            Complex Vc_0 = Va_0;
            Complex Vc_1 = a1 * Va_1;
            Complex Vc_2 = a2 * Va_2;
            Complex Vab_mem = Va_mem - Vb_mem;
            Complex Vbc_mem = Vb_mem - Vc_mem;
            Complex Vca_mem = Vc_mem - Va_mem;
            Complex Vab = Va - Vb;            
            Complex Vbc = Vb - Vc;
            Complex Vca = Vc - Va;
            Ia = new Complex(Convert.ToDouble(textBox8.Text) * Math.Cos(Convert.ToDouble(textBox9.Text) * Math.PI / 180), Convert.ToDouble(textBox8.Text) * Math.Sin(Convert.ToDouble(textBox9.Text) * Math.PI / 180));
            Ib = new Complex(Convert.ToDouble(textBox10.Text) * Math.Cos(Convert.ToDouble(textBox11.Text) * Math.PI / 180), Convert.ToDouble(textBox10.Text) * Math.Sin(Convert.ToDouble(textBox11.Text) * Math.PI / 180));
            Ic = new Complex(Convert.ToDouble(textBox12.Text) * Math.Cos(Convert.ToDouble(textBox13.Text) * Math.PI / 180), Convert.ToDouble(textBox12.Text) * Math.Sin(Convert.ToDouble(textBox13.Text) * Math.PI / 180));
            Complex Ia_0 = (Ia + Ib + Ic) / 3;
            Complex Ia_1 = ((a1 * Ib) + (a2 * Ic) + Ia) / 3;
            Complex Ia_2 = ((a2 * Ib) + (a1 * Ic) + Ia) / 3;
            Complex Ib_0 = Ia_0;
            Complex Ib_1 = a2 * Ia_1;
            Complex Ib_2 = a1 * Ia_2;
            Complex Ic_0 = Ia_0;
            Complex Ic_1 = a1 * Ia_1;
            Complex Ic_2 = a2 * Va_2;
            Complex Iab = Ia - Ib;
            Complex Ibc = Ib - Ic;
            Complex Ica = Ic - Ia;
            Complex IaG = Ia + (k0 * In);
            Complex IbG = Ib + (k0 * In);
            Complex IcG = Ic + (k0 * In);
            In = Ia + Ib + Ic;
            
            label19.Text = "" + Math.Round(k0.Real,3)+"+i"+Math.Round(k0.Imaginary,3);

            //Plot LT
            chart1.Series["LT"].Points.AddXY(0, 0);
            chart1.Series["LT"].Points.AddXY(LT_Z1.Real, LT_Z1.Imaginary);
            chart2.Series["LT"].Points.AddXY(0, 0);
            chart2.Series["LT"].Points.AddXY(LT_Z1.Real, LT_Z1.Imaginary);
            //Plot MHO
            double radius = Complex.Abs(Zr) / 2;
            for (int k = 0; k <= 1000; k++)
            {
                double x = (radius * Math.Cos(Zr.Phase)) + (radius) * Math.Cos(k * 2 * Math.PI / 1000);
                double y = (radius * Math.Sin(Zr.Phase)) + radius * Math.Sin(k * 2 * Math.PI / 1000);
                chart1.Series[0].Points.AddXY(x, y);
            }
            // loop AG
            //std mho            
            Complex[] Zm_AG =  { 0,Va / IaG};
            chart1.Series["Zm_AG"].Points.AddXY(0, 0);
            chart1.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
            Complex[] Sop_AG = { Zm_AG[1], (Zm_AG[1] + (((Zr * IaG) - Va) / IaG)) };
            chart1.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
            chart1.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
            // loop BG            
            Complex[] Zm_BG = { 0, Vb / IbG };
            chart1.Series["Zm_BG"].Points.AddXY(0, 0);
            chart1.Series["Zm_BG"].Points.AddXY(Zm_BG[1].Real, Zm_BG[1].Imaginary);
            Complex[] Sop_BG = { Zm_BG[1], (Zm_BG[1] + (((Zr * IbG) - Vb) / IbG)) };
            chart1.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
            chart1.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
            // loop CG            
            Complex[] Zm_CG = { 0, Vc / IbG };
            chart1.Series["Zm_CG"].Points.AddXY(0, 0);
            chart1.Series["Zm_CG"].Points.AddXY(Zm_CG[1].Real, Zm_CG[1].Imaginary);
            Complex[] Sop_CG = { Zm_CG[1], (Zm_CG[1] + (((Zr * IcG) - Vc) / IcG)) };
            chart1.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
            chart1.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
            // loop AB            
            Complex[] Zm_AB = { 0, Vab / Iab };
            chart1.Series["Zm_AB"].Points.AddXY(0, 0);
            chart1.Series["Zm_AB"].Points.AddXY(Zm_AB[1].Real, Zm_AB[1].Imaginary);
            Complex[] Sop_AB = { Zm_AB[1], (Zm_AB[1] + (((Zr * Iab) - Vab) / Iab)) };
            chart1.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
            chart1.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
            // loop BC            
            Complex[] Zm_BC = { 0, Vbc / Ibc };
            chart1.Series["Zm_BC"].Points.AddXY(0, 0);
            chart1.Series["Zm_BC"].Points.AddXY(Zm_BC[1].Real, Zm_BC[1].Imaginary);
            Complex[] Sop_BC = { Zm_BC[1], (Zm_BC[1] + (((Zr * Ibc) - Vbc) / Ibc)) };
            chart1.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
            chart1.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
            // loop CA            
            Complex[] Zm_CA = { 0, Vca / Ica };
            chart1.Series["Zm_CA"].Points.AddXY(0, 0);
            chart1.Series["Zm_CA"].Points.AddXY(Zm_CA[1].Real, Zm_CA[1].Imaginary);
            Complex[] Sop_CA = { Zm_CA[1], (Zm_CA[1] + (((Zr * Ica) - Vca) / Ica)) };
            chart1.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
            chart1.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
            // Dynamic Polarization loops
            switch (comboBox1.SelectedIndex)
            {
                case 0: // Dual
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
                        if (radioButton1.Checked == true)
                        {                            
                            //dual mho AB                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_dual_AB = Vab + (k_pol * Vc);
                            Complex Zpol_dual_AB = k_pol * Vc / Iab;
                            Complex[] Zs_dual_AB = { -Zpol_dual_AB, Zm_AB[1] };
                            double r_dual_AB = (Complex.Abs(Zr - Zs_dual_AB[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_AB[1].Real + ((Zr * Iab - Vab) / Iab).Real + Zs_dual_AB[0].Real) / 2)) + r_dual_AB * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_AB[1].Imaginary + ((Zr * Iab - Vab) / Iab).Imaginary + Zs_dual_AB[0].Imaginary) / 2)) + r_dual_AB * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AB"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_AB"].Points.AddXY(0, 0);
                            chart2.Series["Zs_AB"].Points.AddXY(Zs_dual_AB[0].Real, Zs_dual_AB[0].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_dual_AB[0].Real, Zs_dual_AB[0].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_dual_AB[1].Real, Zs_dual_AB[1].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
                        }
                        if (radioButton2.Checked == true)
                        {
                            //dual mho BC                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_dual_BC = Vbc + (k_pol * Va);
                            Complex Zpol_dual_BC = k_pol * Va / Ibc;
                            Complex[] Zs_dual_BC = { -Zpol_dual_BC, Zm_BC[1] };
                            double r_dual_BC = (Complex.Abs(Zr - Zs_dual_BC[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_BC[1].Real + ((Zr * Ibc - Vbc) / Ibc).Real + Zs_dual_BC[0].Real) / 2)) + r_dual_BC * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_BC[1].Imaginary + ((Zr * Ibc - Vbc) / Ibc).Imaginary + Zs_dual_BC[0].Imaginary) / 2)) + r_dual_BC * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BC"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_BC"].Points.AddXY(0, 0);
                            chart2.Series["Zs_BC"].Points.AddXY(Zs_dual_BC[0].Real, Zs_dual_BC[0].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_dual_BC[0].Real, Zs_dual_BC[0].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_dual_BC[1].Real, Zs_dual_BC[1].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
                        }
                        if (radioButton3.Checked == true)
                        {
                            //dual mho CA                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_dual_CA = Vca + (k_pol * Vb);
                            Complex Zpol_dual_CA = k_pol * Vb / Ica;
                            Complex[] Zs_dual_CA = { -Zpol_dual_CA, Zm_CA[1] };
                            double r_dual_CA = (Complex.Abs(Zr - Zs_dual_CA[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_CA[1].Real + ((Zr * Ica - Vca) / Ica).Real + Zs_dual_CA[0].Real) / 2)) + r_dual_CA * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_CA[1].Imaginary + ((Zr * Ica - Vca) / Ica).Imaginary + Zs_dual_CA[0].Imaginary) / 2)) + r_dual_CA * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CA"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_CA"].Points.AddXY(0, 0);
                            chart2.Series["Zs_CA"].Points.AddXY(Zs_dual_CA[0].Real, Zs_dual_CA[0].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_dual_CA[0].Real, Zs_dual_CA[0].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_dual_CA[1].Real, Zs_dual_CA[1].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
                        }
                        if (radioButton4.Checked == true)
                        {
                            //dual mho AG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_dual_AG = Va + (k_pol * Vbc);
                            Complex Zpol_dual_AG = k_pol * Vbc / IaG;
                            Complex[] Zs_dual_AG = { -Zpol_dual_AG, Zm_AG[1] };
                            radius = (Complex.Abs(Zr - Zs_dual_AG[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_AG[1].Real + ((Zr * IaG - Va) / IaG).Real + Zs_dual_AG[0].Real) / 2)) + radius * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_AG[1].Imaginary + ((Zr * IaG - Va) / IaG).Imaginary + Zs_dual_AG[0].Imaginary) / 2)) + radius * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_AG"].Points.AddXY(0, 0);
                            chart2.Series["Zs_AG"].Points.AddXY(Zs_dual_AG[0].Real, Zs_dual_AG[0].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_dual_AG[0].Real, Zs_dual_AG[0].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_dual_AG[1].Real, Zs_dual_AG[1].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
                        }
                        if (radioButton5.Checked == true)
                        {
                            //dual mho BG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_dual_BG = Vb + (k_pol * Vca);
                            Complex Zpol_dual_BG = k_pol * Vca / IbG;
                            Complex[] Zs_dual_BG = { -Zpol_dual_BG, Zm_BG[1] };
                            double r_dual_BG = (Complex.Abs(Zr - Zs_dual_BG[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_BG[1].Real + ((Zr * IbG - Vb) / IbG).Real + Zs_dual_BG[0].Real) / 2)) + r_dual_BG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_BG[1].Imaginary + ((Zr * IbG - Vb) / IbG).Imaginary + Zs_dual_BG[0].Imaginary) / 2)) + r_dual_BG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_BG"].Points.AddXY(0, 0);
                            chart2.Series["Zs_BG"].Points.AddXY(Zs_dual_BG[0].Real, Zs_dual_BG[0].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_dual_BG[0].Real, Zs_dual_BG[0].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_dual_BG[1].Real, Zs_dual_BG[1].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
                        }
                        if (radioButton6.Checked == true)
                        {
                            //dual mho CG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_dual_CG = Vc + (k_pol * Vab);
                            Complex Zpol_dual_CG = k_pol * Vab / IcG;
                            Complex[] Zs_dual_CG = { -Zpol_dual_CG, Zm_CG[1] };
                            double r_dual_CG = (Complex.Abs(Zr - Zs_dual_CG[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_CG[1].Real + ((Zr * IcG - Vc) / IcG).Real + Zs_dual_CG[0].Real) / 2)) + r_dual_CG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_CG[1].Imaginary + ((Zr * IbG - Vc) / IcG).Imaginary + Zs_dual_CG[0].Imaginary) / 2)) + r_dual_CG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_CG"].Points.AddXY(0, 0);
                            chart2.Series["Zs_CG"].Points.AddXY(Zs_dual_CG[0].Real, Zs_dual_CG[0].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_dual_CG[0].Real, Zs_dual_CG[0].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_dual_CG[1].Real, Zs_dual_CG[1].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
                        }
                    }
                    break;
                case 1: // Dual alternative
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
                        if (radioButton1.Checked == true)
                        {
                            //dual mho AB                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-120 * Math.PI / 180), kpol_abs * Math.Sin(-120 * Math.PI / 180));
                            Complex Sop_dual_alt_AB = Vab + (k_pol * Vca);
                            Complex Zpol_dual_alt_AB = k_pol * Vca / Iab;
                            Complex[] Zs_dual_alt_AB = { -Zpol_dual_alt_AB, Zm_AB[1] };
                            double r_dual_alt_AB = (Complex.Abs(Zr - Zs_dual_alt_AB[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_AB[1].Real + ((Zr * Iab - Vab) / Iab).Real + Zs_dual_alt_AB[0].Real) / 2)) + r_dual_alt_AB * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_AB[1].Imaginary + ((Zr * Iab - Vab) / Iab).Imaginary + Zs_dual_alt_AB[0].Imaginary) / 2)) + r_dual_alt_AB * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AB"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_AB"].Points.AddXY(0, 0);
                            chart2.Series["Zs_AB"].Points.AddXY(Zs_dual_alt_AB[0].Real, Zs_dual_alt_AB[0].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_dual_alt_AB[0].Real, Zs_dual_alt_AB[0].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_dual_alt_AB[1].Real, Zs_dual_alt_AB[1].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
                        }
                        if (radioButton2.Checked == true)
                        {
                            //dual mho BC                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-120 * Math.PI / 180), kpol_abs * Math.Sin(-120 * Math.PI / 180));
                            Complex Sop_dual_alt_BC = Vbc + (k_pol * Vab);
                            Complex Zpol_dual_alt_BC = k_pol * Vab / Ibc;
                            Complex[] Zs_dual_alt_BC = { -Zpol_dual_alt_BC, Zm_BC[1] };
                            double r_dual_alt_BC = (Complex.Abs(Zr - Zs_dual_alt_BC[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_BC[1].Real + ((Zr * Ibc - Vbc) / Ibc).Real + Zs_dual_alt_BC[0].Real) / 2)) + r_dual_alt_BC * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_BC[1].Imaginary + ((Zr * Ibc - Vbc) / Ibc).Imaginary + Zs_dual_alt_BC[0].Imaginary) / 2)) + r_dual_alt_BC * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BC"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_BC"].Points.AddXY(0, 0);
                            chart2.Series["Zs_BC"].Points.AddXY(Zs_dual_alt_BC[0].Real, Zs_dual_alt_BC[0].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_dual_alt_BC[0].Real, Zs_dual_alt_BC[0].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_dual_alt_BC[1].Real, Zs_dual_alt_BC[1].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
                        }
                        if (radioButton3.Checked == true)
                        {
                            //dual mho CA                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-120 * Math.PI / 180), kpol_abs * Math.Sin(-120 * Math.PI / 180));
                            Complex Sop_dual_alt_CA = Vca + (k_pol * Vbc);
                            Complex Zpol_dual_alt_CA = k_pol * Vbc / Ica;
                            Complex[] Zs_dual_alt_CA = { -Zpol_dual_alt_CA, Zm_CA[1] };
                            double r_dual_alt_CA = (Complex.Abs(Zr - Zs_dual_alt_CA[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_CA[1].Real + ((Zr * Ica - Vca) / Ica).Real + Zs_dual_alt_CA[0].Real) / 2)) + r_dual_alt_CA * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_CA[1].Imaginary + ((Zr * Ica - Vca) / Ica).Imaginary + Zs_dual_alt_CA[0].Imaginary) / 2)) + r_dual_alt_CA * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CA"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_CA"].Points.AddXY(0, 0);
                            chart2.Series["Zs_CA"].Points.AddXY(Zs_dual_alt_CA[0].Real, Zs_dual_alt_CA[0].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_dual_alt_CA[0].Real, Zs_dual_alt_CA[0].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_dual_alt_CA[1].Real, Zs_dual_alt_CA[1].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
                        }
                        if (radioButton4.Checked == true)
                        {
                            //dual mho AG
                            k_pol = new Complex(kpol_abs * Math.Cos(120 * Math.PI / 180), kpol_abs * Math.Sin(120 * Math.PI / 180));
                            Complex Sop_dual_alt_AG = Va + (k_pol * Vb);
                            Complex Zpol_dual_alt_AG = k_pol * Vb / IaG;
                            Complex[] Zs_dual_alt_AG = { -Zpol_dual_alt_AG, Zm_AG[1] };
                            double r_dual_alt = (Complex.Abs(Zr - Zs_dual_alt_AG[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_AG[1].Real + ((Zr * IaG - Va) / IaG).Real + Zs_dual_alt_AG[0].Real) / 2)) + r_dual_alt * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_AG[1].Imaginary + ((Zr * IaG - Va) / IaG).Imaginary + Zs_dual_alt_AG[0].Imaginary) / 2)) + r_dual_alt * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_AG"].Points.AddXY(0, 0);
                            chart2.Series["Zs_AG"].Points.AddXY(Zs_dual_alt_AG[0].Real, Zs_dual_alt_AG[0].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_dual_alt_AG[0].Real, Zs_dual_alt_AG[0].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_dual_alt_AG[1].Real, Zs_dual_alt_AG[1].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
                        }
                        if (radioButton5.Checked == true)
                        {
                            //dual mho BG
                            k_pol = new Complex(kpol_abs * Math.Cos(120 * Math.PI / 180), kpol_abs * Math.Sin(120 * Math.PI / 180));
                            Complex Sop_dual_alt_BG = Vb + (k_pol * Vc);
                            Complex Zpol_dual_alt_BG = k_pol * Vc / IbG;
                            Complex[] Zs_dual_alt_BG = { -Zpol_dual_alt_BG, Zm_BG[1] };
                            double r_dual_alt_BG = (Complex.Abs(Zr - Zs_dual_alt_BG[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_BG[1].Real + ((Zr * IbG - Vb) / IbG).Real + Zs_dual_alt_BG[0].Real) / 2)) + r_dual_alt_BG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_BG[1].Imaginary + ((Zr * IbG - Vb) / IbG).Imaginary + Zs_dual_alt_BG[0].Imaginary) / 2)) + r_dual_alt_BG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_BG"].Points.AddXY(0, 0);
                            chart2.Series["Zs_BG"].Points.AddXY(Zs_dual_alt_BG[0].Real, Zs_dual_alt_BG[0].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_dual_alt_BG[0].Real, Zs_dual_alt_BG[0].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_dual_alt_BG[1].Real, Zs_dual_alt_BG[1].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
                        }
                        if (radioButton6.Checked == true)
                        {
                            //dual mho CG
                            k_pol = new Complex(kpol_abs * Math.Cos(120 * Math.PI / 180), kpol_abs * Math.Sin(120 * Math.PI / 180));
                            Complex Sop_dual_alt_CG = Vc + (k_pol * Va);
                            Complex Zpol_dual_alt_CG = k_pol * Va / IcG;
                            Complex[] Zs_dual_alt_CG = { -Zpol_dual_alt_CG, Zm_CG[1] };
                            double r_dual_alt_CG = (Complex.Abs(Zr - Zs_dual_alt_CG[0])) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zm_CG[1].Real + ((Zr * IcG - Vc) / IcG).Real + Zs_dual_alt_CG[0].Real) / 2)) + r_dual_alt_CG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zm_CG[1].Imaginary + ((Zr * IbG - Vc) / IcG).Imaginary + Zs_dual_alt_CG[0].Imaginary) / 2)) + r_dual_alt_CG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zs_CG"].Points.AddXY(0, 0);
                            chart2.Series["Zs_CG"].Points.AddXY(Zs_dual_alt_CG[0].Real, Zs_dual_alt_CG[0].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_dual_alt_CG[0].Real, Zs_dual_alt_CG[0].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_dual_alt_CG[1].Real, Zs_dual_alt_CG[1].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
                        }
                    }
                    break;
                case 2: // Cross polarization
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
                        if (radioButton1.Checked == true)
                        {
                            //cross mho AB                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_crz_AB = Vab + (k_pol * Vc);
                            Complex Zpol_crz_AB = k_pol * Vc / Iab;
                            Complex Zs_crz_AB = Zm_AB[1] - Zpol_crz_AB;
                            double r_crz_AB = (Complex.Abs(-Zr + Zm_AB[1] - Zpol_crz_AB)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AB[1] - Zpol_crz_AB).Real / 2)) + r_crz_AB * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AB[1] - Zpol_crz_AB).Imaginary / 2)) + r_crz_AB * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AB"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AB"].Points.AddXY(Zm_AB[1].Real, Zm_AB[1].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_crz_AB.Real, Zs_crz_AB.Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
                        }
                        if (radioButton2.Checked == true)
                        {
                            //cross mho BC                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_crz_BC = Vbc + (k_pol * Va);
                            Complex Zpol_crz_BC = k_pol * Va / Ibc;
                            Complex Zs_crz_BC = Zm_BC[1] - Zpol_crz_BC;
                            double r_crz_BC = (Complex.Abs(-Zr + Zm_BC[1] - Zpol_crz_BC)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BC[1] - Zpol_crz_BC).Real / 2)) + r_crz_BC * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BC[1] - Zpol_crz_BC).Imaginary / 2)) + r_crz_BC * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BC"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BC"].Points.AddXY(Zm_BC[1].Real, Zm_BC[1].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_crz_BC.Real, Zs_crz_BC.Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
                        }
                        if (radioButton3.Checked == true)
                        {
                            //cross mho CA
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_crz_CA = (k_pol * Vb);
                            Complex Zpol_crz_CA = k_pol * Vb / Ica;
                            Complex Zs_crz_CA = Zm_CA[1] - Zpol_crz_CA;
                            double r_crz_CA = (Complex.Abs(-Zr + Zm_CA[1] - Zpol_crz_CA)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CA[1] - Zpol_crz_CA).Real / 2)) + r_crz_CA * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CA[1] - Zpol_crz_CA).Imaginary / 2)) + r_crz_CA * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CA"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CA"].Points.AddXY(Zm_CA[1].Real, Zm_CA[1].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_crz_CA.Real, Zs_crz_CA.Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
                        }
                        if (radioButton4.Checked == true)
                        {
                            //Cross mho AG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_crz_AG = Va + (k_pol * Vbc);
                            Complex Zpol_crz_AG = k_pol * Vbc / IaG;
                            Complex Zs_crz_AG = Zm_AG[1] - Zpol_crz_AG;
                            double r_crz_AG = (Complex.Abs(-Zr + Zm_AG[1] - Zpol_crz_AG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AG[1] - Zpol_crz_AG).Real / 2)) + r_crz_AG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AG[1] - Zpol_crz_AG).Imaginary / 2)) + r_crz_AG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_crz_AG.Real, Zs_crz_AG.Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
                        }
                        if (radioButton5.Checked == true)
                        {
                            //cross mho BG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_crz_BG = Vb + (k_pol * Vca);
                            Complex Zpol_crz_BG = k_pol * Vca / IbG;
                            Complex Zs_crz_BG = Zm_BG[1] - Zpol_crz_BG;
                            double r_crz_BG = (Complex.Abs(-Zr + Zm_BG[1] - Zpol_crz_BG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BG[1] - Zpol_crz_BG).Real / 2)) + r_crz_BG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BG[1] - Zpol_crz_BG).Imaginary / 2)) + r_crz_BG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BG"].Points.AddXY(Zm_BG[1].Real, Zm_BG[1].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_crz_BG.Real, Zs_crz_BG.Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
                        }
                        if (radioButton6.Checked == true)
                        {
                            //cross mho CG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_crz_CG = Vc + (k_pol * Vab);
                            Complex Zpol_crz_CG = k_pol * Vab / IcG;
                            Complex Zs_crz_CG = Zm_CG[1] - Zpol_crz_CG;
                            double r_crz_CG = (Complex.Abs(-Zr + Zm_CG[1] - Zpol_crz_CG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CG[1] - Zpol_crz_CG).Real / 2)) + r_crz_CG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CG[1] - Zpol_crz_CG).Imaginary / 2)) + r_crz_CG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CG"].Points.AddXY(Zm_CG[1].Real, Zm_CG[1].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_crz_CG.Real, Zs_crz_CG.Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
                        }
                    }
                    break;
                case 3:// Cross polarization with voltage memory
                    {
                        textBox21.Font = textBox20.Font;
                        if (radioButton1.Checked == true)
                        {
                            //Cross with voltage memory mho AB
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex V1ab_mem = (alpha * Vb) + ((1 - alpha) * Vb_mem);
                            Complex Zpol_crz_mem_AB = k_pol * V1ab_mem / Iab;
                            Complex Zs_crz_mem_AB = Zm_AB[1] - Zpol_crz_mem_AB;
                            double r_crz_mem_AB = (Complex.Abs(-Zr + Zm_AB[1] - Zpol_crz_mem_AB)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AB[1] - Zpol_crz_mem_AB).Real / 2)) + r_crz_mem_AB * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AB[1] - Zpol_crz_mem_AB).Imaginary / 2)) + r_crz_mem_AB * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AB"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AB"].Points.AddXY(Zm_AB[1].Real, Zm_AB[1].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_crz_mem_AB.Real, Zs_crz_mem_AB.Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
                        }
                        if (radioButton2.Checked == true)
                        {
                            //Cross with voltage memory mho BC
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex V1bc_mem = (alpha * Va) + ((1 - alpha) * Va_mem);
                            Complex Zpol_crz_mem_BC = k_pol * V1bc_mem / Ibc;
                            Complex Zs_crz_mem_BC = Zm_BC[1] - Zpol_crz_mem_BC;
                            double r_crz_mem_BC = (Complex.Abs(-Zr + Zm_BC[1] - Zpol_crz_mem_BC)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BC[1] - Zpol_crz_mem_BC).Real / 2)) + r_crz_mem_BC * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BC[1] - Zpol_crz_mem_BC).Imaginary / 2)) + r_crz_mem_BC * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BC"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BC"].Points.AddXY(Zm_BC[1].Real, Zm_BC[1].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_crz_mem_BC.Real, Zs_crz_mem_BC.Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
                        }
                        if (radioButton3.Checked == true)
                        {
                            //Cross with voltage memory mho CA
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex V1ca_mem = (alpha * Vb) + ((1 - alpha) * Vb_mem);
                            Complex Zpol_crz_mem_CA = k_pol * V1ca_mem / Ica;
                            Complex Zs_crz_mem_CA = Zm_CA[1] - Zpol_crz_mem_CA;
                            double r_crz_mem_CA = (Complex.Abs(-Zr + Zm_CA[1] - Zpol_crz_mem_CA)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CA[1] - Zpol_crz_mem_CA).Real / 2)) + r_crz_mem_CA * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CA[1] - Zpol_crz_mem_CA).Imaginary / 2)) + r_crz_mem_CA * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CA"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CA"].Points.AddXY(Zm_CA[1].Real, Zm_CA[1].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_crz_mem_CA.Real, Zs_crz_mem_CA.Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
                        }
                        if (radioButton4.Checked == true)
                        {
                            //Cross with voltage memory mho AG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Vag_mem = (alpha * Vbc) + ((1 - alpha) * Vbc_mem);
                            Complex Zpol_crz_mem_AG = k_pol * Vag_mem / IaG;
                            Complex Zs_crz_mem_AG = Zm_AG[1] - Zpol_crz_mem_AG;
                            double r_crz_mem_AG = (Complex.Abs(-Zr + Zm_AG[1] - Zpol_crz_mem_AG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AG[1] - Zpol_crz_mem_AG).Real / 2)) + r_crz_mem_AG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AG[1] - Zpol_crz_mem_AG).Imaginary / 2)) + r_crz_mem_AG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_crz_mem_AG.Real, Zs_crz_mem_AG.Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
                        }
                        if (radioButton5.Checked == true)
                        {
                            //Cross with voltage memory mho BG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Vbg_mem = (alpha * Vca) + ((1 - alpha) * Vca_mem);
                            Complex Zpol_crz_mem_BG = k_pol * Vbg_mem / IbG;
                            Complex Zs_crz_mem_BG = Zm_BG[1] - Zpol_crz_mem_BG;
                            double r_crz_mem_BG = (Complex.Abs(-Zr + Zm_BG[1] - Zpol_crz_mem_BG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BG[1] - Zpol_crz_mem_BG).Real / 2)) + r_crz_mem_BG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BG[1] - Zpol_crz_mem_BG).Imaginary / 2)) + r_crz_mem_BG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BG"].Points.AddXY(Zm_BG[1].Real, Zm_BG[1].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_crz_mem_BG.Real, Zs_crz_mem_BG.Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
                        }
                        if (radioButton6.Checked == true)
                        {
                            //Cross with voltage memory mho CG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Vcg_mem = (alpha * Vab) + ((1 - alpha) * Vab_mem);
                            Complex Zpol_crz_mem_CG = k_pol * Vcg_mem / IcG;
                            Complex Zs_crz_mem_CG = Zm_CG[1] - Zpol_crz_mem_CG;
                            double r_crz_mem_CG = (Complex.Abs(-Zr + Zm_CG[1] - Zpol_crz_mem_CG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CG[1] - Zpol_crz_mem_CG).Real / 2)) + r_crz_mem_CG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CG[1] - Zpol_crz_mem_CG).Imaginary / 2)) + r_crz_mem_CG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CG"].Points.AddXY(Zm_CG[1].Real, Zm_CG[1].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_crz_mem_CG.Real, Zs_crz_mem_CG.Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
                        }
                    }
                    break;
                case 4:// positive sequence
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
                        if (radioButton1.Checked == true)
                        {
                            //Pos mho AB                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_pos_AB = Vab + (k_pol * Vc_1);
                            Complex Zpol_pos_AB = k_pol * Vc_1 / Iab;
                            Complex Zs_pos_AB = Zm_AB[1] - Zpol_pos_AB;
                            double r_pos_AB = (Complex.Abs(-Zr + Zm_AB[1] - Zpol_pos_AB)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AB[1] - Zpol_pos_AB).Real / 2)) + r_pos_AB * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AB[1] - Zpol_pos_AB).Imaginary / 2)) + r_pos_AB * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AB"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AB"].Points.AddXY(Zm_AB[1].Real, Zm_AB[1].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_pos_AB.Real, Zs_pos_AB.Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
                        }
                        if (radioButton2.Checked == true)
                        {
                            //pos mho BC                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_pos_BC = Vbc + (k_pol * Va_1);
                            Complex Zpol_pos_BC = k_pol * Va_1 / Ibc;
                            Complex Zs_pos_BC = Zm_BC[1] - Zpol_pos_BC;
                            double r_pos_BC = (Complex.Abs(-Zr + Zm_BC[1] - Zpol_pos_BC)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BC[1] - Zpol_pos_BC).Real / 2)) + r_pos_BC * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BC[1] - Zpol_pos_BC).Imaginary / 2)) + r_pos_BC * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BC"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BC"].Points.AddXY(Zm_BC[1].Real, Zm_BC[1].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_pos_BC.Real, Zs_pos_BC.Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
                        }
                        if (radioButton3.Checked == true)
                        {
                            //pos mho CA
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex Sop_pos_CA = (k_pol * Vb_1);
                            Complex Zpol_pos_CA = k_pol * Vb_1 / Ica;
                            Complex Zs_pos_CA = Zm_CA[1] - Zpol_pos_CA;
                            double r_pos_CA = (Complex.Abs(-Zr + Zm_CA[1] - Zpol_pos_CA)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CA[1] - Zpol_pos_CA).Real / 2)) + r_pos_CA * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CA[1] - Zpol_pos_CA).Imaginary / 2)) + r_pos_CA * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CA"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CA"].Points.AddXY(Zm_CA[1].Real, Zm_CA[1].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_pos_CA.Real, Zs_pos_CA.Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
                        }
                        if (radioButton4.Checked == true)
                        {
                            //pos mho AG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_pos_AG = Va + (k_pol * Va_1);
                            Complex Zpol_pos_AG = k_pol * Va_1 / IaG;
                            Complex Zs_pos_AG = Zm_AG[1] - Zpol_pos_AG;
                            double r_pos_AG = (Complex.Abs(-Zr + Zm_AG[1] - Zpol_pos_AG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AG[1] - Zpol_pos_AG).Real / 2)) + r_pos_AG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AG[1] - Zpol_pos_AG).Imaginary / 2)) + r_pos_AG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_pos_AG.Real, Zs_pos_AG.Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
                        }
                        if (radioButton5.Checked == true)
                        {
                            //cross mho BG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_pos_BG = Vb + (k_pol * Vb_1);
                            Complex Zpol_pos_BG = k_pol * Vb_1 / IbG;
                            Complex Zs_pos_BG = Zm_BG[1] - Zpol_pos_BG;
                            double r_pos_BG = (Complex.Abs(-Zr + Zm_BG[1] - Zpol_pos_BG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BG[1] - Zpol_pos_BG).Real / 2)) + r_pos_BG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BG[1] - Zpol_pos_BG).Imaginary / 2)) + r_pos_BG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BG"].Points.AddXY(Zm_BG[1].Real, Zm_BG[1].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_pos_BG.Real, Zs_pos_BG.Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
                        }
                        if (radioButton6.Checked == true)
                        {
                            //cross mho CG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex Sop_pos_CG = Vc + (k_pol * Vc_1);
                            Complex Zpol_pos_CG = k_pol * Vc_1 / IcG;
                            Complex Zs_pos_CG = Zm_CG[1] - Zpol_pos_CG;
                            double r_pos_CG = (Complex.Abs(-Zr + Zm_CG[1] - Zpol_pos_CG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CG[1] - Zpol_pos_CG).Real / 2)) + r_pos_CG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CG[1] - Zpol_pos_CG).Imaginary / 2)) + r_pos_CG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CG"].Points.AddXY(Zm_CG[1].Real, Zm_CG[1].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_pos_CG.Real, Zs_pos_CG.Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
                        }
                    }
                    break;
                case 5:// positive sequence with voltage memory
                    {
                        textBox21.Font = textBox21.Font;
                        if (radioButton1.Checked == true)
                        {
                            //Pos mho AB                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex V1ab_mem = (alpha * Vc_1) + ((1 - alpha) * Vc_mem);
                            Complex Sop_pos_mem_AB = Vab + (k_pol * V1ab_mem);
                            Complex Zpol_pos_mem_AB = k_pol * V1ab_mem / Iab;
                            Complex Zs_pos_mem_AB = Zm_AB[1] - Zpol_pos_mem_AB;
                            double r_pos_mem_AB = (Complex.Abs(-Zr + Zm_AB[1] - Zpol_pos_mem_AB)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AB[1] - Zpol_pos_mem_AB).Real / 2)) + r_pos_mem_AB * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AB[1] - Zpol_pos_mem_AB).Imaginary / 2)) + r_pos_mem_AB * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AB"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AB"].Points.AddXY(Zm_AB[1].Real, Zm_AB[1].Imaginary);
                            chart2.Series["Zm_AB"].Points.AddXY(Zs_pos_mem_AB.Real, Zs_pos_mem_AB.Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
                            chart2.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
                        }
                        if (radioButton2.Checked == true)
                        {
                            //pos mho BC                       
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex V1bc_mem = (alpha * Vb_1) + ((1 - alpha) * Vb_mem);
                            Complex Sop_pos_mem_BC = Vbc + (k_pol * V1bc_mem);
                            Complex Zpol_pos_mem_BC = k_pol * V1bc_mem / Ibc;
                            Complex Zs_pos_mem_BC = Zm_BC[1] - Zpol_pos_mem_BC;
                            double r_pos_mem_BC = (Complex.Abs(-Zr + Zm_BC[1] - Zpol_pos_mem_BC)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BC[1] - Zpol_pos_mem_BC).Real / 2)) + r_pos_mem_BC * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BC[1] - Zpol_pos_mem_BC).Imaginary / 2)) + r_pos_mem_BC * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BC"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BC"].Points.AddXY(Zm_BC[1].Real, Zm_BC[1].Imaginary);
                            chart2.Series["Zm_BC"].Points.AddXY(Zs_pos_mem_BC.Real, Zs_pos_mem_BC.Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                            chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
                        }
                        if (radioButton3.Checked == true)
                        {
                            //pos mho CA
                            k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                            Complex V1ca_mem = (alpha * Vc_1) + ((1 - alpha) * Vc_mem);
                            Complex Sop_pos_mem_CA = (k_pol * V1ca_mem);
                            Complex Zpol_pos_mem_CA = k_pol * V1ca_mem / Ica;
                            Complex Zs_pos_mem_CA = Zm_CA[1] - Zpol_pos_mem_CA;
                            double r_pos_mem_CA = (Complex.Abs(-Zr + Zm_CA[1] - Zpol_pos_mem_CA)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CA[1] - Zpol_pos_mem_CA).Real / 2)) + r_pos_mem_CA * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CA[1] - Zpol_pos_mem_CA).Imaginary / 2)) + r_pos_mem_CA * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CA"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CA"].Points.AddXY(Zm_CA[1].Real, Zm_CA[1].Imaginary);
                            chart2.Series["Zm_CA"].Points.AddXY(Zs_pos_mem_CA.Real, Zs_pos_mem_CA.Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
                            chart2.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
                        }
                        if (radioButton4.Checked == true)
                        {
                            //pos mho AG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex V1ag_mem = (alpha * Va_1) + ((1 - alpha) * Va_mem);
                            Complex Sop_pos_mem_AG = Va + (k_pol * V1ag_mem);
                            Complex Zpol_pos_mem_AG = k_pol * V1ag_mem / IaG;
                            Complex Zs_pos_mem_AG = Zm_AG[1] - Zpol_pos_mem_AG;
                            double r_pos_mem_AG = (Complex.Abs(-Zr + Zm_AG[1] - Zpol_pos_mem_AG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_AG[1] - Zpol_pos_mem_AG).Real / 2)) + r_pos_mem_AG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_AG[1] - Zpol_pos_mem_AG).Imaginary / 2)) + r_pos_mem_AG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_AG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
                            chart2.Series["Zm_AG"].Points.AddXY(Zs_pos_mem_AG.Real, Zs_pos_mem_AG.Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                            chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
                        }
                        if (radioButton5.Checked == true)
                        {
                            //cross mho BG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex V1bg_mem = (alpha * Vb_1) + ((1 - alpha) * Vb_mem);
                            Complex Sop_pos_mem_BG = Vb + (k_pol * V1bg_mem);
                            Complex Zpol_pos_mem_BG = k_pol * V1bg_mem / IbG;
                            Complex Zs_pos_mem_BG = Zm_BG[1] - Zpol_pos_mem_BG;
                            double r_pos_mem_BG = (Complex.Abs(-Zr + Zm_BG[1] - Zpol_pos_mem_BG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_BG[1] - Zpol_pos_mem_BG).Real / 2)) + r_pos_mem_BG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_BG[1] - Zpol_pos_mem_BG).Imaginary / 2)) + r_pos_mem_BG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_BG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_BG"].Points.AddXY(Zm_BG[1].Real, Zm_BG[1].Imaginary);
                            chart2.Series["Zm_BG"].Points.AddXY(Zs_pos_mem_BG.Real, Zs_pos_mem_BG.Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
                            chart2.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
                        }
                        if (radioButton6.Checked == true)
                        {
                            //cross mho CG
                            k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                            Complex V1cg_mem = (alpha * Vc_1) + ((1 - alpha) * Vc_mem);
                            Complex Sop_pos_mem_CG = Vc + (k_pol * V1cg_mem);
                            Complex Zpol_pos_mem_CG = k_pol * V1cg_mem / IcG;
                            Complex Zs_pos_mem_CG = Zm_CG[1] - Zpol_pos_mem_CG;
                            double r_pos_mem_CG = (Complex.Abs(-Zr + Zm_CG[1] - Zpol_pos_mem_CG)) / 2;
                            for (int k = 0; k <= 1000; k++)
                            {
                                double x = (((Zr + Zm_CG[1] - Zpol_pos_mem_CG).Real / 2)) + r_pos_mem_CG * Math.Cos(k * 2 * Math.PI / 1000);
                                double y = (((Zr + Zm_CG[1] - Zpol_pos_mem_CG).Imaginary / 2)) + r_pos_mem_CG * Math.Sin(k * 2 * Math.PI / 1000);
                                chart2.Series["Mho_CG"].Points.AddXY(x, y);
                            }
                            chart2.Series["Zm_CG"].Points.AddXY(Zm_CG[1].Real, Zm_CG[1].Imaginary);
                            chart2.Series["Zm_CG"].Points.AddXY(Zs_pos_mem_CG.Real, Zs_pos_mem_CG.Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
                            chart2.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
                        }
                    }
                    break;
            }
            chart1.ChartAreas[0].AxisY.Interval = chart1.ChartAreas[0].AxisX.Interval;
            chart2.ChartAreas[0].AxisY.Interval = chart2.ChartAreas[0].AxisX.Interval;
        }
        void loops_fault()
        {
            //AB
            if (radioButton1.Checked == true)
            {
                chart1.Series["Zpol_AB"].Enabled = true;
                chart1.Series["Zm_AB"].Enabled = true;
                chart2.Series["Zpol_AB"].Enabled = true;
                chart2.Series["Zm_AB"].Enabled = true;
                chart2.Series["Zs_AB"].Enabled = true;
                chart2.Series["Mho_AB"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_AB"].Enabled = false;
                chart1.Series["Zm_AB"].Enabled = false;
                chart2.Series["Zpol_AB"].Enabled = false;
                chart2.Series["Zm_AB"].Enabled = false;
                chart2.Series["Zs_AB"].Enabled = false;
                chart2.Series["Mho_AB"].Enabled = false;
            }
            //BC
            if (radioButton2.Checked == true)
            {
                chart1.Series["Zpol_BC"].Enabled = true;
                chart1.Series["Zm_BC"].Enabled = true;
                chart2.Series["Zpol_BC"].Enabled = true;
                chart2.Series["Zm_BC"].Enabled = true;
                chart2.Series["Zs_BC"].Enabled = true;
                chart2.Series["Mho_BC"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_BC"].Enabled = false;
                chart1.Series["Zm_BC"].Enabled = false;
                chart2.Series["Zpol_BC"].Enabled = false;
                chart2.Series["Zm_BC"].Enabled = false;
                chart2.Series["Zs_BC"].Enabled = false;
                chart2.Series["Mho_BC"].Enabled = false;

            }
            //CA
            if (radioButton3.Checked == true)
            {
                chart1.Series["Zpol_CA"].Enabled = true;
                chart1.Series["Zm_CA"].Enabled = true;
                chart2.Series["Zpol_CA"].Enabled = true;
                chart2.Series["Zm_CA"].Enabled = true;
                chart2.Series["Zs_CA"].Enabled = true;
                chart2.Series["Mho_CA"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_CA"].Enabled = false;
                chart1.Series["Zm_CA"].Enabled = false;
                chart2.Series["Zpol_CA"].Enabled = false;
                chart2.Series["Zm_CA"].Enabled = false;
                chart2.Series["Zs_CA"].Enabled = false;
                chart2.Series["Mho_CA"].Enabled = false;
            }
            //AG
            if (radioButton4.Checked == true)
            {
                chart1.Series["Zpol_AG"].Enabled = true;
                chart1.Series["Zm_AG"].Enabled = true;
                chart2.Series["Zpol_AG"].Enabled = true;
                chart2.Series["Zm_AG"].Enabled = true;
                chart2.Series["Zs_AG"].Enabled = true;
                chart2.Series["Mho_AG"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_AG"].Enabled = false;
                chart1.Series["Zm_AG"].Enabled = false;
                chart2.Series["Zpol_AG"].Enabled = false;
                chart2.Series["Zm_AG"].Enabled = false;
                chart2.Series["Zs_AG"].Enabled = false;
                chart2.Series["Mho_AG"].Enabled = false;
            }
            //BG
            if (radioButton5.Checked == true)
            {
                chart1.Series["Zpol_BG"].Enabled = true;
                chart1.Series["Zm_BG"].Enabled = true;
                chart2.Series["Zpol_BG"].Enabled = true;
                chart2.Series["Zm_BG"].Enabled = true;
                chart2.Series["Zs_BG"].Enabled = true;
                chart2.Series["Mho_BG"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_BG"].Enabled = false;
                chart1.Series["Zm_BG"].Enabled = false;
                chart2.Series["Zpol_BG"].Enabled = false;
                chart2.Series["Zm_BG"].Enabled = false;
                chart2.Series["Zs_BG"].Enabled = false;
                chart2.Series["Mho_BG"].Enabled = false;
            }
            //CG
            if (radioButton6.Checked == true)
            {
                chart1.Series["Zpol_CG"].Enabled = true;
                chart1.Series["Zm_CG"].Enabled = true;
                chart2.Series["Zpol_CG"].Enabled = true;
                chart2.Series["Zm_CG"].Enabled = true;
                chart2.Series["Zs_CG"].Enabled = true;
                chart2.Series["Mho_CG"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_CG"].Enabled = false;
                chart1.Series["Zm_CG"].Enabled = false;
                chart2.Series["Zpol_CG"].Enabled = false;
                chart2.Series["Zm_CG"].Enabled = false;
                chart2.Series["Zs_CG"].Enabled = false;
                chart2.Series["Mho_CG"].Enabled = false;
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            polarization_calc();
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            loops_fault();
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            polarization_calc();
            
        }
        double f_sist = 0;
        double f_ams = 0;
        int datagridlengh = 0;
        string import_format;
        string HT_VA;
        string HT_VB;
        string HT_VC;
        string HT_IA;
        string HT_IB;
        string HT_IC;
        double g_Zl1_re = 0;
        double g_Zl0_re = 0;
        double g_Zl1_im = 0;
        double g_Zl0_im = 0;
        double g_Leng_line = 0;
        Complex g_LT1 = new Complex(0, 0);
        double[] g_time = new double[10000];
        double[] g_va = new double[10000];
        double[] g_vb = new double[10000];
        double[] g_vc = new double[10000];
        double[] g_ia = new double[10000];
        double[] g_ib = new double[10000];
        double[] g_ic = new double[10000];
        Complex[] g_za = new Complex[10000];
        Complex[] g_zb = new Complex[10000];
        Complex[] g_zc = new Complex[10000];
        Complex[] g_zab = new Complex[10000];
        Complex[] g_zbc = new Complex[10000];
        Complex[] g_zca = new Complex[10000];
        void import()
        {
            OpenFileDialog OF = new OpenFileDialog();
            OF.Filter = "COMTRADE|*.cfg|ASCII |*.*";
            if (OF.ShowDialog() == DialogResult.OK)
            {
                if (OF.FilterIndex == 1)
                {                  
                    
                    dataGridView1.Rows.Clear();
                    dataGridView2.Rows.Clear();
                    dataGridView3.Rows.Clear();
                    string fn = OF.FileName;
                    string[] line = File.ReadAllLines(fn);
                    string[] contrade = new string[line.Length];
                    string CFGpath = Path.GetFullPath(fn);
                    int lenght = line.Length;
                    string[] entradas = new string[3];

                    int count = 0;
                    foreach (string str in line[1].Split('\t', ',', ';', ' '))
                        if (count < 3)
                        {
                            entradas[count] = str.Replace(@"A", "");
                            count++;

                        }
                    int AnalogLengh = Convert.ToInt32(entradas[1]) + 2;
                    int amostragemVET = Convert.ToInt32(entradas[0]) + 4;
                    int freqsistVET = Convert.ToInt32(entradas[0]) + 2;
                    dataGridView2.ColumnCount = lenght;
                    int j = 0;
                    foreach (string s1 in line)
                    {
                        string[] linha = new string[lenght];
                        int contagem = 0;
                        if (s1 != "")
                        {
                            foreach (string coluna in s1.Split(',', '\t', ';'))
                                if (coluna != null)
                                {
                                    linha[contagem] = coluna.Replace(".", ",");
                                    contagem++;
                                }
                            contrade[j] = linha[1];
                            j++;
                            dataGridView2.Rows.Add(linha);
                        }
                    }
                    dataGridView1.ColumnCount = AnalogLengh;
                    dataGridView3.ColumnCount = AnalogLengh;
                    string[] nome = new string[AnalogLengh - 2];
                    for (int x = 2; x < AnalogLengh; x++)
                    {
                        if (dataGridView2.Rows[x].Cells[1].Value.ToString() != null)
                        {
                            nome[x - 2] = dataGridView2.Rows[x].Cells[1].Value.ToString();
                            dataGridView1.Columns[x].HeaderText = nome[x - 2].ToString();
                        }
                    }
                    Form1.ActiveForm.Text = "Power Systems Analysis - " + CFGpath;
                    count = 0;
                    string DATpath = CFGpath.Replace(".CFG", ".DAT");
                    using (var DATreader = new StreamReader(DATpath))
                    {
                        List<string> listA = new List<string>();
                        List<string> listB = new List<string>();
                        while (!DATreader.EndOfStream)
                        {
                            var lines = DATreader.ReadLine();
                            var values = lines.Split(',');
                            dataGridView1.Rows.Add(values);
                            count++;
                        }
                    }
                    string[] f1_rt_p = new string[100];
                    string[] f1_rt_s = new string[100];
                    string[] unidade = new string[100];
                    string unidade_V = "";
                    string unidade_A = "";
                    import_format = "comtrade";
                    Form2 form2 = new Form2();
                    form2.import_fmt = import_format;
                    form2.Analogico_A_forms2 = analogic_A;
                    form2.Analogico_V_forms2 = analogic_V;
                    for (int x = 2; x < AnalogLengh; x++)
                    {
                        if (dataGridView2.Rows[x].Cells[1].Value.ToString() != null)
                        {
                            unidade[x - 2] = dataGridView2.Rows[x].Cells[4].Value.ToString();
                            if (unidade[x - 2] == "kA".ToString() || unidade[x - 2] == "A".ToString())
                            {
                                nome[x - 2] = dataGridView2.Rows[x].Cells[1].Value.ToString();
                                analogic_A[x - 2] = nome[x - 2];
                            }
                            if (unidade[x - 2] == "kV".ToString() || unidade[x - 2] == "V".ToString())
                            {
                                nome[x - 2] = dataGridView2.Rows[x].Cells[1].Value.ToString();
                                analogic_V[x - 2] = nome[x - 2];
                            }
                            if (dataGridView2.Rows[x].Cells[11].Value != null)
                            {
                                f1_rt_p[x - 2] = dataGridView2.Rows[x].Cells[10].Value.ToString();
                                f1_rt_s[x - 2] = dataGridView2.Rows[x].Cells[11].Value.ToString();
                            }
                            else
                            {
                                f1_rt_p[x - 2] = 1.ToString();
                                f1_rt_s[x - 2] = 1.ToString();
                            }
                            unidade[x - 2] = dataGridView2.Rows[x].Cells[4].Value.ToString();
                        }


                        if (unidade[x - 2] == "kA".ToString() || unidade[x - 2] == "A".ToString())
                        {
                            form2.RTC_P = Math.Round(Convert.ToDouble(f1_rt_p[x - 2]), 2).ToString();
                            form2.RTC_S = Math.Round(Convert.ToDouble(f1_rt_s[x - 2]), 2).ToString();
                            unidade_A = unidade[x - 2];
                        }
                        if (unidade[x - 2] == "kV".ToString() || unidade[x - 2] == "V".ToString())
                        {
                            form2.RTP_P = Math.Round(Convert.ToDouble(f1_rt_p[x - 2]), 2).ToString();
                            form2.RTP_S = Math.Round(Convert.ToDouble(f1_rt_s[x - 2]), 2).ToString();
                            unidade_V = unidade[x - 2];
                        }
                    }
                    string Fase_A = 0.ToString();
                    string Fase_B = 0.ToString();
                    string Fase_C = 0.ToString();
                    string Corrente_A = 0.ToString();
                    string Corrente_B = 0.ToString();
                    string Corrente_C = 0.ToString();
                    double rtp = 0;
                    double rtc = 0;
                    if (form2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        Fase_A = form2.Tensao_VA;
                        Fase_B = form2.Tensao_VB;
                        Fase_C = form2.Tensao_VC;
                        Corrente_A = form2.Corrente_IA;
                        Corrente_B = form2.Corrente_IB;
                        Corrente_C = form2.Corrente_IC;
                        rtp = form2.F2_RTP;
                        rtc = form2.F2_RTC;
                        g_Leng_line = Convert.ToDouble(form2.Lengh_Line);
                        g_Zl1_re = Convert.ToDouble(form2.ZL1_re) * g_Leng_line;
                        g_Zl1_im = Convert.ToDouble(form2.ZL1_im) * g_Leng_line;
                        g_Zl0_re = Convert.ToDouble(form2.ZL0_re) * g_Leng_line;
                        g_Zl0_im = Convert.ToDouble(form2.ZL0_im) * g_Leng_line;

                        string file = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\TCC_DadosLinha.txt";
                        using (StreamWriter bw = File.CreateText(file))
                        {
                            bw.Write("" + g_Leng_line);//0
                            bw.Write("\r\n" + (form2.ZL1_re));//1
                            bw.Write("\r\n" + form2.ZL1_im);//2
                            bw.Write("\r\n" + form2.ZL0_re);//3
                            bw.Write("\r\n" + form2.ZL0_im);
                        }
                    }
                    g_LT1 = new Complex(g_Zl1_re, g_Zl1_im);
                    HT_VA = Fase_A;
                    HT_VB = Fase_B;
                    HT_VC = Fase_C;
                    HT_IA = Corrente_A;
                    HT_IB = Corrente_B;
                    HT_IC = Corrente_C;
                    double amostragem = 0;
                    double freqsist = 0;
                    for (int x = 0; x < dataGridView2.ColumnCount; x++)
                    {
                        // freq sist
                        if (x == freqsistVET)
                        {
                            freqsist = Convert.ToDouble(dataGridView2.Rows[x].Cells[0].Value);
                        }
                        // taxa de amostragem
                        if (x == amostragemVET)
                        {
                            amostragem = Convert.ToDouble(dataGridView2.Rows[x].Cells[0].Value);
                        }
                    }
                    double adder_tensao_va = 0;
                    double adder_tensao_vb = 0;
                    double adder_tensao_vc = 0;
                    double adder_corrente_ia = 0;
                    double adder_corrente_ib = 0;
                    double adder_corrente_ic = 0;
                    double multi_tensao_va = 0;
                    double multi_tensao_vb = 0;
                    double multi_tensao_vc = 0;
                    double multi_corrente_ia = 0;
                    double multi_corrente_ib = 0;
                    double multi_corrente_ic = 0;
                    for (int x = 0; x < dataGridView2.Rows.Count - 2; x++)
                    {
                        if ((dataGridView2.Rows[x].Cells[1].Value != null) && (Fase_A == dataGridView2.Rows[x].Cells[1].Value.ToString()))
                        {
                            multi_tensao_va = Convert.ToDouble(dataGridView2.Rows[x].Cells[5].Value);
                            adder_tensao_va = Convert.ToDouble(dataGridView2.Rows[x].Cells[6].Value);
                        }
                        if ((dataGridView2.Rows[x].Cells[1].Value != null) && (Fase_B == dataGridView2.Rows[x].Cells[1].Value.ToString()))
                        {
                            multi_tensao_vb = Convert.ToDouble(dataGridView2.Rows[x].Cells[5].Value);
                            adder_tensao_vb = Convert.ToDouble(dataGridView2.Rows[x].Cells[6].Value);
                        }
                        if ((dataGridView2.Rows[x].Cells[1].Value != null) && (Fase_C == dataGridView2.Rows[x].Cells[1].Value.ToString()))
                        {
                            multi_tensao_vc = Convert.ToDouble(dataGridView2.Rows[x].Cells[5].Value);
                            adder_tensao_vc = Convert.ToDouble(dataGridView2.Rows[x].Cells[6].Value);
                        }
                        if ((dataGridView2.Rows[x].Cells[1].Value != null) && (Corrente_A == dataGridView2.Rows[x].Cells[1].Value.ToString()))
                        {
                            multi_corrente_ia = Convert.ToDouble(dataGridView2.Rows[x].Cells[5].Value);
                            adder_corrente_ia = Convert.ToDouble(dataGridView2.Rows[x].Cells[6].Value);
                        }
                        if ((dataGridView2.Rows[x].Cells[1].Value != null) && (Corrente_B == dataGridView2.Rows[x].Cells[1].Value.ToString()))
                        {
                            multi_corrente_ib = Convert.ToDouble(dataGridView2.Rows[x].Cells[5].Value);
                            adder_corrente_ib = Convert.ToDouble(dataGridView2.Rows[x].Cells[6].Value);
                        }
                        if ((dataGridView2.Rows[x].Cells[1].Value != null) && (Corrente_C == dataGridView2.Rows[x].Cells[1].Value.ToString()))
                        {
                            multi_corrente_ic = Convert.ToDouble(dataGridView2.Rows[x].Cells[5].Value);
                            adder_corrente_ic = Convert.ToDouble(dataGridView2.Rows[x].Cells[6].Value);
                        }
                    }
                    f_sist = freqsist;
                    f_ams = amostragem;
                    chart3.Visible = true;
                    double time = 0;
                    double[] tempo = new double[dataGridView1.Rows.Count - 1];
                    double[] va = new double[dataGridView1.Rows.Count - 1];
                    double[] vb = new double[dataGridView1.Rows.Count - 1];
                    double[] vc = new double[dataGridView1.Rows.Count - 1];
                    double[] ia = new double[dataGridView1.Rows.Count - 1];
                    double[] ib = new double[dataGridView1.Rows.Count - 1];
                    double[] ic = new double[dataGridView1.Rows.Count - 1];
                    float[] VA_ = new float[dataGridView1.Rows.Count - 1];
                    double nf = 0;
                    g_time = new double[dataGridView1.Rows.Count - 1];
                    g_va = new double[dataGridView1.Rows.Count - 1];
                    g_vb = new double[dataGridView1.Rows.Count - 1];
                    g_vc = new double[dataGridView1.Rows.Count - 1];
                    g_ia = new double[dataGridView1.Rows.Count - 1];
                    g_ib = new double[dataGridView1.Rows.Count - 1];
                    g_ic = new double[dataGridView1.Rows.Count - 1];
                    datagridlengh = dataGridView1.Rows.Count;
                    toolStripLabel1.Text = "f: " + amostragem + " Hz";
                    toolStripLabel2.Text = "fn: " + freqsist + " Hz";
                    for (int x = 2; x < AnalogLengh; x++)
                    {
                        if (dataGridView1.Columns[x].HeaderText == Fase_A)
                        {
                            chart3.Series[0].Name = "" + Fase_A;
                            time = 0;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                //series 1                            
                                g_time[y] = Math.Round((1 / amostragem) + time, 6);
                                nf = g_time[y];
                                time = Math.Round((1 / amostragem) + time, 6);
                                va[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * rtp;
                                if (multi_tensao_va != 0)
                                {
                                    va[y] = (Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * multi_tensao_va * rtp) + adder_tensao_va;
                                }
                                if (unidade_V == "kV".ToString())
                                {
                                    va[y] = va[y] * 1000;
                                }
                                g_va[y] = va[y];
                                VA_[y] = (float)va[y];
                                chart3.Series[0].Points.AddXY(time, va[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Fase_B)
                        {
                            chart3.Series[1].Name = "" + Fase_B;
                            time = 0;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                //series 1
                                g_time[y] = Math.Round((1 / amostragem) + time, 6);
                                time = Math.Round((1 / amostragem) + time, 6);
                                vb[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * rtp;
                                if (multi_tensao_va != 0)
                                {
                                    vb[y] = (Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * multi_tensao_vb * rtp) + adder_tensao_vb;
                                }
                                if (unidade_V == "kV".ToString())
                                {
                                    vb[y] = vb[y] * 1000;
                                }
                                g_vb[y] = vb[y];
                                chart3.Series[1].Points.AddXY(time, vb[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Fase_C)
                        {
                            chart3.Series[2].Name = "" + Fase_C;
                            time = 0;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                //series 1
                                g_time[y] = Math.Round((1 / amostragem) + time, 6);
                                time = Math.Round((1 / amostragem) + time, 6);
                                vc[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * rtp;
                                if (multi_tensao_va != 0)
                                {
                                    vc[y] = (Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * multi_tensao_vc * rtp) + adder_tensao_vc;
                                }
                                if (unidade_V == "kV".ToString())
                                {
                                    vc[y] = vc[y] * 1000;
                                }
                                g_vc[y] = vc[y];
                                chart3.Series[2].Points.AddXY(time, vc[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Corrente_A)
                        {
                            time = 0;
                            chart3.Series[3].Name = "" + Corrente_A;
                            //chart12.Series[0].Name = "" + Corrente_A + " RMS";
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                g_time[y] = Math.Round((1 / amostragem) + time, 6);
                                time = Math.Round((1 / amostragem) + time, 6);
                                ia[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value.ToString()) * rtc;
                                if (multi_corrente_ia != 0)
                                {
                                    ia[y] = (Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value.ToString()) * multi_corrente_ia * rtc) + adder_corrente_ia;
                                }
                                g_ia[y] = ia[y];
                                chart3.Series[3].Points.AddXY(time, ia[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Corrente_B)
                        {
                            time = 0;
                            chart3.Series[4].Name = "" + Corrente_B;
                            //chart12.Series[1].Name = "" + Corrente_B + " RMS";
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                g_time[y] = Math.Round((1 / amostragem) + time, 6);
                                time = Math.Round((1 / amostragem) + time, 6);
                                ib[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * rtc;
                                if (multi_corrente_ia != 0)
                                {
                                    ib[y] = (Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * multi_corrente_ib * rtc) + adder_corrente_ib;
                                }
                                g_ib[y] = ib[y];
                                chart3.Series[4].Points.AddXY(time, ib[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Corrente_C)
                        {
                            time = 0;
                            chart3.Series[5].Name = "" + Corrente_C;
                            //chart12.Series[2].Name = "" + Corrente_C + " RMS";
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                g_time[y] = Math.Round((1 / amostragem) + time, 6);
                                time = Math.Round((1 / amostragem) + time, 6);
                                ic[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * rtc;
                                if (multi_corrente_ia != 0)
                                {
                                    ic[y] = (Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value) * multi_corrente_ic * rtc) + adder_corrente_ic;
                                }
                                g_ic[y] = ic[y];
                                chart3.Series[5].Points.AddXY(time, ic[y]);
                            }
                        }
                    }
                    dataGridView1.Rows.Clear();
                    int N = Convert.ToInt32(f_ams / f_sist);
                    for (int x = 0; x < g_time.Length; x++)
                    {
                        dataGridView1.Rows.Add(g_time[x], va[x], vb[x], vc[x], ia[x], ib[x], ic[x]);
                    }
                    dataGridView1.Columns[0].HeaderText = "tempo";
                    dataGridView1.Columns[1].HeaderText = "Va";
                    dataGridView1.Columns[2].HeaderText = "Vb";
                    dataGridView1.Columns[3].HeaderText = "Vc";
                    dataGridView1.Columns[4].HeaderText = "Ia";
                    dataGridView1.Columns[5].HeaderText = "Ib";
                    dataGridView1.Columns[6].HeaderText = "Ic";
                }


                if (OF.FilterIndex == 2)
                {
                    dataGridView1.Rows.Clear();
                    string fn = OF.FileName;
                    string readfile = File.ReadAllText(fn);
                    string[] line = readfile.Split('\n');
                    string CFGpath = Path.GetFullPath(fn);
                    Form1.ActiveForm.Text = "Power Systems Analysis - " + CFGpath;
                    int count = 0;
                    int lenght = line.Length;                    
                    foreach (string str in line[2].Split(',', ';', '\t')) // escolhe o separador das linhas (pode mudar para qualquer separador ";" ou " ")
                    {
                        if (str != "")
                            count++;                      
                        
                    }
                    dataGridView1.ColumnCount = count;
                    dataGridView3.ColumnCount = count;
                    List<string> list = new List<string>();
                    int startIndexofcomm;
                    
                    foreach (string s1 in line)
                    {
                        string[] linha = new string[lenght];
                        int contagem = 0;
                        if (s1.Contains(@"//"))
                        {
                            startIndexofcomm = s1.IndexOf(@"//");
                            list.Add(s1.Substring(0, startIndexofcomm));
                            continue;
                        }
                        if (s1 != "")
                        {
                            foreach (string coluna in s1.Split(',', '\t', ';'))
                                if (coluna != null)
                                {
                                    linha[contagem] = coluna.Replace(".", ",");
                                    contagem++;
                                }
                            dataGridView1.Rows.Add(linha);
                        }
                    }
                    import_format = "ascii";
                    Form2 form2 = new Form2();
                    form2.import_fmt = import_format;
                    form2.Analogico_A_forms2 = analogic_A;
                    form2.Analogico_V_forms2 = analogic_V;
                    string[] unidade = new string[count];
                    for (int x = 0; x < count; x++)
                    {
                        if (dataGridView1.Rows[0].Cells[x].Value.ToString() != null)
                        {
                            unidade[x] = dataGridView1.Rows[0].Cells[x].Value.ToString();
                            if (unidade[x] == "i".ToString() || unidade[x] == "A".ToString())
                            {
                                
                            }
                            if (unidade[x] == "v".ToString() || unidade[x] == "V".ToString())
                            {
                                
                            }
                            analogic_A[x] = dataGridView1.Rows[0].Cells[x].Value.ToString();
                            analogic_V[x] = dataGridView1.Rows[0].Cells[x].Value.ToString();
                            dataGridView1.Columns[x].HeaderText = dataGridView1.Rows[0].Cells[x].Value.ToString();
                        }
                    }
                    string Fase_A = 0.ToString();
                    string Fase_B = 0.ToString();
                    string Fase_C = 0.ToString();
                    string Corrente_A = 0.ToString();
                    string Corrente_B = 0.ToString();
                    string Corrente_C = 0.ToString();
                    string time = 0.ToString();
                    double[] t = new double[dataGridView1.Rows.Count - 1];
                    double[] va = new double[dataGridView1.Rows.Count - 1];
                    double[] vb = new double[dataGridView1.Rows.Count - 1];
                    double[] vc = new double[dataGridView1.Rows.Count - 1];
                    double[] ia = new double[dataGridView1.Rows.Count - 1];
                    double[] ib = new double[dataGridView1.Rows.Count - 1];
                    double[] ic = new double[dataGridView1.Rows.Count - 1];
                    g_time = new double[dataGridView1.Rows.Count - 1];
                    g_va = new double[dataGridView1.Rows.Count - 1];
                    g_vb = new double[dataGridView1.Rows.Count - 1];
                    g_vc = new double[dataGridView1.Rows.Count - 1];
                    g_ia = new double[dataGridView1.Rows.Count - 1];
                    g_ib = new double[dataGridView1.Rows.Count - 1];
                    g_ic = new double[dataGridView1.Rows.Count - 1];
                    dataGridView1.Rows.RemoveAt(this.dataGridView1.Rows[0].Index);
                    if (form2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        Fase_A = form2.Tensao_VA;
                        Fase_B = form2.Tensao_VB;
                        Fase_C = form2.Tensao_VC;
                        Corrente_A = form2.Corrente_IA;
                        Corrente_B = form2.Corrente_IB;
                        Corrente_C = form2.Corrente_IC;
                        time = form2.g_time;
                    }
                    chart3.Visible = true;
                    for (int x = 0; x < count; x++)
                    {
                        if (dataGridView1.Columns[x].HeaderText == time)
                        {
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {                               
                                t[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value);
                                g_time[y] = t[y];
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Fase_A)
                        {
                            chart3.Series[0].Name = "" + Fase_A;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                //series 1                                
                                va[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value);
                                g_va[y] = va[y];
                                chart3.Series[0].Points.AddXY(t[y], va[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Fase_B)
                        {
                            chart3.Series[1].Name = "" + Fase_B;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                //series 1
                                vb[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value);
                                g_vb[y] = vb[y];
                                chart3.Series[1].Points.AddXY(t[y], vb[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Fase_C)
                        {
                            chart3.Series[2].Name = "" + Fase_C;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                //series 1
                                vc[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value);
                                g_vc[y] = vc[y];
                                chart3.Series[2].Points.AddXY(t[y], vc[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Corrente_A)
                        {
                            chart3.Series[3].Name = "" + Corrente_A;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                ia[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value.ToString());                                
                                g_ia[y] = ia[y];
                                chart3.Series[3].Points.AddXY(t[y], ia[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Corrente_B)
                        {
                            chart3.Series[4].Name = "" + Corrente_B;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                ib[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value);
                                g_ib[y] = ib[y];
                                chart3.Series[4].Points.AddXY(t[y], ib[y]);
                            }
                        }
                        if (dataGridView1.Columns[x].HeaderText == Corrente_C)
                        {
                            chart3.Series[5].Name = "" + Corrente_C;
                            for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                            {
                                ic[y] = Convert.ToDouble(dataGridView1.Rows[y].Cells[x].Value);
                                g_ic[y] = ic[y];
                                chart3.Series[5].Points.AddXY(t[y], ic[y]);
                            }
                        }
                        f_ams = 1 / g_time[1];
                        f_sist = Convert.ToDouble(form2.g_fn);
                    }
                    
                }
                trackBar1.Visible = true;
                trackBar1.Value = 1;
            }

            // Discretização do sinal


        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            /*
            Form2 form2 = new Form2();
            form2.ShowDialog();
            */
            import();
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            int AnalogLengh = dataGridView1.ColumnCount;
            int pn = dataGridView1.RowCount - 1;

            int N = Convert.ToInt32(f_ams / f_sist);
            trackBar1.Maximum = g_time.Length - N-1;
            DataPoint Va = new DataPoint(0, 0);
            DataPoint Vy = new DataPoint(0, 0);
            if (chart3.Series[0].Points.Count != 0)
            {
                chart3.Annotations[0].AnchorDataPoint = chart3.Series[0].Points[0];
                Va = chart3.Series[0].Points[trackBar1.Value];
                Vy = chart3.Series[0].Points[trackBar1.Value + N - 1];
            }
            if (chart3.Series[3].Points.Count != 0)
            {
                chart3.Annotations[1].AnchorDataPoint = chart3.Series[3].Points[0];
                Va = chart3.Series[3].Points[trackBar1.Value];
                Vy = chart3.Series[3].Points[trackBar1.Value + N - 1];
            }
            chart3.Annotations[0].AnchorDataPoint = chart3.Series[0].Points[0];
            chart3.Annotations[0].Y = Math.Max(chart3.ChartAreas[0].AxisY.Maximum, Math.Abs(chart3.ChartAreas[0].AxisY.Minimum));
            chart3.Annotations[0].Height = Math.Max(chart3.ChartAreas[0].AxisY.Maximum, Math.Abs(chart3.ChartAreas[0].AxisY.Minimum)) * -2;
            //chart3.Annotations[0].Width = g_time[N];
            chart3.Annotations[0].X = g_time[trackBar1.Value];

            chart3.Annotations[1].Y = Math.Max(chart3.ChartAreas[1].AxisY.Maximum, Math.Abs(chart3.ChartAreas[1].AxisY.Minimum));
            chart3.Annotations[1].Height = Math.Max(chart3.ChartAreas[1].AxisY.Maximum, Math.Abs(chart3.ChartAreas[1].AxisY.Minimum)) * -2;
            //chart3.Annotations[1].Width = g_time[N];
            chart3.Annotations[1].X = g_time[trackBar1.Value];


            double px = Va.XValue;
            double py = Vy.XValue;
            dataGridView3.Rows.Clear();
            for (int o = 0; o < pn; o++)
            {
                if (double.Parse(dataGridView1.Rows[o].Cells[0].Value.ToString()) >= px && double.Parse(dataGridView1.Rows[o].Cells[0].Value.ToString()) <= py)
                {
                    string Nfasor0, Nfasor1, Nfasor2;
                    string Ifasor0, Ifasor1, Ifasor2;
                    Nfasor0 = dataGridView1.Rows[o].Cells[1].Value.ToString();
                    Nfasor1 = dataGridView1.Rows[o].Cells[2].Value.ToString();
                    Nfasor2 = dataGridView1.Rows[o].Cells[3].Value.ToString();
                    Ifasor0 = dataGridView1.Rows[o].Cells[4].Value.ToString();
                    Ifasor1 = dataGridView1.Rows[o].Cells[5].Value.ToString();
                    Ifasor2 = dataGridView1.Rows[o].Cells[6].Value.ToString();
                    dataGridView3.Rows.Add(Nfasor0, Nfasor1, Nfasor2, Ifasor0, Ifasor1, Ifasor2);
                }
            }
            //CALCULO DA DFT
            double pi_div = 2.0 * Math.PI / N;
            double[] funcao_va = new double[N];
            double[] N1_va = new double[N];
            double[] N2_va = new double[N];
            double[] funcao_vb = new double[N];
            double[] N1_vb = new double[N];
            double[] N2_vb = new double[N];
            double[] funcao_vc = new double[N];
            double[] N1_vc = new double[N];
            double[] N2_vc = new double[N];
            double[] funcao_ia = new double[N];
            double[] N1_ia = new double[N];
            double[] N2_ia = new double[N];
            double[] funcao_ib = new double[N];
            double[] N1_ib = new double[N];
            double[] N2_ib = new double[N];
            double[] funcao_ic = new double[N];
            double[] N1_ic = new double[N];
            double[] N2_ic = new double[N];

            for (int p = 0; p < N; p++)
            {
                for (int j = 0; j < N; j++)
                {
                    funcao_va[j] = Convert.ToDouble(dataGridView3.Rows[j].Cells[0].Value); // todos os valores da função a ser analisada                    
                    N1_va[p] += funcao_va[j] * Math.Cos((pi_div * p * j)); // Reais
                    N2_va[p] -= funcao_va[j] * Math.Sin((pi_div * p * j)); // Imaginarios

                    funcao_vb[j] = Convert.ToDouble(dataGridView3.Rows[j].Cells[1].Value); // todos os valores da função a ser analisada                    
                    N1_vb[p] += funcao_vb[j] * Math.Cos((pi_div * p * j)); // Reais
                    N2_vb[p] -= funcao_vb[j] * Math.Sin((pi_div * p * j)); // Imaginarios

                    funcao_vc[j] = Convert.ToDouble(dataGridView3.Rows[j].Cells[2].Value); // todos os valores da função a ser analisada                    
                    N1_vc[p] += funcao_vc[j] * Math.Cos((pi_div * p * j)); // Reais
                    N2_vc[p] -= funcao_vc[j] * Math.Sin((pi_div * p * j)); // Imaginarios   

                    funcao_ia[j] = Convert.ToDouble(dataGridView3.Rows[j].Cells[3].Value); // todos os valores da função a ser analisada                    
                    N1_ia[p] += funcao_ia[j] * Math.Cos((pi_div * p * j)); // Reais
                    N2_ia[p] -= funcao_ia[j] * Math.Sin((pi_div * p * j)); // Imaginarios

                    funcao_ib[j] = Convert.ToDouble(dataGridView3.Rows[j].Cells[4].Value); // todos os valores da função a ser analisada                    
                    N1_ib[p] += funcao_ib[j] * Math.Cos((pi_div * p * j)); // Reais
                    N2_ib[p] -= funcao_ib[j] * Math.Sin((pi_div * p * j)); // Imaginarios

                    funcao_ic[j] = Convert.ToDouble(dataGridView3.Rows[j].Cells[5].Value); // todos os valores da função a ser analisada                    
                    N1_ic[p] += funcao_ic[j] * Math.Cos((pi_div * p * j)); // Reais
                    N2_ic[p] -= funcao_ic[j] * Math.Sin((pi_div * p * j)); // Imaginarios
                }

            }
            // CALCULO VALOR RMS

            double[] Va_ABS = new double[N];
            double[] Vb_ABS = new double[N];
            double[] Vc_ABS = new double[N];
            double casas_a = 0, RMS_VA = 0, somatorio_a = 0;
            double casas_b = 0, RMS_VB = 0, somatorio_b = 0;
            double casas_c = 0, RMS_VC = 0, somatorio_c = 0;

            double[] Ia_ABS = new double[N];
            double[] Ib_ABS = new double[N];
            double[] Ic_ABS = new double[N];
            double casas_ia = 0, RMS_IA = 0, somatorio_ia = 0;
            double casas_ib = 0, RMS_IB = 0, somatorio_ib = 0;
            double casas_ic = 0, RMS_IC = 0, somatorio_ic = 0;
            for (int p = 0; p < N; p++)
            {
                while (p < N / 2)
                {
                    p++;
                    double V0a_re = N1_va[0];//RMS VA
                    double V0a_im = N2_va[0];
                    double V0a_abs = Math.Sqrt(Math.Pow(V0a_re, 2) + Math.Pow(V0a_im, 2)) / 2;
                    double V0b_re = N1_vb[0];
                    double V0b_im = N2_vb[0];
                    double V0b_abs = Math.Sqrt(Math.Pow(V0b_re, 2) + Math.Pow(V0b_im, 2)) / 2;
                    double V0c_re = N1_vc[0];
                    double V0c_im = N2_vc[0];
                    double V0c_abs = Math.Sqrt(Math.Pow(V0c_re, 2) + Math.Pow(V0c_im, 2)) / 2;

                    double I0a_re = N1_ia[0];//RMS IA
                    double I0a_im = N2_ia[0];
                    double I0a_abs = Math.Sqrt(Math.Pow(I0a_re, 2) + Math.Pow(I0a_im, 2)) / 2;
                    double I0b_re = N1_ib[0];
                    double I0b_im = N2_ib[0];
                    double I0b_abs = Math.Sqrt(Math.Pow(I0b_re, 2) + Math.Pow(I0b_im, 2)) / 2;
                    double I0c_re = N1_ic[0];
                    double I0c_im = N2_ic[0];
                    double I0c_abs = Math.Sqrt(Math.Pow(I0c_re, 2) + Math.Pow(I0c_im, 2)) / 2;
                    if (p > 0)
                    {
                        double Va_re = N1_va[p];//RMS VA
                        double Va_im = N2_va[p];
                        Va_ABS[p] = Math.Sqrt(Math.Pow(Va_re, 2) + Math.Pow(Va_im, 2));
                        casas_a += Math.Pow(Va_ABS[p] * 2 / N, 2);
                        double Vb_re = N1_vb[p];//RMS VB
                        double Vb_im = N2_vb[p];
                        Vb_ABS[p] = Math.Sqrt(Math.Pow(Vb_re, 2) + Math.Pow(Vb_im, 2));
                        casas_b += Math.Pow(Vb_ABS[p] * 2 / N, 2);
                        double Vc_re = N1_vc[p];//RMS VC
                        double Vc_im = N2_vc[p];
                        Vc_ABS[p] = Math.Sqrt(Math.Pow(Vc_re, 2) + Math.Pow(Vc_im, 2));
                        casas_c += Math.Pow(Vc_ABS[p] * 2 / N, 2);

                        double Ia_re = N1_ia[p];//RMS IA
                        double Ia_im = N2_ia[p];
                        Ia_ABS[p] = Math.Sqrt(Math.Pow(Ia_re, 2) + Math.Pow(Ia_im, 2));
                        casas_ia += Math.Pow(Ia_ABS[p] * 2 / N, 2);
                        double Ib_re = N1_ib[p];//RMS IB
                        double Ib_im = N2_ib[p];
                        Ib_ABS[p] = Math.Sqrt(Math.Pow(Ib_re, 2) + Math.Pow(Ib_im, 2));
                        casas_ib += Math.Pow(Ib_ABS[p] * 2 / N, 2);
                        double Ic_re = N1_ic[p];//RMS IC
                        double Ic_im = N2_ic[p];
                        Ic_ABS[p] = Math.Sqrt(Math.Pow(Ic_re, 2) + Math.Pow(Ic_im, 2));
                        casas_ic += Math.Pow(Ic_ABS[p] * 2 / N, 2);
                    }
                    somatorio_a = Math.Sqrt(V0a_abs + (0.5 * casas_a));
                    RMS_VA = Math.Round(somatorio_a, 2);
                    somatorio_b = Math.Sqrt(V0b_abs + (0.5 * casas_b));
                    RMS_VB = Math.Round(somatorio_b, 2);
                    somatorio_c = Math.Sqrt(V0c_abs + (0.5 * casas_c));
                    RMS_VC = Math.Round(somatorio_c, 2);

                    somatorio_ia = Math.Sqrt(I0a_abs + (0.5 * casas_ia));
                    RMS_IA = Math.Round(somatorio_ia, 2);
                    somatorio_ib = Math.Sqrt(I0b_abs + (0.5 * casas_ib));
                    RMS_IB = Math.Round(somatorio_ib, 2);
                    somatorio_ic = Math.Sqrt(I0c_abs + (0.5 * casas_ic));
                    RMS_IC = Math.Round(somatorio_ic, 2);
                    // Calculo de Z
                    Complex Va_Amp = new Complex(N1_va[1], N2_va[1]);
                    Complex Va_ = (Va_Amp * 2 / N);
                    Complex Ia_Amp = new Complex(N1_ia[1], N2_ia[1]);
                    Complex Ia_ = (Ia_Amp * 2 / N);

                    Complex Vb_Amp = new Complex(N1_vb[1], N2_vb[1]);
                    Complex Vb_ = (Vb_Amp * 2 / N);
                    Complex Ib_Amp = new Complex(N1_ib[1], N2_ib[1]);
                    Complex Ib_ = (Ib_Amp * 2 / N);

                    Complex Vc_Amp = new Complex(N1_vc[1], N2_vc[1]);
                    Complex Vc_ = (Vc_Amp * 2 / N);
                    Complex Ic_Amp = new Complex(N1_ic[1], N2_ic[1]);
                    Complex Ic_ = (Ic_Amp * 2 / N);
                    //CALCULO DAS COMPONENTES DE SEQUENCIA
                    Complex complex_va = Va_;   //new Complex(va_true_Re, va_Im);
                    Complex complex_vb = Vb_;   //new Complex(vb_true_Re, vb_Im);
                    Complex complex_vc = Vc_;   //new Complex(vc_true_Re, vc_Im);
                    Complex complex_ia = Ia_;
                    Complex complex_ib = Ib_;
                    Complex complex_ic = Ic_;
                    Complex alpha_1 = Complex.FromPolarCoordinates(1, 120 * (Math.PI / 180)); // new Complex(-(1 / 2), Math.Sqrt(3) / 2)
                    Complex alpha_2 = Complex.FromPolarCoordinates(1, 240 * (Math.PI / 180)); // new Complex(-(1 / 2), -Math.Sqrt(3) / 2)

                    Complex VA_Seq_0_01 = complex_va + complex_vb + complex_vc;
                    Complex VA_Seq_0 = VA_Seq_0_01 / 3;
                    Complex IA_Seq_0_01 = complex_ia + complex_ib + complex_ic;
                    Complex IA_Seq_0 = IA_Seq_0_01 / 3;

                    Complex VA_Seq_1_01 = alpha_1 * complex_vb;
                    Complex VA_Seq_1_02 = alpha_2 * complex_vc;
                    Complex VA_Seq_1_03 = VA_Seq_1_02 + VA_Seq_1_01 + complex_va;
                    Complex VA_Seq_1 = VA_Seq_1_03 / 3;
                    Complex IA_Seq_1_01 = alpha_1 * complex_ib;
                    Complex IA_Seq_1_02 = alpha_2 * complex_ic;
                    Complex IA_Seq_1_03 = IA_Seq_1_02 + IA_Seq_1_01 + complex_ia;
                    Complex IA_Seq_1 = IA_Seq_1_03 / 3;

                    Complex VA_Seq_2_01 = alpha_2 * complex_vb;
                    Complex VA_Seq_2_02 = alpha_1 * complex_vc;
                    Complex VA_Seq_2_03 = VA_Seq_2_01 + VA_Seq_2_02 + complex_va;
                    Complex VA_Seq_2 = VA_Seq_2_03 / 3;
                    Complex IA_Seq_2_01 = alpha_2 * complex_ib;
                    Complex IA_Seq_2_02 = alpha_1 * complex_ic;
                    Complex IA_Seq_2_03 = IA_Seq_2_01 + IA_Seq_2_02 + complex_ia;
                    Complex IA_Seq_2 = IA_Seq_2_03 / 3;

                    Complex VB_Seq_0 = VA_Seq_0;
                    Complex VB_Seq_1 = alpha_2 * VA_Seq_1;
                    Complex VB_Seq_2 = alpha_1 * VA_Seq_2;
                    Complex IB_Seq_0 = IA_Seq_0;
                    Complex IB_Seq_1 = alpha_2 * IA_Seq_1;
                    Complex IB_Seq_2 = alpha_1 * IA_Seq_2;

                    Complex VC_Seq_0 = VA_Seq_0;
                    Complex VC_Seq_1 = alpha_1 * VA_Seq_1;
                    Complex VC_Seq_2 = alpha_2 * VA_Seq_2;
                    Complex IC_Seq_0 = IA_Seq_0;
                    Complex IC_Seq_1 = alpha_1 * IA_Seq_1;
                    Complex IC_Seq_2 = alpha_2 * IA_Seq_2;

                    Complex I0_comp = IA_Seq_0 + IB_Seq_0 + IC_Seq_0;
                    double I0_ABS = I0_comp.Magnitude;


                    textBox2.Text = "" + Va_.Magnitude / (1000 * Math.Sqrt(2));
                    textBox3.Text = "" + Va_.Phase* 180 / Math.PI;
                    textBox4.Text = "" + Vb_.Magnitude/ (1000 * Math.Sqrt(2));
                    textBox5.Text = "" + Vb_.Phase * 180 / Math.PI;
                    textBox6.Text = "" + Vc_.Magnitude/ (1000 * Math.Sqrt(2));
                    textBox7.Text = "" + Vc_.Phase * 180 / Math.PI;

                    textBox8.Text = "" + Ia_.Magnitude / Math.Sqrt(2);
                    textBox9.Text = "" + Ia_.Phase * 180 / Math.PI;
                    textBox10.Text = "" + Ib_.Magnitude / Math.Sqrt(2);
                    textBox11.Text = "" + Ib_.Phase * 180 / Math.PI;
                    textBox12.Text = "" + Ic_.Magnitude / Math.Sqrt(2);
                    textBox13.Text = "" + Ic_.Phase * 180 / Math.PI;
                    polarization_calc();
                    /*
                    for (int x = 0; x < g_time.Length; x++)
                    {
                        try
                        {
                            chart4.Series[0].Points.Clear();
                            chart4.Series[0].Points.AddXY(0, 0);
                            chart4.Series[0].Points.AddXY(-1 * (Va_.Phase * 180 / Math.PI), Va_.Magnitude);
                            chart4.Series[0].Name = "Va: " + " " + Math.Round(Va_.Magnitude, 2) + " V";
                            chart4.Series[1].Points.Clear();
                            chart4.Series[1].Points.AddXY(0, 0);
                            chart4.Series[1].Points.AddXY(-1 * (Vb_.Phase * 180 / Math.PI), Vb_.Magnitude);
                            chart4.Series[1].Name = "Vb: " + " " + Math.Round(Vb_.Magnitude, 2) + " V";
                            chart4.Series[2].Points.Clear();
                            chart4.Series[2].Points.AddXY(0, 0);
                            chart4.Series[2].Points.AddXY(-1 * (Vc_.Phase * 180 / Math.PI), Vc_.Magnitude);
                            chart4.Series[2].Name = "Vc: " + " " + Math.Round(Vc_.Magnitude, 2) + " V";

                            chart4.Series[3].Points.Clear();
                            chart4.Series[3].Points.AddXY(0, 0);
                            chart4.Series[3].Points.AddXY(-1 * (Ia_.Phase * 180 / Math.PI), Ia_.Magnitude);
                            chart4.Series[3].Name = "Ia: " + " " + Math.Round(Ia_.Magnitude, 2) + " A";
                            chart4.Series[4].Points.Clear();
                            chart4.Series[4].Points.AddXY(0, 0);
                            chart4.Series[4].Points.AddXY(-1 * (Ib_.Phase * 180 / Math.PI), Ib_.Magnitude);
                            chart4.Series[4].Name = "Ib: " + " " + Math.Round(Ib_.Magnitude, 2) + " A";
                            chart4.Series[5].Points.Clear();
                            chart4.Series[5].Points.AddXY(0, 0);
                            chart4.Series[5].Points.AddXY(-1 * (Ic_.Phase * 180 / Math.PI), Ic_.Magnitude);
                            chart4.Series[5].Name = "Ic: " + " " + Math.Round(Ic_.Magnitude, 2) + " A";


                        }
                        catch
                        { }
                    */
                    
                }                    
                
            }

        }
    }
}
