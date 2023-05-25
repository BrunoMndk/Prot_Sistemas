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

namespace Prot_Sistemas
{
    public partial class Form1 : Form
    {
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
            comboBox1.SelectedItem = comboBox1.Items[0];

            chart1.ChartAreas[0].AxisX.Maximum = 50;
            chart1.ChartAreas[0].AxisX.Minimum = -30;
            chart1.ChartAreas[0].AxisY.Maximum = 50;
            chart1.ChartAreas[0].AxisY.Minimum = -30;
            chart2.ChartAreas[0].AxisX.Maximum = 50;
            chart2.ChartAreas[0].AxisX.Minimum = -30;
            chart2.ChartAreas[0].AxisY.Maximum = 50;
            chart2.ChartAreas[0].AxisY.Minimum = -30;

            

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
            if (!e.Button.HasFlag(MouseButtons.Left)) return;
            Axis ax = chart1.ChartAreas[0].AxisX;
            Axis ay = chart1.ChartAreas[0].AxisY;
            double range_x = ax.Maximum - ax.Minimum;
            double range_y = ay.Maximum - ay.Minimum;
            double xv = ax.PixelPositionToValue(e.Location.X);
            double yv = ay.PixelPositionToValue(e.Location.Y);
            ax.Minimum -= Math.Round(xv - mDown_x,0);
            ax.Maximum = Math.Round(ax.Minimum + range_x,0);

            ay.Minimum -= Math.Round(yv - mDown_y, 0);
            ay.Maximum = Math.Round(ay.Minimum + range_y, 0);
        }
        private void chart2_MouseDown(object sender, MouseEventArgs e)
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
        private void chart2_MouseMove(object sender, MouseEventArgs e)
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
        void import()
        {

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
            Complex Vab_mem = Va_mem - Vb_mem;
            Complex Vbc_mem = Vb_mem - Vc_mem;
            Complex Vca_mem = Vc_mem - Va_mem;
            Complex Vab = Va - Vb;            
            Complex Vbc = Vb - Vc;
            Complex Vca = Vc - Va;
            Complex Va_zero = (1 / 3) * (Va + Vb + Vc);
            Ia = new Complex(Convert.ToDouble(textBox8.Text) * Math.Cos(Convert.ToDouble(textBox9.Text) * Math.PI / 180), Convert.ToDouble(textBox8.Text) * Math.Sin(Convert.ToDouble(textBox9.Text) * Math.PI / 180));
            Ib = new Complex(Convert.ToDouble(textBox10.Text) * Math.Cos(Convert.ToDouble(textBox11.Text) * Math.PI / 180), Convert.ToDouble(textBox10.Text) * Math.Sin(Convert.ToDouble(textBox11.Text) * Math.PI / 180));
            Ic = new Complex(Convert.ToDouble(textBox12.Text) * Math.Cos(Convert.ToDouble(textBox13.Text) * Math.PI / 180), Convert.ToDouble(textBox12.Text) * Math.Sin(Convert.ToDouble(textBox13.Text) * Math.PI / 180));
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

                        //dual mho BC                       
                        k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                        Complex Sop_dual_BC = Vbc + (k_pol * Va);
                        Complex Zpol_dual_BC = k_pol * Va / Ibc;
                        Complex[] Zs_dual_BC = { -Zpol_dual_BC, Zm_BC[1] };
                        double r_dual_BC = (Complex.Abs(Zr - Zs_dual_BC[0])) / 2;
                        for (int k = 0; k <= 1000; k++)
                        {
                            double x = (((Zm_BC[1].Real + ((Zr * Ibc - Vbc) / Ibc).Real + Zs_dual_AB[0].Real) / 2)) + r_dual_BC * Math.Cos(k * 2 * Math.PI / 1000);
                            double y = (((Zm_BC[1].Imaginary + ((Zr * Ibc - Vbc) / Ibc).Imaginary + Zs_dual_BC[0].Imaginary) / 2)) + r_dual_BC * Math.Sin(k * 2 * Math.PI / 1000);
                            chart2.Series["Mho_BC"].Points.AddXY(x, y);
                        }
                        chart2.Series["Zs_BC"].Points.AddXY(0, 0);
                        chart2.Series["Zs_BC"].Points.AddXY(Zs_dual_BC[0].Real, Zs_dual_BC[0].Imaginary);
                        chart2.Series["Zm_BC"].Points.AddXY(Zs_dual_BC[0].Real, Zs_dual_BC[0].Imaginary);
                        chart2.Series["Zm_BC"].Points.AddXY(Zs_dual_BC[1].Real, Zs_dual_BC[1].Imaginary);
                        chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
                        chart2.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);

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
                    break;
                case 1: // Dual alternative
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
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
                    break;
                case 2: // Cross polarization
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
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


                        //Cross mho AG
                        k_pol = new Complex(kpol_abs * Math.Cos(90 * Math.PI / 180), kpol_abs * Math.Sin(90 * Math.PI / 180));
                        Complex Sop_crz_AG = Va + (k_pol * Vbc);
                        Complex Zpol_crz_AG = k_pol * Vbc / IaG;
                        Complex Zs_crz_AG = Zm_AG[1] - Zpol_crz_AG;
                        double r_crz_AG = (Complex.Abs(-Zr + Zm_AG[1] - Zpol_crz_AG)) / 2;
                        for (int k = 0; k <= 1000; k++)
                        {
                            double x = (((Zr + Zm_AG[1] -Zpol_crz_AG).Real / 2)) + r_crz_AG * Math.Cos(k * 2 * Math.PI / 1000);
                            double y = (((Zr + Zm_AG[1] - Zpol_crz_AG).Imaginary / 2)) + r_crz_AG * Math.Sin(k * 2 * Math.PI / 1000);
                            chart2.Series["Mho_AG"].Points.AddXY(x, y);
                        }
                        chart2.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
                        chart2.Series["Zm_AG"].Points.AddXY(Zs_crz_AG.Real, Zs_crz_AG.Imaginary);
                        chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
                        chart2.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
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
                    break;
                case 3:// Cross polarization with voltage memory
                    {
                        textBox21.Font = textBox20.Font;
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
                    break;
                case 4:// positive sequence
                    {
                        textBox21.Font = new Font(textBox20.Font, FontStyle.Strikeout);
                        //cross mho AB                       
                        k_pol = new Complex(kpol_abs * Math.Cos(-90 * Math.PI / 180), kpol_abs * Math.Sin(-90 * Math.PI / 180));
                        Complex Sop_pos_AB = Vab + (k_pol * Vc);
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
                    break;
                case 5:// positive sequence with voltage memory
                    {

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

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }
    }
}
