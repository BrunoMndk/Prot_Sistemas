﻿using System;
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

namespace Prot_Sistemas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true;
            polarization_calc();            
        }
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                polarization_calc();
            }
        }

        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            var chart = (Chart)sender;
            var xAxis = chart.ChartAreas[0].AxisX;
            var yAxis = chart.ChartAreas[0].AxisY;
            try
            {
                if (chart1.ChartAreas[0].AxisX.Maximum <= 60)
                {
                    chart1.ChartAreas[0].AxisX.Interval = 10;
                    chart1.ChartAreas[0].AxisY.Interval = 10;
                }
                if ((chart1.ChartAreas[0].AxisX.Maximum > 61)&&(chart1.ChartAreas[0].AxisX.Maximum < 90))
                {
                    chart1.ChartAreas[0].AxisX.Interval = 20;
                    chart1.ChartAreas[0].AxisY.Interval = 20;
                }
                if ((chart1.ChartAreas[0].AxisX.Maximum > 91)&&(chart1.ChartAreas[0].AxisX.Maximum > 200))
                {
                    chart1.ChartAreas[0].AxisX.Interval = 50;
                    chart1.ChartAreas[0].AxisY.Interval = 50;
                }
                if (chart1.ChartAreas[0].AxisX.Maximum > 20)
                {
                    if (e.Delta < 0) // Scrolled down.
                    {
                        chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Maximum - 10;
                        chart1.ChartAreas[0].AxisX.Minimum = chart1.ChartAreas[0].AxisX.Minimum + 10;
                        chart1.ChartAreas[0].AxisY.Maximum = chart1.ChartAreas[0].AxisY.Maximum - 10;
                        chart1.ChartAreas[0].AxisY.Minimum = chart1.ChartAreas[0].AxisY.Minimum + 10;

                    }                    
                }
                if (chart1.ChartAreas[0].AxisX.Maximum >= 20)
                {
                    if (e.Delta > 0) // Scrolled up.
                    {
                        chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Maximum + 10;
                        chart1.ChartAreas[0].AxisX.Minimum = chart1.ChartAreas[0].AxisX.Minimum - 10;
                        chart1.ChartAreas[0].AxisY.Maximum = chart1.ChartAreas[0].AxisY.Maximum + 10;
                        chart1.ChartAreas[0].AxisY.Minimum = chart1.ChartAreas[0].AxisY.Minimum - 10;
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
        void polarization_calc()
        {
            for (int x = 0; x < 14; x++)
            {
                chart1.Series[x].Points.Clear();
            }
            

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

            //Plot LT
            chart1.Series["LT"].Points.AddXY(0, 0);
            chart1.Series["LT"].Points.AddXY(LT_Z1.Real, LT_Z1.Imaginary);
            //Plot MHO
            double radius = Complex.Abs(Zr) / 2;
            for (int k = 0; k <= 1000; k++)
            {
                double x = (radius * Math.Cos(Zr.Phase)) + (radius) * Math.Cos(k * 2 * Math.PI / 1000);
                double y = (radius * Math.Sin(Zr.Phase)) + radius * Math.Sin(k * 2 * Math.PI / 1000);
                chart1.Series[0].Points.AddXY(x, y);
            }
            // loop AG
            Complex IaG = Ia + (k0 * In);
            Complex[] Zm_AG =  { 0,Va / IaG};
            chart1.Series["Zm_AG"].Points.AddXY(0, 0);
            chart1.Series["Zm_AG"].Points.AddXY(Zm_AG[1].Real, Zm_AG[1].Imaginary);
            Complex[] Sop_AG = { Zm_AG[1], (Zm_AG[1] + (((Zr * IaG) - Va) / IaG)) };
            chart1.Series["Zpol_AG"].Points.AddXY(Sop_AG[0].Real, Sop_AG[0].Imaginary);
            chart1.Series["Zpol_AG"].Points.AddXY(Sop_AG[1].Real, Sop_AG[1].Imaginary);
            // loop BG
            Complex IbG = Ib + (k0 * In);
            Complex[] Zm_BG = { 0, Vb / IbG };
            chart1.Series["Zm_BG"].Points.AddXY(0, 0);
            chart1.Series["Zm_BG"].Points.AddXY(Zm_BG[1].Real, Zm_BG[1].Imaginary);
            Complex[] Sop_BG = { Zm_BG[1], (Zm_BG[1] + (((Zr * IbG) - Vb) / IbG)) };
            chart1.Series["Zpol_BG"].Points.AddXY(Sop_BG[0].Real, Sop_BG[0].Imaginary);
            chart1.Series["Zpol_BG"].Points.AddXY(Sop_BG[1].Real, Sop_BG[1].Imaginary);
            // loop CG
            Complex IcG = Ic + (k0 * In);
            Complex[] Zm_CG = { 0, Vc / IbG };
            chart1.Series["Zm_CG"].Points.AddXY(0, 0);
            chart1.Series["Zm_CG"].Points.AddXY(Zm_CG[1].Real, Zm_CG[1].Imaginary);
            Complex[] Sop_CG = { Zm_CG[1], (Zm_CG[1] + (((Zr * IcG) - Vc) / IcG)) };
            chart1.Series["Zpol_CG"].Points.AddXY(Sop_CG[0].Real, Sop_CG[0].Imaginary);
            chart1.Series["Zpol_CG"].Points.AddXY(Sop_CG[1].Real, Sop_CG[1].Imaginary);
            // loop AB
            Complex Vab = Va - Vb;
            Complex Iab = Ia - Ib;
            Complex[] Zm_AB = { 0, Vab / Iab };
            chart1.Series["Zm_AB"].Points.AddXY(0, 0);
            chart1.Series["Zm_AB"].Points.AddXY(Zm_AB[1].Real, Zm_AB[1].Imaginary);
            Complex[] Sop_AB = { Zm_AB[1], (Zm_AB[1] + (((Zr * Iab) - Vab) / Iab)) };
            chart1.Series["Zpol_AB"].Points.AddXY(Sop_AB[0].Real, Sop_AB[0].Imaginary);
            chart1.Series["Zpol_AB"].Points.AddXY(Sop_AB[1].Real, Sop_AB[1].Imaginary);
            // loop BC
            Complex Vbc = Vb - Vc;
            Complex Ibc = Ib - Ic;
            Complex[] Zm_BC = { 0, Vbc / Ibc };
            chart1.Series["Zm_BC"].Points.AddXY(0, 0);
            chart1.Series["Zm_BC"].Points.AddXY(Zm_BC[1].Real, Zm_BC[1].Imaginary);
            Complex[] Sop_BC = { Zm_BC[1], (Zm_BC[1] + (((Zr * Ibc) - Vbc) / Ibc)) };
            chart1.Series["Zpol_BC"].Points.AddXY(Sop_BC[0].Real, Sop_BC[0].Imaginary);
            chart1.Series["Zpol_BC"].Points.AddXY(Sop_BC[1].Real, Sop_BC[1].Imaginary);
            // loop CA
            Complex Vca = Vc - Va;
            Complex Ica = Ic - Ia;
            Complex[] Zm_CA = { 0, Vca / Ica };
            chart1.Series["Zm_CA"].Points.AddXY(0, 0);
            chart1.Series["Zm_CA"].Points.AddXY(Zm_CA[1].Real, Zm_CA[1].Imaginary);
            Complex[] Sop_CA = { Zm_CA[1], (Zm_CA[1] + (((Zr * Ica) - Vca) / Ica)) };
            chart1.Series["Zpol_CA"].Points.AddXY(Sop_CA[0].Real, Sop_CA[0].Imaginary);
            chart1.Series["Zpol_CA"].Points.AddXY(Sop_CA[1].Real, Sop_CA[1].Imaginary);
        }
        void loops_fault()
        {
            //AB
            if (radioButton1.Checked == true)
            {
                chart1.Series["Zpol_AB"].Enabled = true;
                chart1.Series["Zm_AB"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_AB"].Enabled = false;
                chart1.Series["Zm_AB"].Enabled = false;
            }
            //BC
            if (radioButton2.Checked == true)
            {
                chart1.Series["Zpol_BC"].Enabled = true;
                chart1.Series["Zm_BC"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_BC"].Enabled = false;
                chart1.Series["Zm_BC"].Enabled = false;
            }
            //CA
            if (radioButton3.Checked == true)
            {
                chart1.Series["Zpol_CA"].Enabled = true;
                chart1.Series["Zm_CA"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_CA"].Enabled = false;
                chart1.Series["Zm_CA"].Enabled = false;
            }
            //AG
            if (radioButton4.Checked == true)
            {
                chart1.Series["Zpol_AG"].Enabled = true;
                chart1.Series["Zm_AG"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_AG"].Enabled = false;
                chart1.Series["Zm_AG"].Enabled = false;
            }
            //BG
            if (radioButton5.Checked == true)
            {
                chart1.Series["Zpol_BG"].Enabled = true;
                chart1.Series["Zm_BG"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_BG"].Enabled = false;
                chart1.Series["Zm_BG"].Enabled = false;
            }
            //CG
            if (radioButton6.Checked == true)
            {
                chart1.Series["Zpol_CG"].Enabled = true;
                chart1.Series["Zm_CG"].Enabled = true;
            }
            else
            {
                chart1.Series["Zpol_CG"].Enabled = false;
                chart1.Series["Zm_CG"].Enabled = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            loops_fault();
        }

       
    }
}
