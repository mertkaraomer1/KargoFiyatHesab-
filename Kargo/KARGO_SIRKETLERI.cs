﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Kargo
{
    public partial class KARGO_SIRKETLERI : Form
    {


        public KARGO_SIRKETLERI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            YURTICI_KARGO YK = new YURTICI_KARGO();
            YK.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ARAS_KARGO ARAS = new ARAS_KARGO();
            ARAS.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SURAT_KARGO SURAT = new SURAT_KARGO();
            SURAT.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            MNG_KARGO MNG = new MNG_KARGO();
            MNG.Show();
            this.Hide();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ANKARA_KARGO ANKR = new ANKARA_KARGO();
            ANKR.Show();
            this.Hide();

        }


        private void KARGO_SIRKETLERI_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else
                e.Cancel = true;//Çıkışı durdur

        }





        private void button6_Click(object sender, EventArgs e)
        {
            FILTER FTR = new FILTER();
            FTR.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            CAN_KARGO CN = new CAN_KARGO();
            CN.Show();
            this.Hide();
        }


        private void button9_Click(object sender, EventArgs e)
        {
            KULLANICIGİRİSİ KG = new KULLANICIGİRİSİ();
            KG.Show();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            UPS_KARGO UPS=new UPS_KARGO();
            UPS.Show();
            this.Hide();
        }
    }
}
