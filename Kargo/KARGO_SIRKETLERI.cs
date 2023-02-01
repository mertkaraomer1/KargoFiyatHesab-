using System;
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
            YURTICI_KARGO YK=new YURTICI_KARGO();
            YK.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ARAS_KARGO ARAS=new ARAS_KARGO();
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


        private void KARGO_SIRKETLERI_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            label1.Text = "ERMED TIP MEDİKAL KARGO FİYAT HESABI...";

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = label1.Text.Substring(1) + label1.Text.Substring(0, 1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FILTER FTR = new FILTER();
            FTR.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            CAN_KARGO CN= new CAN_KARGO();
            CN.Show();
            this.Hide();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ERGULKARGO ERG= new ERGULKARGO();
            ERG.Show();
            this.Hide();
        }
    }
}
