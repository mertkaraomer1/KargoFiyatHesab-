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

        private void KARGO_SIRKETLERI_FormClosing(object sender, FormClosingEventArgs e)
        {
             if (MessageBox.Show("Çıkmak istediğinize emin misiniz?", "www.kaizen40.com",
                MessageBoxButtons.YesNo) == DialogResult.No)
            {
                e.Cancel = true;

                // iptal ederseniz ne yapacağınızı buraya yazın
            }

            // Evet' i tıklarsanız çıkarsınız

        }
    }
}
