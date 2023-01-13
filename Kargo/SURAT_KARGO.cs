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
    public partial class SURAT_KARGO : Form
    {
        public SURAT_KARGO()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double sonuc = 0;
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Alanları Doldur.");
            }
            else
            {
                double en = Convert.ToDouble(textBox1.Text);
                double boy = Convert.ToDouble(textBox2.Text);
                double yukseklik = Convert.ToDouble(textBox3.Text);

                sonuc = (en * boy * yukseklik) / 3000;

                textBox4.Text = sonuc.ToString();
                textBox4.ForeColor = Color.Red;
                //var satir = new ListViewItem(sonuc.ToString());

                //dataGridView1.Rows.Add(sonuc.ToString());
                //dataGridView1.Visible = true;

                double ekdesı = 70 + (sonuc - 30) * 2.7;
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";

                if (sonuc < 1)
                    textBox5.Text = 25.45.ToString() + " TL";

                else if (sonuc >= 1 && sonuc <= 5)
                    textBox5.Text = 25.65.ToString() + " TL";

                else if (sonuc > 5 && sonuc <= 10)
                    textBox5.Text = 32.33.ToString() + " TL";

                else if (sonuc > 10 && sonuc <= 15)
                    textBox5.Text = 40.37.ToString() + " TL";

                else if (sonuc > 15 && sonuc <= 20)
                    textBox5.Text = 46.44.ToString() + " TL";


                else if (sonuc > 20 && sonuc <= 25)
                    textBox5.Text = 53.53.ToString() + " TL";

                else if (sonuc > 25 && sonuc <= 30)
                    textBox5.Text = 63.ToString() + " TL";

                else if (sonuc > 30)
                    textBox5.Text = ekdesı.ToString() + " TL";


            }
        }

        private void SURAT_KARGO_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "DESI";
            dataGridView1.Columns[1].Name = "FIYAT";
        }

        private void TEMIZLE_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            KARGO_SIRKETLERI KS = new KARGO_SIRKETLERI();
            KS.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(textBox4.Text, textBox5.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
            }
            textBox6.Text = toplam.ToString() + "TL";
        }
    }
}
