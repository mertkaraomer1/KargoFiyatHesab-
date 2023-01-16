using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
namespace Kargo
{
    public partial class MNG_KARGO : Form
    {
        public MNG_KARGO()
        {
            InitializeComponent();
        }
        double sonuc = 0;
        double ekdesı = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            
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

                ekdesı = 80 + (sonuc - 30) * 3;
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";

                if (sonuc == 0 && sonuc < 1)
                    textBox5.Text = 30.ToString();

                else if (sonuc >= 1 && sonuc <= 5)
                    textBox5.Text = 32.ToString();

                else if (sonuc > 5 && sonuc <= 10)
                    textBox5.Text = 36.ToString();

                else if (sonuc > 10 && sonuc <= 20)
                    textBox5.Text = 47.ToString();

                else if (sonuc > 20 && sonuc <= 30)
                    textBox5.Text = 62.ToString();

                else if (sonuc > 30 && sonuc <=40)
                    textBox5.Text = 80.ToString();

                else if ( sonuc > 40)
                    textBox5.Text = ekdesı.ToString();

            }
        }

        private void MNG_KARGO_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].Name = "DESI";
            dataGridView1.Columns[1].Name = "FIYAT TL";
            dataGridView1.Columns[2].Name = "ADET";
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
            int adet = 1;
            adet = Convert.ToInt32(textBox7.Text);
            double desı =Convert.ToDouble(textBox4.Text);
            if (textBox4 != null)
            {
                ekdesı = 80 + (desı - 40) * 3;
                if (desı == 0 && desı < 1)
                    textBox5.Text = (adet*30).ToString();

                else if (desı >= 1 && desı <= 5)
                    textBox5.Text = (adet*32).ToString();

                else if (desı > 5 && desı <= 10)
                    textBox5.Text = (adet * 36).ToString();

                else if (desı > 10 && desı <= 20)
                    textBox5.Text = (adet * 47).ToString();

                else if (desı > 20 && desı <= 30)
                    textBox5.Text = (adet * 62).ToString();

                else if (desı > 30 && desı <= 40)
                    textBox5.Text = (adet * 80).ToString();

                else if ( desı > 40)
                    textBox5.Text = (adet*ekdesı).ToString();
            }
            dataGridView1.Rows.Add(desı,textBox5.Text,adet);
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

        private void MNG_KARGO_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Çıkmak istediğinize emin misiniz?", "www.kaizen40.com",
                MessageBoxButtons.YesNo) == DialogResult.No)
            {
                e.Cancel = true;

                // iptal ederseniz ne yapacağınızı buraya yazın
            }

            // Evet' i tıklarsanız çıkarsınız
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook kitap = excel.Workbooks.Add(Missing);
            Worksheet sayfa = (Worksheet)kitap.Sheets[1];
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                Range alan = (Range)sayfa.Cells[1, 1];
                alan.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    Range alan2 = (Range)sayfa.Cells[j + 1, i + 1];
                    alan2.Cells[2, 1] = dataGridView1[i, j].Value;
                }
            }
        }
    }
}
