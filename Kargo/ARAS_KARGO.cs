using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
namespace Kargo
{
    public partial class ARAS_KARGO : Form
    {
        public ARAS_KARGO()
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

                double ekdesı = 70 + (sonuc - 30) * 2.94;
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";

                if (sonuc < 1)
                    textBox5.Text = 21.71.ToString();

                else if (sonuc >= 1 && sonuc <= 5)
                    textBox5.Text = 39.69.ToString();

                else if (sonuc > 5 && sonuc <= 10)
                    textBox5.Text = 58.57.ToString();

                else if (sonuc > 10 && sonuc <= 15)
                    textBox5.Text = 62.36.ToString();

                else if (sonuc > 15 && sonuc <= 20)
                    textBox5.Text = 67.9.ToString();


                else if (sonuc > 20 && sonuc <= 25)
                    textBox5.Text = 78.04.ToString();

                else if (sonuc > 25 && sonuc <= 30)
                    textBox5.Text = 88.65.ToString();

                else if (sonuc > 30)
                    textBox5.Text = ekdesı.ToString();


            }
        }

        private void ARAS_KARGO_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].Name = "FİRMA ADI";
            dataGridView1.Columns[1].Name = "DESI";
            dataGridView1.Columns[2].Name = "FIYAT TL";
            dataGridView1.Columns[3].Name = "ADET";
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
            double desı = Convert.ToDouble(textBox4.Text);

            if (textBox4 != null)
            {
                double ekdesı = 70 + (desı - 30) * 2.94;
                if (desı < 1)
                    textBox5.Text = (adet*21.71).ToString();

                else if (desı >= 1 && desı <= 5)
                    textBox5.Text = (adet*39.69).ToString();

                else if (desı > 5 && desı <= 10)
                    textBox5.Text = (adet * 58.57).ToString();

                else if (desı > 10 && desı <= 15)
                    textBox5.Text = (adet * 62.36).ToString();

                else if (desı > 15 && desı <= 20)
                    textBox5.Text = (adet * 67.9).ToString();


                else if (desı > 20 && desı <= 25)
                    textBox5.Text = (adet * 78.04).ToString();

                else if (desı > 25 && desı <= 30)
                    textBox5.Text = (adet * 88.65).ToString();

                else if (desı > 30)
                    textBox5.Text = (adet*ekdesı).ToString();
            }
            dataGridView1.Rows.Add(textBox8.Text,desı, textBox5.Text,adet);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
            }
            textBox6.Text = toplam.ToString() + "TL";

        
        }

        private void ARAS_KARGO_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else
                e.Cancel = true;//Çıkışı durdur
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
                Range alan4 = (Range)sayfa.Cells[1, 6];
                alan4.Value2 = "TOPLAM FİYAT";
                Range alan3 = (Range)sayfa.Cells[2, 6];
                alan3.Value2 = textBox6.Text;
            }

        }
    }
}
