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
    public partial class YURTICI_KARGO : Form
    {
        public YURTICI_KARGO()
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

                //textBox4.Text = sonuc.ToString() + " desi";
                //textBox4.ForeColor = Color.Red;
                var satir = new ListViewItem(sonuc.ToString());

                //dataGridView1.Rows.Add(sonuc.ToString());
                dataGridView1.Visible = true;


                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";

                if (sonuc < 1)
                    dataGridView1.Rows.Add(sonuc, "22.6");
                else if (sonuc >= 1 || sonuc <= 4)
                    dataGridView1.Rows.Add(sonuc, "27.55");
                else if (sonuc == 5)
                    dataGridView1.Rows.Add(sonuc, "30,80");
                else if (sonuc >= 6 || sonuc <= 10)
                    dataGridView1.Rows.Add(sonuc, "33.85");
                else if (sonuc >= 11 || sonuc <= 15)
                    dataGridView1.Rows.Add(sonuc, "38.40");
                else if (sonuc >= 16 || sonuc <= 20)
                    dataGridView1.Rows.Add(sonuc, "47");
                else if (sonuc >= 21 || sonuc <= 25)
                    dataGridView1.Rows.Add(sonuc, "58.75");
                else if (sonuc >= 26 || sonuc <= 30)
                    dataGridView1.Rows.Add(sonuc, "70");


            }
        }

        private void YURTICI_KARGO_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "DESI";
            dataGridView1.Columns[1].Name = "FIYAT";
        }

        private void TEMIZLE_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }
    }
}
