using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kargo
{
    public partial class KULLANICIGİRİSİ : Form
    {
        public KULLANICIGİRİSİ()
        {
            InitializeComponent();
        }
        SqlConnection baglanti;
        SqlCommand cmd;
        SqlDataReader dr;

        private void button1_Click(object sender, EventArgs e)
        {
           
            baglanti = new SqlConnection("Data Source=DELLSRV;Initial Catalog=ermed_kargo;User ID=sa;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            cmd = new SqlCommand();
            baglanti.Open();
            cmd.Connection= baglanti;
            cmd .CommandText = "select* From KullaniciGirisi where K_Adi = '" + textBox1.Text + "' and Sifre='" + textBox2.Text + "'";
            dr=cmd.ExecuteReader();
            if (dr.Read())
            {
                LISTELEME LST = new LISTELEME();
                LST.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Kullanıcı Adı veya Şifre hatalı!!!");
            }
            baglanti.Close();



        }
    }
}
