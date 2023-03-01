﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Data.SqlClient;

namespace Kargo
{
    public partial class ERGULKARGO : Form
    {
        public ERGULKARGO()
        {
            InitializeComponent();
        }
        DataSet ds;
        OleDbCommand cmd;
        OleDbConnection con;
        double adet;
        double fiyat;
        double desi;
        SqlConnection baglanti;

        private void ERGULKARGO_Load(object sender, EventArgs e)
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=fiyat_listesi.accdb");
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter("select * from ERGULKARGOİLLER ORDER BY ID ASC ", con);
            da.Fill(dt);
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "iller";
            comboBox1.DataSource = dt;




            dataGridView1.ColumnCount = 8;
            dataGridView1.Columns[0].Name = "FİRMA ADI";
            dataGridView1.Columns[1].Name = "DESI/KİLO";
            dataGridView1.Columns[2].Name = "FIYAT";
            dataGridView1.Columns[3].Name = "TL";
            dataGridView1.Columns[4].Name = "ADET";
            dataGridView1.Columns[5].Name = "DEPO";
            dataGridView1.Columns[6].Name = "İL";
            dataGridView1.Columns[7].Name = "TARİH";
            comboBox2.Items.Add("ANA DEPO");
            comboBox2.Items.Add("DMO");
        }

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



                textBox4.Text = Math.Round((en * boy * yukseklik) / 3000, 0).ToString();
                desi = Convert.ToDouble(textBox4.Text);
            }
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {

            adet=Convert.ToDouble(textBox7.Text);
            desi = Convert.ToDouble(textBox4.Text);
            if (desi > 1 && desi <= 33)
            {
                if (comboBox1.Text == "Bursa")
                {
                    if(adet > 0 && adet <= 5) 
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (56 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (44 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet >  40)
                    {
                        textBox5.Text = Math.Round(adet * (39 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "İstanbul Avrupa")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (50 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (40 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (35 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "İstanbul Anadolu")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (50 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (40 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (35 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Düzce")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (50 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (40 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (35 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Ankara")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (56 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (44 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (39 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Denizli")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (80 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (49 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "İzmir")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (80 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (49 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Balıkesir")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (80 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (49 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Manisa")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (60 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Tekirdağ")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (56 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (44 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (39 * 1.18 * 1.0235), 2).ToString();
                    }
                }

            }
            else if (desi > 33 && desi <= 50)
            {
                if (comboBox1.Text == "Bursa")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100  * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "İstanbul Avrupa")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (125 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (90 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (60 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "İstanbul Anadolu")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (125 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (90 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (60 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Düzce")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (125 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (90 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (85 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (60 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (55 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Ankara")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Denizli")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "İzmir")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Balıkesir")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Manisa")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
                else if (comboBox1.Text == "Tekirdağ")
                {
                    if (adet > 0 && adet <= 5)
                    {
                        textBox5.Text = Math.Round(adet * (140 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 5 && adet <= 10)
                    {
                        textBox5.Text = Math.Round(adet * (100 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 10 && adet <= 20)
                    {
                        textBox5.Text = Math.Round(adet * (95 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 20 && adet <= 40)
                    {
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.0235), 2).ToString();
                    }
                    else if (adet > 40)
                    {
                        textBox5.Text = Math.Round(adet * (65 * 1.18 * 1.0235), 2).ToString();
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string Depo=comboBox2.SelectedItem.ToString();
            string TL = "TL";
            dataGridView1.Rows.Add(textBox8.Text, desi, textBox5.Text, TL, adet,Depo, comboBox1.Text,DateTime.Now.ToString("yyyy-MM-dd"));
        }  
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DELLSRV;Initial Catalog=ermed_kargo;User ID=sa;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (baglanti.State == ConnectionState.Open)
                    {
                        baglanti.Close();
                    }

                    var G_firma = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value); // 2. kolon
                    var K_Firma = Convert.ToString("ERGÜL KARGO"); // 3. kolon
                    var Desi = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value).ToString(); // 4. kolon
                    var Fiyat = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value).ToString(); // 5. kolon
                    var Adet = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value).ToString(); // 6. kolon
                    var Depo = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                    var il = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value); // 7. kolon
                    var Tarih = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); // 9. kolon

                    baglanti.Open();
                    SqlCommand komut = new SqlCommand("INSERT INTO Kargolar (Gonderilcek_firma,Kargo_Sirketi,Desi_KG,Fiyat,Adet,Depo,İL,Tarih) VALUES ('" + G_firma + "' , '" + K_Firma + "','" + Desi + "' , '" + Fiyat + "' , '" + Adet + "','" + Depo + "','" + il + "' ,'" + Tarih + "')", baglanti);
                    komut.ExecuteNonQuery();
                }
                baglanti.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("HATA VAR!!!");
            }
            finally
            {
                MessageBox.Show("BAŞARI İLE KAYDEDİLDİ...");
            }
        }
        StringFormat strFormat;
        ArrayList arrColumnLefts = new ArrayList();
        ArrayList arrColumnWidths = new ArrayList();
        int iCellHeight = 0;
        int iTotalWidth = 0;
        int iRow = 0;
        bool bFirstPage = false;
        bool bNewPage = false;
        int iHeaderHeight = 0;

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            System.Drawing.Image Logo = imageList1.Images["ASPİRASYON SONDASI.PNG"];
            Pen kalem = new Pen(Color.Black);
            Font font = new Font("Arial", 15);
            SolidBrush firca = new SolidBrush(Color.Black);
            int iCount = 0;
            int iTopMargin = e.MarginBounds.Top;
            try
            {
                int iLeftMargin = e.MarginBounds.Left;

                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;
                bFirstPage = true;


                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dataGridView1.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;


                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }

                while (iRow <= dataGridView1.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView1.Rows[iRow];

                    iCellHeight = GridRow.Height + 5;


                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {


                            e.Graphics.DrawString("ERGÜL KARGO FİYAT HESABI", new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("ERGÜL KARGO FİYAT HESABI", new Font(dataGridView1.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            e.Graphics.DrawString("ERMED TIP MEDİKAL", font,
                                  Brushes.Black, e.MarginBounds.Top + 225, e.MarginBounds.Top -
                                  e.Graphics.MeasureString("ERMED TIP MEDİKAL", font, e.MarginBounds.Left).Height + 10);

                            String strDate = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                            e.Graphics.DrawString(strDate, new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView1.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("KARGO FİYAT HESABI", new Font(new Font(dataGridView1.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);


                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView1.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;

                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }

                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            iCount++;
                        }

                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }


                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Graphics mg = Graphics.FromImage(Logo);
            e.Graphics.DrawImage(Logo, 720, 10, 80, 80);






            e.Graphics.DrawString("TOPLAM FİYAT=", font, firca, iCellHeight + 390, iTopMargin + 10);
            e.Graphics.DrawString(textBox6.Text, font, firca, iCellHeight + 570, iTopMargin + 10);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin, iCellHeight + 360, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 695, iTopMargin, iCellHeight + 695, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 570, iTopMargin, iCellHeight + 570, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin, iCellHeight + 695, iTopMargin);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin + 40, iCellHeight + 695, iTopMargin + 40);
        }

        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dataGridView1.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            PrintPreviewDialog onizleme = new PrintPreviewDialog();
            onizleme.Document = printDocument1;
            ((Form)onizleme).WindowState = FormWindowState.Maximized; // Tam ekran olması için
            onizleme.PrintPreviewControl.Zoom = 1.0; //Sayfanın %100 boyutunda olması için
            onizleme.ShowDialog();
            PrintDialog yazdir = new PrintDialog();
            yazdir.Document = printDocument1;
            yazdir.UseEXDialog = true;
            if (yazdir.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
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
                Range alan4 = (Range)sayfa.Cells[1, 11];
                alan4.Value2 = "TOPLAM FİYAT";
                Range alan3 = (Range)sayfa.Cells[2, 11];
                alan3.Value2 = textBox6.Text;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                toplam = Math.Round(toplam, 2);
            }
            toplam = Math.Round(toplam, 2);
            textBox6.Text = toplam.ToString() + " TL";
        }

        private void TEMIZLE_Click(object sender, EventArgs e)
        {

            dataGridView1.Rows.Clear();
            textBox6.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            KARGO_SIRKETLERI KS = new KARGO_SIRKETLERI();
            KS.Show();
            this.Hide();
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.Green;

        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;

        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            button5.BackColor = Color.Blue;

        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            button5.BackColor = Color.White;

        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.BackColor = Color.Green;

        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackColor = Color.White;

        }

        private void TEMIZLE_MouseEnter(object sender, EventArgs e)
        {
            TEMIZLE.BackColor = Color.Gold;

        }

        private void TEMIZLE_MouseLeave(object sender, EventArgs e)
        {
            TEMIZLE.BackColor = Color.White;

        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackColor = Color.Green;

        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;

        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.Red;

        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;

        }

        private void ERGULKARGO_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else
                e.Cancel = true;//Çıkışı durdur
        }


    }
}