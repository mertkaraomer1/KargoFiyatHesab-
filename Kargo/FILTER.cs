using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;
using System.Runtime.ConstrainedExecution;
using Microsoft.VisualBasic.Logging;
using static System.Net.Mime.MediaTypeNames;
namespace Kargo
{
    public partial class FILTER : Form
    {
        public FILTER()
        {
            InitializeComponent();
        }

        private void FILTER_Load(object sender, EventArgs e)
        {



            dataGridView2.ColumnCount = 7;
            dataGridView2.Columns[0].Name = "FİRMA ADI";
            dataGridView2.Columns[1].Name = "KARGO FİRMASI";
            dataGridView2.Columns[2].Name = "DESİ";
            dataGridView2.Columns[3].Name = "FİYAT TL";
            dataGridView2.Columns[4].Name = "ADET";
            dataGridView2.Columns[5].Name = "İL";
            dataGridView2.Columns[6].Name = "İLÇE";

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
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Double desi = Convert.ToDouble(textBox4.Text);
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 11;
            dataGridView1.Columns[0].Name = "KARGO ŞİRKETLERİ";
            dataGridView1.Columns[1].Name = "0 DESİ/KG";
            dataGridView1.Columns[2].Name = "1-4 DESİ/KG";
            dataGridView1.Columns[3].Name = "5 DESİ/KG";
            dataGridView1.Columns[4].Name = "6-10 DESİ/KG";
            dataGridView1.Columns[5].Name = "11-15 DESİ/KG";
            dataGridView1.Columns[6].Name = "16-20 DESİ/KG";
            dataGridView1.Columns[7].Name = "21-25 DESİ/KG";
            dataGridView1.Columns[8].Name = "26-30 DESİ/KG";
            dataGridView1.Columns[9].Name = "31-40 DESİ/KG";
            dataGridView1.Columns[10].Name = "41-50 DESİ/KG";



            if (textBox4.Text != null)
            {
                dataGridView1.Rows.Add("MNG KARGO", 32.ToString(), 32.ToString(), 32.ToString(), 35.ToString(), 38.ToString(), 45.ToString(), 55.ToString(), 60.ToString(), 75.ToString(), (75 + (desi - 40) * 2.30).ToString());
                dataGridView1.Rows.Add("ARAS KARGO", 21.71.ToString(), 39.69.ToString(), 39.69.ToString(), 58.57.ToString(), 62.36.ToString(), 67.9.ToString(), 78.04.ToString(), 88.ToString(), (88.65 + (desi - 30) * 2.94).ToString(), (88.65 + (desi - 30) * 2.94).ToString());
                dataGridView1.Rows.Add("SÜRAT KARGO", 25.45.ToString(), 25.45.ToString(), 25.45.ToString(), 32.33.ToString(), 40.37.ToString(), 46.44.ToString(), 53.53.ToString(), 63.ToString(), (63 + (desi - 30) * 2.7).ToString(), (63 + (desi - 30) * 2.7).ToString());
                dataGridView1.Rows.Add("YURTİÇİ KARGO", 22.6.ToString(), 27.55.ToString(), 30.80.ToString(), 33.85.ToString(), 38.40.ToString(), 47.ToString(), 58.75.ToString(), 70.ToString(), (63 + (desi - 30) * 2.7).ToString(), (63 + (desi - 30) * 2.7).ToString());
                dataGridView1.Rows.Add("ANKARA KARGO", 22.74.ToString(), 22.74.ToString(), 22.74.ToString(), 22.74.ToString(), 34.99.ToString(), 34.99.ToString(), 50.74.ToString(), 50.74.ToString(), 59.48.ToString(), 73.48.ToString());
                dataGridView1.Rows.Add("CAN KARGO", 28.ToString(), 28.ToString(), 28.ToString(), 28.ToString(), 28.ToString(), 28.ToString(), 42.ToString(), 42.ToString(), 56.ToString(), 70.ToString());

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (desi == 0)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                    }
                    else if (desi > 0 && desi <= 4)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[2], ListSortDirection.Ascending);
                    }
                    else if (desi == 5)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[3], ListSortDirection.Ascending);
                    }
                    else if (desi > 5 && desi <= 10)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[4], ListSortDirection.Ascending);
                    }
                    else if (desi > 10 && desi <= 15)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[5], ListSortDirection.Ascending);
                    }
                    else if (desi > 15 && desi <= 20)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[6], ListSortDirection.Ascending);
                    }
                    else if (desi > 20 && desi <= 25)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[7], ListSortDirection.Ascending);
                    }
                    else if (desi > 25 && desi <= 30)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[8], ListSortDirection.Ascending);
                    }
                    else if (desi > 30 && desi <= 40)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[9], ListSortDirection.Ascending);
                    }
                    else if (desi > 40 && desi <= 50)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[10], ListSortDirection.Ascending);
                    }

                }
            }
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }



        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (textBox9.Text == "ANKARA KARGO")
            {
                label8.Visible = true;
                label9.Visible = true;
                comboBox1.Visible = true;
                comboBox2.Visible = true;
                label13.Visible = true;
                AutoCompleteStringCollection collection = new AutoCompleteStringCollection();

                comboBox1.Text = "Seçiniz...";
                object[] sehirler = new object[] { "Ankara", "Bursa", "Kocaeli", "Sakarya", "Eskişehir", "Manisa", "Adana", "İzmir", "Gaziantep", "İstanbul Avrupa", "İstanbul Anadolu", "Konya", "Mersin", "Tekirdağ" };
                comboBox1.Items.AddRange(sehirler);
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    collection.Add(comboBox1.Items[i].ToString());
                }
                //AutoCompleteStringCollection'u comboBox'un AutoCompleteCustomSource özelliğine atıyoruz.
                comboBox1.AutoCompleteCustomSource = collection;

                //comboBox'un otomatik tamamlama türünü seçiyoruz.
                comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

                //comboBox'un AutoCompleteSource özelliğinin CustomSource türünde olacağını belirtiyoruz.
                comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            }
            else if (textBox9.Text == "CAN KARGO")
            {
                label8.Visible = true;
                label9.Visible = true;
                comboBox3.Visible = true;
                comboBox4.Visible = true;
                button7.Visible = true;
                AutoCompleteStringCollection collection = new AutoCompleteStringCollection();

                comboBox3.Text = "Seçiniz...";

                object[] sehirler1 = new object[] { "Ankara", "Bursa", "Kocaeli", "Sakarya", "Eskişehir", "Manisa", "İzmir", "İstanbul Avrupa", "İstanbul Anadolu", "Balıkesir" };
                comboBox3.Items.AddRange(sehirler1);
                for (int i = 0; i < comboBox3.Items.Count; i++)
                {
                    collection.Add(comboBox3.Items[i].ToString());
                }
                //AutoCompleteStringCollection'u comboBox'un AutoCompleteCustomSource özelliğine atıyoruz.
                comboBox3.AutoCompleteCustomSource = collection;

                //comboBox'un otomatik tamamlama türünü seçiyoruz.
                comboBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

                //comboBox'un AutoCompleteSource özelliğinin CustomSource türünde olacağını belirtiyoruz.
                comboBox3.AutoCompleteSource = AutoCompleteSource.CustomSource;
            }
            else
            {
                label8.Visible = false;
                label9.Visible = false;
                comboBox1.Visible = false;
                comboBox2.Visible = false;
                button7.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                comboBox3.Visible = false;
                comboBox4.Visible = false;
            }




        }

        private void button5_Click(object sender, EventArgs e)
        {
            double desı = Convert.ToDouble(textBox4.Text);
            double adet = 1;
            double fıyat = Convert.ToDouble(textBox11.Text);
            adet = Convert.ToDouble(textBox7.Text);

            double netfıyat = (fıyat * adet);
            if (textBox9.Text != "ANKARA KARGO")
                dataGridView2.Rows.Add(textBox8.Text, textBox9.Text, textBox4.Text, netfıyat, adet);
            else if (textBox9.Text != "CAN KARGO")
                dataGridView2.Rows.Add(textBox8.Text, textBox9.Text, textBox4.Text, netfıyat, adet);
            else
                dataGridView2.Rows.Add(textBox8.Text, textBox9.Text, textBox4.Text, netfıyat, adet, comboBox1.Text, comboBox2.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox10.Text = textBox4.Text.ToString();

            //adet = Convert.ToInt32(textBox7.Text);
            double ekdesı;
            double desı = Convert.ToDouble(textBox4.Text);
            if (textBox4.Text != null)
            {
                if (textBox9.Text == "YURTİÇİ KARGO")
                {

                    ekdesı = ((70 + (desı - 30) * 2.35) * 1.18 * 1.0235);
                    if (desı < 1)
                        textBox11.Text = Math.Round((22.6 * 1.18) + (22.6 * 0.0235), 2).ToString();

                    else if (desı >= 1 && desı <= 4)
                        textBox11.Text = (27.55 * 1.18 * 1.0235).ToString();

                    else if (desı > 4 && desı < 6)
                        textBox11.Text = Math.Round((30.80 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 6 && desı <= 10)
                        textBox11.Text = Math.Round((33.85 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 10 && desı <= 15)
                        textBox11.Text = Math.Round((38.40 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 15 && desı <= 20)
                        textBox11.Text = Math.Round((47 * 1.18 * 1.0235), 2).ToString();


                    else if (desı > 20 && desı <= 25)
                        textBox11.Text = Math.Round((58.75 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 25 && desı <= 30)
                        textBox11.Text = Math.Round((70 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 30)
                        textBox11.Text = Math.Round(ekdesı, 2).ToString();
                }
                else if (textBox9.Text == "MNG KARGO")
                {
                    ekdesı = ((75 + (desı - 40) * 2.30) * 1.18 * 1.0235);
                    if (desı == 0 && desı < 1)
                        textBox11.Text = Math.Round((32 * 1.18 * 1.0235), 2).ToString();

                    else if (desı >= 1 && desı <= 5)
                        textBox11.Text = Math.Round((32 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 5 && desı <= 10)
                        textBox11.Text = Math.Round((35 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 10 && desı <= 15)
                        textBox11.Text = Math.Round((38 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 15 && desı <= 20)
                        textBox11.Text = Math.Round((45 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 20 && desı <= 25)
                        textBox11.Text = Math.Round((55 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 25 && desı <= 30)
                        textBox11.Text = Math.Round((60 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 30 && desı <= 40)
                        textBox11.Text = Math.Round((75 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 40)
                        textBox11.Text = Math.Round(ekdesı, 2).ToString();
                }
                else if (textBox9.Text == "ARAS KARGO")
                {
                    ekdesı = ((88.65 + (desı - 30) * 2.94) * 1.18 * 1.0235);
                    if (desı < 1)
                        textBox11.Text = Math.Round((21.71 * 1.18 * 1.0235), 2).ToString();

                    else if (desı >= 1 && desı <= 5)
                        textBox11.Text = Math.Round((39.69 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 5 && desı <= 10)
                        textBox11.Text = Math.Round((58.57 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 10 && desı <= 15)
                        textBox11.Text = Math.Round((62.36 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 15 && desı <= 20)
                        textBox11.Text = Math.Round((67.9 * 1.18 * 1.0235), 2).ToString();


                    else if (desı > 20 && desı <= 25)
                        textBox11.Text = Math.Round((78.04 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 25 && desı <= 30)
                        textBox11.Text = Math.Round((88.65 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 30)
                        textBox11.Text = Math.Round(ekdesı, 2).ToString();
                }
                else if (textBox9.Text == "SÜRAT KARGO")
                {
                    ekdesı = ((63 + (desı - 30) * 2.7) * 1.18 * 1.0235);
                    if (desı < 1)
                        textBox11.Text = Math.Round((25.45 * 1.18 * 1.0235), 2).ToString();

                    else if (desı >= 1 && desı <= 5)
                        textBox11.Text = Math.Round((25.65 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 5 && desı <= 10)
                        textBox11.Text = Math.Round((32.33 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 10 && desı <= 15)
                        textBox11.Text = Math.Round((40.37 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 15 && desı <= 20)
                        textBox11.Text = Math.Round((46.44 * 1.18 * 1.0235), 2).ToString();


                    else if (desı > 20 && desı <= 25)
                        textBox11.Text = Math.Round((53.53 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 25 && desı <= 30)
                        textBox11.Text = Math.Round((63 * 1.18 * 1.0235), 2).ToString();

                    else if (desı > 30)
                        textBox11.Text = Math.Round(ekdesı, 2).ToString();
                }
                else if (textBox9.Text == "ANKARA KARGO")
                {
                    ekdesı = (73.48 + (desı - 50) * 1.91) * 1.18 * 1.06;
                    if (desı > 0 && desı <= 10)
                        textBox11.Text = Math.Round((22.74 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 10 && desı <= 20)
                        textBox11.Text = Math.Round((34.99 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 20 && desı <= 30)
                        textBox11.Text = Math.Round((50.74 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 30 && desı <= 40)
                        textBox11.Text = Math.Round((59.48 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 40 && desı <= 50)
                        textBox11.Text = Math.Round((73.48 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 50)
                        textBox11.Text = Math.Round(ekdesı, 2).ToString();
                }
                else if (textBox9.Text == "CAN KARGO")
                {

                    ekdesı = (70 + (desı - 50) * 1.40) * 1.18 * 1.06;
                    if (desı > 0 && desı <= 20)
                        textBox11.Text = Math.Round((28 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 20 && desı <= 30)
                        textBox11.Text = Math.Round((42 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 30 && desı <= 40)
                        textBox11.Text = Math.Round((56 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 40 && desı <= 50)
                        textBox11.Text = Math.Round((70 * 1.18 * 1.06), 2).ToString();

                    else if (desı > 50)
                        textBox11.Text = Math.Round(ekdesı, 2).ToString();
                }


            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Ankara")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Akyurt", "Altındağ", "Çankaya", "Etimesgut", "Gölbaşı", "Kazan", "Keçiören", "Mamak", "Pursaklar", "Sincan", " Yenimahalle", "Hasanoğlan", " Etlik", "Kızılay", "Balgat", "Dikmen", "Şaşmaz", "Yenikent", "Öveçler", "Batıkent", "İskitler", "Akköprü", "Polatlı", "Maliköy", "Temelli", "Ulus", "Dışkapı", "karapürçekler", "Aydınlıkevler", "Saray", "Eryaman", "Subayevler", "Ostim", "İvedik", "Demetevler", "Gersan", "Abidinpaşa", "Cebeci", "Ümitköy", "Çayyolu", "Söğütözü", "Gimat", "Macunköy", "Çıbuk", "Siteler", "Gülveren", "Oran" };
                comboBox2.Items.AddRange(ilceler);
            }

            else if (comboBox1.Text == "Bursa")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Osmangazi", "Nilüfer", "Yıldırım", "Gürsu", "Kestel", "İnegöl" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Kocaeli")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Merkez", "Derince", "İzmit", "Kartepe", "Başiskele", "Körfez", "Kuruçeşme", "Gebze" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Sakarya")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Merkez", "Serdivan", "Erenler" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Eskişehir")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Odunpazarı", "Tepebaşı" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Manisa")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Yunusemre", "Şehzadeler" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Adana")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Yüreğir", "Seyhan", "Çukurova", "Sarıçam" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "İzmir")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Bornova", "Çankaya", "Balçova", "Pınarbaşı", "Kemalpaşa", "Işıkkent", "Yenişehir", "Konak", "Kemeraltı", "Alsancak", "Çiğle", "Karşıyaka", "Bostanlı", "Gaziemir", "Karabağlar", "Menderes", "Sarnıç", "Kısıkköy", "Torbalı", "Bayraklı", "Mavişehir", "5.Sanayi", "Yeşilyurt", "Hatay", "Buca", "Narlıdere", "Menemen" };
                comboBox2.Items.AddRange(ilceler);
            }

            else if (comboBox1.Text == "Gaziantep")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Şahinbey", "Şehitkemal", "Nizip" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "İstanbul Avrupa")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Avcılar", "Büyükçekmece", "Küçükçekmece", "Florya", "Yeşilköy", "Bakırköy", "Başakşehir", "Güneşli", "Bağcılar", "Esenler", "Zeytinburnu", "Güngören", "Aksaray", "Fatih", "Eminönü", "Şişli", "Levent", "Sarıyer", "Kağıthane", "Mecidiyeköy", "Beyoğlu", "Topkapı", "Kocamustafapaşa" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "İstanbul Anadolu")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Üsküdar", "Ümraniye(Merkez)", "Kadıköy", "Bostancı", "Beykoz(Merkez)", "Pendik", "Kartal", "Maltepe", "Ataşehir", "Tuzla", "Sultanbeyli", "Sancaktepe", "Çekmeköy", "Dilovası" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Konya")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Selçuklu", "Karatay", "Meram" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Mersin")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Akdeniz", "Toros", "Tarsus", "Yenişehir", "Erdemli", "Mezitli" };
                comboBox2.Items.AddRange(ilceler);
            }
            else if (comboBox1.Text == "Tekirdağ")
            {
                comboBox2.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Çorlu", "Çerkezköy" };
                comboBox2.Items.AddRange(ilceler);
            }

        }

    

        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
                Math.Round(toplam, 2);
            }
            Math.Round(toplam, 2);
            textBox6.Text = toplam.ToString() + " TL";
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
        StringFormat strFormat;
        ArrayList arrColumnLefts = new ArrayList();
        ArrayList arrColumnWidths = new ArrayList();
        int iCellHeight = 0;
        int iTotalWidth = 0;
        int iRow = 0;
        bool bFirstPage = false;
        bool bNewPage = false;
        int iHeaderHeight = 0;
        private void printDocument1_BeginPrint(object sender, PrintEventArgs e)
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
                foreach (DataGridViewColumn dgvGridCol in dataGridView2.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
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
                    foreach (DataGridViewColumn GridCol in dataGridView2.Columns)
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

                while (iRow <= dataGridView2.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dataGridView2.Rows[iRow];

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


                            e.Graphics.DrawString("TÜM KARGO ŞİRKETLERİ FİYAT HESABI", new Font(dataGridView2.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("YURTİÇİ KARGO FİYAT HESABI", new Font(dataGridView2.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            e.Graphics.DrawString("ERMED TIP MEDİKAL", font,
                                  Brushes.Black, e.MarginBounds.Top + 225, e.MarginBounds.Top -
                                  e.Graphics.MeasureString("ERMED TIP MEDİKAL", font, e.MarginBounds.Left).Height + 10);

                            String strDate = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                            e.Graphics.DrawString(strDate, new Font(dataGridView2.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dataGridView2.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("KARGO FİYAT HESABI", new Font(new Font(dataGridView2.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);


                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dataGridView2.Columns)
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
            e.Graphics.DrawImage(Logo, 730, 10, 80, 80);






            e.Graphics.DrawString("TOPLAM FİYAT=", font, firca, iCellHeight + 390, iTopMargin + 10);
            e.Graphics.DrawString(textBox6.Text, font, firca, iCellHeight + 570, iTopMargin + 10);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin, iCellHeight + 360, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 695, iTopMargin, iCellHeight + 695, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 570, iTopMargin, iCellHeight + 570, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin, iCellHeight + 695, iTopMargin);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin + 40, iCellHeight + 695, iTopMargin + 40);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook kitap = excel.Workbooks.Add(Missing);
            Worksheet sayfa = (Worksheet)kitap.Sheets[1];
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                Range alan = (Range)sayfa.Cells[1, 1];
                alan.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView2.Rows.Count; j++)
                {
                    Range alan2 = (Range)sayfa.Cells[j + 1, i + 1];
                    alan2.Cells[2, 1] = dataGridView2[i, j].Value;
                }
                Range alan4 = (Range)sayfa.Cells[1, 6];
                alan4.Value2 = "TOPLAM FİYAT";
                Range alan3 = (Range)sayfa.Cells[2, 6];
                alan3.Value2 = textBox6.Text;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            KARGO_SIRKETLERI KS = new KARGO_SIRKETLERI();
            KS.Show();
            this.Hide();
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            double desi = Convert.ToDouble(textBox4.Text);
            if (desi != null)
            {
                if (desi == 0)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                else if (desi > 0 && desi <= 4)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                else if (desi == 5)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                else if (desi > 5 && desi <= 10)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                else if (desi > 10 && desi <= 15)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                else if (desi > 16 && desi <= 20)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                else if (desi > 21 && desi <= 25)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                else if (desi >= 26 && desi <= 30)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                else if (desi >= 31 && desi <= 40)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                else if (desi >= 41 && desi <= 50)
                    textBox5.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            }

            textBox9.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            if (textBox9.Text == "ANKARA KARGO")
            {
                comboBox1.Text = "Seçiniz...";
                AutoCompleteStringCollection collection = new AutoCompleteStringCollection();
                object[] sehirler = new object[] { "Ankara", "Bursa", "Kocaeli", "Sakarya", "Eskişehir", "Manisa", "Adana", "İzmir", "Gaziantep", "İstanbul Avrupa", "İstanbul Anadolu", "Konya", "Mersin", "Tekirdağ" };
                comboBox1.Items.AddRange(sehirler);
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    collection.Add(comboBox1.Items[i].ToString());
                }
                //AutoCompleteStringCollection'u comboBox'un AutoCompleteCustomSource özelliğine atıyoruz.
                comboBox1.AutoCompleteCustomSource = collection;

                //comboBox'un otomatik tamamlama türünü seçiyoruz.
                comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

                //comboBox'un AutoCompleteSource özelliğinin CustomSource türünde olacağını belirtiyoruz.
                comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;

            }
            else if(textBox9.Text!="ANKARA KARGO")
            {
                comboBox1.Items.Clear();

            }
            else if (textBox9.Text == "CAN KARGO")
            {
                comboBox1.Text = "Seçiniz...";
                AutoCompleteStringCollection collection = new AutoCompleteStringCollection();
                object[] sehirler1 = new object[] { "Ankara", "Bursa", "Kocaeli", "Sakarya", "Eskişehir", "Manisa", "İzmir", "İstanbul Avrupa", "İstanbul Anadolu", "Balıkesir" };
                comboBox1.Items.AddRange(sehirler1);
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    collection.Add(comboBox1.Items[i].ToString());
                }
                //AutoCompleteStringCollection'u comboBox'un AutoCompleteCustomSource özelliğine atıyoruz.
                comboBox1.AutoCompleteCustomSource = collection;

                //comboBox'un otomatik tamamlama türünü seçiyoruz.
                comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

                //comboBox'un AutoCompleteSource özelliğinin CustomSource türünde olacağını belirtiyoruz.
                comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            }
            else if (textBox9.Text != "CAN KARGO")
            {
                comboBox1.Items.Clear();

            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            double desı = Convert.ToDouble(textBox4.Text);
            double adet = 1;
            adet = Convert.ToDouble(textBox7.Text);
            if (desı >= 0 && desı <= 20)
                if (adet >= 6)
                    textBox11.Text = Math.Round(adet * (28 * 1.18 * 1.06), 2).ToString();
                else
                    MessageBox.Show("adet en az 6 girilmeli...");

            else if (desı > 20 && desı <= 30)
                if (adet >= 4)
                    textBox11.Text = Math.Round(adet * (42 * 1.18 * 1.06), 2).ToString();
                else
                    MessageBox.Show("adet en az 4 girilmeli...");

            else if (desı > 30 && desı <= 40)
                if (adet >= 3)
                    textBox11.Text = Math.Round(adet * (56 * 1.18 * 1.06), 2).ToString();
                else
                    MessageBox.Show("adet en az 4 girilmeli...");

            else if (desı > 40 && desı <= 50)
                if (adet >= 3)
                    textBox11.Text = Math.Round(adet * (70 * 1.18 * 1.06), 2).ToString();
                else
                    MessageBox.Show("adet en az 3 girilmeli...");

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text == "Ankara")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Akyurt", "Altındağ", "Çankaya", "Etimesgut", "Gölbaşı", "Kazan", "Keçiören", "Mamak", "Pursaklar", "Sincan", " Yenimahalle" };
                comboBox4.Items.AddRange(ilceler);
            }

            else if (comboBox3.Text == "Bursa")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Osmangazi", "Nilüfer", "Yıldırım", "Gürsu", "Kestel", "İnegöl", "Hasanağa OSB", "Mudanya" };
                comboBox4.Items.AddRange(ilceler);
            }
            else if (comboBox3.Text == "Kocaeli")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Merkez", "Derince", "İzmit", "Kartepe", "Başiskele", "Körfez", "Kuruçeşme", "Gebze", "Gölcük", "Karamürsel", "Darıca", "Çayırova", "Dilovası" };
                comboBox4.Items.AddRange(ilceler);
            }
            else if (comboBox3.Text == "Sakarya")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Merkez", "Serdivan", "Erenler", "Ferizli", "Söğütlü", "Sapanca", "Arifiye" };
                comboBox4.Items.AddRange(ilceler);
            }
            else if (comboBox3.Text == "Eskişehir")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Odunpazarı", "Tepebaşı" };
                comboBox4.Items.AddRange(ilceler);
            }
            else if (comboBox3.Text == "Manisa")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Merkez", "Yunusemre", "Şehzadeler" };
                comboBox4.Items.AddRange(ilceler);
            }

            else if (comboBox3.Text == "İzmir")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Bornova", "Balçova", "Pınarbaşı", "Kemalpaşa", "Çiğle", "Karşıyaka", "Gaziemir", "Karabağlar", "Menderes", "Torbalı", "Bayraklı", "Buca", "Menemen", "Kemalpaşa", "Konak" };
                comboBox4.Items.AddRange(ilceler);
            }

            else if (comboBox3.Text == "İstanbul Avrupa")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Avcılar", "Büyükçekmece", "Küçükçekmece", "Florya", "Yeşilköy", "Bakırköy", "Başakşehir", "Güneşli", "Bağcılar", "Esenler", "Zeytinburnu", "Güngören", "Aksaray", "Fatih", "Eminönü", "Şişli", "Levent", "Sarıyer", "Kağıthane", "Mecidiyeköy", "Beyoğlu", "Topkapı", "Kocamustafapaşa" };
                comboBox4.Items.AddRange(ilceler);
            }
            else if (comboBox3.Text == "İstanbul Anadolu")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Üsküdar", "Ümraniye(Merkez)", "Kadıköy", "Bostancı", "Beykoz(Merkez)", "Pendik", "Kartal", "Maltepe", "Ataşehir", "Tuzla", "Sultanbeyli", "Sancaktepe", "Çekmeköy", "Dilovası" };
                comboBox4.Items.AddRange(ilceler);
            }
            else if (comboBox3.Text == "Balıkesir")
            {
                comboBox4.Text = "Seçiniz...";
                object[] ilceler = new object[] { "Altıeylül", "Karesi" };
                comboBox4.Items.AddRange(ilceler);
            }
        }

        private void FILTER_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else
                e.Cancel = true;//Çıkışı durdur
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
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

        private void button8_MouseEnter(object sender, EventArgs e)
        {
            button8.BackColor = Color.Gold;
        }

        private void button8_MouseLeave(object sender, EventArgs e)
        {
            button8.BackColor = Color.Gold;
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.Green;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.BackColor = Color.SandyBrown;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackColor = Color.White;
        }

        private void button6_MouseEnter(object sender, EventArgs e)
        {
            button6.BackColor = Color.Blue;
        }

        private void button6_MouseLeave(object sender, EventArgs e)
        {
            button6.BackColor = Color.White;
        }

        private void button7_MouseEnter(object sender, EventArgs e)
        {
            button7.BackColor = Color.Green;
        }

        private void button7_MouseLeave(object sender, EventArgs e)
        {
            button7.BackColor = Color.White;
        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            button5.BackColor = Color.Green;
        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            button5.BackColor = Color.White;
        }
    }
}
