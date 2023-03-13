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
using System.Data.OleDb;
using DataTable = System.Data.DataTable;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Kargo
{
    public partial class FILTER : Form
    {
        public FILTER()
        {
            InitializeComponent();
        }
        int row;
        double desi;
        double fiyat;
        OleDbConnection con;
        OleDbConnection con1;
        OleDbConnection cone;
        OleDbDataAdapter daa;
        OleDbCommand cmd;
        DataSet ds;
        SqlConnection baglanti;
        SqlCommand komut;
        SqlDataAdapter da;
        string Depo;
        void griddoldur()
        {
            cone = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
            daa = new OleDbDataAdapter("SElect *from KargoFiltre", cone);
            ds = new DataSet();
            cone.Open();
            daa.Fill(ds, "KargoFiltre");
            dataGridView1.DataSource = ds.Tables["KargoFiltre"];
            cone.Close();
        }
        public void FILTER_Load(object sender, EventArgs e)
        {
            griddoldur();

            //dataGridView2.ColumnCount = 9;
            //dataGridView2.Columns[0].Name = "GÖNDERİLECEK FİRMA ADI";
            //dataGridView2.Columns[1].Name = "KARGO FİRMASI";
            //dataGridView2.Columns[2].Name = "DESİ";
            //dataGridView2.Columns[3].Name = "FİYAT TL";
            //dataGridView2.Columns[4].Name = "ADET";
            //dataGridView2.Columns[5].Name = "DEPO";
            //dataGridView2.Columns[6].Name = "İL";
            //dataGridView2.Columns[7].Name = "İLÇE";
            //dataGridView2.Columns[8].Name = "TARİH";

            if (textBox9.Text == null)
            {
                con.Close();
                con1.Close();
            }
            else if (textBox9.Text == "ANKARA KARGO")
            {

                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from ANKARAKARGOİL ORDER BY ID ASC ", con);
                da.Fill(dt);
                comboBox1.ValueMember = "ID";

                comboBox1.DisplayMember = "il";

                comboBox1.DataSource = dt;
                dt.EndInit();

            }
            else if (textBox9.Text == "CAN KARGO")
            {
                con1 = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
                DataTable dt1 = new DataTable();
                OleDbDataAdapter da1 = new OleDbDataAdapter("select * from CANKARGOİL ORDER BY ID ASC ", con1);
                da1.Fill(dt1);
                comboBox3.ValueMember = "ID";
                comboBox3.DisplayMember = "il";
                comboBox3.DataSource = dt1;
                dt1.EndInit();
            }
            comboBox5.Items.Add("ANA DEPO");
            comboBox5.Items.Add("DMO");



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

                textBox4.Text = Math.Round((en * boy * yukseklik) / 3000,0).ToString();

                desi= Convert.ToDouble(textBox4.Text);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Double desi = Convert.ToDouble(textBox4.Text);




            if (textBox4.Text != null)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (desi == 0)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[2], ListSortDirection.Ascending);
                    }
                    else if (desi > 0 && desi <= 4)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[3], ListSortDirection.Ascending);
                    }
                    else if (desi == 5)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[4], ListSortDirection.Ascending);
                    }
                    else if (desi > 5 && desi <= 10)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[5], ListSortDirection.Ascending);
                    }
                    else if (desi > 10 && desi <= 15)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[6], ListSortDirection.Ascending);
                    }
                    else if (desi > 15 && desi <= 20)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[7], ListSortDirection.Ascending);
                    }
                    else if (desi > 20 && desi <= 25)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[8], ListSortDirection.Ascending);
                    }
                    else if (desi > 25 && desi <= 30)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[9], ListSortDirection.Ascending);
                    }
                    else if (desi > 30 && desi <= 40)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[10], ListSortDirection.Ascending);
                    }
                    else if (desi > 40 && desi <= 50)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[11], ListSortDirection.Ascending);
                    }
                    else if(desi>50)
                    {
                        dataGridView1.Sort(dataGridView1.Columns[11], ListSortDirection.Ascending);
                    }

                }


            }
 
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }
        private void button5_Click(object sender, EventArgs e)
        {

            dataGridView2.ColumnCount = 9;
            dataGridView2.Columns[0].Name = "GÖNDERİLECEK FİRMA ADI";
            dataGridView2.Columns[1].Name = "KARGO FİRMASI";
            dataGridView2.Columns[2].Name = "DESİ";
            dataGridView2.Columns[3].Name = "FİYAT TL";
            dataGridView2.Columns[4].Name = "ADET";
            dataGridView2.Columns[5].Name = "DEPO";
            dataGridView2.Columns[6].Name = "İL";
            dataGridView2.Columns[7].Name = "İLÇE";
            dataGridView2.Columns[8].Name = "TARİH";

            Depo = comboBox5.SelectedItem.ToString();
            DateTime zaman = DateTime.Now;
            string format = "yyyy-MM-dd";
            var zamanim = zaman.ToString(format);
            double desı = Convert.ToDouble(textBox4.Text);
            double fıyat1 = Convert.ToDouble(textBox11.Text);
            double adet = Convert.ToDouble(textBox7.Text);

            if (textBox9.Text == "ANKARA KARGO" && textBox8.Text != null)
                dataGridView2.Rows.Add(textBox8.Text, textBox9.Text, textBox4.Text, fıyat1, adet, Depo, comboBox1.Text, comboBox2.Text, zamanim);
            else if (textBox9.Text == "CAN KARGO" && textBox8.Text != null)
                dataGridView2.Rows.Add(textBox8.Text, textBox9.Text, textBox4.Text, fıyat1, adet, Depo, comboBox3.Text, comboBox4.Text, zamanim);
            else if (textBox8.Text != null)
                dataGridView2.Rows.Add(textBox8.Text, textBox9.Text, textBox4.Text, fıyat1, adet, Depo, null, null, zamanim);
            else
                MessageBox.Show("hata");


        }
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DELLSRV;Initial Catalog=ermed_kargo;User ID=sa;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            try
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (baglanti.State == ConnectionState.Open)
                    {
                        baglanti.Close();
                    }

                    var G_firma = Convert.ToString(dataGridView2.Rows[i].Cells[0].Value); // 2. kolon
                    var K_Firma = Convert.ToString(dataGridView2.Rows[i].Cells[1].Value); // 3. kolon
                    var Desi = Convert.ToInt32(dataGridView2.Rows[i].Cells[2].Value).ToString(); // 4. kolon
                    var Fiyat = Convert.ToInt32(dataGridView2.Rows[i].Cells[3].Value).ToString(); // 5. kolon
                    var Adet = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value).ToString(); // 6. kolon
                    var Depo = Convert.ToString(dataGridView2.Rows[i].Cells[5].Value);
                    var il = Convert.ToString(dataGridView2.Rows[i].Cells[6].Value); // 7. kolon
                    var ilce = Convert.ToString(dataGridView2.Rows[i].Cells[7].Value); // 8. kolon
                    var Tarih = Convert.ToDateTime(dataGridView2.Rows[i].Cells[8].Value).ToString("yyyy-MM-dd"); // 9. kolon

                    baglanti.Open();
                    SqlCommand komut = new SqlCommand("INSERT INTO Kargolar (Gonderilcek_firma,Kargo_Sirketi,Desi_KG,Fiyat,Adet,Depo,İL,İLCE,Tarih) VALUES ('" + G_firma + "' , '" + K_Firma + "','" + Desi + "' , '" + Fiyat + "' , '" + Adet + "','" + Depo + "','" + il + "' , '" + ilce + "','" + Tarih + "')", baglanti);
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

        private void button6_Click(object sender, EventArgs e)
        {
            desi=Convert.ToDouble(textBox4.Text);
            textBox10.Text = textBox4.Text.ToString();
            double ekdesi;

            if (textBox4.Text != null)
            {
                if (textBox9.Text == "YURTİÇİ KARGO")
                {

                    ekdesi = (desi - 30) * 3.29 * 1.18 * 1.0235;
                    double desifiyat = (fiyat * 1.18 * 1.0235) + ekdesi;
                    if (desi < 1)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi >= 1 && desi <= 4)
                        textBox11.Text = Math.Round(fiyat * 1.18 * 1.0235).ToString();

                    else if (desi > 4 && desi < 6)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 6 && desi <= 10)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 10 && desi <= 15)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 15 && desi <= 20)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();


                    else if (desi > 20 && desi <= 25)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 25 && desi <= 30)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 30)
                        textBox11.Text = Math.Round(desifiyat, 2).ToString();
                }
                else if (textBox9.Text == "MNG KARGO")
                {
                    ekdesi =  (desi - 40) * 2.30 * 1.18 * 1.0235;
                    double desifiyat = fiyat + ekdesi;
                    if (desi == 0 && desi < 1)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi >= 1 && desi <= 5)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 5 && desi <= 10)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 10 && desi <= 15)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 15 && desi <= 20)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 20 && desi <= 25)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 25 && desi <= 30)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 30 && desi <= 40)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 40)
                        textBox11.Text = Math.Round(desifiyat, 2).ToString();
                }
                else if (textBox9.Text == "A KARGO")
                {
                    ekdesi = (desi - 30) * 2.94 * 1.18 * 1.0235;
                    double desifiyat = fiyat + ekdesi;
                    if (desi < 1)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi >= 1 && desi <= 5)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 5 && desi <= 10)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 10 && desi <= 15)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 15 && desi <= 20)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();


                    else if (desi > 20 && desi <= 25)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 25 && desi <= 30)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 30)
                        textBox11.Text = Math.Round(desifiyat, 2).ToString();
                }
                else if (textBox9.Text == "SÜRAT KARGO")
                {
                    ekdesi =  (desi - 30) * 2.7 * 1.18 * 1.0235;
                    double desifiyat = fiyat + ekdesi;
                    if (desi < 1)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi >= 1 && desi <= 5)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 5 && desi <= 10)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 10 && desi <= 15)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 15 && desi <= 20)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();


                    else if (desi > 20 && desi <= 25)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 25 && desi <= 30)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.0235), 2).ToString();

                    else if (desi > 30)
                        textBox11.Text = Math.Round(desifiyat, 2).ToString();
                }
                else if (textBox9.Text == "ANKARA KARGO")
                {
                    ekdesi = (desi - 50) * 1.91 * 1.18 * 1.06;
                    double desifiyat = fiyat + ekdesi;

                    if (desi > 0 && desi <= 10)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 10 && desi <= 20)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 20 && desi <= 30)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 30 && desi <= 40)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 40 && desi <= 50)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 50)
                        textBox11.Text = Math.Round(desifiyat, 2).ToString();
                }
                else if (textBox9.Text == "CAN KARGO")
                {

                    ekdesi =(desi - 50) * 1.40 * 1.18 * 1.06;
                    double desifiyat = fiyat + ekdesi;
                    if (desi > 0 && desi <= 20)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 20 && desi <= 30)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 30 && desi <= 40)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 40 && desi <= 50)
                        textBox11.Text = Math.Round((fiyat * 1.18 * 1.06), 2).ToString();

                    else if (desi > 50)
                        textBox11.Text = Math.Round(desifiyat, 2).ToString();
                }


            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox1.SelectedIndex != -1)
            {
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from ANKARAKARGOİLCE where il = " + comboBox1.SelectedValue.ToString(), con);
                da.Fill(dt);
                comboBox2.ValueMember = "ID";
                comboBox2.DisplayMember = "ilce";
                comboBox2.DataSource = dt;
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
            textBox6.Text = Math.Round(toplam, 2).ToString() + " TL";
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
                Range alan4 = (Range)sayfa.Cells[1, 11];
                alan4.Value2 = "TOPLAM FİYAT";
                Range alan3 = (Range)sayfa.Cells[2, 11];
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

            desi=Convert.ToDouble(textBox4.Text);
            if (desi != null)
            {
                     row = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                    textBox9.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
                    if (desi == 0)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[2].Value);
                    else if (desi >= 1 && desi <= 4)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[3].Value);
                    else if (desi == 5)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[4].Value);
                    else if (desi > 5 && desi <= 10)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[5].Value);
                    else if (desi > 10 && desi <= 15)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[6].Value);
                    else if (desi > 1 && desi <= 20)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[7].Value);
                    else if (desi > 20 && desi <= 25)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[8].Value);
                    else if (desi > 25 && desi <= 30)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[9].Value);
                    else if (desi > 30 && desi <= 40)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[10].Value);
                    else if (desi > 40 && desi <= 50)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[11].Value);
                    else if (desi > 50)
                        fiyat = Convert.ToDouble(dataGridView1.SelectedRows[0].Cells[11].Value);
            }

            if (textBox9.Text == null)
            {
                con.Close();
                con1.Close();
            }
            else if (textBox9.Text == "ANKARA KARGO")
            {

                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from ANKARAKARGOİL ORDER BY ID ASC ", con);
                da.Fill(dt);
                comboBox1.ValueMember = "ID";
                
                comboBox1.DisplayMember = "il";

                comboBox1.DataSource = dt;
                dt.EndInit();

            }
            else if (textBox9.Text == "CAN KARGO")
            {
                con1 = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
                DataTable dt1 = new DataTable();
                OleDbDataAdapter da1 = new OleDbDataAdapter("select * from CANKARGOİL ORDER BY ID ASC ", con1);
                da1.Fill(dt1);
                comboBox3.ValueMember = "ID";
                comboBox3.DisplayMember = "il";
                comboBox3.DataSource = dt1;
                dt1.EndInit();
            }

            if (textBox9.Text == "ANKARA KARGO")
            {
                label8.Visible = true;
                label9.Visible = true;
                comboBox1.Visible = true;
                comboBox2.Visible = true;
                label13.Visible = true;
                comboBox3.Visible = false;
                comboBox4.Visible = false;
                label14.Visible = false;
            }
            else if (textBox9.Text == "CAN KARGO")
            {
                label14.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                comboBox3.Visible = true;
                comboBox4.Visible = true;
                comboBox1.Visible = false;
                comboBox2.Visible = false;
            }
            else
            {
                comboBox1.Visible = false;
                comboBox2.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                comboBox3.Visible = false;
                comboBox4.Visible = false;

            }
        }

            private void button7_Click(object sender, EventArgs e)
            {
                double desı = Convert.ToDouble(textBox4.Text);
                double adet = 1;
                adet = Convert.ToDouble(textBox7.Text);
                if (textBox9.Text == "CAN KARGO")
                {
                    if (desı >= 0 && desı <= 20)
                        if (adet >= 6)
                            textBox11.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 2).ToString();
                        else
                            MessageBox.Show("adet en az 6 girilmeli...");

                    else if (desı > 20 && desı <= 30)
                        if (adet >= 4)
                            textBox11.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 2).ToString();
                        else
                            MessageBox.Show("adet en az 4 girilmeli...");

                    else if (desı > 30 && desı <= 40)
                        if (adet >= 3)
                            textBox11.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 2).ToString();
                        else
                            MessageBox.Show("adet en az 4 girilmeli...");

                    else if (desı > 40 && desı <= 50)
                        if (adet >= 3)
                            textBox11.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 2).ToString();
                        else
                            MessageBox.Show("adet en az 3 girilmeli...");
                }
                else
                {
                   double gfiyat=Convert.ToDouble(textBox11.Text);
                   textBox11.Text=(gfiyat*adet).ToString();
                }
            }

            private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) 
            {

                if (comboBox3.SelectedIndex != -1)
                {
                    DataTable dt1 = new DataTable();
                    OleDbDataAdapter da1 = new OleDbDataAdapter("select * from CANKARGOİLCE where il = " + comboBox3.SelectedValue, con1);
                    da1.Fill(dt1);
                    comboBox4.ValueMember = "ID";
                    comboBox4.DisplayMember = "ilce";
                    comboBox4.DataSource = dt1;
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
            textBox6.Clear();
            textBox8.Clear();
            textBox4.Clear();
            textBox3.Clear();
            textBox2.Clear();
            textBox1.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox11.Clear();
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
                button8.BackColor = Color.White;
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

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            try
            {
                // Dosya seçme penceresi açmak için
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                // seçtiğimiz excel'in tam yolu
                string tamYol = file.FileName;

                //Excel bağlantı adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                //bağlantı 
                OleDbConnection baglanti = new(baglantiAdresi);

                //tüm verileri seçmek için select sorgumuz. Sayfa1 olan kısmı Excel'de hangi sayfayı açmak istiyosanız orayı yazabilirsiniz.
                OleDbCommand komut = new OleDbCommand("Select * From [" + "Sheet1" + "$]", baglanti);

                //bağlantıyı açıyoruz.
                baglanti.Open();

                //Gelen verileri DataAdapter'e atıyoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz için bir DataTable oluşturuyoruz.
                DataTable data = new DataTable();

                //DataAdapter'da ki verileri data adındaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kaynağını oluşturduğumuz DataTable ile dolduruyoruz.
                dataGridView2.DataSource = data;
                
            }
            catch (Exception ex)
            {
                // Hata alırsak ekrana bastırıyoruz.
                MessageBox.Show(ex.Message);
            }
        }
    }
}

