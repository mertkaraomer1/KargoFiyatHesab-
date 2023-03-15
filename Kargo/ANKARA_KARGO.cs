using System;
using System.Collections;
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
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;
using System.Runtime.ConstrainedExecution;
using Microsoft.VisualBasic.Logging;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing.Imaging;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Globalization;
using System.Data.SqlClient;

namespace Kargo
{
    public partial class ANKARA_KARGO : Form
    {

        public ANKARA_KARGO()
        {
            InitializeComponent();
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
        OleDbConnection cone;
        DataSet ds;
        double fiyat;
        OleDbCommand cmd;
        double desi;
        SqlConnection baglanti;
        string Depo;
        void griddoldur()
        {
            double desi = Convert.ToDouble(textBox4.Text);
            cone = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
            if (desi <=50)
            {
                 cmd = new OleDbCommand("Select fiyat from ANKARAKARGO where desi Like '" + textBox4.Text + "'", cone);
            }
            else if (desi > 50)
            {
                 cmd = new OleDbCommand("Select fiyat from ANKARAKARGO where desi Like '50'", cone);
            }
            ds = new DataSet();
            cone.Open();
            OleDbDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {

                fiyat = (double)dr["fiyat"];
                //desiList.Add(new Dictionary<double, double>((double)dr["desi"], (double)dr["fiyat"]));

            }

            cone.Close();
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            System.Drawing.Image Logo = imageList1.Images["ASPİRASYON SONDASI.PNG"];
            Pen kalem = new Pen(Color.Black);
            Font font = new Font("Arial", 14);
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


                            e.Graphics.DrawString("ANKARA KARGO FİYAT HESABI", new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("YURTİÇİ KARGO FİYAT HESABI", new Font(dataGridView1.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            e.Graphics.DrawString("ERMED TIP MEDİKAL", font,
                                   Brushes.Black, e.MarginBounds.Top + 225, e.MarginBounds.Top -
                                   e.Graphics.MeasureString("ERMED TIP MEDİKAL", font, e.MarginBounds.Left).Height - 15);

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
            e.Graphics.DrawImage(Logo, 730, 10, 80, 80);

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



                textBox4.Text =Math.Round((en * boy * yukseklik) / 3000,0).ToString();
                desi = Convert.ToDouble(textBox4.Text);
            }

            griddoldur();

            double ekdesı = (desi - 50) * 1.91*1.18*1.06;
            double desifiyat = (fiyat * 1.18 * 1.0235) + ekdesı;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";



            if (desi > 0 && desi <= 10)
                textBox5.Text = Math.Round((fiyat * 1.18 * 1.06), 0).ToString();

            else if (desi > 10 && desi < 20)
                textBox5.Text = Math.Round((fiyat * 1.18 * 1.06), 0).ToString();

            else if (desi > 20 && desi <= 30)
                textBox5.Text = Math.Round((fiyat * 1.18 * 1.06), 0).ToString();

            else if (desi > 30 && desi <= 40)
                textBox5.Text = Math.Round((fiyat * 1.18 * 1.06), 0).ToString();

            else if (desi > 40 && desi <= 50)
                textBox5.Text = Math.Round((fiyat * 1.18 * 1.06), 0).ToString();

            else if (desi > 50)
                textBox5.Text = Math.Round(desifiyat, 0).ToString();


        }



        OleDbConnection con;
        private void ANKARA_KARGO_Load(object sender, EventArgs e)
        {

            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=\\\\hpserver\\PROGRAM_PAYLASIM\\KARGO_FIYAT\\Kargo\\bin\\Debug\\net6.0-windows\\fiyat_listesi.accdb");
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter("select * from ANKARAKARGOİL ORDER BY ID ASC ", con);
            da.Fill(dt);
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "il";
            comboBox1.DataSource = dt;




            dataGridView1.ColumnCount = 9;
            dataGridView1.Columns[0].Name = "FİRMA ADI";
            dataGridView1.Columns[1].Name = "DESI/KİLO";
            dataGridView1.Columns[2].Name = "FIYAT";
            dataGridView1.Columns[3].Name = "TL";
            dataGridView1.Columns[4].Name = "ADET";
            dataGridView1.Columns[5].Name = "DEPO";
            dataGridView1.Columns[6].Name = "İL"; 
            dataGridView1.Columns[7].Name = "İLÇE";
            dataGridView1.Columns[8].Name = "TARİH";

            comboBox3.Items.Add("ANA DEPO");
            comboBox3.Items.Add("DMO");
        }

        private void TEMIZLE_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox8.Clear();
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

            griddoldur();
            Depo = comboBox3.SelectedItem.ToString();
            int adet = 1;
            adet = Convert.ToInt32(textBox7.Text);
            string TL = "TL";
            double desi = Convert.ToDouble(textBox4.Text);
            if (textBox4 != null)
            {

                double ekdesı = (desi - 50) * 1.91 * 1.18 * 1.06;
                double desifiyat = (fiyat * 1.18 * 1.0235) + ekdesı;

                if (desi > 1 && desi <= 10)
                    textBox5.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 0).ToString();


                else if (desi > 10 && desi <= 20)
                    textBox5.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 0).ToString();

                else if (desi > 20 && desi <= 30)
                    textBox5.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 0).ToString();

                else if (desi > 30 && desi <= 40)
                    textBox5.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 0).ToString();

                else if (desi > 40 && desi <= 50)
                    textBox5.Text = Math.Round(adet * (fiyat * 1.18 * 1.06), 0).ToString();

                else if (desi > 50)
                    textBox5.Text = Math.Round(adet * desifiyat, 0).ToString();

            }
            dataGridView1.Rows.Add(textBox8.Text, desi, textBox5.Text, TL, adet, Depo, comboBox1.Text, comboBox2.Text, DateTime.Now.ToString("yyyy-MM-dd"));



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
                    var K_Firma = Convert.ToString("ANKARA KARGO"); // 3. kolon
                    var Desi = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value).ToString(); // 4. kolon
                    var Fiyat = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value).ToString(); // 5. kolon
                    var Adet = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value).ToString(); // 6. kolon
                    var Depo = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value); // 7. kolon
                    var il = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value); // 7. kolon
                    var ilce = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value); // 8. kolon
                    var Tarih = Convert.ToDateTime(dataGridView1.Rows[i].Cells[8].Value).ToString("yyyy-MM-dd"); // 9. kolon

                    baglanti.Open();
                    SqlCommand komut = new SqlCommand("INSERT INTO Kargolarr (Gonderilcek_firma,Kargo_Sirketi,Desi_KG,Fiyat,Adet,Depo,İL,İLCE,Tarih) VALUES ('" + G_firma + "' , '" + K_Firma + "','" + Desi + "' , '" + Fiyat + "' , '" + Adet + "','" + Depo + "','" + il + "' , '" + ilce + "','" + Tarih + "')", baglanti);
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
        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                Math.Round(toplam, 0);
            }
            Math.Round(toplam, 0);
            textBox6.Text = Math.Round(toplam, 0).ToString() + " TL";
        }

        private void ANKARA_KARGO_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else
                
                e.Cancel = true;//Çıkışı durdur
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

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.Green;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.Red;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.BackColor = Color.Green;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackColor = Color.White;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackColor = Color.Green;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;
        }

        private void TEMIZLE_MouseLeave(object sender, EventArgs e)
        {
            TEMIZLE.BackColor = Color.White;
        }

        private void TEMIZLE_MouseEnter(object sender, EventArgs e)
        {
            TEMIZLE.BackColor = Color.Gold;
        }


    }
}
