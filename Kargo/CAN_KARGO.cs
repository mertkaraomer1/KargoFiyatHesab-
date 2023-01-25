using Microsoft.Office.Interop.Excel;
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
using Range = Microsoft.Office.Interop.Excel.Range;
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;
using System.Runtime.ConstrainedExecution;
using Microsoft.VisualBasic.Logging;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing.Imaging;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Kargo
{
    public partial class CAN_KARGO : Form
    {
        public CAN_KARGO()
        {
            InitializeComponent();
        }

        private void CAN_KARGO_Load(object sender, EventArgs e)
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






            dataGridView1.ColumnCount = 7;
            dataGridView1.Columns[0].Name = "FİRMA ADI";
            dataGridView1.Columns[1].Name = "DESI/KİLO";
            dataGridView1.Columns[2].Name = "FIYAT";
            dataGridView1.Columns[3].Name = "TL";
            dataGridView1.Columns[4].Name = "ADET";
            dataGridView1.Columns[5].Name = "İL";
            dataGridView1.Columns[6].Name = "İLÇE";

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


                            e.Graphics.DrawString("CAN KARGO FİYAT HESABI", new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("CAN KARGO FİYAT HESABI", new Font(dataGridView1.Font,
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
                Range alan4 = (Range)sayfa.Cells[1, 6];
                alan4.Value2 = "TOPLAM FİYAT";
                Range alan3 = (Range)sayfa.Cells[2, 6];
                alan3.Value2 = textBox6.Text;
            }
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


                double ekdesı = (70 + (sonuc - 50) * 1.40) * 1.18 * 1.06;
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";



                ekdesı = (70 + (sonuc - 50) * 1.40) * 1.18 * 1.06;
                if (sonuc > 0 && sonuc <= 20)
                    textBox5.Text = Math.Round((28 * 1.18 * 1.06), 2).ToString();

                else if (sonuc > 20 && sonuc <= 30)
                    textBox5.Text = Math.Round((42 * 1.18 * 1.06), 2).ToString();

                else if (sonuc > 30 && sonuc <= 40)
                    textBox5.Text = Math.Round((56 * 1.18 * 1.06), 2).ToString();

                else if (sonuc > 40 && sonuc <= 50)
                    textBox5.Text = Math.Round((70 * 1.18 * 1.06), 2).ToString();

                else if (sonuc > 50)
                    textBox5.Text = Math.Round(ekdesı, 2).ToString();

            }
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
            string TL = "TL";
            double desı = Convert.ToDouble(textBox4.Text);
           

            dataGridView1.Rows.Add(textBox8.Text, desı, textBox5.Text, TL, adet, comboBox1.Text, comboBox2.Text);
            comboBox2.Items.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                Math.Round(toplam, 2);
            }
            Math.Round(toplam, 2);
            textBox6.Text = toplam.ToString() + " TL";
        }

        private void CAN_KARGO_FormClosing(object sender, FormClosingEventArgs e)
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
            if (comboBox1.Text == "Ankara")
                {
                    comboBox2.Text = "Seçiniz...";
                    object[] ilceler = new object[] { "Akyurt", "Altındağ", "Çankaya", "Etimesgut", "Gölbaşı", "Kazan", "Keçiören", "Mamak", "Pursaklar", "Sincan", " Yenimahalle"};
                    comboBox2.Items.AddRange(ilceler);
                }

                else if (comboBox1.Text == "Bursa")
                {
                    comboBox2.Text = "Seçiniz...";
                    object[] ilceler = new object[] { "Osmangazi", "Nilüfer", "Yıldırım", "Gürsu", "Kestel", "İnegöl","Hasanağa OSB","Mudanya" };
                    comboBox2.Items.AddRange(ilceler);
                }
                else if (comboBox1.Text == "Kocaeli")
                {
                    comboBox2.Text = "Seçiniz...";
                    object[] ilceler = new object[] { "Merkez", "Derince", "İzmit", "Kartepe", "Başiskele", "Körfez", "Kuruçeşme", "Gebze","Gölcük","Karamürsel","Darıca","Çayırova","Dilovası"};
                    comboBox2.Items.AddRange(ilceler);
                }
                else if (comboBox1.Text == "Sakarya")
                {
                    comboBox2.Text = "Seçiniz...";
                    object[] ilceler = new object[] { "Merkez", "Serdivan", "Erenler","Ferizli","Söğütlü","Sapanca","Arifiye"};
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
                    object[] ilceler = new object[] { "Merkez","Yunusemre", "Şehzadeler" };
                    comboBox2.Items.AddRange(ilceler);
                }

                else if (comboBox1.Text == "İzmir")
                {
                    comboBox2.Text = "Seçiniz...";
                    object[] ilceler = new object[] { "Bornova",  "Balçova", "Pınarbaşı", "Kemalpaşa", "Çiğle", "Karşıyaka", "Gaziemir", "Karabağlar", "Menderes", "Torbalı", "Bayraklı", "Buca", "Menemen","Kemalpaşa","Konak" };
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
                else if (comboBox1.Text == "Balıkesir")
                {
                    comboBox2.Text = "Seçiniz...";
                    object[] ilceler = new object[] { "Altıeylül","Karesi" };
                    comboBox2.Items.AddRange(ilceler);
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

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.Red;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackColor = Color.Green;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int adet = 1;
            adet = Convert.ToInt32(textBox7.Text);
            string TL = "TL";
            double desı = Convert.ToDouble(textBox4.Text);
            if (textBox4 != null)
            {

                double ekdesı = ((70 + (desı - 50) * 1.40) * 1.18 * 1.06);

                if (desı >= 0 && desı <= 20)
                    if (adet >= 6)
                        textBox5.Text = Math.Round(adet * (28 * 1.18 * 1.06), 2).ToString();
                    else
                        MessageBox.Show("adet en az 6 girilmeli...");

                else if (desı > 20 && desı <= 30)
                    if (adet >= 4)
                        textBox5.Text = Math.Round(adet * (42 * 1.18 * 1.06), 2).ToString();
                    else
                        MessageBox.Show("adet en az 4 girilmeli...");

                else if (desı > 30 && desı <= 40)
                    if (adet >= 3)
                        textBox5.Text = Math.Round(adet * (56 * 1.18 * 1.06), 2).ToString();
                    else
                        MessageBox.Show("adet en az 4 girilmeli...");

                else if (desı > 40 && desı <= 50)
                    if (adet >= 3)
                        textBox5.Text = Math.Round(adet * (70 * 1.18 * 1.06), 2).ToString();
                    else
                        MessageBox.Show("adet en az 3 girilmeli...");

                else if (desı > 50)
                    textBox5.Text = Math.Round(adet * ekdesı, 2).ToString();

            }
        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            button5.BackColor = Color.Blue;

        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            button5.BackColor = Color.White;

        }
    }
}
