﻿using System;
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
using System.Collections;
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;
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
            System.Drawing.Image Logo = imageList1.Images["ERMED2.png"];
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

                            e.Graphics.DrawString("KARGO FİYAT HESABI", new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("KARGO FİYAT HESABI", new Font(dataGridView1.Font,
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
            e.Graphics.DrawImage(Logo, 730, 10, 100, 100);
            e.Graphics.DrawString("TOPLAM FİYAT=", font, firca, iCellHeight + 390, iTopMargin + 10);
            e.Graphics.DrawString(textBox6.Text, font, firca, iCellHeight + 590, iTopMargin + 10);
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
            onizleme.ShowDialog();
            PrintDialog yazdir = new PrintDialog();
            yazdir.Document = printDocument1;
            yazdir.UseEXDialog = true;
            if (yazdir.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }
    }
}
