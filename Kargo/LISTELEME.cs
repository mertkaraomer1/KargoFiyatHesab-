﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using Font = System.Drawing.Font;
using Range = Microsoft.Office.Interop.Excel.Range;
using Rectangle = System.Drawing.Rectangle;
using TextBox = System.Windows.Forms.TextBox;

namespace Kargo
{
    public partial class LISTELEME : Form
    {
        public LISTELEME()
        {
            InitializeComponent();
        }
        SqlConnection baglanti;
        SqlDataAdapter da;
        DataSet ds;
        System.Data.DataTable dt;
        void griddoldur()
        {
            baglanti = new SqlConnection("Data Source=MERTSANAL;Initial Catalog=ermed_kargo;User ID=sa;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            da = new SqlDataAdapter("select* From Kargolarr", baglanti);
            ds = new DataSet();
            baglanti.Open();
            da.Fill(ds, "Kargolarr");
            dataGridView1.DataSource = ds.Tables["Kargolarr"];
            baglanti.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            KARGO_SIRKETLERI KS = new KARGO_SIRKETLERI();
            KS.Show();
            this.Hide();
        }

        private void LISTELEME_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            griddoldur();
            comboBox1.Items.Add("ANA DEPO");
            comboBox1.Items.Add("DMO");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            griddoldur();
            if (comboBox1.Text.ToString() == "" && comboBox6.Text.ToString() == "" && textBox9.Text == "")
            {
                da = new SqlDataAdapter("select * From Kargolarr where Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'AND'" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'", baglanti);
                ds = new DataSet();
            }

            else if (comboBox6.Text.ToString() == "" && textBox9.Text == "")
            {
                da = new SqlDataAdapter("select * From Kargolarr where Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'AND'" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'AND Depo = '" + comboBox1.SelectedItem.ToString() + "'", baglanti);
                ds = new DataSet();
            }

            else if (comboBox1.Text.ToString() == "" && textBox9.Text == "")
            {
                da = new SqlDataAdapter("select * From Kargolarr where Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'AND'" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' AND İL='" + comboBox6.SelectedItem.ToString() + "'", baglanti);
                ds = new DataSet();
            }
            else if (comboBox1.Text.ToString() == "" && dateTimePicker1.Value.ToString() != null && dateTimePicker2.Value.ToString() != null && textBox9.Text != null)
            {
                da = new SqlDataAdapter("select * From Kargolarr where Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'AND'" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'AND Kargo_sirketi = '" + textBox9.Text + "' ", baglanti);
                ds = new DataSet();
            }
            else if (textBox9.Text == "")
            {
                da = new SqlDataAdapter("select * From Kargolarr where Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'AND'" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'AND Depo = '" + comboBox1.SelectedItem.ToString() + "'AND İL='" + comboBox6.SelectedItem.ToString() + "'", baglanti);
                ds = new DataSet();

            }
            else if (dateTimePicker1.Value.ToString() == null && dateTimePicker2.Value.ToString() == null && textBox9.Text != null)
            {
                da = new SqlDataAdapter("select* From Kargolarr where Kargo_sirketi = '" + textBox9.Text + "'", baglanti);
                ds = new DataSet();
            }

            else if (comboBox6.Text.ToString() == "" && dateTimePicker1.Value.ToString() != null && dateTimePicker2.Value.ToString() != null && textBox9.Text != null && comboBox1.SelectedItem.ToString() != null)
            {

                da = new SqlDataAdapter("select* From Kargolarr where Kargo_sirketi = '" + textBox9.Text + "' AND Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'AND Depo = '" + comboBox1.SelectedItem.ToString() + "' ", baglanti);
                ds = new DataSet();
            }
            else if (comboBox6.SelectedItem.ToString() != null && dateTimePicker1.Value.ToString() != null && dateTimePicker2.Value.ToString() != null && textBox9.Text != null && comboBox1.SelectedItem.ToString() != null)
            {

                da = new SqlDataAdapter("select* From Kargolarr where Kargo_sirketi = '" + textBox9.Text + "' AND Tarih BETWEEN'" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' AND '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'AND Depo = '" + comboBox1.SelectedItem.ToString() + "'AND İL='" + comboBox6.SelectedItem.ToString() + "' ", baglanti);
                ds = new DataSet();
            }

            baglanti.Open();
            da.Fill(ds, "Kargolar");
            dataGridView1.DataSource = ds.Tables["Kargolar"];
            baglanti.Close();

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
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
        StringFormat strFormat;
        ArrayList arrColumnLefts = new ArrayList();
        ArrayList arrColumnWidths = new ArrayList();
        int iCellHeight = 0;
        int iTotalWidth = 0;
        int iRow = 0;
        bool bFirstPage = false;
        bool bNewPage = false;
        int iHeaderHeight = 0;
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

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
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


                            e.Graphics.DrawString("KARGO LİSTESİ", new Font(dataGridView1.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString("KARGO LİSTESİ", new Font(dataGridView1.Font,
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
            e.Graphics.DrawLine(kalem, iCellHeight + 687, iTopMargin, iCellHeight + 687, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 570, iTopMargin, iCellHeight + 570, iTopMargin + 40);
            e.Graphics.DrawLine(kalem, iCellHeight + 350, iTopMargin, iCellHeight + 687, iTopMargin);
            e.Graphics.DrawLine(kalem, iCellHeight + 360, iTopMargin + 40, iCellHeight + 687, iTopMargin + 40);
        }

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

        private void button4_Click(object sender, EventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                Math.Round(toplam, 0);
            }
            Math.Round(toplam, 0);
            textBox6.Text = Math.Round(toplam, 0).ToString() + " TL";
        }

        private void LISTELEME_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult c;
            c = MessageBox.Show("Çıkmakistediğinizden eminmisiniz ? ", "KargoFiyatHesaplama Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (c == DialogResult.Yes)
                Environment.Exit(0);
            else

                e.Cancel = true;//Çıkışı durdur
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.Red;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.Green;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackColor = Color.Blue;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;
        }


    }
}
